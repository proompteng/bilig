import { formatAddress, parseCellAddress } from '@bilig/formula'
import type { GridEngineLike, GridRenderRevisionSnapshot } from '@bilig/grid'
import {
  ValueTag,
  type CellRangeRef,
  type CellSnapshot,
  type CellStyleField,
  type CellStylePatch,
  type CellStyleRecord,
  type Viewport,
  type WorkbookAxisEntrySnapshot,
  type WorkbookMergeRangeSnapshot,
} from '@bilig/protocol'
import {
  decodeWorkbookDeltaBatchV3,
  type RenderTileDeltaSubscription,
  type ViewportPatch,
  type WorkbookDeltaBatchV3,
  type WorkerEngineClient,
} from '@bilig/worker-transport'
import { resolveProjectedStyleDirtyMask } from './projected-style-dirty-mask.js'
import type { ProjectedRenderTile, ProjectedTileSceneChange, ProjectedTileSceneStore } from './projected-tile-scene-store.js'
import { ProjectedViewportAxisStore } from './projected-viewport-axis-store.js'
import { DEFAULT_MAX_CACHED_CELLS_PER_SHEET, ProjectedViewportCellCache } from './projected-viewport-cell-cache.js'
import { ProjectedViewportPatchCoordinator, type ProjectedViewportPatchApplied } from './projected-viewport-patch-coordinator.js'
import { ProjectedViewportPatchRevisionGate } from './projected-viewport-patch-revision-gate.js'
import {
  normalizeViewportRange,
  ProjectedViewportRangeOverlayStore,
  viewportRangeCellCount,
  type NormalizedViewportRange,
} from './projected-viewport-range-overlay.js'
import {
  applyProjectedStylePatch,
  assertValidProjectedAxisMutation,
  buildAxisEntries,
  clearProjectedStyleFields,
  DEFAULT_STYLE_ID,
  normalizeProjectedCellStyle,
  projectedCellStyleIdForKey,
  projectedCellStyleKey,
} from './projected-viewport-style-helpers.js'
import { ProjectedWorkbookLocalDeltaPublisher } from './projected-workbook-local-delta-publisher.js'
import { LOCAL_CELL_VISUAL_DIRTY_MASK } from './projected-workbook-local-delta.js'
import { OPTIMISTIC_CELL_SNAPSHOT_FLAG } from './workbook-optimistic-cell-flags.js'
import { createContentClearedOptimisticSnapshot } from './workbook-optimistic-range.js'
import { normalizeWorkbookMergeRange } from './worker-runtime-support.js'

export interface ProjectedViewportStoreOptions {
  readonly maxCachedCellsPerSheet?: number
}
interface ProjectedCellSnapshotWriteOptions {
  readonly force?: boolean
  readonly forceOptimistic?: boolean
  readonly allowOptimisticClearResurrection?: boolean
  readonly emitLocalDelta?: boolean
  readonly localDirtyMask?: number | ((snapshot: CellSnapshot) => number) | undefined
  readonly suppressRangeOverlays?: boolean
}
interface ProjectedAxisMutationOptions {
  readonly emitLocalDelta?: boolean
}
type CellItem = readonly [number, number]
type SheetViewportChannel = 'columnWidths' | 'rowHeights' | 'hiddenColumns' | 'hiddenRows' | 'freeze' | 'merges'
type SheetIdentity = { readonly sheetId: number; readonly sheetOrdinal: number }
const MAX_MATERIALIZED_OPTIMISTIC_STYLE_CELLS = 512

export class ProjectedViewportStore implements GridEngineLike {
  private readonly options: ProjectedViewportStoreOptions
  private readonly cellCache: ProjectedViewportCellCache
  private readonly axisStore: ProjectedViewportAxisStore
  private readonly patchCoordinator: ProjectedViewportPatchCoordinator
  private readonly rangeOverlayStore: ProjectedViewportRangeOverlayStore
  private readonly patchRevisionGate = new ProjectedViewportPatchRevisionGate()
  private readonly localDeltaPublisher: ProjectedWorkbookLocalDeltaPublisher
  private tileSceneStore: ProjectedTileSceneStore | null = null
  private readonly sheetIdentitiesByName = new Map<string, SheetIdentity>()
  private readonly sheetChannelListeners = new Map<string, Map<SheetViewportChannel, Set<() => void>>>()
  private readonly mergeRangesBySheet = new Map<string, Map<string, WorkbookMergeRangeSnapshot>>()
  private localRevision = 0

  readonly workbook = {
    getSheet: (sheetName: string) => this.cellCache.getSheet(sheetName),
  }

  constructor(
    private readonly client?: WorkerEngineClient,
    options: ProjectedViewportStoreOptions = {},
  ) {
    this.options = options
    this.cellCache = new ProjectedViewportCellCache({
      maxCachedCellsPerSheet: this.options.maxCachedCellsPerSheet ?? DEFAULT_MAX_CACHED_CELLS_PER_SHEET,
    })
    this.rangeOverlayStore = new ProjectedViewportRangeOverlayStore({
      deleteCellSnapshots: (targets) => {
        this.cellCache.deleteCellSnapshots(targets)
      },
      forEachCachedOrVisibleCellSnapshotInRange: (range, listener) => {
        this.cellCache.forEachCachedOrVisibleCellSnapshotInRange(range, listener)
      },
      getCell: (sheetName, address) => this.cellCache.getCell(sheetName, address),
      hasCellSnapshot: (sheetName, address) => this.cellCache.hasCellSnapshot(sheetName, address),
      setCellSnapshots: (snapshots) => {
        this.setCellSnapshots(snapshots, { force: true, forceOptimistic: true, suppressRangeOverlays: false })
      },
      restoreCellSnapshots: (snapshots) => {
        this.setCellSnapshots(snapshots, {
          force: true,
          forceOptimistic: true,
          allowOptimisticClearResurrection: true,
          suppressRangeOverlays: false,
        })
      },
    })
    this.axisStore = new ProjectedViewportAxisStore({
      markSheetKnown: (sheetName) => this.cellCache.markSheetKnown(sheetName),
      notifyListeners: () => this.cellCache.notifyListeners(),
    })
    this.patchCoordinator = new ProjectedViewportPatchCoordinator({
      cellCache: this.cellCache,
      axisStore: this.axisStore,
      mergeRangesBySheet: this.mergeRangesBySheet,
      ...(client ? { client } : {}),
      shouldApplyViewportPatch: (patch) => this.patchRevisionGate.shouldApplyViewportPatch(patch),
      resolveSheetIdentity: (sheetName) => this.resolveSheetIdentity(sheetName),
      onViewportPatchApplied: (patch, result) => this.handleViewportPatchApplied(patch, result),
    })
    this.localDeltaPublisher = new ProjectedWorkbookLocalDeltaPublisher({
      getLastAuthoritativeRevision: () => this.patchRevisionGate.getLastAuthoritativeRevision(),
      getLastBatchId: () => this.patchRevisionGate.getLastBatchId(),
      resolveSheetIdentity: (sheetName) => this.resolveSheetIdentity(sheetName),
    })
  }

  subscribe(listener: () => void): () => void {
    return this.cellCache.subscribe(listener)
  }

  subscribeSheetChannel(sheetName: string, channel: SheetViewportChannel, listener: () => void): () => void {
    const channels = this.sheetChannelListeners.get(sheetName) ?? new Map<SheetViewportChannel, Set<() => void>>()
    const listeners = channels.get(channel) ?? new Set<() => void>()
    listeners.add(listener)
    channels.set(channel, listeners)
    this.sheetChannelListeners.set(sheetName, channels)
    return () => {
      listeners.delete(listener)
      if (listeners.size === 0) {
        channels.delete(channel)
      }
      if (channels.size === 0) {
        this.sheetChannelListeners.delete(sheetName)
      }
    }
  }

  subscribeCell(sheetName: string, address: string, listener: () => void): () => void {
    return this.cellCache.subscribeCells(sheetName, [address], listener)
  }

  peekCell(sheetName: string, address: string): CellSnapshot | undefined {
    const snapshot = this.cellCache.peekCell(sheetName, address)
    if (snapshot) {
      return this.rangeOverlayStore.apply(sheetName, address, snapshot)
    }
    return this.rangeOverlayStore.hasOverlayForCell(sheetName, address)
      ? this.rangeOverlayStore.apply(sheetName, address, this.cellCache.getCell(sheetName, address))
      : undefined
  }

  peekBaseCell(sheetName: string, address: string): CellSnapshot | undefined {
    return this.cellCache.peekCell(sheetName, address)
  }

  getColumnWidths(sheetName: string): Readonly<Record<number, number>> {
    return this.axisStore.getColumnWidths(sheetName)
  }

  getColumnSizes(sheetName: string): Readonly<Record<number, number>> {
    return this.axisStore.getColumnSizes(sheetName)
  }

  getRowHeights(sheetName: string): Readonly<Record<number, number>> {
    return this.axisStore.getRowHeights(sheetName)
  }

  getRowSizes(sheetName: string): Readonly<Record<number, number>> {
    return this.axisStore.getRowSizes(sheetName)
  }

  getHiddenColumns(sheetName: string): Readonly<Record<number, true>> {
    return this.axisStore.getHiddenColumns(sheetName)
  }

  getHiddenRows(sheetName: string): Readonly<Record<number, true>> {
    return this.axisStore.getHiddenRows(sheetName)
  }

  getFreezeRows(sheetName: string): number {
    return this.axisStore.getFreezeRows(sheetName)
  }

  getFreezeCols(sheetName: string): number {
    return this.axisStore.getFreezeCols(sheetName)
  }

  getCell(sheetName: string, address: string): CellSnapshot {
    return this.rangeOverlayStore.apply(sheetName, address, this.cellCache.getCell(sheetName, address))
  }

  hasCellSnapshot(sheetName: string, address: string): boolean {
    return this.cellCache.hasCellSnapshot(sheetName, address)
  }

  forEachCellSnapshotInRange(range: CellRangeRef, listener: (snapshot: CellSnapshot) => void): void {
    this.cellCache.forEachCellSnapshotInRange(range, listener)
  }

  forEachCachedOrVisibleCellSnapshotInRange(range: CellRangeRef, listener: (snapshot: CellSnapshot) => void): void {
    this.cellCache.forEachCachedOrVisibleCellSnapshotInRange(range, listener)
  }

  getCellStyle(styleId: string | undefined): CellStyleRecord | undefined {
    return this.cellCache.getCellStyle(styleId)
  }

  setRangeStyle(range: CellRangeRef, patch: CellStylePatch): (() => void) | null {
    return this.applyCachedRangeStyleMutation(range, (baseStyle) => applyProjectedStylePatch(baseStyle, patch))
  }

  clearRangeStyle(range: CellRangeRef, fields?: readonly CellStyleField[]): (() => void) | null {
    return this.applyCachedRangeStyleMutation(range, (baseStyle) => clearProjectedStyleFields(baseStyle, fields))
  }

  getMergeRange(sheetName: string, address: string): WorkbookMergeRangeSnapshot | undefined {
    const parsed = parseCellAddress(address, sheetName)
    for (const range of this.mergeRangesBySheet.get(sheetName)?.values() ?? []) {
      const normalized = normalizeWorkbookMergeRange(range)
      if (
        parsed.row >= normalized.startRow &&
        parsed.row <= normalized.endRow &&
        parsed.col >= normalized.startCol &&
        parsed.col <= normalized.endCol
      ) {
        return {
          sheetName: normalized.sheetName,
          startAddress: normalized.startAddress,
          endAddress: normalized.endAddress,
        }
      }
    }
    return undefined
  }

  listMergeRanges(sheetName: string): WorkbookMergeRangeSnapshot[] {
    return [...(this.mergeRangesBySheet.get(sheetName)?.values() ?? [])].map((range) => {
      const normalized = normalizeWorkbookMergeRange(range)
      return {
        sheetName: normalized.sheetName,
        startAddress: normalized.startAddress,
        endAddress: normalized.endAddress,
      }
    })
  }

  getColumnAxisEntries(sheetName: string): WorkbookAxisEntrySnapshot[] {
    return buildAxisEntries(this.axisStore.getColumnSizes(sheetName), this.axisStore.getHiddenColumns(sheetName), 'col')
  }

  getRowAxisEntries(sheetName: string): WorkbookAxisEntrySnapshot[] {
    return buildAxisEntries(this.axisStore.getRowSizes(sheetName), this.axisStore.getHiddenRows(sheetName), 'row')
  }

  getLastMetrics(): Pick<NonNullable<ViewportPatch['metrics']>, 'batchId'> {
    return { batchId: this.patchRevisionGate.getLastBatchId() }
  }

  getLastAuthoritativeRevision(): number | null {
    return this.patchRevisionGate.getLastAuthoritativeRevision()
  }

  getRenderRevisionSnapshot(): GridRenderRevisionSnapshot {
    return {
      authoritativeRevision: this.patchRevisionGate.getLastAuthoritativeRevision(),
      localRevision: this.localRevision,
      projectedRevision: this.patchRevisionGate.getLastBatchId(),
      tileSceneCameraSeq: this.tileSceneStore?.getLastCameraSeq() ?? null,
      tileSceneRevision: this.tileSceneStore?.getLastBatchId() ?? null,
    }
  }

  setCellSnapshot(snapshot: CellSnapshot, options: ProjectedCellSnapshotWriteOptions = {}): void {
    this.setCellSnapshots([snapshot], options)
  }

  private setCellSnapshots(snapshots: readonly CellSnapshot[], options: ProjectedCellSnapshotWriteOptions = {}): void {
    const acceptedSnapshotsBySheetAndMask = new Map<string, Map<number, CellSnapshot[]>>()
    snapshots.forEach((snapshot) => {
      if (options.suppressRangeOverlays !== false) {
        this.rangeOverlayStore.suppressExistingOverlaysForCell(snapshot.sheetName, snapshot.address)
      }
    })
    this.cellCache.writeCellSnapshots(snapshots, options).forEach((result) => {
      if (result.changed && result.acceptedSnapshot && options.emitLocalDelta !== false) {
        const acceptedSnapshotsByMask =
          acceptedSnapshotsBySheetAndMask.get(result.acceptedSnapshot.sheetName) ?? new Map<number, CellSnapshot[]>()
        const dirtyMask = resolveProjectedCellLocalDirtyMask(result.acceptedSnapshot, options.localDirtyMask)
        const acceptedSnapshots = acceptedSnapshotsByMask.get(dirtyMask) ?? []
        acceptedSnapshots.push(result.acceptedSnapshot)
        acceptedSnapshotsByMask.set(dirtyMask, acceptedSnapshots)
        acceptedSnapshotsBySheetAndMask.set(result.acceptedSnapshot.sheetName, acceptedSnapshotsByMask)
      }
    })
    if (acceptedSnapshotsBySheetAndMask.size > 0) {
      this.localRevision += 1
      acceptedSnapshotsBySheetAndMask.forEach((acceptedSnapshotsByMask, sheetName) => {
        acceptedSnapshotsByMask.forEach((acceptedSnapshots, dirtyMask) => {
          this.localDeltaPublisher.emitCellSnapshots(sheetName, acceptedSnapshots, dirtyMask)
        })
      })
    }
  }

  clearOptimisticCellFlagsForSheet(sheetName: string): void {
    this.rangeOverlayStore.dropSheets([sheetName])
    const changedSnapshots = this.cellCache.clearOptimisticCellFlagsForSheet(sheetName)
    if (changedSnapshots.length > 0) {
      this.localRevision += 1
      this.localDeltaPublisher.emitCellSnapshots(sheetName, changedSnapshots)
    }
  }

  beginOptimisticClearRange(range: CellRangeRef): (() => void) | null {
    return this.rangeOverlayStore.register(range, createIdempotentContentClearedSnapshot)
  }

  beginOptimisticRangeOverlay(range: CellRangeRef, apply: (snapshot: CellSnapshot) => CellSnapshot): (() => void) | null {
    return this.rangeOverlayStore.register(range, apply)
  }

  setColumnWidth(sheetName: string, columnIndex: number, width: number, options: ProjectedAxisMutationOptions = {}): void {
    assertValidProjectedAxisMutation('column', columnIndex, width)
    const previousWidth = this.axisStore.getColumnWidths(sheetName)[columnIndex]
    this.axisStore.setColumnWidth(sheetName, columnIndex, width)
    this.notifySheetChannels(sheetName, ['columnWidths'])
    if (options.emitLocalDelta !== false && this.axisStore.getColumnWidths(sheetName)[columnIndex] !== previousWidth) {
      this.localDeltaPublisher.emitAxis(sheetName, 'column', columnIndex)
    }
  }

  ackColumnWidth(sheetName: string, columnIndex: number, width: number): void {
    assertValidProjectedAxisMutation('column', columnIndex, width)
    this.axisStore.ackColumnWidth(sheetName, columnIndex, width)
    this.notifySheetChannels(sheetName, ['columnWidths'])
  }

  rollbackColumnWidth(sheetName: string, columnIndex: number, width: number | undefined): void {
    assertValidProjectedAxisMutation('column', columnIndex, width)
    const previousWidth = this.axisStore.getColumnWidths(sheetName)[columnIndex]
    this.axisStore.rollbackColumnWidth(sheetName, columnIndex, width)
    this.notifySheetChannels(sheetName, ['columnWidths'])
    if (this.axisStore.getColumnWidths(sheetName)[columnIndex] !== previousWidth) {
      this.localDeltaPublisher.emitAxis(sheetName, 'column', columnIndex)
    }
  }

  setColumnHidden(sheetName: string, columnIndex: number, hidden: boolean, size: number): void {
    assertValidProjectedAxisMutation('column', columnIndex, size)
    const previousWidth = this.axisStore.getColumnWidths(sheetName)[columnIndex]
    this.axisStore.setColumnHidden(sheetName, columnIndex, hidden, size)
    this.notifySheetChannels(sheetName, ['columnWidths', 'hiddenColumns'])
    if (this.axisStore.getColumnWidths(sheetName)[columnIndex] !== previousWidth) {
      this.localDeltaPublisher.emitAxis(sheetName, 'column', columnIndex)
    }
  }

  rollbackColumnHidden(sheetName: string, columnIndex: number, previous: { hidden: boolean; size: number | undefined }): void {
    assertValidProjectedAxisMutation('column', columnIndex, previous.size)
    const previousWidth = this.axisStore.getColumnWidths(sheetName)[columnIndex]
    this.axisStore.rollbackColumnHidden(sheetName, columnIndex, previous)
    this.notifySheetChannels(sheetName, ['columnWidths', 'hiddenColumns'])
    if (this.axisStore.getColumnWidths(sheetName)[columnIndex] !== previousWidth) {
      this.localDeltaPublisher.emitAxis(sheetName, 'column', columnIndex)
    }
  }

  setRowHeight(sheetName: string, rowIndex: number, height: number, options: ProjectedAxisMutationOptions = {}): void {
    assertValidProjectedAxisMutation('row', rowIndex, height)
    const previousHeight = this.axisStore.getRowHeights(sheetName)[rowIndex]
    this.axisStore.setRowHeight(sheetName, rowIndex, height)
    this.notifySheetChannels(sheetName, ['rowHeights'])
    if (options.emitLocalDelta !== false && this.axisStore.getRowHeights(sheetName)[rowIndex] !== previousHeight) {
      this.localDeltaPublisher.emitAxis(sheetName, 'row', rowIndex)
    }
  }

  ackRowHeight(sheetName: string, rowIndex: number, height: number): void {
    assertValidProjectedAxisMutation('row', rowIndex, height)
    this.axisStore.ackRowHeight(sheetName, rowIndex, height)
    this.notifySheetChannels(sheetName, ['rowHeights'])
  }

  rollbackRowHeight(sheetName: string, rowIndex: number, height: number | undefined): void {
    assertValidProjectedAxisMutation('row', rowIndex, height)
    const previousHeight = this.axisStore.getRowHeights(sheetName)[rowIndex]
    this.axisStore.rollbackRowHeight(sheetName, rowIndex, height)
    this.notifySheetChannels(sheetName, ['rowHeights'])
    if (this.axisStore.getRowHeights(sheetName)[rowIndex] !== previousHeight) {
      this.localDeltaPublisher.emitAxis(sheetName, 'row', rowIndex)
    }
  }

  setRowHidden(sheetName: string, rowIndex: number, hidden: boolean, size: number): void {
    assertValidProjectedAxisMutation('row', rowIndex, size)
    const previousHeight = this.axisStore.getRowHeights(sheetName)[rowIndex]
    this.axisStore.setRowHidden(sheetName, rowIndex, hidden, size)
    this.notifySheetChannels(sheetName, ['rowHeights', 'hiddenRows'])
    if (this.axisStore.getRowHeights(sheetName)[rowIndex] !== previousHeight) {
      this.localDeltaPublisher.emitAxis(sheetName, 'row', rowIndex)
    }
  }

  rollbackRowHidden(sheetName: string, rowIndex: number, previous: { hidden: boolean; size: number | undefined }): void {
    assertValidProjectedAxisMutation('row', rowIndex, previous.size)
    const previousHeight = this.axisStore.getRowHeights(sheetName)[rowIndex]
    this.axisStore.rollbackRowHidden(sheetName, rowIndex, previous)
    this.notifySheetChannels(sheetName, ['rowHeights', 'hiddenRows'])
    if (this.axisStore.getRowHeights(sheetName)[rowIndex] !== previousHeight) {
      this.localDeltaPublisher.emitAxis(sheetName, 'row', rowIndex)
    }
  }

  setKnownSheets(sheetNames: readonly string[]): void {
    const removedSheets = this.cellCache.setKnownSheets(sheetNames)
    this.axisStore.dropSheets(removedSheets)
    this.tileSceneStore?.dropSheets(removedSheets)
    this.rangeOverlayStore.dropSheets(removedSheets)
    removedSheets.forEach((sheetName) => {
      this.mergeRangesBySheet.delete(sheetName)
      this.sheetChannelListeners.delete(sheetName)
      this.sheetIdentitiesByName.delete(sheetName)
    })
  }

  setSheetIdentities(sheets: readonly { readonly id: number; readonly name: string; readonly order: number }[]): void {
    this.sheetIdentitiesByName.clear()
    sheets.forEach((sheet) => {
      this.sheetIdentitiesByName.set(sheet.name, {
        sheetId: sheet.id,
        sheetOrdinal: sheet.order,
      })
    })
  }

  resetProjectionState(sheetNames: readonly string[] = this.cellCache.getKnownSheetNames()): void {
    this.cellCache.resetSheets(sheetNames)
    this.axisStore.dropSheets(sheetNames)
    this.tileSceneStore?.dropSheets(sheetNames)
    this.rangeOverlayStore.dropSheets(sheetNames)
    sheetNames.forEach((sheetName) => {
      this.mergeRangesBySheet.delete(sheetName)
      this.notifySheetChannels(sheetName, ['columnWidths', 'rowHeights', 'hiddenColumns', 'hiddenRows', 'freeze', 'merges'])
    })
  }

  subscribeCells(sheetName: string, addresses: readonly string[], listener: () => void): () => void {
    return this.cellCache.subscribeCells(sheetName, addresses, listener)
  }

  subscribeViewport(
    sheetName: string,
    viewport: Viewport,
    listener: (damage?: readonly { cell: readonly [number, number] }[]) => void,
    options: { readonly initialPatch?: 'full' | 'none'; readonly notifyOnProofRevision?: boolean } = {},
  ): () => void {
    const subscriptionOptions: { initialPatch?: 'full' | 'none'; notifyOnProofRevision?: boolean } = {}
    if (options.initialPatch !== undefined) {
      subscriptionOptions.initialPatch = options.initialPatch
    }
    if (options.notifyOnProofRevision !== undefined) {
      subscriptionOptions.notifyOnProofRevision = options.notifyOnProofRevision
    }
    const unsubscribe = this.patchCoordinator.subscribeViewport(sheetName, viewport, listener, subscriptionOptions)
    this.rangeOverlayStore.materializeViewport(sheetName, viewport)
    return unsubscribe
  }

  subscribeAuxiliaryViewport(
    sheetName: string,
    viewport: Viewport,
    listener: (damage?: readonly { cell: readonly [number, number] }[]) => void,
    options: { readonly initialPatch?: 'full' | 'none' } = {},
  ): () => void {
    return this.patchCoordinator.subscribeViewport(sheetName, viewport, listener, {
      initialPatch: options.initialPatch ?? 'full',
    })
  }

  subscribeRenderTileDeltas(subscription: RenderTileDeltaSubscription, listener: (change: ProjectedTileSceneChange) => void): () => void {
    this.sheetIdentitiesByName.set(subscription.sheetName, {
      sheetId: subscription.sheetId,
      sheetOrdinal: subscription.sheetOrdinal ?? subscription.tileInterest?.sheetOrdinal ?? subscription.sheetId,
    })
    let disposed = false
    let unsubscribe: (() => void) | null = null
    void (async () => {
      const store = await this.getTileSceneStore()
      if (disposed) {
        return
      }
      unsubscribe = store.subscribe(subscription, (change) => {
        this.noteObservedBatchId(change.batchId)
        listener(change)
      })
    })()
    return () => {
      disposed = true
      unsubscribe?.()
      unsubscribe = null
    }
  }

  subscribeWorkbookDeltas(listener: (batch: WorkbookDeltaBatchV3) => void): () => void {
    if (!this.client) {
      throw new Error('Workbook delta subscriptions require a worker engine client')
    }
    const unsubscribeLocal = this.localDeltaPublisher.subscribe(listener)
    const unsubscribeClient = this.client.subscribeWorkbookDeltas((bytes) => {
      listener(decodeWorkbookDeltaBatchV3(bytes))
    })
    return () => {
      unsubscribeLocal()
      unsubscribeClient()
    }
  }

  peekRenderTile(tileId: number): ProjectedRenderTile | null {
    return this.tileSceneStore?.peekTile(tileId) ?? null
  }

  applyViewportPatch(patch: ViewportPatch): readonly { cell: CellItem }[] {
    const result = this.patchCoordinator.applyViewportPatchDetailed(patch)
    this.handleViewportPatchApplied(patch, result)
    return result.damage
  }

  private handleViewportPatchApplied(patch: ViewportPatch, result: ProjectedViewportPatchApplied): void {
    this.patchRevisionGate.noteAppliedViewportPatch(patch)
    const channels: SheetViewportChannel[] = []
    if (result.columnsChanged) {
      channels.push('columnWidths', 'hiddenColumns')
    }
    if (result.rowsChanged) {
      channels.push('rowHeights', 'hiddenRows')
    }
    if (result.freezeChanged) {
      channels.push('freeze')
    }
    if (result.mergesChanged) {
      channels.push('merges')
    }
    if (channels.length > 0) {
      this.notifySheetChannels(patch.viewport.sheetName, channels)
    }
  }

  private async getTileSceneStore(): Promise<ProjectedTileSceneStore> {
    if (this.tileSceneStore) {
      return this.tileSceneStore
    }
    const { ProjectedTileSceneStore } = await import('./projected-tile-scene-store.js')
    this.tileSceneStore = new ProjectedTileSceneStore(this.client)
    return this.tileSceneStore
  }

  private notifySheetChannels(sheetName: string, channels: readonly SheetViewportChannel[]): void {
    const sheetChannels = this.sheetChannelListeners.get(sheetName)
    if (!sheetChannels) {
      return
    }
    const visited = new Set<() => void>()
    for (const channel of channels) {
      const listeners = sheetChannels.get(channel)
      if (!listeners) {
        continue
      }
      for (const listener of listeners) {
        if (visited.has(listener)) {
          continue
        }
        visited.add(listener)
        listener()
      }
    }
  }

  private resolveSheetIdentity(sheetName: string): SheetIdentity | null {
    return this.sheetIdentitiesByName.get(sheetName) ?? null
  }

  getSheetIdentity(sheetName: string): SheetIdentity | null {
    return this.resolveSheetIdentity(sheetName)
  }

  private noteObservedBatchId(batchId: number): void {
    this.patchRevisionGate.noteObservedBatchId(batchId)
  }

  private applyCachedRangeStyleMutation(
    range: CellRangeRef,
    mutateStyle: (baseStyle: CellStyleRecord) => Omit<CellStyleRecord, 'id'>,
  ): (() => void) | null {
    const normalizedRange = normalizeViewportRange(range)
    if (viewportRangeCellCount(normalizedRange) > MAX_MATERIALIZED_OPTIMISTIC_STYLE_CELLS) {
      const revisionBeforeRegister = this.localRevision
      const rollback = this.rangeOverlayStore.register(range, (snapshot) => this.applyStyleMutationToSnapshot(snapshot, mutateStyle))
      if (!rollback) {
        return null
      }
      if (this.localRevision === revisionBeforeRegister) {
        this.emitLocalRangeVisualDelta(normalizedRange)
      }
      return () => {
        const revisionBeforeRollback = this.localRevision
        rollback()
        if (this.localRevision === revisionBeforeRollback) {
          this.emitLocalRangeVisualDelta(normalizedRange)
        }
      }
    }
    const previousSnapshots: Array<{
      readonly existed: boolean
      readonly overlaySuppressionCutoff: number
      readonly snapshot: CellSnapshot
    }> = []
    for (let row = normalizedRange.startRow; row <= normalizedRange.endRow; row += 1) {
      for (let col = normalizedRange.startCol; col <= normalizedRange.endCol; col += 1) {
        const address = formatAddress(row, col)
        const snapshot = this.getCell(range.sheetName, address)
        const mutation = this.applyStyleMutationToSnapshotDetailed(snapshot, mutateStyle)
        if (snapshotStyleId(snapshot) === snapshotStyleId(mutation.snapshot)) {
          continue
        }
        previousSnapshots.push({
          existed: this.cellCache.hasCellSnapshot(snapshot.sheetName, snapshot.address),
          overlaySuppressionCutoff: this.rangeOverlayStore.getSuppressedOverlayMaxId(snapshot.sheetName, snapshot.address),
          snapshot: structuredClone(snapshot),
        })
        this.setCellSnapshot(mutation.snapshot, { force: true, forceOptimistic: true, localDirtyMask: mutation.dirtyMask })
      }
    }
    if (previousSnapshots.length === 0) {
      return null
    }
    return () => {
      previousSnapshots.forEach(({ existed, overlaySuppressionCutoff, snapshot }) => {
        this.rangeOverlayStore.restoreOverlaySuppression(snapshot.sheetName, snapshot.address, overlaySuppressionCutoff)
        const currentSnapshot = this.getCell(snapshot.sheetName, snapshot.address)
        const currentStyle = this.cellCache.getCellStyle(currentSnapshot.styleId) ?? { id: DEFAULT_STYLE_ID }
        const restoreStyle = this.cellCache.getCellStyle(snapshot.styleId) ?? { id: DEFAULT_STYLE_ID }
        this.setCellSnapshot(snapshot, {
          force: true,
          forceOptimistic: true,
          localDirtyMask: resolveProjectedStyleDirtyMask({ baseStyle: currentStyle, nextStyle: restoreStyle, snapshot: currentSnapshot }),
          suppressRangeOverlays: false,
        })
        if (!existed) {
          this.cellCache.deleteCellSnapshot(snapshot.sheetName, snapshot.address)
        }
      })
    }
  }

  private emitLocalRangeVisualDelta(range: NormalizedViewportRange): void {
    this.localRevision += 1
    this.localDeltaPublisher.emitRange(range.sheetName, range)
  }

  private internLocalCellStyle(style: Omit<CellStyleRecord, 'id'>): CellStyleRecord {
    const normalized = normalizeProjectedCellStyle(style)
    const key = projectedCellStyleKey(normalized)
    const id = key === projectedCellStyleKey({}) ? DEFAULT_STYLE_ID : projectedCellStyleIdForKey(key)
    const record = { id, ...normalized }
    this.cellCache.upsertCellStyle(record)
    return record
  }

  private applyStyleMutationToSnapshot(
    snapshot: CellSnapshot,
    mutateStyle: (baseStyle: CellStyleRecord) => Omit<CellStyleRecord, 'id'>,
  ): CellSnapshot {
    return this.applyStyleMutationToSnapshotDetailed(snapshot, mutateStyle).snapshot
  }

  private applyStyleMutationToSnapshotDetailed(
    snapshot: CellSnapshot,
    mutateStyle: (baseStyle: CellStyleRecord) => Omit<CellStyleRecord, 'id'>,
  ): { readonly dirtyMask: number; readonly snapshot: CellSnapshot } {
    const baseStyle = this.cellCache.getCellStyle(snapshot.styleId) ?? { id: DEFAULT_STYLE_ID }
    const nextStyle = this.internLocalCellStyle(mutateStyle(baseStyle))
    return {
      dirtyMask: resolveProjectedStyleDirtyMask({ baseStyle, nextStyle, snapshot }),
      snapshot: nextStyle.id === DEFAULT_STYLE_ID ? omitSnapshotStyleId(snapshot) : { ...snapshot, styleId: nextStyle.id },
    }
  }
}

function snapshotStyleId(snapshot: CellSnapshot): string {
  return snapshot.styleId ?? DEFAULT_STYLE_ID
}

function resolveProjectedCellLocalDirtyMask(
  snapshot: CellSnapshot,
  dirtyMask: number | ((snapshot: CellSnapshot) => number) | undefined,
): number {
  const resolved = typeof dirtyMask === 'function' ? dirtyMask(snapshot) : dirtyMask
  return Number.isInteger(resolved) && resolved !== undefined && resolved >= 0 ? resolved : LOCAL_CELL_VISUAL_DIRTY_MASK
}

function createIdempotentContentClearedSnapshot(snapshot: CellSnapshot): CellSnapshot {
  if (
    snapshot.value.tag === ValueTag.Empty &&
    snapshot.formula === undefined &&
    snapshot.input === undefined &&
    (snapshot.flags & OPTIMISTIC_CELL_SNAPSHOT_FLAG) !== 0
  ) {
    return snapshot
  }
  return createContentClearedOptimisticSnapshot(snapshot)
}

function omitSnapshotStyleId(snapshot: CellSnapshot): CellSnapshot {
  if (snapshot.styleId === undefined) {
    return snapshot
  }
  const next = { ...snapshot }
  delete next.styleId
  return next
}
