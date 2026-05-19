import type { GridEngineLike, GridRenderRevisionSnapshot } from '@bilig/grid'
import { parseCellAddress } from '@bilig/formula'
import {
  MAX_COLS,
  MAX_ROWS,
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
import { ProjectedViewportAxisStore } from './projected-viewport-axis-store.js'
import { DEFAULT_MAX_CACHED_CELLS_PER_SHEET, ProjectedViewportCellCache } from './projected-viewport-cell-cache.js'
import { ProjectedViewportPatchCoordinator, type ProjectedViewportPatchApplied } from './projected-viewport-patch-coordinator.js'
import type { ProjectedRenderTile, ProjectedTileSceneChange, ProjectedTileSceneStore } from './projected-tile-scene-store.js'
import { getWorkbookScrollPerfCollector } from './perf/workbook-scroll-perf.js'
import { normalizeWorkbookMergeRange } from './worker-runtime-support.js'
import { buildLocalAxisWorkbookDelta, buildLocalCellSnapshotWorkbookDelta } from './projected-workbook-local-delta.js'
import { ProjectedViewportPatchRevisionGate } from './projected-viewport-patch-revision-gate.js'
import { ProjectedViewportRangeOverlayStore } from './projected-viewport-range-overlay.js'
import { createContentClearedOptimisticSnapshot } from './workbook-optimistic-range.js'
import { OPTIMISTIC_CELL_SNAPSHOT_FLAG } from './workbook-optimistic-cell-flags.js'

export interface ProjectedViewportStoreOptions {
  readonly maxCachedCellsPerSheet?: number
}
interface ProjectedCellSnapshotWriteOptions {
  readonly force?: boolean
  readonly forceOptimistic?: boolean
  readonly allowOptimisticClearResurrection?: boolean
  readonly emitLocalDelta?: boolean
}
type CellItem = readonly [number, number]
type SheetViewportChannel = 'columnWidths' | 'rowHeights' | 'hiddenColumns' | 'hiddenRows' | 'freeze' | 'merges'
type SheetIdentity = { readonly sheetId: number; readonly sheetOrdinal: number }
const DEFAULT_STYLE_ID = 'style-0'

export class ProjectedViewportStore implements GridEngineLike {
  private readonly options: ProjectedViewportStoreOptions
  private readonly cellCache: ProjectedViewportCellCache
  private readonly axisStore: ProjectedViewportAxisStore
  private readonly patchCoordinator: ProjectedViewportPatchCoordinator
  private readonly rangeOverlayStore: ProjectedViewportRangeOverlayStore
  private readonly patchRevisionGate = new ProjectedViewportPatchRevisionGate()
  private tileSceneStore: ProjectedTileSceneStore | null = null
  private readonly localWorkbookDeltaListeners = new Set<(batch: WorkbookDeltaBatchV3) => void>()
  private readonly sheetIdentitiesByName = new Map<string, SheetIdentity>()
  private readonly sheetChannelListeners = new Map<string, Map<SheetViewportChannel, Set<() => void>>>()
  private readonly mergeRangesBySheet = new Map<string, Map<string, WorkbookMergeRangeSnapshot>>()
  private localRevision = 0
  private localWorkbookDeltaSeq = 0

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
      deleteCellSnapshot: (sheetName, address) => {
        this.cellCache.deleteCellSnapshot(sheetName, address)
      },
      forEachCachedOrVisibleCellSnapshotInRange: (range, listener) => {
        this.cellCache.forEachCachedOrVisibleCellSnapshotInRange(range, listener)
      },
      getCell: (sheetName, address) => this.cellCache.getCell(sheetName, address),
      hasCellSnapshot: (sheetName, address) => this.cellCache.hasCellSnapshot(sheetName, address),
      setCellSnapshot: (snapshot) => {
        this.setCellSnapshot(snapshot, { force: true, forceOptimistic: true })
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
      onViewportPatchApplied: (patch, result) => this.handleViewportPatchApplied(patch, result),
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
    const result = this.cellCache.writeCellSnapshot(snapshot, options)
    if (result.changed && result.acceptedSnapshot && options.emitLocalDelta !== false) {
      this.localRevision += 1
      this.emitLocalCellSnapshotDelta(result.acceptedSnapshot)
    }
  }

  clearOptimisticCellFlagsForSheet(sheetName: string): void {
    this.rangeOverlayStore.dropSheets([sheetName])
    if (this.cellCache.clearOptimisticCellFlagsForSheet(sheetName)) {
      this.localRevision += 1
    }
  }

  beginOptimisticClearRange(range: CellRangeRef): (() => void) | null {
    return this.rangeOverlayStore.register(range, createIdempotentContentClearedSnapshot)
  }

  setColumnWidth(sheetName: string, columnIndex: number, width: number): void {
    assertValidProjectedAxisMutation('column', columnIndex, width)
    const previousWidth = this.axisStore.getColumnWidths(sheetName)[columnIndex]
    this.axisStore.setColumnWidth(sheetName, columnIndex, width)
    this.notifySheetChannels(sheetName, ['columnWidths'])
    if (this.axisStore.getColumnWidths(sheetName)[columnIndex] !== previousWidth) {
      this.emitLocalAxisDelta(sheetName, 'column', columnIndex)
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
      this.emitLocalAxisDelta(sheetName, 'column', columnIndex)
    }
  }

  setColumnHidden(sheetName: string, columnIndex: number, hidden: boolean, size: number): void {
    assertValidProjectedAxisMutation('column', columnIndex, size)
    const previousWidth = this.axisStore.getColumnWidths(sheetName)[columnIndex]
    this.axisStore.setColumnHidden(sheetName, columnIndex, hidden, size)
    this.notifySheetChannels(sheetName, ['columnWidths', 'hiddenColumns'])
    if (this.axisStore.getColumnWidths(sheetName)[columnIndex] !== previousWidth) {
      this.emitLocalAxisDelta(sheetName, 'column', columnIndex)
    }
  }

  rollbackColumnHidden(sheetName: string, columnIndex: number, previous: { hidden: boolean; size: number | undefined }): void {
    assertValidProjectedAxisMutation('column', columnIndex, previous.size)
    const previousWidth = this.axisStore.getColumnWidths(sheetName)[columnIndex]
    this.axisStore.rollbackColumnHidden(sheetName, columnIndex, previous)
    this.notifySheetChannels(sheetName, ['columnWidths', 'hiddenColumns'])
    if (this.axisStore.getColumnWidths(sheetName)[columnIndex] !== previousWidth) {
      this.emitLocalAxisDelta(sheetName, 'column', columnIndex)
    }
  }

  setRowHeight(sheetName: string, rowIndex: number, height: number): void {
    assertValidProjectedAxisMutation('row', rowIndex, height)
    const previousHeight = this.axisStore.getRowHeights(sheetName)[rowIndex]
    this.axisStore.setRowHeight(sheetName, rowIndex, height)
    this.notifySheetChannels(sheetName, ['rowHeights'])
    if (this.axisStore.getRowHeights(sheetName)[rowIndex] !== previousHeight) {
      this.emitLocalAxisDelta(sheetName, 'row', rowIndex)
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
      this.emitLocalAxisDelta(sheetName, 'row', rowIndex)
    }
  }

  setRowHidden(sheetName: string, rowIndex: number, hidden: boolean, size: number): void {
    assertValidProjectedAxisMutation('row', rowIndex, size)
    const previousHeight = this.axisStore.getRowHeights(sheetName)[rowIndex]
    this.axisStore.setRowHidden(sheetName, rowIndex, hidden, size)
    this.notifySheetChannels(sheetName, ['rowHeights', 'hiddenRows'])
    if (this.axisStore.getRowHeights(sheetName)[rowIndex] !== previousHeight) {
      this.emitLocalAxisDelta(sheetName, 'row', rowIndex)
    }
  }

  rollbackRowHidden(sheetName: string, rowIndex: number, previous: { hidden: boolean; size: number | undefined }): void {
    assertValidProjectedAxisMutation('row', rowIndex, previous.size)
    const previousHeight = this.axisStore.getRowHeights(sheetName)[rowIndex]
    this.axisStore.rollbackRowHidden(sheetName, rowIndex, previous)
    this.notifySheetChannels(sheetName, ['rowHeights', 'hiddenRows'])
    if (this.axisStore.getRowHeights(sheetName)[rowIndex] !== previousHeight) {
      this.emitLocalAxisDelta(sheetName, 'row', rowIndex)
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
    options: { readonly initialPatch?: 'full' | 'none' } = {},
  ): () => void {
    const unsubscribe = this.patchCoordinator.subscribeViewport(
      sheetName,
      viewport,
      listener,
      options.initialPatch === undefined ? {} : { initialPatch: options.initialPatch },
    )
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
    this.localWorkbookDeltaListeners.add(listener)
    const unsubscribeClient = this.client.subscribeWorkbookDeltas((bytes) => {
      listener(decodeWorkbookDeltaBatchV3(bytes))
    })
    return () => {
      this.localWorkbookDeltaListeners.delete(listener)
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

  private emitLocalCellSnapshotDelta(snapshot: CellSnapshot): void {
    if (this.localWorkbookDeltaListeners.size === 0) {
      return
    }
    const startedAt = nowMs()
    const identity = this.resolveSheetIdentity(snapshot.sheetName)
    if (!identity) {
      return
    }
    const seq = this.nextLocalWorkbookDeltaSeq()
    const batch = buildLocalCellSnapshotWorkbookDelta({ identity, seq, snapshot })
    this.localWorkbookDeltaListeners.forEach((listener) => {
      listener(batch)
    })
    getWorkbookScrollPerfCollector()?.noteRendererDeltaApply({
      dirtyTileCount: 1,
      durationMs: Math.max(0, nowMs() - startedAt),
      mutationCount: 1,
    })
  }

  private resolveSheetIdentity(sheetName: string): SheetIdentity | null {
    return this.sheetIdentitiesByName.get(sheetName) ?? null
  }

  private emitLocalAxisDelta(sheetName: string, axis: 'column' | 'row', index: number): void {
    if (this.localWorkbookDeltaListeners.size === 0) {
      return
    }
    const startedAt = nowMs()
    const identity = this.resolveSheetIdentity(sheetName)
    if (!identity) {
      return
    }
    const seq = this.nextLocalWorkbookDeltaSeq()
    const batch = buildLocalAxisWorkbookDelta({ axis, identity, index, seq })
    this.localWorkbookDeltaListeners.forEach((listener) => {
      listener(batch)
    })
    getWorkbookScrollPerfCollector()?.noteRendererDeltaApply({
      dirtyTileCount: 1,
      durationMs: Math.max(0, nowMs() - startedAt),
      mutationCount: 1,
    })
  }

  private nextLocalWorkbookDeltaSeq(): number {
    this.localWorkbookDeltaSeq = Math.max(
      this.localWorkbookDeltaSeq,
      this.patchRevisionGate.getLastBatchId(),
      this.patchRevisionGate.getLastAuthoritativeRevision() ?? 0,
    )
    return ++this.localWorkbookDeltaSeq
  }

  private noteObservedBatchId(batchId: number): void {
    this.patchRevisionGate.noteObservedBatchId(batchId)
  }

  private applyCachedRangeStyleMutation(
    range: CellRangeRef,
    mutateStyle: (baseStyle: CellStyleRecord) => Omit<CellStyleRecord, 'id'>,
  ): (() => void) | null {
    return this.rangeOverlayStore.register(range, (snapshot) => this.applyStyleMutationToSnapshot(snapshot, mutateStyle))
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
    const baseStyle = this.cellCache.getCellStyle(snapshot.styleId) ?? { id: DEFAULT_STYLE_ID }
    const nextStyle = this.internLocalCellStyle(mutateStyle(baseStyle))
    return nextStyle.id === DEFAULT_STYLE_ID ? omitSnapshotStyleId(snapshot) : { ...snapshot, styleId: nextStyle.id }
  }
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

function normalizeProjectedCellStyle(style: Omit<CellStyleRecord, 'id'>): Omit<CellStyleRecord, 'id'> {
  return {
    ...(style.fill?.backgroundColor ? { fill: { backgroundColor: style.fill.backgroundColor } } : {}),
    ...(style.font && Object.keys(style.font).length > 0 ? { font: { ...style.font } } : {}),
    ...(style.alignment && Object.keys(style.alignment).length > 0 ? { alignment: { ...style.alignment } } : {}),
    ...(style.borders && Object.keys(style.borders).length > 0
      ? {
          borders: {
            ...(style.borders.top ? { top: { ...style.borders.top } } : {}),
            ...(style.borders.right ? { right: { ...style.borders.right } } : {}),
            ...(style.borders.bottom ? { bottom: { ...style.borders.bottom } } : {}),
            ...(style.borders.left ? { left: { ...style.borders.left } } : {}),
          },
        }
      : {}),
    ...(style.protection ? { protection: { ...style.protection } } : {}),
  }
}

function projectedCellStyleKey(style: Omit<CellStyleRecord, 'id'>): string {
  return JSON.stringify({
    alignment: style.alignment ?? null,
    borders: style.borders ?? null,
    fill: style.fill?.backgroundColor ?? null,
    font: style.font ?? null,
    protection: style.protection ?? null,
  })
}

function projectedCellStyleIdForKey(key: string): string {
  let hash = 2166136261
  for (let index = 0; index < key.length; index += 1) {
    hash ^= key.charCodeAt(index)
    hash = Math.imul(hash, 16777619)
  }
  return `style-local-${(hash >>> 0).toString(16)}`
}

function cloneProjectedStyleWithoutId(style: CellStyleRecord): Omit<CellStyleRecord, 'id'> {
  return normalizeProjectedCellStyle(style)
}

function applyProjectedStylePatch(baseStyle: CellStyleRecord, patch: CellStylePatch): Omit<CellStyleRecord, 'id'> {
  const next = cloneProjectedStyleWithoutId(baseStyle)
  if (patch.fill === null) {
    delete next.fill
  } else if (patch.fill !== undefined) {
    const backgroundColor = patch.fill.backgroundColor
    if (backgroundColor === null) {
      delete next.fill
    } else if (backgroundColor !== undefined) {
      next.fill = { backgroundColor }
    }
  }
  if (patch.font === null) {
    delete next.font
  } else if (patch.font) {
    const font = { ...next.font }
    applyOptionalProjectedField(font, 'family', patch.font.family)
    applyOptionalProjectedField(font, 'size', patch.font.size)
    applyOptionalProjectedField(font, 'bold', patch.font.bold)
    applyOptionalProjectedField(font, 'italic', patch.font.italic)
    applyOptionalProjectedField(font, 'underline', patch.font.underline)
    applyOptionalProjectedField(font, 'color', patch.font.color)
    if (Object.keys(font).length > 0) {
      next.font = font
    } else {
      delete next.font
    }
  }
  if (patch.alignment === null) {
    delete next.alignment
  } else if (patch.alignment) {
    const alignment = { ...next.alignment }
    applyOptionalProjectedField(alignment, 'horizontal', patch.alignment.horizontal)
    applyOptionalProjectedField(alignment, 'vertical', patch.alignment.vertical)
    applyOptionalProjectedField(alignment, 'wrap', patch.alignment.wrap)
    applyOptionalProjectedField(alignment, 'indent', patch.alignment.indent)
    applyOptionalProjectedField(alignment, 'shrinkToFit', patch.alignment.shrinkToFit)
    applyOptionalProjectedField(alignment, 'readingOrder', patch.alignment.readingOrder)
    applyOptionalProjectedField(alignment, 'textRotation', patch.alignment.textRotation)
    applyOptionalProjectedField(alignment, 'justifyLastLine', patch.alignment.justifyLastLine)
    if (Object.keys(alignment).length > 0) {
      next.alignment = alignment
    } else {
      delete next.alignment
    }
  }
  if (patch.borders === null) {
    delete next.borders
  } else if (patch.borders) {
    const borders = { ...next.borders }
    applyProjectedBorderSidePatch(borders, 'top', patch.borders.top)
    applyProjectedBorderSidePatch(borders, 'right', patch.borders.right)
    applyProjectedBorderSidePatch(borders, 'bottom', patch.borders.bottom)
    applyProjectedBorderSidePatch(borders, 'left', patch.borders.left)
    if (Object.keys(borders).length > 0) {
      next.borders = borders
    } else {
      delete next.borders
    }
  }
  return normalizeProjectedCellStyle(next)
}

function applyProjectedBorderSidePatch(
  borders: NonNullable<CellStyleRecord['borders']>,
  side: keyof NonNullable<CellStyleRecord['borders']>,
  patch: NonNullable<CellStylePatch['borders']>['top'] | null | undefined,
): void {
  if (patch === undefined) {
    return
  }
  if (patch === null) {
    delete borders[side]
    return
  }
  const nextSide: Partial<NonNullable<CellStyleRecord['borders']>['top']> = { ...borders[side] }
  applyOptionalProjectedField(nextSide, 'style', patch.style)
  applyOptionalProjectedField(nextSide, 'weight', patch.weight)
  applyOptionalProjectedField(nextSide, 'color', patch.color)
  if (nextSide.style && nextSide.weight && nextSide.color) {
    borders[side] = {
      color: nextSide.color,
      style: nextSide.style,
      weight: nextSide.weight,
    }
  } else {
    delete borders[side]
  }
}

function clearProjectedStyleFields(baseStyle: CellStyleRecord, fields: readonly CellStyleField[] | undefined): Omit<CellStyleRecord, 'id'> {
  if (fields === undefined || fields.length === 0) {
    return {}
  }
  const next = cloneProjectedStyleWithoutId(baseStyle)
  const cleared = new Set(fields)
  if (cleared.has('backgroundColor')) {
    delete next.fill
  }
  const font = filterProjectedStyleSection(
    next.font,
    [
      ['fontFamily', 'family'],
      ['fontSize', 'size'],
      ['fontBold', 'bold'],
      ['fontItalic', 'italic'],
      ['fontUnderline', 'underline'],
      ['fontColor', 'color'],
    ],
    cleared,
  )
  if (font) {
    next.font = font
  } else {
    delete next.font
  }
  const alignment = filterProjectedStyleSection(
    next.alignment,
    [
      ['alignmentHorizontal', 'horizontal'],
      ['alignmentVertical', 'vertical'],
      ['alignmentWrap', 'wrap'],
      ['alignmentIndent', 'indent'],
      ['alignmentShrinkToFit', 'shrinkToFit'],
      ['alignmentReadingOrder', 'readingOrder'],
      ['alignmentTextRotation', 'textRotation'],
      ['alignmentJustifyLastLine', 'justifyLastLine'],
    ],
    cleared,
  )
  if (alignment) {
    next.alignment = alignment
  } else {
    delete next.alignment
  }
  const borders = filterProjectedStyleSection(
    next.borders,
    [
      ['borderTop', 'top'],
      ['borderRight', 'right'],
      ['borderBottom', 'bottom'],
      ['borderLeft', 'left'],
    ],
    cleared,
  )
  if (borders) {
    next.borders = borders
  } else {
    delete next.borders
  }
  return normalizeProjectedCellStyle(next)
}

function filterProjectedStyleSection<T extends object>(
  section: T | undefined,
  fields: ReadonlyArray<readonly [CellStyleField, keyof T]>,
  cleared: ReadonlySet<CellStyleField>,
): T | undefined {
  if (!section) {
    return undefined
  }
  const nextSection = { ...section }
  fields.forEach(([field, key]) => {
    if (cleared.has(field)) {
      delete nextSection[key]
    }
  })
  return Object.keys(nextSection).length > 0 ? nextSection : undefined
}

function applyOptionalProjectedField<T extends object, K extends keyof T>(target: T, key: K, value: T[K] | null | undefined): void {
  if (value === undefined) {
    return
  }
  if (value === null) {
    delete target[key]
    return
  }
  target[key] = value
}

function nowMs(): number {
  return typeof performance === 'undefined' ? Date.now() : performance.now()
}

function assertValidProjectedAxisMutation(axis: 'column' | 'row', index: number, size: number | undefined): void {
  const axisLength = axis === 'column' ? MAX_COLS : MAX_ROWS
  if (!Number.isInteger(index) || index < 0 || index >= axisLength) {
    throw new Error(`Invalid projected ${axis} index: ${index}`)
  }
  if (size !== undefined && (!Number.isFinite(size) || size < 0)) {
    throw new Error(`Invalid projected ${axis} size: ${size}`)
  }
}

function buildAxisEntries(
  sizes: Readonly<Record<number, number>>,
  hiddenAxes: Readonly<Record<number, true>>,
  idPrefix: 'col' | 'row',
): WorkbookAxisEntrySnapshot[] {
  const indexes = new Set<number>()
  for (const key of Object.keys(sizes)) {
    const index = Number(key)
    if (Number.isInteger(index) && index >= 0) {
      indexes.add(index)
    }
  }
  for (const key of Object.keys(hiddenAxes)) {
    const index = Number(key)
    if (Number.isInteger(index) && index >= 0) {
      indexes.add(index)
    }
  }
  return [...indexes]
    .toSorted((left, right) => left - right)
    .map((index) => ({
      id: `${idPrefix}-${index}`,
      index,
      size: sizes[index] ?? null,
      hidden: hiddenAxes[index] === true ? true : null,
    }))
}
