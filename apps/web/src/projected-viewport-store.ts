import type { GridEngineLike, GridRenderRevisionSnapshot } from '@bilig/grid'
import { parseCellAddress } from '@bilig/formula'
import {
  MAX_COLS,
  MAX_ROWS,
  type CellRangeRef,
  type CellSnapshot,
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
export class ProjectedViewportStore implements GridEngineLike {
  private readonly options: ProjectedViewportStoreOptions
  private readonly cellCache: ProjectedViewportCellCache
  private readonly axisStore: ProjectedViewportAxisStore
  private readonly patchCoordinator: ProjectedViewportPatchCoordinator
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
    return this.cellCache.getCell(sheetName, address)
  }

  forEachCellSnapshotInRange(range: CellRangeRef, listener: (snapshot: CellSnapshot) => void): void {
    this.cellCache.forEachCellSnapshotInRange(range, listener)
  }

  getCellStyle(styleId: string | undefined): CellStyleRecord | undefined {
    return this.cellCache.getCellStyle(styleId)
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
    if (this.cellCache.clearOptimisticCellFlagsForSheet(sheetName)) {
      this.localRevision += 1
    }
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
    return this.patchCoordinator.subscribeViewport(
      sheetName,
      viewport,
      listener,
      options.initialPatch === undefined ? {} : { initialPatch: options.initialPatch },
    )
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
