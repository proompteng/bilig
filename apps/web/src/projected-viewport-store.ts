import type { GridEngineLike } from '@bilig/grid'
import { parseCellAddress } from '@bilig/formula'
import type { CellSnapshot, CellStyleRecord, Viewport, WorkbookAxisEntrySnapshot, WorkbookMergeRangeSnapshot } from '@bilig/protocol'
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
import { DirtyMaskV3 } from '../../../packages/grid/src/renderer-v3/tile-damage-index.js'
import { normalizeWorkbookMergeRange } from './worker-runtime-support.js'

export interface ProjectedViewportStoreOptions {
  readonly maxCachedCellsPerSheet?: number
}
type CellItem = readonly [number, number]
type SheetViewportChannel = 'columnWidths' | 'rowHeights' | 'hiddenColumns' | 'hiddenRows' | 'freeze' | 'merges'
type SheetIdentity = { readonly sheetId: number; readonly sheetOrdinal: number }
export class ProjectedViewportStore implements GridEngineLike {
  private readonly options: ProjectedViewportStoreOptions
  private readonly cellCache: ProjectedViewportCellCache
  private readonly axisStore: ProjectedViewportAxisStore
  private readonly patchCoordinator: ProjectedViewportPatchCoordinator
  private tileSceneStore: ProjectedTileSceneStore | null = null
  private readonly localWorkbookDeltaListeners = new Set<(batch: WorkbookDeltaBatchV3) => void>()
  private readonly sheetIdentitiesByName = new Map<string, SheetIdentity>()
  private readonly sheetChannelListeners = new Map<string, Map<SheetViewportChannel, Set<() => void>>>()
  private readonly mergeRangesBySheet = new Map<string, Map<string, WorkbookMergeRangeSnapshot>>()
  private lastBatchId = 0
  private lastAuthoritativeRevision: number | null = null
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
    return { batchId: this.lastBatchId }
  }

  getLastAuthoritativeRevision(): number | null {
    return this.lastAuthoritativeRevision
  }

  setCellSnapshot(snapshot: CellSnapshot, options: { force?: boolean } = {}): void {
    if (this.cellCache.setCellSnapshot(snapshot, options)) {
      this.emitLocalCellSnapshotDelta(snapshot)
    }
  }

  setColumnWidth(sheetName: string, columnIndex: number, width: number): void {
    this.axisStore.setColumnWidth(sheetName, columnIndex, width)
    this.notifySheetChannels(sheetName, ['columnWidths'])
  }

  ackColumnWidth(sheetName: string, columnIndex: number, width: number): void {
    this.axisStore.ackColumnWidth(sheetName, columnIndex, width)
    this.notifySheetChannels(sheetName, ['columnWidths'])
  }

  rollbackColumnWidth(sheetName: string, columnIndex: number, width: number | undefined): void {
    this.axisStore.rollbackColumnWidth(sheetName, columnIndex, width)
    this.notifySheetChannels(sheetName, ['columnWidths'])
  }

  setColumnHidden(sheetName: string, columnIndex: number, hidden: boolean, size: number): void {
    this.axisStore.setColumnHidden(sheetName, columnIndex, hidden, size)
    this.notifySheetChannels(sheetName, ['columnWidths', 'hiddenColumns'])
  }

  rollbackColumnHidden(sheetName: string, columnIndex: number, previous: { hidden: boolean; size: number | undefined }): void {
    this.axisStore.rollbackColumnHidden(sheetName, columnIndex, previous)
    this.notifySheetChannels(sheetName, ['columnWidths', 'hiddenColumns'])
  }

  setRowHeight(sheetName: string, rowIndex: number, height: number): void {
    this.axisStore.setRowHeight(sheetName, rowIndex, height)
    this.notifySheetChannels(sheetName, ['rowHeights'])
  }

  ackRowHeight(sheetName: string, rowIndex: number, height: number): void {
    this.axisStore.ackRowHeight(sheetName, rowIndex, height)
    this.notifySheetChannels(sheetName, ['rowHeights'])
  }

  rollbackRowHeight(sheetName: string, rowIndex: number, height: number | undefined): void {
    this.axisStore.rollbackRowHeight(sheetName, rowIndex, height)
    this.notifySheetChannels(sheetName, ['rowHeights'])
  }

  setRowHidden(sheetName: string, rowIndex: number, hidden: boolean, size: number): void {
    this.axisStore.setRowHidden(sheetName, rowIndex, hidden, size)
    this.notifySheetChannels(sheetName, ['rowHeights', 'hiddenRows'])
  }

  rollbackRowHidden(sheetName: string, rowIndex: number, previous: { hidden: boolean; size: number | undefined }): void {
    this.axisStore.rollbackRowHidden(sheetName, rowIndex, previous)
    this.notifySheetChannels(sheetName, ['rowHeights', 'hiddenRows'])
  }

  setKnownSheets(sheetNames: readonly string[]): void {
    const removedSheets = this.cellCache.setKnownSheets(sheetNames)
    this.axisStore.dropSheets(removedSheets)
    this.tileSceneStore?.dropSheets(removedSheets)
    removedSheets.forEach((sheetName) => {
      this.mergeRangesBySheet.delete(sheetName)
      this.sheetChannelListeners.delete(sheetName)
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
      unsubscribe = store.subscribe(subscription, listener)
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
    const batchId = patch.metrics?.batchId
    if (Number.isInteger(batchId) && batchId >= 0) {
      this.lastBatchId = batchId
    }
    const authoritativeRevision = patch.authoritativeRevision
    if (typeof authoritativeRevision === 'number' && Number.isInteger(authoritativeRevision) && authoritativeRevision >= 0) {
      this.lastAuthoritativeRevision = authoritativeRevision
    }
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
    const parsed = parseCellAddress(snapshot.address, snapshot.sheetName)
    const valueSeq = Math.max(0, snapshot.version)
    const batch: WorkbookDeltaBatchV3 = {
      axisSeqX: 0,
      axisSeqY: 0,
      calcSeq: valueSeq,
      dirty: {
        axisX: new Uint32Array(),
        axisY: new Uint32Array(),
        cellRanges: new Uint32Array([
          parsed.row,
          parsed.row,
          parsed.col,
          parsed.col,
          DirtyMaskV3.Value | DirtyMaskV3.Style | DirtyMaskV3.Text | DirtyMaskV3.Rect,
        ]),
      },
      freezeSeq: 0,
      magic: 'bilig.workbook.delta.v3',
      seq: ++this.localWorkbookDeltaSeq,
      sheetId: identity.sheetId,
      sheetOrdinal: identity.sheetOrdinal,
      source: 'localOptimistic',
      styleSeq: valueSeq,
      valueSeq,
      version: 1,
    }
    this.localWorkbookDeltaListeners.forEach((listener) => {
      listener(batch)
    })
    getWorkbookScrollPerfCollector()?.noteRendererDeltaApply({
      dirtyTileCount: 1,
      durationMs: Math.max(0, nowMs() - startedAt),
      mutationCount: 1,
    })
  }

  private resolveSheetIdentity(sheetName: string): SheetIdentity {
    return this.sheetIdentitiesByName.get(sheetName) ?? { sheetId: 0, sheetOrdinal: 0 }
  }
}

function nowMs(): number {
  return typeof performance === 'undefined' ? Date.now() : performance.now()
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
