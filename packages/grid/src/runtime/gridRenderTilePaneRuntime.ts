import { ValueTag, type CellSnapshot, type Viewport } from '@bilig/protocol'
import { noteRendererTileReadiness, noteTypeGpuTileCacheStaleLookups, noteTypeGpuTileCacheVisibleMark } from '../grid-render-counters.js'
import type { GridEngineLike } from '../grid-engine.js'
import type { GridMetrics } from '../gridMetrics.js'
import type { Item } from '../gridTypes.js'
import { buildLocalFixedRenderTiles } from '../renderer-v3/local-render-tile-materializer.js'
import { GRID_RECT_INSTANCE_FLOAT_COUNT_V3 } from '../renderer-v3/rect-instance-buffer.js'
import { buildFixedRenderTilePaneStates } from '../renderer-v3/render-tile-pane-builder.js'
import type { GridRenderTile, GridRenderTileSceneChange, GridRenderTileSource } from '../renderer-v3/render-tile-source.js'
import type { WorkbookRenderTilePaneState } from '../renderer-v3/render-tile-pane-state.js'
import type { WorkbookDeltaBatchLikeV3 } from '../renderer-v3/tile-damage-index.js'
import type { TileKey53 } from '../renderer-v3/tile-key.js'
import type { GridTileInterestBatchV3, GridTileReadinessSnapshotV3 } from './gridTileCoordinator.js'
import type { GridRuntimeHost } from './gridRuntimeHost.js'
import {
  buildRetainedFixedRenderTileDataPanesCompatibility,
  isResidentGridRenderTile,
  matchesRenderTileSheetIdentity,
  resolveGridRenderTileInputSheetOrdinal,
  resolveRenderTileInterestTileKeys,
  resolveRenderTileViewportUnion,
  resolveWarmRenderTileKeys,
  sameLocalInvalidationConnectionIdentity,
  sameRenderTileDeltaConnectionIdentity,
  sameRetainedFixedRenderTileDataPanesCompatibility,
  sameWorkbookDeltaConnectionIdentity,
  shouldForceLocalTilesForWorkbookDelta,
  upsertRenderTileIntoHost,
  type GridRenderTileInterestRuntimeInput,
  type LocalInvalidationConnectionIdentity,
  type RenderTileDeltaConnectionIdentity,
  type RetainedFixedRenderTileDataPanesCompatibility,
  type WorkbookDeltaConnectionIdentity,
} from './gridRenderTilePaneRuntimeHelpers.js'
import { GridVisibleTextRefreshCache } from './gridVisibleTextRefreshCache.js'

type SortedAxisOverrides = readonly (readonly [number, number])[]

export interface GridRenderTilePaneRuntimeState {
  readonly needsLocalCellInvalidation: boolean
  readonly preloadDataPanes: readonly WorkbookRenderTilePaneState[]
  readonly renderTilePanes: readonly WorkbookRenderTilePaneState[]
  readonly residentBodyPane: WorkbookRenderTilePaneState | null
  readonly residentDataPanes: readonly WorkbookRenderTilePaneState[]
  readonly tileReadiness: GridTileReadinessSnapshotV3
}

export interface GridRenderTilePaneBridgeState {
  readonly forceLocalTiles: boolean
  readonly localFallbackRevision: number
  readonly renderTileRevision: number
}

export interface GridRenderTilePaneRuntimeInput {
  readonly columnWidths: Readonly<Record<number, number>>
  readonly dprBucket: number
  readonly editingCell?: Item | null | undefined
  readonly engine: GridEngineLike
  readonly freezeCols: number
  readonly freezeRows: number
  readonly forceLocalTiles?: boolean | undefined
  readonly frozenColumnWidth: number
  readonly frozenRowHeight: number
  readonly gridMetrics: GridMetrics
  readonly gridRuntimeHost: GridRuntimeHost
  readonly hostClientHeight: number
  readonly hostClientWidth: number
  readonly hostReady: boolean
  readonly renderTileSource?: GridRenderTileSource | undefined
  readonly renderTileViewport: Viewport
  readonly residentViewport: Viewport
  readonly rowHeights: Readonly<Record<number, number>>
  readonly sceneRevision: number
  readonly selectedCell?: Item | undefined
  readonly selectedCellSnapshot?: CellSnapshot | null | undefined
  readonly sheetId?: number | undefined
  readonly sheetOrdinal?: number | undefined
  readonly sheetName: string
  readonly sortedColumnWidthOverrides: SortedAxisOverrides
  readonly sortedRowHeightOverrides: SortedAxisOverrides
  readonly visibleViewport: Viewport
}

export interface GridRenderTileDeltaRuntimeInput {
  readonly dprBucket: number
  readonly freezeCols?: number | undefined
  readonly freezeRows?: number | undefined
  readonly gridRuntimeHost: GridRuntimeHost
  readonly renderTileSource?: GridRenderTileSource | undefined
  readonly renderTileViewport: Viewport
  readonly residentViewport?: Viewport | undefined
  readonly sheetId?: number | undefined
  readonly sheetOrdinal?: number | undefined
  readonly sheetName: string
  readonly visibleViewport?: Viewport | undefined
}

export interface GridRenderTileDamageRuntimeInput {
  readonly dprBucket: number
  readonly gridRuntimeHost: GridRuntimeHost
  readonly renderTileSource?: GridRenderTileSource | undefined
  readonly sheetId?: number | undefined
  readonly sheetOrdinal?: number | undefined
}

export interface GridRenderTileLocalInvalidationRuntimeInput {
  readonly engine: GridEngineLike
  readonly needsLocalCellInvalidation: boolean
  readonly sheetName: string
  readonly visibleAddresses: readonly string[]
}

export interface GridRenderTileConnectionRuntimeInput {
  readonly dprBucket: number
  readonly engine: GridEngineLike
  readonly freezeCols?: number | undefined
  readonly freezeRows?: number | undefined
  readonly gridRuntimeHost: GridRuntimeHost
  readonly needsLocalCellInvalidation: boolean
  readonly renderTileSource?: GridRenderTileSource | undefined
  readonly renderTileViewport: Viewport
  readonly residentViewport?: Viewport | undefined
  readonly sheetId?: number | undefined
  readonly sheetOrdinal?: number | undefined
  readonly sheetName: string
  readonly visibleAddresses: readonly string[]
  readonly visibleViewport?: Viewport | undefined
}

const EMPTY_TILE_PANE_RUNTIME_STATE: GridRenderTilePaneRuntimeState = Object.freeze({
  needsLocalCellInvalidation: false,
  preloadDataPanes: [],
  renderTilePanes: [],
  residentBodyPane: null,
  residentDataPanes: [],
  tileReadiness: {
    exactHits: [],
    misses: [],
    staleHits: [],
    staleLookupCount: 0,
    staleLookupScannedEntries: 0,
    visibleDirtyTileKeys: [],
    visibleMarkedTiles: 0,
    warmDirtyTileKeys: [],
  },
})

const INITIAL_RENDER_TILE_PANE_BRIDGE_STATE: GridRenderTilePaneBridgeState = Object.freeze({
  forceLocalTiles: false,
  localFallbackRevision: 0,
  renderTileRevision: 0,
})

type TileResolutionSource = 'local' | 'remote'

interface FixedRenderTileDataPanesResolution {
  readonly panes: readonly WorkbookRenderTilePaneState[]
  readonly source: TileResolutionSource
}

interface GridRenderTileResolution {
  readonly tiles: readonly GridRenderTile[]
  readonly source: TileResolutionSource
}

interface GridRenderTilePreloadResolution {
  readonly panes: readonly WorkbookRenderTilePaneState[]
  readonly tiles: readonly GridRenderTile[]
}

function expectedRenderTileGridBorderCount(tile: GridRenderTile): number {
  const rowCount = tile.bounds.rowEnd - tile.bounds.rowStart + 1
  const colCount = tile.bounds.colEnd - tile.bounds.colStart + 1
  return rowCount > 0 && colCount > 0 ? rowCount + colCount : 0
}

function countRenderTileGridBorderRects(tile: GridRenderTile): number {
  const readableRectCount = Math.min(tile.rectCount, Math.floor(tile.rectInstances.length / GRID_RECT_INSTANCE_FLOAT_COUNT_V3))
  let count = 0
  for (let index = 0; index < readableRectCount; index += 1) {
    const offset = index * GRID_RECT_INSTANCE_FLOAT_COUNT_V3
    const width = tile.rectInstances[offset + 2] ?? 0
    const height = tile.rectInstances[offset + 3] ?? 0
    const borderAlpha = tile.rectInstances[offset + 11] ?? 0
    const borderThickness = tile.rectInstances[offset + 13] ?? 0
    if (borderAlpha > 0 && borderThickness > 0 && ((width <= 1.5 && height > 0) || (height <= 1.5 && width > 0))) {
      count += 1
    }
  }
  return count
}

function hasCompleteRenderTileGrid(tile: GridRenderTile): boolean {
  const expectedBorderCount = expectedRenderTileGridBorderCount(tile)
  return expectedBorderCount === 0 || countRenderTileGridBorderRects(tile) >= expectedBorderCount
}

function selectedSnapshotTextHint(snapshot: CellSnapshot | null | undefined): string | null {
  if (!snapshot) {
    return null
  }
  if (snapshot.input !== undefined && snapshot.input !== null && snapshot.input !== '') {
    return String(snapshot.input)
  }
  if (snapshot.formula !== undefined && snapshot.formula.length > 0) {
    return snapshot.formula
  }
  if (snapshot.value.tag === ValueTag.String) {
    return snapshot.value.value
  }
  if (snapshot.value.tag === ValueTag.Number) {
    return String(snapshot.value.value)
  }
  return null
}

function findSelectedTextRun(tile: GridRenderTile | null, selectedCell: Item | undefined): { readonly text: string } | null {
  if (!tile || !selectedCell) {
    return null
  }
  return tile.textRuns.find((run) => run.col === selectedCell[0] && run.row === selectedCell[1] && run.text.length > 0) ?? null
}

function tileSelectedTextNeedsLocalRefresh(
  tile: GridRenderTile | null,
  selectedCell: Item | undefined,
  selectedCellSnapshot: CellSnapshot | null | undefined,
): boolean {
  if (!selectedCell) {
    return false
  }
  const selectedRun = findSelectedTextRun(tile, selectedCell)
  const expectedText = selectedSnapshotTextHint(selectedCellSnapshot)
  if (expectedText === null) {
    return selectedRun !== null
  }
  return selectedRun?.text !== expectedText
}

interface RuntimeConnection<Identity> {
  readonly identity: Identity
  readonly unsubscribe: (() => void) | undefined
}

export class GridRenderTilePaneRuntime {
  private retainedFixedRenderTileDataPanes: {
    readonly compatibility: RetainedFixedRenderTileDataPanesCompatibility
    readonly panes: readonly WorkbookRenderTilePaneState[]
  } | null = null
  private bridgeState = INITIAL_RENDER_TILE_PANE_BRIDGE_STATE
  private readonly bridgeListeners = new Set<() => void>()
  private readonly lastWorkbookDeltaSeqBySheetAndSource = new Map<string, number>()
  private renderTileDeltaConnection: RuntimeConnection<RenderTileDeltaConnectionIdentity> | null = null
  private workbookDeltaConnection: RuntimeConnection<WorkbookDeltaConnectionIdentity> | null = null
  private localInvalidationConnection: RuntimeConnection<LocalInvalidationConnectionIdentity> | null = null
  private readonly visibleTextRefreshCache = new GridVisibleTextRefreshCache()

  resolve(input: GridRenderTilePaneRuntimeInput): GridRenderTilePaneRuntimeState {
    if (!input.hostReady) {
      return EMPTY_TILE_PANE_RUNTIME_STATE
    }
    const resolution = this.resolveTiles(input)
    const preloadResolution = this.resolvePreloadPanes(input, resolution?.tiles ?? [])
    const tileReadiness = this.resolveTileReadiness(input, resolution?.tiles ?? [], preloadResolution.tiles)
    const fixedRenderTileDataPanes = resolution ? this.buildFixedRenderTileDataPanes(input, resolution) : null
    if (input.sheetId !== undefined && fixedRenderTileDataPanes?.source === 'remote') {
      this.retainedFixedRenderTileDataPanes = {
        compatibility: buildRetainedFixedRenderTileDataPanesCompatibility(input),
        panes: fixedRenderTileDataPanes.panes,
      }
    }

    const shouldUseRemoteRenderTileSource = input.renderTileSource !== undefined && input.sheetId !== undefined
    const retainedFixedRenderTileDataPanes =
      fixedRenderTileDataPanes?.panes ??
      (shouldUseRemoteRenderTileSource &&
      this.retainedFixedRenderTileDataPanes &&
      sameRetainedFixedRenderTileDataPanesCompatibility(
        this.retainedFixedRenderTileDataPanes.compatibility,
        buildRetainedFixedRenderTileDataPanesCompatibility(input),
      )
        ? this.retainedFixedRenderTileDataPanes.panes
        : null)
    const residentDataPanes = retainedFixedRenderTileDataPanes ?? []
    const usesLocalTilePanes = fixedRenderTileDataPanes?.source === 'local'
    return {
      needsLocalCellInvalidation: usesLocalTilePanes || !shouldUseRemoteRenderTileSource,
      preloadDataPanes: preloadResolution.panes,
      renderTilePanes: residentDataPanes,
      residentBodyPane: residentDataPanes.find((pane) => pane.paneId === 'body') ?? null,
      residentDataPanes,
      tileReadiness,
    }
  }

  clearRetainedPanes(): void {
    this.retainedFixedRenderTileDataPanes = null
  }

  snapshotBridgeState(): GridRenderTilePaneBridgeState {
    return this.bridgeState
  }

  subscribeBridgeState(listener: () => void): () => void {
    this.bridgeListeners.add(listener)
    return () => {
      this.bridgeListeners.delete(listener)
    }
  }

  noteRenderTileDelta(): GridRenderTilePaneBridgeState {
    const previous = this.bridgeState
    this.bridgeState = {
      forceLocalTiles: false,
      localFallbackRevision: previous.localFallbackRevision,
      renderTileRevision: previous.renderTileRevision + 1,
    }
    this.emitBridgeState()
    return this.bridgeState
  }

  noteWorkbookDeltaDamage(input: { readonly forceLocalTiles?: boolean | undefined } = {}): GridRenderTilePaneBridgeState {
    const previous = this.bridgeState
    const forceLocalTiles = input.forceLocalTiles ?? false
    this.bridgeState = {
      forceLocalTiles,
      localFallbackRevision: forceLocalTiles ? previous.localFallbackRevision + 1 : previous.localFallbackRevision,
      renderTileRevision: previous.renderTileRevision + 1,
    }
    this.emitBridgeState()
    return this.bridgeState
  }

  noteLocalFallbackInvalidation(): GridRenderTilePaneBridgeState {
    const previous = this.bridgeState
    this.bridgeState = {
      forceLocalTiles: true,
      localFallbackRevision: previous.localFallbackRevision + 1,
      renderTileRevision: previous.renderTileRevision,
    }
    this.emitBridgeState()
    return this.bridgeState
  }

  noteTileReadiness(readiness: GridTileReadinessSnapshotV3): void {
    const exactHits = readiness.exactHits.length
    const staleHits = readiness.staleHits.length
    const misses = readiness.misses.length
    const visibleDirtyTiles = readiness.visibleDirtyTileKeys.length
    const warmDirtyTiles = readiness.warmDirtyTileKeys.length
    if (exactHits + staleHits + misses + visibleDirtyTiles + warmDirtyTiles === 0) {
      return
    }
    if (readiness.visibleMarkedTiles > 0) {
      noteTypeGpuTileCacheVisibleMark(readiness.visibleMarkedTiles)
    }
    if (readiness.staleLookupCount > 0) {
      noteTypeGpuTileCacheStaleLookups({
        hits: staleHits,
        lookups: readiness.staleLookupCount,
        scannedEntries: readiness.staleLookupScannedEntries,
      })
    }
    noteRendererTileReadiness({
      exactHits,
      misses,
      staleHits,
      visibleDirtyTiles,
      warmDirtyTiles,
    })
  }

  connectLocalCellInvalidation(input: GridRenderTileLocalInvalidationRuntimeInput, listener?: () => void): (() => void) | undefined {
    if (!input.needsLocalCellInvalidation || input.visibleAddresses.length === 0) {
      return undefined
    }
    const invalidate = () => {
      this.clearRetainedPanes()
      this.noteLocalFallbackInvalidation()
      listener?.()
    }
    const unsubscribeCells = input.engine.subscribeCells(input.sheetName, input.visibleAddresses, invalidate)
    const unsubscribeMerges = input.engine.subscribeSheetChannel?.(input.sheetName, 'merges', invalidate)
    if (!unsubscribeMerges) {
      return unsubscribeCells
    }
    return () => {
      unsubscribeCells()
      unsubscribeMerges()
    }
  }

  syncConnections(input: GridRenderTileConnectionRuntimeInput): void {
    this.syncRenderTileDeltaConnection(input)
    this.syncWorkbookDeltaConnection(input)
    this.syncLocalInvalidationConnection(input)
  }

  disconnectConnections(): void {
    this.renderTileDeltaConnection?.unsubscribe?.()
    this.renderTileDeltaConnection = null
    this.workbookDeltaConnection?.unsubscribe?.()
    this.workbookDeltaConnection = null
    this.localInvalidationConnection?.unsubscribe?.()
    this.localInvalidationConnection = null
  }

  connectWorkbookDeltaDamage(
    input: GridRenderTileDamageRuntimeInput,
    listener?: (batch: WorkbookDeltaBatchLikeV3) => void,
  ): (() => void) | undefined {
    const renderTileSource = input.renderTileSource
    if (!renderTileSource?.subscribeWorkbookDeltas || input.sheetId === undefined) {
      return undefined
    }
    const sheetOrdinal = resolveGridRenderTileInputSheetOrdinal(input)
    return renderTileSource.subscribeWorkbookDeltas((batch) => {
      if (!matchesRenderTileSheetIdentity(batch, { sheetId: input.sheetId, sheetOrdinal })) {
        return
      }
      if (!this.applyWorkbookDeltaDamage(input, batch)) {
        return
      }
      this.noteWorkbookDeltaDamage({
        forceLocalTiles: shouldForceLocalTilesForWorkbookDelta(batch),
      })
      listener?.(batch)
    })
  }

  connectRenderTileDeltas(
    input: GridRenderTileDeltaRuntimeInput,
    listener?: (change: GridRenderTileSceneChange) => void,
  ): (() => void) | undefined {
    if (!input.renderTileSource || input.sheetId === undefined) {
      return undefined
    }
    const tileInterest = this.buildViewportTileInterest({
      ...input,
      sheetId: input.sheetId,
      visibleViewport: input.visibleViewport ?? input.residentViewport ?? input.renderTileViewport,
    })
    return input.renderTileSource.subscribeRenderTileDeltas(
      {
        ...input.renderTileViewport,
        cameraSeq: tileInterest.cameraSeq,
        dprBucket: input.dprBucket,
        initialDelta: 'full',
        sheetId: input.sheetId,
        sheetOrdinal: tileInterest.sheetOrdinal,
        sheetName: input.sheetName,
        tileInterest: {
          axisSeqX: tileInterest.axisSeqX,
          axisSeqY: tileInterest.axisSeqY,
          freezeSeq: tileInterest.freezeSeq,
          pinnedTileKeys: tileInterest.pinnedTileKeys,
          reason: tileInterest.reason,
          seq: tileInterest.seq,
          sheetOrdinal: tileInterest.sheetOrdinal,
          visibleTileKeys: tileInterest.visibleTileKeys,
          warmTileKeys: tileInterest.warmTileKeys,
        },
        warmTileKeys: tileInterest.warmTileKeys,
      },
      (change) => {
        if (change) {
          this.applyRenderTileSceneChange(input, change)
        }
        this.noteRenderTileDelta()
        listener?.(change)
      },
    )
  }

  private applyWorkbookDeltaDamage(input: GridRenderTileDamageRuntimeInput, batch: WorkbookDeltaBatchLikeV3): boolean {
    if (batch.seq !== undefined) {
      const source = 'source' in batch && typeof batch.source === 'string' ? batch.source : 'unknown'
      const sequenceKey = `${batch.sheetId ?? 'x'}:${batch.sheetOrdinal ?? 'x'}:${source}`
      const lastSeq = this.lastWorkbookDeltaSeqBySheetAndSource.get(sequenceKey) ?? -1
      if (batch.seq <= lastSeq) {
        return false
      }
      this.lastWorkbookDeltaSeqBySheetAndSource.set(sequenceKey, batch.seq)
    }
    input.gridRuntimeHost.tiles.applyWorkbookDelta(batch, { dprBucket: input.dprBucket })
    return true
  }

  private syncRenderTileDeltaConnection(input: GridRenderTileConnectionRuntimeInput): void {
    const identity: RenderTileDeltaConnectionIdentity = {
      dprBucket: input.dprBucket,
      freezeCols: input.freezeCols ?? 0,
      freezeRows: input.freezeRows ?? 0,
      renderTileSource: input.renderTileSource,
      residentViewport: input.residentViewport,
      sheetId: input.sheetId,
      sheetName: input.sheetName,
      sheetOrdinal: input.sheetOrdinal,
      viewport: input.renderTileViewport,
    }
    if (this.renderTileDeltaConnection && sameRenderTileDeltaConnectionIdentity(this.renderTileDeltaConnection.identity, identity)) {
      return
    }
    this.renderTileDeltaConnection?.unsubscribe?.()
    this.renderTileDeltaConnection = {
      identity,
      unsubscribe: this.connectRenderTileDeltas({
        dprBucket: input.dprBucket,
        gridRuntimeHost: input.gridRuntimeHost,
        freezeCols: input.freezeCols,
        freezeRows: input.freezeRows,
        renderTileSource: input.renderTileSource,
        renderTileViewport: input.renderTileViewport,
        residentViewport: input.residentViewport,
        sheetId: input.sheetId,
        sheetName: input.sheetName,
        sheetOrdinal: input.sheetOrdinal,
        visibleViewport: input.visibleViewport,
      }),
    }
  }

  private syncWorkbookDeltaConnection(input: GridRenderTileConnectionRuntimeInput): void {
    const identity: WorkbookDeltaConnectionIdentity = {
      dprBucket: input.dprBucket,
      renderTileSource: input.renderTileSource,
      sheetId: input.sheetId,
      sheetOrdinal: input.sheetOrdinal,
    }
    if (this.workbookDeltaConnection && sameWorkbookDeltaConnectionIdentity(this.workbookDeltaConnection.identity, identity)) {
      return
    }
    this.workbookDeltaConnection?.unsubscribe?.()
    this.workbookDeltaConnection = {
      identity,
      unsubscribe: this.connectWorkbookDeltaDamage({
        dprBucket: input.dprBucket,
        gridRuntimeHost: input.gridRuntimeHost,
        renderTileSource: input.renderTileSource,
        sheetId: input.sheetId,
        sheetOrdinal: input.sheetOrdinal,
      }),
    }
  }

  private syncLocalInvalidationConnection(input: GridRenderTileConnectionRuntimeInput): void {
    const identity: LocalInvalidationConnectionIdentity = {
      engine: input.engine,
      needsLocalCellInvalidation: input.needsLocalCellInvalidation,
      sheetName: input.sheetName,
      visibleAddresses: input.visibleAddresses,
    }
    if (this.localInvalidationConnection && sameLocalInvalidationConnectionIdentity(this.localInvalidationConnection.identity, identity)) {
      return
    }
    this.localInvalidationConnection?.unsubscribe?.()
    this.localInvalidationConnection = {
      identity,
      unsubscribe: this.connectLocalCellInvalidation({
        engine: input.engine,
        needsLocalCellInvalidation: input.needsLocalCellInvalidation,
        sheetName: input.sheetName,
        visibleAddresses: input.visibleAddresses,
      }),
    }
  }

  private buildFixedRenderTileDataPanes(
    input: GridRenderTilePaneRuntimeInput,
    resolution: GridRenderTileResolution,
  ): FixedRenderTileDataPanesResolution | null {
    const panes = buildFixedRenderTilePaneStates({
      freezeCols: input.freezeCols,
      freezeRows: input.freezeRows,
      frozenColumnWidth: input.frozenColumnWidth,
      frozenRowHeight: input.frozenRowHeight,
      gridMetrics: input.gridMetrics,
      hostHeight: input.hostClientHeight,
      hostWidth: input.hostClientWidth,
      residentViewport: input.residentViewport,
      sortedColumnWidthOverrides: input.sortedColumnWidthOverrides,
      sortedRowHeightOverrides: input.sortedRowHeightOverrides,
      tiles: resolution.tiles,
      visibleViewport: input.visibleViewport,
    })
    return panes.length > 0 ? { panes, source: resolution.source } : null
  }

  private resolveTileReadiness(
    input: GridRenderTilePaneRuntimeInput,
    tiles: readonly GridRenderTile[],
    preloadTiles: readonly GridRenderTile[] = [],
  ): GridTileReadinessSnapshotV3 {
    if (input.sheetId === undefined) {
      return EMPTY_TILE_PANE_RUNTIME_STATE.tileReadiness
    }
    const sheetId = input.sheetId
    for (const tile of tiles) {
      this.upsertHostTile(input, tile)
    }
    for (const tile of preloadTiles) {
      this.upsertHostTile(input, tile)
    }
    const interest = this.buildViewportTileInterest({
      ...input,
      sheetId,
    })
    return input.gridRuntimeHost.tiles.reconcileInterest(interest)
  }

  private buildViewportTileInterest(input: GridRenderTileInterestRuntimeInput & { readonly sheetId: number }): GridTileInterestBatchV3 {
    const sheetOrdinal = resolveGridRenderTileInputSheetOrdinal(input)
    const snapshot = input.gridRuntimeHost.snapshot()
    const visibleTileKeys = resolveRenderTileInterestTileKeys(input)
    return input.gridRuntimeHost.tiles.buildInterest({
      axisSeqX: snapshot.axisSeqX,
      axisSeqY: snapshot.axisSeqY,
      cameraSeq: snapshot.camera.seq,
      freezeSeq: snapshot.freezeSeq,
      pinnedTileKeys: [],
      reason: 'scroll',
      sheetId: input.sheetId,
      sheetOrdinal,
      visibleTileKeys,
      warmTileKeys: this.resolveWarmTileKeys(input, visibleTileKeys),
    })
  }

  private resolveWarmTileKeys(
    input: GridRenderTileInterestRuntimeInput & { readonly sheetId: number },
    visibleTileKeys = resolveRenderTileInterestTileKeys(input),
  ): readonly TileKey53[] {
    return resolveWarmRenderTileKeys(input, visibleTileKeys)
  }

  private resolvePreloadPanes(
    input: GridRenderTilePaneRuntimeInput,
    visibleTiles: readonly GridRenderTile[],
  ): GridRenderTilePreloadResolution {
    if (!input.renderTileSource || input.sheetId === undefined) {
      return { panes: EMPTY_TILE_PANE_RUNTIME_STATE.preloadDataPanes, tiles: [] }
    }
    const sheetOrdinal = resolveGridRenderTileInputSheetOrdinal(input)
    const visibleTileIds = new Set(visibleTiles.map((tile) => tile.tileId))
    const tiles: GridRenderTile[] = []
    for (const tileKey of this.resolveWarmTileKeys({
      dprBucket: input.dprBucket,
      gridRuntimeHost: input.gridRuntimeHost,
      renderTileViewport: input.renderTileViewport,
      visibleViewport: input.visibleViewport,
      sheetId: input.sheetId,
      sheetOrdinal,
    })) {
      if (visibleTileIds.has(tileKey)) {
        continue
      }
      const tile = input.renderTileSource.peekRenderTile(tileKey)
      if (!tile || !matchesRenderTileSheetIdentity(tile.coord, { sheetId: input.sheetId, sheetOrdinal })) {
        continue
      }
      tiles.push(tile)
    }
    if (tiles.length === 0) {
      return { panes: EMPTY_TILE_PANE_RUNTIME_STATE.preloadDataPanes, tiles }
    }
    return {
      panes: this.buildPreloadPanes(input, tiles),
      tiles,
    }
  }

  private buildPreloadPanes(
    input: GridRenderTilePaneRuntimeInput,
    tiles: readonly GridRenderTile[],
  ): readonly WorkbookRenderTilePaneState[] {
    const preloadViewport = resolveRenderTileViewportUnion(tiles)
    if (!preloadViewport) {
      return EMPTY_TILE_PANE_RUNTIME_STATE.preloadDataPanes
    }
    return buildFixedRenderTilePaneStates({
      freezeCols: input.freezeCols,
      freezeRows: input.freezeRows,
      frozenColumnWidth: input.frozenColumnWidth,
      frozenRowHeight: input.frozenRowHeight,
      gridMetrics: input.gridMetrics,
      hostHeight: input.hostClientHeight,
      hostWidth: input.hostClientWidth,
      residentViewport: preloadViewport,
      sortedColumnWidthOverrides: input.sortedColumnWidthOverrides,
      sortedRowHeightOverrides: input.sortedRowHeightOverrides,
      tiles,
      visibleViewport: input.visibleViewport,
    })
  }

  private upsertHostTile(input: GridRenderTilePaneRuntimeInput, tile: GridRenderTile): void {
    upsertRenderTileIntoHost(input.gridRuntimeHost, tile)
  }

  private applyRenderTileSceneChange(input: GridRenderTileDeltaRuntimeInput, change: GridRenderTileSceneChange): void {
    change.invalidatedTileIds.forEach((tileId) => {
      input.gridRuntimeHost.tiles.deleteTile(tileId)
    })
    const renderTileSource = input.renderTileSource
    if (!renderTileSource) {
      return
    }
    change.changedTileIds.forEach((tileId) => {
      const tile = renderTileSource.peekRenderTile(tileId)
      if (
        !tile ||
        !matchesRenderTileSheetIdentity(tile.coord, {
          sheetId: input.sheetId,
          sheetOrdinal: resolveGridRenderTileInputSheetOrdinal(input),
        })
      ) {
        return
      }
      upsertRenderTileIntoHost(input.gridRuntimeHost, tile)
    })
  }

  private resolveTiles(input: GridRenderTilePaneRuntimeInput): GridRenderTileResolution | null {
    if (input.forceLocalTiles) {
      return this.buildLocalTiles(input, { mergeCleanRemoteTiles: true })
    }

    if (input.renderTileSource && input.sheetId !== undefined) {
      return this.buildHybridLocalDirtyTiles(input, input.renderTileSource, input.sheetId, {
        localizeDirtyVisibleTiles: false,
      })
    }

    return this.buildLocalTiles(input)
  }

  private buildLocalTiles(
    input: GridRenderTilePaneRuntimeInput,
    options: { readonly mergeCleanRemoteTiles?: boolean } = {},
  ): GridRenderTileResolution {
    if (options.mergeCleanRemoteTiles && input.renderTileSource && input.sheetId !== undefined) {
      const hybrid = this.buildHybridLocalDirtyTiles(input, input.renderTileSource, input.sheetId)
      if (hybrid) {
        return hybrid
      }
    }
    return {
      source: 'local',
      tiles: buildLocalFixedRenderTiles({
        cameraSeq: input.gridRuntimeHost.snapshot().camera.seq,
        columnWidths: input.columnWidths,
        dirtySpansForTile: (tileId) => input.gridRuntimeHost.tiles.dirtyTiles.getSpans(tileId),
        dprBucket: input.dprBucket,
        editingCell: input.editingCell ?? null,
        engine: input.engine,
        freezeSeq: input.gridRuntimeHost.snapshot().freezeSeq,
        generation: input.sceneRevision,
        gridMetrics: input.gridMetrics,
        rowHeights: input.rowHeights,
        selectedCell: input.selectedCell,
        selectedCellSnapshot: input.selectedCellSnapshot,
        sheetId: input.sheetId ?? 0,
        sheetOrdinal: resolveGridRenderTileInputSheetOrdinal(input),
        sheetName: input.sheetName,
        sortedColumnWidthOverrides: input.sortedColumnWidthOverrides,
        sortedRowHeightOverrides: input.sortedRowHeightOverrides,
        viewport: input.renderTileViewport,
      }),
    }
  }

  private buildHybridLocalDirtyTiles(
    input: GridRenderTilePaneRuntimeInput,
    renderTileSource: GridRenderTileSource,
    sheetId: number,
    options: { readonly localizeDirtyVisibleTiles?: boolean | undefined } = {},
  ): GridRenderTileResolution | null {
    const sheetOrdinal = resolveGridRenderTileInputSheetOrdinal(input)
    const tileKeys = input.gridRuntimeHost.viewportTileKeys({
      dprBucket: input.dprBucket,
      sheetOrdinal,
      viewport: input.renderTileViewport,
    })
    if (tileKeys.length === 0) {
      return { source: 'local', tiles: [] }
    }

    const remoteTiles = new Map<number, GridRenderTile>()
    const dirtyBaseTiles = new Map<number, GridRenderTile>()
    const dirtyTileKeys: number[] = []
    const visibleTileKeys = new Set(
      input.gridRuntimeHost.viewportTileKeys({
        dprBucket: input.dprBucket,
        sheetOrdinal,
        viewport: input.visibleViewport,
      }),
    )
    const selectedCellTileKey = input.selectedCell
      ? input.gridRuntimeHost.viewportTileKeys({
          dprBucket: input.dprBucket,
          sheetOrdinal,
          viewport: {
            colEnd: input.selectedCell[0],
            colStart: input.selectedCell[0],
            rowEnd: input.selectedCell[1],
            rowStart: input.selectedCell[1],
          },
        })[0]
      : undefined
    const editingCellTileKey = input.editingCell
      ? input.gridRuntimeHost.viewportTileKeys({
          dprBucket: input.dprBucket,
          sheetOrdinal,
          viewport: {
            colEnd: input.editingCell[0],
            colStart: input.editingCell[0],
            rowEnd: input.editingCell[1],
            rowStart: input.editingCell[1],
          },
        })[0]
      : undefined
    for (const tileKey of tileKeys) {
      const sourceTile = renderTileSource.peekRenderTile(tileKey)
      const tile =
        sourceTile && sourceTile.coord.sheetId === sheetId && sourceTile.coord.sheetOrdinal === sheetOrdinal
          ? sourceTile
          : this.resolveResidentRenderTile(input, tileKey, sheetId, sheetOrdinal)
      const isDirty = visibleTileKeys.has(tileKey) && input.gridRuntimeHost.tiles.dirtyTiles.getUnconsumedMask(tileKey) !== 0
      const isMissingResidentTile = !tile
      const isMissingGridPayload = tile !== null && !hasCompleteRenderTileGrid(tile)
      const shouldLocalizeDirty = (options.localizeDirtyVisibleTiles ?? true) && isDirty
      const shouldLocalizeSelectedCellText =
        selectedCellTileKey === tileKey && tileSelectedTextNeedsLocalRefresh(tile, input.selectedCell, input.selectedCellSnapshot)
      const shouldLocalizeVisibleText = visibleTileKeys.has(tileKey) && this.visibleTextRefreshCache.needsLocalRefresh(tileKey, tile, input)
      const shouldLocalizeEditingCellText = editingCellTileKey === tileKey
      if (
        shouldLocalizeDirty ||
        isMissingResidentTile ||
        isMissingGridPayload ||
        shouldLocalizeSelectedCellText ||
        shouldLocalizeVisibleText ||
        shouldLocalizeEditingCellText
      ) {
        if (
          (shouldLocalizeDirty || shouldLocalizeSelectedCellText || shouldLocalizeVisibleText || shouldLocalizeEditingCellText) &&
          tile &&
          hasCompleteRenderTileGrid(tile)
        ) {
          dirtyBaseTiles.set(tileKey, tile)
        }
        dirtyTileKeys.push(tileKey)
        continue
      }
      if (tile) {
        remoteTiles.set(tileKey, tile)
      }
    }

    if (dirtyTileKeys.length === 0 && remoteTiles.size === tileKeys.length) {
      return { source: 'remote', tiles: tileKeys.map((tileKey) => remoteTiles.get(tileKey)!) }
    }

    const localTiles = new Map(
      buildLocalFixedRenderTiles({
        cameraSeq: input.gridRuntimeHost.snapshot().camera.seq,
        columnWidths: input.columnWidths,
        dirtySpansForTile: (tileId) => input.gridRuntimeHost.tiles.dirtyTiles.getSpans(tileId),
        dprBucket: input.dprBucket,
        editingCell: input.editingCell ?? null,
        engine: input.engine,
        generation: input.sceneRevision,
        gridMetrics: input.gridMetrics,
        rowHeights: input.rowHeights,
        selectedCell: input.selectedCell,
        selectedCellSnapshot: input.selectedCellSnapshot,
        sheetId,
        sheetOrdinal,
        sheetName: input.sheetName,
        sortedColumnWidthOverrides: input.sortedColumnWidthOverrides,
        sortedRowHeightOverrides: input.sortedRowHeightOverrides,
        tileKeys: dirtyTileKeys,
        reuseStaticGridRectsByTileId: dirtyBaseTiles,
        viewport: input.renderTileViewport,
      }).map((tile) => [tile.tileId, tile] as const),
    )

    return {
      source: 'local',
      tiles: tileKeys.flatMap((tileKey) => {
        const localTile = localTiles.get(tileKey)
        if (localTile) {
          return [localTile]
        }
        const remoteTile = remoteTiles.get(tileKey)
        return remoteTile ? [remoteTile] : []
      }),
    }
  }

  private resolveResidentRenderTile(
    input: GridRenderTilePaneRuntimeInput,
    tileKey: number,
    sheetId: number,
    sheetOrdinal: number,
  ): GridRenderTile | null {
    const packet = input.gridRuntimeHost.tiles.residency.getExact(tileKey)?.packet
    if (!isResidentGridRenderTile(packet)) {
      return null
    }
    return packet.coord.sheetId === sheetId && packet.coord.sheetOrdinal === sheetOrdinal ? packet : null
  }

  private emitBridgeState(): void {
    this.bridgeListeners.forEach((listener) => {
      listener()
    })
  }
}

export function getGridRenderTilePaneRuntime(current: unknown): GridRenderTilePaneRuntime {
  return current instanceof GridRenderTilePaneRuntime ? current : new GridRenderTilePaneRuntime()
}
