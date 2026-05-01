import type { Viewport } from '@bilig/protocol'
import { noteRendererTileReadiness } from '../grid-render-counters.js'
import type { GridEngineLike } from '../grid-engine.js'
import type { GridMetrics } from '../gridMetrics.js'
import { buildLocalFixedRenderTiles } from '../renderer-v3/local-render-tile-materializer.js'
import { buildFixedRenderTilePaneStates } from '../renderer-v3/render-tile-pane-builder.js'
import type { GridRenderTile, GridRenderTileSceneChange, GridRenderTileSource } from '../renderer-v3/render-tile-source.js'
import type { WorkbookRenderTilePaneState } from '../renderer-v3/render-tile-pane-state.js'
import type { WorkbookDeltaBatchLikeV3 } from '../renderer-v3/tile-damage-index.js'
import { MAX_TILE_COLUMN_INDEX, MAX_TILE_ROW_INDEX, packTileKey53, unpackTileKey53, type TileKey53 } from '../renderer-v3/tile-key.js'
import type { GridTileInterestBatchV3, GridTileReadinessSnapshotV3 } from './gridTileCoordinator.js'
import type { GridRuntimeHost } from './gridRuntimeHost.js'

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
  readonly sheetId?: number | undefined
  readonly sheetOrdinal?: number | undefined
  readonly sheetName: string
  readonly sortedColumnWidthOverrides: SortedAxisOverrides
  readonly sortedRowHeightOverrides: SortedAxisOverrides
  readonly visibleViewport: Viewport
}

export interface GridRenderTileDeltaRuntimeInput {
  readonly dprBucket: number
  readonly gridRuntimeHost: GridRuntimeHost
  readonly renderTileSource?: GridRenderTileSource | undefined
  readonly renderTileViewport: Viewport
  readonly sheetId?: number | undefined
  readonly sheetOrdinal?: number | undefined
  readonly sheetName: string
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
  readonly gridRuntimeHost: GridRuntimeHost
  readonly needsLocalCellInvalidation: boolean
  readonly renderTileSource?: GridRenderTileSource | undefined
  readonly renderTileViewport: Viewport
  readonly sheetId?: number | undefined
  readonly sheetOrdinal?: number | undefined
  readonly sheetName: string
  readonly visibleAddresses: readonly string[]
}

interface GridRenderTileInterestRuntimeInput {
  readonly dprBucket: number
  readonly gridRuntimeHost: GridRuntimeHost
  readonly renderTileViewport: Viewport
  readonly sheetId?: number | undefined
  readonly sheetOrdinal?: number | undefined
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
    visibleDirtyTileKeys: [],
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

interface RuntimeConnection<Identity> {
  readonly identity: Identity
  readonly unsubscribe: (() => void) | undefined
}

interface RenderTileDeltaConnectionIdentity {
  readonly dprBucket: number
  readonly renderTileSource: GridRenderTileSource | undefined
  readonly sheetId: number | undefined
  readonly sheetOrdinal: number | undefined
  readonly sheetName: string
  readonly viewport: Viewport
}

interface WorkbookDeltaConnectionIdentity {
  readonly dprBucket: number
  readonly renderTileSource: GridRenderTileSource | undefined
  readonly sheetId: number | undefined
  readonly sheetOrdinal: number | undefined
}

interface LocalInvalidationConnectionIdentity {
  readonly engine: GridEngineLike
  readonly needsLocalCellInvalidation: boolean
  readonly sheetName: string
  readonly visibleAddresses: readonly string[]
}

export class GridRenderTilePaneRuntime {
  private retainedFixedRenderTileDataPanes: {
    readonly sheetId: number
    readonly panes: readonly WorkbookRenderTilePaneState[]
  } | null = null
  private bridgeState = INITIAL_RENDER_TILE_PANE_BRIDGE_STATE
  private readonly bridgeListeners = new Set<() => void>()
  private readonly lastWorkbookDeltaSeqBySheetOrdinal = new Map<number, number>()
  private renderTileDeltaConnection: RuntimeConnection<RenderTileDeltaConnectionIdentity> | null = null
  private workbookDeltaConnection: RuntimeConnection<WorkbookDeltaConnectionIdentity> | null = null
  private localInvalidationConnection: RuntimeConnection<LocalInvalidationConnectionIdentity> | null = null

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
        panes: fixedRenderTileDataPanes.panes,
        sheetId: input.sheetId,
      }
    }

    const shouldUseRemoteRenderTileSource = input.renderTileSource !== undefined && input.sheetId !== undefined
    const retainedFixedRenderTileDataPanes =
      fixedRenderTileDataPanes?.panes ??
      (shouldUseRemoteRenderTileSource && input.sheetId !== undefined && this.retainedFixedRenderTileDataPanes?.sheetId === input.sheetId
        ? this.retainedFixedRenderTileDataPanes.panes
        : null)
    const residentDataPanes = retainedFixedRenderTileDataPanes ?? []
    return {
      needsLocalCellInvalidation: !shouldUseRemoteRenderTileSource,
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

  noteWorkbookDeltaDamage(): GridRenderTilePaneBridgeState {
    const previous = this.bridgeState
    this.bridgeState = {
      forceLocalTiles: previous.forceLocalTiles,
      localFallbackRevision: previous.localFallbackRevision,
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
      if (batch.sheetId !== input.sheetId && batch.sheetOrdinal !== sheetOrdinal) {
        return
      }
      if (!this.applyWorkbookDeltaDamage(input, batch)) {
        return
      }
      this.noteWorkbookDeltaDamage()
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
      const lastSeq = this.lastWorkbookDeltaSeqBySheetOrdinal.get(batch.sheetOrdinal) ?? -1
      if (batch.seq <= lastSeq) {
        return false
      }
      this.lastWorkbookDeltaSeqBySheetOrdinal.set(batch.sheetOrdinal, batch.seq)
    }
    input.gridRuntimeHost.tiles.applyWorkbookDelta(batch, { dprBucket: input.dprBucket })
    return true
  }

  private syncRenderTileDeltaConnection(input: GridRenderTileConnectionRuntimeInput): void {
    const identity: RenderTileDeltaConnectionIdentity = {
      dprBucket: input.dprBucket,
      renderTileSource: input.renderTileSource,
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
        renderTileSource: input.renderTileSource,
        renderTileViewport: input.renderTileViewport,
        sheetId: input.sheetId,
        sheetName: input.sheetName,
        sheetOrdinal: input.sheetOrdinal,
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
    return input.gridRuntimeHost.buildViewportTileInterest({
      dprBucket: input.dprBucket,
      reason: 'scroll',
      sheetId: input.sheetId,
      sheetOrdinal,
      viewport: input.renderTileViewport,
      warmTileKeys: this.resolveWarmTileKeys(input),
    })
  }

  private resolveWarmTileKeys(input: GridRenderTileInterestRuntimeInput & { readonly sheetId: number }): readonly TileKey53[] {
    const sheetOrdinal = resolveGridRenderTileInputSheetOrdinal(input)
    const visibleTileKeys = input.gridRuntimeHost.viewportTileKeys({
      dprBucket: input.dprBucket,
      sheetOrdinal,
      viewport: input.renderTileViewport,
    })
    if (visibleTileKeys.length === 0) {
      return []
    }
    const visibleSet = new Set(visibleTileKeys)
    let minRowTile = Number.POSITIVE_INFINITY
    let maxRowTile = Number.NEGATIVE_INFINITY
    let minColTile = Number.POSITIVE_INFINITY
    let maxColTile = Number.NEGATIVE_INFINITY
    for (const key of visibleTileKeys) {
      const fields = unpackTileKey53(key)
      minRowTile = Math.min(minRowTile, fields.rowTile)
      maxRowTile = Math.max(maxRowTile, fields.rowTile)
      minColTile = Math.min(minColTile, fields.colTile)
      maxColTile = Math.max(maxColTile, fields.colTile)
    }
    const warmTileKeys: number[] = []
    for (let rowTile = Math.max(0, minRowTile - 1); rowTile <= Math.min(MAX_TILE_ROW_INDEX, maxRowTile + 1); rowTile += 1) {
      for (let colTile = Math.max(0, minColTile - 1); colTile <= Math.min(MAX_TILE_COLUMN_INDEX, maxColTile + 1); colTile += 1) {
        const key = packTileKey53({
          colTile,
          dprBucket: input.dprBucket,
          rowTile,
          sheetOrdinal,
        })
        if (!visibleSet.has(key)) {
          warmTileKeys.push(key)
        }
      }
    }
    return warmTileKeys
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
      sheetId: input.sheetId,
      sheetOrdinal,
    })) {
      if (visibleTileIds.has(tileKey)) {
        continue
      }
      const tile = input.renderTileSource.peekRenderTile(tileKey)
      if (!tile || tile.coord.sheetId !== input.sheetId || tile.coord.sheetOrdinal !== sheetOrdinal) {
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
      if (!tile || tile.coord.sheetId !== input.sheetId || tile.coord.sheetOrdinal !== resolveGridRenderTileInputSheetOrdinal(input)) {
        return
      }
      upsertRenderTileIntoHost(input.gridRuntimeHost, tile)
    })
  }

  private resolveTiles(input: GridRenderTilePaneRuntimeInput): GridRenderTileResolution | null {
    if (input.forceLocalTiles) {
      return this.buildLocalTiles(input)
    }

    if (input.renderTileSource && input.sheetId !== undefined) {
      const tiles: GridRenderTile[] = []
      const sheetOrdinal = resolveGridRenderTileInputSheetOrdinal(input)
      const tileKeys = input.gridRuntimeHost.viewportTileKeys({
        dprBucket: input.dprBucket,
        sheetOrdinal,
        viewport: input.renderTileViewport,
      })
      for (const tileKey of tileKeys) {
        const tile = input.renderTileSource.peekRenderTile(tileKey)
        if (!tile || tile.coord.sheetId !== input.sheetId || tile.coord.sheetOrdinal !== sheetOrdinal) {
          return this.retainedFixedRenderTileDataPanes?.sheetId === input.sheetId ? null : this.buildLocalTiles(input)
        }
        tiles.push(tile)
      }
      return { source: 'remote', tiles }
    }

    return this.buildLocalTiles(input)
  }

  private buildLocalTiles(input: GridRenderTilePaneRuntimeInput): GridRenderTileResolution {
    return {
      source: 'local',
      tiles: buildLocalFixedRenderTiles({
        cameraSeq: input.gridRuntimeHost.snapshot().camera.seq,
        columnWidths: input.columnWidths,
        dprBucket: input.dprBucket,
        engine: input.engine,
        generation: input.sceneRevision,
        gridMetrics: input.gridMetrics,
        rowHeights: input.rowHeights,
        sheetId: input.sheetId ?? 0,
        sheetOrdinal: resolveGridRenderTileInputSheetOrdinal(input),
        sheetName: input.sheetName,
        sortedColumnWidthOverrides: input.sortedColumnWidthOverrides,
        sortedRowHeightOverrides: input.sortedRowHeightOverrides,
        viewport: input.renderTileViewport,
      }),
    }
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

function resolveGridRenderTileInputSheetOrdinal(input: {
  readonly sheetId?: number | undefined
  readonly sheetOrdinal?: number | undefined
}): number {
  return input.sheetOrdinal ?? input.sheetId ?? 0
}

function sameRenderTileDeltaConnectionIdentity(left: RenderTileDeltaConnectionIdentity, right: RenderTileDeltaConnectionIdentity): boolean {
  return (
    left.dprBucket === right.dprBucket &&
    left.renderTileSource === right.renderTileSource &&
    left.sheetId === right.sheetId &&
    left.sheetName === right.sheetName &&
    left.sheetOrdinal === right.sheetOrdinal &&
    sameViewportIdentity(left.viewport, right.viewport)
  )
}

function sameWorkbookDeltaConnectionIdentity(left: WorkbookDeltaConnectionIdentity, right: WorkbookDeltaConnectionIdentity): boolean {
  return (
    left.dprBucket === right.dprBucket &&
    left.renderTileSource === right.renderTileSource &&
    left.sheetId === right.sheetId &&
    left.sheetOrdinal === right.sheetOrdinal
  )
}

function sameLocalInvalidationConnectionIdentity(
  left: LocalInvalidationConnectionIdentity,
  right: LocalInvalidationConnectionIdentity,
): boolean {
  return (
    left.engine === right.engine &&
    left.needsLocalCellInvalidation === right.needsLocalCellInvalidation &&
    left.sheetName === right.sheetName &&
    sameStringListIdentity(left.visibleAddresses, right.visibleAddresses)
  )
}

function sameStringListIdentity(left: readonly string[], right: readonly string[]): boolean {
  if (left === right) {
    return true
  }
  if (left.length !== right.length) {
    return false
  }
  for (let index = 0; index < left.length; index += 1) {
    if (left[index] !== right[index]) {
      return false
    }
  }
  return true
}

function sameViewportIdentity(left: Viewport, right: Viewport): boolean {
  return (
    left.colEnd === right.colEnd && left.colStart === right.colStart && left.rowEnd === right.rowEnd && left.rowStart === right.rowStart
  )
}

function upsertRenderTileIntoHost(gridRuntimeHost: GridRuntimeHost, tile: GridRenderTile): void {
  gridRuntimeHost.tiles.upsertTile({
    axisSeqX: tile.version.axisX,
    axisSeqY: tile.version.axisY,
    byteSizeCpu: estimateTileCpuBytes(tile),
    byteSizeGpu: tile.rectInstances.byteLength + tile.textMetrics.byteLength,
    colTile: tile.coord.colTile,
    dprBucket: tile.coord.dprBucket,
    freezeSeq: tile.version.freeze,
    key: tile.tileId,
    packet: tile,
    rectSeq: tile.version.styles,
    rowTile: tile.coord.rowTile,
    sheetOrdinal: tile.coord.sheetOrdinal,
    state: 'ready',
    styleSeq: tile.version.styles,
    textSeq: tile.version.text,
    valueSeq: tile.version.values,
  })
}

function resolveRenderTileViewportUnion(tiles: readonly GridRenderTile[]): Viewport | null {
  const first = tiles[0]
  if (!first) {
    return null
  }
  let colStart = first.bounds.colStart
  let colEnd = first.bounds.colEnd
  let rowStart = first.bounds.rowStart
  let rowEnd = first.bounds.rowEnd
  for (let index = 1; index < tiles.length; index += 1) {
    const tile = tiles[index]!
    colStart = Math.min(colStart, tile.bounds.colStart)
    colEnd = Math.max(colEnd, tile.bounds.colEnd)
    rowStart = Math.min(rowStart, tile.bounds.rowStart)
    rowEnd = Math.max(rowEnd, tile.bounds.rowEnd)
  }
  return { colEnd, colStart, rowEnd, rowStart }
}

function estimateTileCpuBytes(tile: GridRenderTile): number {
  let textBytes = 0
  for (const run of tile.textRuns) {
    textBytes += run.text.length * 2
    textBytes += run.font.length * 2
    textBytes += run.color.length * 2
  }
  return tile.rectInstances.byteLength + tile.textMetrics.byteLength + textBytes
}
