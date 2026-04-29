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
  readonly sheetName: string
}

export interface GridRenderTileDamageRuntimeInput {
  readonly dprBucket: number
  readonly gridRuntimeHost: GridRuntimeHost
  readonly renderTileSource?: GridRenderTileSource | undefined
  readonly sheetId?: number | undefined
}

export interface GridRenderTileLocalInvalidationRuntimeInput {
  readonly engine: GridEngineLike
  readonly needsLocalCellInvalidation: boolean
  readonly sheetName: string
  readonly visibleAddresses: readonly string[]
}

interface GridRenderTileInterestRuntimeInput {
  readonly dprBucket: number
  readonly gridRuntimeHost: GridRuntimeHost
  readonly renderTileViewport: Viewport
  readonly sheetId?: number | undefined
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

type TileResolutionSource = 'local' | 'remote'

interface FixedRenderTileDataPanesResolution {
  readonly panes: readonly WorkbookRenderTilePaneState[]
  readonly source: TileResolutionSource
}

interface GridRenderTileResolution {
  readonly tiles: readonly GridRenderTile[]
  readonly source: TileResolutionSource
}

export class GridRenderTilePaneRuntime {
  private retainedFixedRenderTileDataPanes: {
    readonly sheetId: number
    readonly panes: readonly WorkbookRenderTilePaneState[]
  } | null = null
  private readonly lastWorkbookDeltaSeqBySheetOrdinal = new Map<number, number>()

  resolve(input: GridRenderTilePaneRuntimeInput): GridRenderTilePaneRuntimeState {
    if (!input.hostReady) {
      return EMPTY_TILE_PANE_RUNTIME_STATE
    }
    const resolution = this.resolveTiles(input)
    const tileReadiness = this.resolveTileReadiness(input, resolution?.tiles ?? [])
    const fixedRenderTileDataPanes = resolution ? this.buildFixedRenderTileDataPanes(input, resolution) : null
    if (input.sheetId !== undefined && fixedRenderTileDataPanes) {
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
      needsLocalCellInvalidation: shouldUseRemoteRenderTileSource,
      preloadDataPanes: EMPTY_TILE_PANE_RUNTIME_STATE.preloadDataPanes,
      renderTilePanes: residentDataPanes,
      residentBodyPane: residentDataPanes.find((pane) => pane.paneId === 'body') ?? null,
      residentDataPanes,
      tileReadiness,
    }
  }

  clearRetainedPanes(): void {
    this.retainedFixedRenderTileDataPanes = null
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

  connectLocalCellInvalidation(input: GridRenderTileLocalInvalidationRuntimeInput, listener: () => void): (() => void) | undefined {
    if (!input.needsLocalCellInvalidation || input.visibleAddresses.length === 0) {
      return undefined
    }
    return input.engine.subscribeCells(input.sheetName, input.visibleAddresses, () => {
      this.clearRetainedPanes()
      listener()
    })
  }

  connectWorkbookDeltaDamage(
    input: GridRenderTileDamageRuntimeInput,
    listener: (batch: WorkbookDeltaBatchLikeV3) => void,
  ): (() => void) | undefined {
    const renderTileSource = input.renderTileSource
    if (!renderTileSource?.subscribeWorkbookDeltas || input.sheetId === undefined) {
      return undefined
    }
    return renderTileSource.subscribeWorkbookDeltas((batch) => {
      if (batch.sheetId !== input.sheetId && batch.sheetOrdinal !== input.sheetId) {
        return
      }
      if (!this.applyWorkbookDeltaDamage(input, batch)) {
        return
      }
      listener(batch)
    })
  }

  connectRenderTileDeltas(
    input: GridRenderTileDeltaRuntimeInput,
    listener: (change: GridRenderTileSceneChange) => void,
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
        listener(change)
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

  private resolveTileReadiness(input: GridRenderTilePaneRuntimeInput, tiles: readonly GridRenderTile[]): GridTileReadinessSnapshotV3 {
    if (input.sheetId === undefined) {
      return EMPTY_TILE_PANE_RUNTIME_STATE.tileReadiness
    }
    const sheetId = input.sheetId
    for (const tile of tiles) {
      this.upsertHostTile(input, tile)
    }
    const interest = this.buildViewportTileInterest({
      ...input,
      sheetId,
    })
    return input.gridRuntimeHost.tiles.reconcileInterest(interest)
  }

  private buildViewportTileInterest(input: GridRenderTileInterestRuntimeInput & { readonly sheetId: number }): GridTileInterestBatchV3 {
    return input.gridRuntimeHost.buildViewportTileInterest({
      dprBucket: input.dprBucket,
      reason: 'scroll',
      sheetId: input.sheetId,
      sheetOrdinal: input.sheetId,
      viewport: input.renderTileViewport,
      warmTileKeys: this.resolveWarmTileKeys(input),
    })
  }

  private resolveWarmTileKeys(input: GridRenderTileInterestRuntimeInput & { readonly sheetId: number }): readonly TileKey53[] {
    const visibleTileKeys = input.gridRuntimeHost.viewportTileKeys({
      dprBucket: input.dprBucket,
      sheetOrdinal: input.sheetId,
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
          sheetOrdinal: input.sheetId,
        })
        if (!visibleSet.has(key)) {
          warmTileKeys.push(key)
        }
      }
    }
    return warmTileKeys
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
      if (!tile || tile.coord.sheetId !== input.sheetId) {
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
      const tileKeys = input.gridRuntimeHost.viewportTileKeys({
        dprBucket: input.dprBucket,
        sheetOrdinal: input.sheetId,
        viewport: input.renderTileViewport,
      })
      for (const tileKey of tileKeys) {
        const tile = input.renderTileSource.peekRenderTile(tileKey)
        if (!tile || tile.coord.sheetId !== input.sheetId) {
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
        sheetName: input.sheetName,
        sortedColumnWidthOverrides: input.sortedColumnWidthOverrides,
        sortedRowHeightOverrides: input.sortedRowHeightOverrides,
        viewport: input.renderTileViewport,
      }),
    }
  }
}

export function getGridRenderTilePaneRuntime(current: unknown): GridRenderTilePaneRuntime {
  return current instanceof GridRenderTilePaneRuntime ? current : new GridRenderTilePaneRuntime()
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
    sheetOrdinal: tile.coord.sheetId,
    state: 'ready',
    styleSeq: tile.version.styles,
    textSeq: tile.version.text,
    valueSeq: tile.version.values,
  })
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
