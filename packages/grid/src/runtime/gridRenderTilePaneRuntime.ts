import type { Viewport } from '@bilig/protocol'
import type { GridEngineLike } from '../grid-engine.js'
import type { GridMetrics } from '../gridMetrics.js'
import { buildLocalFixedRenderTiles } from '../renderer-v3/local-render-tile-materializer.js'
import { buildFixedRenderTilePaneStates } from '../renderer-v3/render-tile-pane-builder.js'
import type { GridRenderTile, GridRenderTileSceneChange, GridRenderTileSource } from '../renderer-v3/render-tile-source.js'
import type { WorkbookRenderTilePaneState } from '../renderer-v3/render-tile-pane-state.js'
import type { GridRuntimeHost } from './gridRuntimeHost.js'

type SortedAxisOverrides = readonly (readonly [number, number])[]

export interface GridRenderTilePaneRuntimeState {
  readonly preloadDataPanes: readonly WorkbookRenderTilePaneState[]
  readonly renderTilePanes: readonly WorkbookRenderTilePaneState[]
  readonly residentBodyPane: WorkbookRenderTilePaneState | null
  readonly residentDataPanes: readonly WorkbookRenderTilePaneState[]
}

export interface GridRenderTilePaneRuntimeInput {
  readonly columnWidths: Readonly<Record<number, number>>
  readonly dprBucket: number
  readonly engine: GridEngineLike
  readonly freezeCols: number
  readonly freezeRows: number
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

const EMPTY_TILE_PANE_RUNTIME_STATE: GridRenderTilePaneRuntimeState = Object.freeze({
  preloadDataPanes: [],
  renderTilePanes: [],
  residentBodyPane: null,
  residentDataPanes: [],
})

export class GridRenderTilePaneRuntime {
  private retainedFixedRenderTileDataPanes: {
    readonly sheetId: number
    readonly panes: readonly WorkbookRenderTilePaneState[]
  } | null = null

  resolve(input: GridRenderTilePaneRuntimeInput): GridRenderTilePaneRuntimeState {
    if (!input.hostReady) {
      return EMPTY_TILE_PANE_RUNTIME_STATE
    }
    const fixedRenderTileDataPanes = this.resolveFixedRenderTileDataPanes(input)
    if (input.sheetId !== undefined && fixedRenderTileDataPanes) {
      this.retainedFixedRenderTileDataPanes = {
        panes: fixedRenderTileDataPanes,
        sheetId: input.sheetId,
      }
    }

    const shouldUseRemoteRenderTileSource = input.renderTileSource !== undefined && input.sheetId !== undefined
    const retainedFixedRenderTileDataPanes =
      fixedRenderTileDataPanes ??
      (shouldUseRemoteRenderTileSource && input.sheetId !== undefined && this.retainedFixedRenderTileDataPanes?.sheetId === input.sheetId
        ? this.retainedFixedRenderTileDataPanes.panes
        : null)
    const residentDataPanes = retainedFixedRenderTileDataPanes ?? []
    return {
      preloadDataPanes: EMPTY_TILE_PANE_RUNTIME_STATE.preloadDataPanes,
      renderTilePanes: residentDataPanes,
      residentBodyPane: residentDataPanes.find((pane) => pane.paneId === 'body') ?? null,
      residentDataPanes,
    }
  }

  clearRetainedPanes(): void {
    this.retainedFixedRenderTileDataPanes = null
  }

  connectRenderTileDeltas(
    input: GridRenderTileDeltaRuntimeInput,
    listener: (change: GridRenderTileSceneChange) => void,
  ): (() => void) | undefined {
    if (!input.renderTileSource || input.sheetId === undefined) {
      return undefined
    }
    const tileInterest = input.gridRuntimeHost.buildViewportTileInterest({
      dprBucket: input.dprBucket,
      reason: 'scroll',
      sheetId: input.sheetId,
      sheetOrdinal: input.sheetId,
      viewport: input.renderTileViewport,
    })
    return input.renderTileSource.subscribeRenderTileDeltas(
      {
        ...input.renderTileViewport,
        cameraSeq: tileInterest.cameraSeq,
        dprBucket: input.dprBucket,
        initialDelta: 'full',
        sheetId: input.sheetId,
        sheetName: input.sheetName,
      },
      listener,
    )
  }

  private resolveFixedRenderTileDataPanes(input: GridRenderTilePaneRuntimeInput): readonly WorkbookRenderTilePaneState[] | null {
    if (!input.hostReady) {
      return null
    }
    const tiles = this.resolveTiles(input)
    if (!tiles) {
      return null
    }
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
      tiles,
      visibleViewport: input.visibleViewport,
    })
    return panes.length > 0 ? panes : null
  }

  private resolveTiles(input: GridRenderTilePaneRuntimeInput): readonly GridRenderTile[] | null {
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
      return tiles
    }

    return this.buildLocalTiles(input)
  }

  private buildLocalTiles(input: GridRenderTilePaneRuntimeInput): readonly GridRenderTile[] {
    return buildLocalFixedRenderTiles({
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
    })
  }
}

export function getGridRenderTilePaneRuntime(current: unknown): GridRenderTilePaneRuntime {
  return current instanceof GridRenderTilePaneRuntime ? current : new GridRenderTilePaneRuntime()
}
