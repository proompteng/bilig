import type { GridMetrics } from '../gridMetrics.js'
import type { OverlayBatchV3, OverlayInstanceV3 } from '../renderer-v3/overlay-layer.js'
import { tileKeysForViewport, type TileKey53 } from '../renderer-v3/tile-key.js'
import type { Item } from '../gridTypes.js'
import type { Viewport } from '@bilig/protocol'
import { viewportFromVisibleRegion } from '../useGridCameraState.js'
import { GridAxisRuntime } from './gridAxisRuntime.js'
import type { AxisEntryOverride } from '../gridAxisIndex.js'
import { GridCameraRuntime, type GridCameraRuntimeSnapshot } from './gridCameraRuntime.js'
import { GridOverlayRuntime } from './gridOverlayRuntime.js'
import {
  GridRenderTilePaneRuntime,
  type GridRenderTilePaneBridgeState,
  type GridRenderTileDamageRuntimeInput,
  type GridRenderTileDeltaRuntimeInput,
  type GridRenderTileLocalInvalidationRuntimeInput,
  type GridRenderTilePaneRuntimeInput,
  type GridRenderTilePaneRuntimeState,
} from './gridRenderTilePaneRuntime.js'
import { GridTileCoordinator, type GridTileInterestBatchV3, type GridTileInterestReasonV3 } from './gridTileCoordinator.js'
import { GridHeaderPaneRuntime, type GridHeaderPaneRuntimeInput } from './gridHeaderPaneRuntime.js'
import {
  GridViewportResidencyRuntime,
  type GridViewportResidencyInvalidationInput,
  type GridViewportResidencyRuntimeInput,
  type GridViewportResidencyState,
} from './gridViewportResidencyRuntime.js'

export interface GridRuntimeHostSnapshot {
  readonly camera: GridCameraRuntimeSnapshot
  readonly axisSeqX: number
  readonly axisSeqY: number
  readonly freezeSeq: number
}

export class GridRuntimeHost {
  readonly columns: GridAxisRuntime
  readonly rows: GridAxisRuntime
  readonly camera: GridCameraRuntime
  readonly overlays = new GridOverlayRuntime()
  readonly tiles = new GridTileCoordinator()
  readonly headers = new GridHeaderPaneRuntime()
  readonly renderTiles = new GridRenderTilePaneRuntime()
  readonly viewportResidency = new GridViewportResidencyRuntime()
  private freezeRows: number
  private freezeCols: number
  private freezeSeq = 1

  constructor(input: {
    readonly columnCount: number
    readonly rowCount: number
    readonly defaultColumnWidth: number
    readonly defaultRowHeight: number
    readonly gridMetrics: GridMetrics
    readonly viewportWidth: number
    readonly viewportHeight: number
    readonly freezeRows?: number | undefined
    readonly freezeCols?: number | undefined
  }) {
    this.freezeRows = Math.max(0, input.freezeRows ?? 0)
    this.freezeCols = Math.max(0, input.freezeCols ?? 0)
    this.columns = new GridAxisRuntime({
      axisLength: input.columnCount,
      defaultSize: input.defaultColumnWidth,
    })
    this.rows = new GridAxisRuntime({
      axisLength: input.rowCount,
      defaultSize: input.defaultRowHeight,
    })
    this.camera = new GridCameraRuntime({
      columns: this.columns,
      freezeCols: this.freezeCols,
      freezeRows: this.freezeRows,
      gridMetrics: input.gridMetrics,
      rows: this.rows,
      viewportHeight: input.viewportHeight,
      viewportWidth: input.viewportWidth,
    })
  }

  snapshot(): GridRuntimeHostSnapshot {
    return {
      axisSeqX: this.columns.snapshot().seq,
      axisSeqY: this.rows.snapshot().seq,
      camera: this.camera.snapshot(),
      freezeSeq: this.freezeSeq,
    }
  }

  updateAxes(input: {
    readonly columns?: readonly AxisEntryOverride[] | undefined
    readonly rows?: readonly AxisEntryOverride[] | undefined
    readonly columnSeq?: number | undefined
    readonly rowSeq?: number | undefined
  }): void {
    if (input.columns) {
      this.columns.update({ overrides: input.columns, seq: input.columnSeq })
    }
    if (input.rows) {
      this.rows.update({ overrides: input.rows, seq: input.rowSeq })
    }
  }

  updateCamera(input: {
    readonly gridMetrics: GridMetrics
    readonly viewportWidth: number
    readonly viewportHeight: number
    readonly scrollLeft: number
    readonly scrollTop: number
    readonly dpr: number
    readonly freezeRows?: number | undefined
    readonly freezeCols?: number | undefined
  }): GridCameraRuntimeSnapshot {
    this.updateFreeze(input.freezeRows, input.freezeCols)
    return this.camera.update({
      columns: this.columns,
      freezeCols: this.freezeCols,
      freezeRows: this.freezeRows,
      gridMetrics: input.gridMetrics,
      rows: this.rows,
      scrollLeft: input.scrollLeft,
      scrollTop: input.scrollTop,
      dpr: input.dpr,
      viewportHeight: input.viewportHeight,
      viewportWidth: input.viewportWidth,
    })
  }

  visibleTileKeys(input: { readonly sheetOrdinal: number; readonly dprBucket: number }): TileKey53[] {
    return tileKeysForViewport({
      dprBucket: input.dprBucket,
      sheetOrdinal: input.sheetOrdinal,
      viewport: viewportFromVisibleRegion(this.camera.snapshot().visibleRegion),
    })
  }

  viewportTileKeys(input: { readonly sheetOrdinal: number; readonly dprBucket: number; readonly viewport: Viewport }): TileKey53[] {
    return tileKeysForViewport({
      dprBucket: input.dprBucket,
      sheetOrdinal: input.sheetOrdinal,
      viewport: input.viewport,
    })
  }

  buildTileInterest(input: {
    readonly sheetId: number
    readonly sheetOrdinal: number
    readonly dprBucket: number
    readonly warmTileKeys?: Iterable<TileKey53> | undefined
    readonly pinnedTileKeys?: Iterable<TileKey53> | undefined
    readonly reason: GridTileInterestReasonV3
  }): GridTileInterestBatchV3 {
    const snapshot = this.snapshot()
    return this.tiles.buildInterest({
      axisSeqX: snapshot.axisSeqX,
      axisSeqY: snapshot.axisSeqY,
      cameraSeq: snapshot.camera.seq,
      freezeSeq: snapshot.freezeSeq,
      pinnedTileKeys: input.pinnedTileKeys,
      reason: input.reason,
      sheetId: input.sheetId,
      sheetOrdinal: input.sheetOrdinal,
      visibleTileKeys: this.visibleTileKeys({
        dprBucket: input.dprBucket,
        sheetOrdinal: input.sheetOrdinal,
      }),
      warmTileKeys: input.warmTileKeys,
    })
  }

  buildViewportTileInterest(input: {
    readonly sheetId: number
    readonly sheetOrdinal: number
    readonly dprBucket: number
    readonly viewport: Viewport
    readonly warmTileKeys?: Iterable<TileKey53> | undefined
    readonly pinnedTileKeys?: Iterable<TileKey53> | undefined
    readonly reason: GridTileInterestReasonV3
  }): GridTileInterestBatchV3 {
    const snapshot = this.snapshot()
    return this.tiles.buildInterest({
      axisSeqX: snapshot.axisSeqX,
      axisSeqY: snapshot.axisSeqY,
      cameraSeq: snapshot.camera.seq,
      freezeSeq: snapshot.freezeSeq,
      pinnedTileKeys: input.pinnedTileKeys,
      reason: input.reason,
      sheetId: input.sheetId,
      sheetOrdinal: input.sheetOrdinal,
      visibleTileKeys: tileKeysForViewport({
        dprBucket: input.dprBucket,
        sheetOrdinal: input.sheetOrdinal,
        viewport: input.viewport,
      }),
      warmTileKeys: input.warmTileKeys,
    })
  }

  resolveViewportResidency(input: GridViewportResidencyRuntimeInput): GridViewportResidencyState {
    return this.viewportResidency.resolve(input)
  }

  connectViewportResidencyInvalidation(
    input: GridViewportResidencyInvalidationInput,
    listener: Parameters<GridViewportResidencyRuntime['connectLocalSceneInvalidation']>[1],
  ): ReturnType<GridViewportResidencyRuntime['connectLocalSceneInvalidation']> {
    return this.viewportResidency.connectLocalSceneInvalidation(input, listener)
  }

  resolveHeaderPanes(input: GridHeaderPaneRuntimeInput): ReturnType<GridHeaderPaneRuntime['resolve']> {
    return this.headers.resolve(input)
  }

  resolveRenderTilePanes(input: Omit<GridRenderTilePaneRuntimeInput, 'gridRuntimeHost'>): GridRenderTilePaneRuntimeState {
    return this.renderTiles.resolve({
      ...input,
      gridRuntimeHost: this,
    })
  }

  snapshotRenderTileBridgeState(): GridRenderTilePaneBridgeState {
    return this.renderTiles.snapshotBridgeState()
  }

  noteRenderTileDelta(): GridRenderTilePaneBridgeState {
    return this.renderTiles.noteRenderTileDelta()
  }

  noteWorkbookDeltaDamage(): GridRenderTilePaneBridgeState {
    return this.renderTiles.noteWorkbookDeltaDamage()
  }

  noteLocalRenderTileFallbackInvalidation(): GridRenderTilePaneBridgeState {
    return this.renderTiles.noteLocalFallbackInvalidation()
  }

  connectRenderTileDeltas(
    input: Omit<GridRenderTileDeltaRuntimeInput, 'gridRuntimeHost'>,
    listener: Parameters<GridRenderTilePaneRuntime['connectRenderTileDeltas']>[1],
  ): ReturnType<GridRenderTilePaneRuntime['connectRenderTileDeltas']> {
    return this.renderTiles.connectRenderTileDeltas(
      {
        ...input,
        gridRuntimeHost: this,
      },
      listener,
    )
  }

  connectWorkbookDeltaDamage(
    input: Omit<GridRenderTileDamageRuntimeInput, 'gridRuntimeHost'>,
    listener: Parameters<GridRenderTilePaneRuntime['connectWorkbookDeltaDamage']>[1],
  ): ReturnType<GridRenderTilePaneRuntime['connectWorkbookDeltaDamage']> {
    return this.renderTiles.connectWorkbookDeltaDamage(
      {
        ...input,
        gridRuntimeHost: this,
      },
      listener,
    )
  }

  noteRenderTileReadiness(readiness: GridRenderTilePaneRuntimeState['tileReadiness']): void {
    this.renderTiles.noteTileReadiness(readiness)
  }

  connectLocalRenderTileCellInvalidation(
    input: GridRenderTileLocalInvalidationRuntimeInput,
    listener: Parameters<GridRenderTilePaneRuntime['connectLocalCellInvalidation']>[1],
  ): ReturnType<GridRenderTilePaneRuntime['connectLocalCellInvalidation']> {
    return this.renderTiles.connectLocalCellInvalidation(input, listener)
  }

  resolveScrollForCellIntoView(input: {
    readonly cell: Item
    readonly viewportWidth: number
    readonly viewportHeight: number
    readonly scrollLeft: number
    readonly scrollTop: number
    readonly gridMetrics: GridMetrics
    readonly freezeRows?: number | undefined
    readonly freezeCols?: number | undefined
  }): { readonly scrollLeft: number; readonly scrollTop: number } {
    this.updateFreeze(input.freezeRows, input.freezeCols)
    const col = input.cell[0]
    const row = input.cell[1]
    const frozenWidth = this.columns.span(0, this.freezeCols)
    const frozenHeight = this.rows.span(0, this.freezeRows)
    const bodyWidth = Math.max(0, input.viewportWidth - input.gridMetrics.rowMarkerWidth - frozenWidth)
    const bodyHeight = Math.max(0, input.viewportHeight - input.gridMetrics.headerHeight - frozenHeight)
    let scrollLeft = input.scrollLeft
    let scrollTop = input.scrollTop

    if (col >= this.freezeCols) {
      const scrollCellLeft = this.columns.offsetOf(col) - frozenWidth
      const cellWidth = this.columns.sizeAt(col)
      if (scrollCellLeft < scrollLeft) {
        scrollLeft = scrollCellLeft
      } else if (scrollCellLeft + cellWidth > scrollLeft + bodyWidth) {
        scrollLeft = scrollCellLeft + cellWidth - bodyWidth
      }
    }

    if (row >= this.freezeRows) {
      const scrollCellTop = this.rows.offsetOf(row) - frozenHeight
      const cellHeight = this.rows.sizeAt(row)
      if (scrollCellTop < scrollTop) {
        scrollTop = scrollCellTop
      } else if (scrollCellTop + cellHeight > scrollTop + bodyHeight) {
        scrollTop = scrollCellTop + cellHeight - bodyHeight
      }
    }

    return {
      scrollLeft: Math.max(0, scrollLeft),
      scrollTop: Math.max(0, scrollTop),
    }
  }

  resolveScrollPositionForViewport(input: {
    readonly viewport: Pick<Viewport, 'colStart' | 'rowStart'>
    readonly freezeRows?: number | undefined
    readonly freezeCols?: number | undefined
  }): { readonly scrollLeft: number; readonly scrollTop: number } {
    this.updateFreeze(input.freezeRows, input.freezeCols)
    const frozenWidth = this.columns.span(0, this.freezeCols)
    const frozenHeight = this.rows.span(0, this.freezeRows)
    return {
      scrollLeft:
        input.viewport.colStart <= this.freezeCols ? 0 : Math.max(0, this.columns.offsetOf(input.viewport.colStart) - frozenWidth),
      scrollTop: input.viewport.rowStart <= this.freezeRows ? 0 : Math.max(0, this.rows.offsetOf(input.viewport.rowStart) - frozenHeight),
    }
  }

  setOverlay(id: string, instance: OverlayInstanceV3): void {
    this.overlays.set(id, instance)
  }

  deleteOverlay(id: string): void {
    this.overlays.delete(id)
  }

  buildOverlayBatch(): OverlayBatchV3 {
    const snapshot = this.snapshot()
    return this.overlays.buildBatch({
      axisSeqX: snapshot.axisSeqX,
      axisSeqY: snapshot.axisSeqY,
      cameraSeq: snapshot.camera.seq,
    })
  }

  private updateFreeze(freezeRows: number | undefined, freezeCols: number | undefined): void {
    const nextRows = Math.max(0, freezeRows ?? this.freezeRows)
    const nextCols = Math.max(0, freezeCols ?? this.freezeCols)
    if (nextRows === this.freezeRows && nextCols === this.freezeCols) {
      return
    }
    this.freezeRows = nextRows
    this.freezeCols = nextCols
    this.freezeSeq += 1
  }
}
