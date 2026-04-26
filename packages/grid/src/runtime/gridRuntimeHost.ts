import type { GridMetrics } from '../gridMetrics.js'
import type { OverlayBatchV3, OverlayInstanceV3 } from '../renderer-v3/overlay-layer.js'
import type { TileKey53 } from '../renderer-v3/tile-key.js'
import { GridAxisRuntime } from './gridAxisRuntime.js'
import { GridCameraRuntime, type GridCameraRuntimeSnapshot } from './gridCameraRuntime.js'
import { GridOverlayRuntime } from './gridOverlayRuntime.js'
import { GridTileCoordinator, type GridTileInterestBatchV3, type GridTileInterestReasonV3 } from './gridTileCoordinator.js'

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
    return this.camera.visibleTileKeys(input)
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
