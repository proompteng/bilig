import type { GridMetrics } from '../gridMetrics.js'
import type { OverlayBatchV3, OverlayInstanceV3 } from '../renderer-v3/overlay-layer.js'
import type { TileKey53 } from '../renderer-v3/tile-key.js'
import { GridAxisRuntime } from './gridAxisRuntime.js'
import { GridCameraRuntime, type GridCameraRuntimeSnapshot } from './gridCameraRuntime.js'
import { GridOverlayRuntime } from './gridOverlayRuntime.js'

export interface GridRuntimeHostSnapshot {
  readonly camera: GridCameraRuntimeSnapshot
  readonly axisSeqX: number
  readonly axisSeqY: number
}

export class GridRuntimeHost {
  readonly columns: GridAxisRuntime
  readonly rows: GridAxisRuntime
  readonly camera: GridCameraRuntime
  readonly overlays = new GridOverlayRuntime()

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
      freezeCols: input.freezeCols,
      freezeRows: input.freezeRows,
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
    return this.camera.update({
      columns: this.columns,
      freezeCols: input.freezeCols,
      freezeRows: input.freezeRows,
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
}
