import { MAX_COLUMN_WIDTH, MAX_ROW_HEIGHT, type GridMetrics } from '../gridMetrics.js'
import type { VisibleRegionState } from '../gridPointer.js'
import { viewportFromVisibleRegion } from '../useGridCameraState.js'
import { tileKeysForViewport, type TileKey53 } from '../renderer-v3/tile-key.js'
import type { GridAxisRuntime } from './gridAxisRuntime.js'

export interface GridCameraRuntimeSnapshot {
  readonly seq: number
  readonly scrollLeft: number
  readonly scrollTop: number
  readonly viewportWidth: number
  readonly viewportHeight: number
  readonly dpr: number
  readonly visibleRegion: VisibleRegionState
}

export class GridCameraRuntime {
  private snapshotValue: GridCameraRuntimeSnapshot

  constructor(input: {
    readonly columns: GridAxisRuntime
    readonly rows: GridAxisRuntime
    readonly gridMetrics: GridMetrics
    readonly viewportWidth: number
    readonly viewportHeight: number
    readonly freezeRows?: number | undefined
    readonly freezeCols?: number | undefined
    readonly scrollLeft?: number | undefined
    readonly scrollTop?: number | undefined
    readonly dpr?: number | undefined
  }) {
    this.snapshotValue = this.createSnapshot({ ...input, previousSeq: 0 })
  }

  snapshot(): GridCameraRuntimeSnapshot {
    return this.snapshotValue
  }

  update(input: {
    readonly columns: GridAxisRuntime
    readonly rows: GridAxisRuntime
    readonly gridMetrics: GridMetrics
    readonly viewportWidth: number
    readonly viewportHeight: number
    readonly freezeRows?: number | undefined
    readonly freezeCols?: number | undefined
    readonly scrollLeft: number
    readonly scrollTop: number
    readonly dpr: number
  }): GridCameraRuntimeSnapshot {
    this.snapshotValue = this.createSnapshot({
      ...input,
      previousSeq: this.snapshotValue.seq,
    })
    return this.snapshotValue
  }

  visibleTileKeys(input: { readonly sheetOrdinal: number; readonly dprBucket: number }): TileKey53[] {
    return tileKeysForViewport({
      dprBucket: input.dprBucket,
      sheetOrdinal: input.sheetOrdinal,
      viewport: viewportFromVisibleRegion(this.snapshotValue.visibleRegion),
    })
  }

  private createSnapshot(input: {
    readonly columns: GridAxisRuntime
    readonly rows: GridAxisRuntime
    readonly gridMetrics: GridMetrics
    readonly viewportWidth: number
    readonly viewportHeight: number
    readonly freezeRows?: number | undefined
    readonly freezeCols?: number | undefined
    readonly scrollLeft?: number | undefined
    readonly scrollTop?: number | undefined
    readonly dpr?: number | undefined
    readonly previousSeq: number
  }): GridCameraRuntimeSnapshot {
    const freezeRows = Math.max(0, input.freezeRows ?? 0)
    const freezeCols = Math.max(0, input.freezeCols ?? 0)
    const scrollLeft = Math.max(0, input.scrollLeft ?? 0)
    const scrollTop = Math.max(0, input.scrollTop ?? 0)
    const frozenWidth = input.columns.span(0, freezeCols)
    const frozenHeight = input.rows.span(0, freezeRows)
    const bodyWidth = Math.max(0, input.viewportWidth - input.gridMetrics.rowMarkerWidth - frozenWidth)
    const bodyHeight = Math.max(0, input.viewportHeight - input.gridMetrics.headerHeight - frozenHeight)
    const horizontalRange = input.columns.visibleRangeForOffset(scrollLeft + frozenWidth, bodyWidth, MAX_COLUMN_WIDTH)
    const verticalRange = input.rows.visibleRangeForOffset(scrollTop + frozenHeight, bodyHeight, MAX_ROW_HEIGHT)
    const horizontalAnchorOffset = scrollLeft + frozenWidth - input.columns.offsetOf(horizontalRange.start)
    const verticalAnchorOffset = scrollTop + frozenHeight - input.rows.offsetOf(verticalRange.start)
    return {
      dpr: input.dpr ?? 1,
      scrollLeft,
      scrollTop,
      seq: input.previousSeq + 1,
      viewportHeight: input.viewportHeight,
      viewportWidth: input.viewportWidth,
      visibleRegion: {
        freezeCols,
        freezeRows,
        range: {
          height: Math.max(1, verticalRange.endExclusive - verticalRange.start),
          width: Math.max(1, horizontalRange.endExclusive - horizontalRange.start),
          x: Math.max(freezeCols, horizontalRange.start),
          y: Math.max(freezeRows, verticalRange.start),
        },
        tx: Math.max(0, horizontalAnchorOffset),
        ty: Math.max(0, verticalAnchorOffset),
      },
    }
  }
}
