import type { Viewport } from '@bilig/protocol'
import { MAX_COLS, MAX_ROWS } from '@bilig/protocol'
import type { GridCameraSnapshot } from './renderer/grid-render-contract.js'
import type { VisibleRegionState } from './gridPointer.js'

export function visibleRegionFromCamera(input: {
  readonly camera: GridCameraSnapshot
  readonly freezeRows: number
  readonly freezeCols: number
}): VisibleRegionState {
  return {
    freezeCols: input.freezeCols,
    freezeRows: input.freezeRows,
    range: {
      height: input.camera.visibleViewport.rowEnd - input.camera.visibleViewport.rowStart + 1,
      width: input.camera.visibleViewport.colEnd - input.camera.visibleViewport.colStart + 1,
      x: input.camera.visibleViewport.colStart,
      y: input.camera.visibleViewport.rowStart,
    },
    tx: input.camera.tx,
    ty: input.camera.ty,
  }
}

export function viewportFromVisibleRegion(region: VisibleRegionState): Viewport {
  return {
    colEnd: Math.min(MAX_COLS - 1, region.range.x + region.range.width - 1),
    colStart: region.range.x,
    rowEnd: Math.min(MAX_ROWS - 1, region.range.y + region.range.height - 1),
    rowStart: region.range.y,
  }
}
