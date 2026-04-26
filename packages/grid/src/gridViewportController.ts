import type { Viewport } from '@bilig/protocol'
import type { GridTileKey } from './gridTileResidencyV2.js'
import type { VisibleRegionState } from './gridPointer.js'
import { resolveRowOffset } from './gridMetrics.js'
import { resolveColumnOffset } from './workbookGridViewport.js'

type SortedAxisOverrides = readonly (readonly [number, number])[]

export function sameVisibleRegionWindow(left: VisibleRegionState, right: VisibleRegionState): boolean {
  return (
    (left.freezeCols ?? 0) === (right.freezeCols ?? 0) &&
    (left.freezeRows ?? 0) === (right.freezeRows ?? 0) &&
    left.range.x === right.range.x &&
    left.range.y === right.range.y &&
    left.range.width === right.range.width &&
    left.range.height === right.range.height
  )
}

export function sameViewportBounds(left: Viewport, right: Viewport): boolean {
  return (
    left.rowStart === right.rowStart && left.rowEnd === right.rowEnd && left.colStart === right.colStart && left.colEnd === right.colEnd
  )
}

export function tileKeyToViewport(tile: GridTileKey): Viewport {
  return {
    colEnd: tile.colEnd,
    colStart: tile.colStart,
    rowEnd: tile.rowEnd,
    rowStart: tile.rowStart,
  }
}

export function resolveGridRenderScrollTransform(options: {
  nextVisibleRegion: VisibleRegionState
  renderViewport: Viewport
  sortedColumnWidthOverrides: SortedAxisOverrides
  sortedRowHeightOverrides: SortedAxisOverrides
  defaultColumnWidth: number
  defaultRowHeight: number
}): { renderTx: number; renderTy: number } {
  const { nextVisibleRegion, renderViewport, sortedColumnWidthOverrides, sortedRowHeightOverrides, defaultColumnWidth, defaultRowHeight } =
    options
  return {
    renderTx:
      resolveColumnOffset(nextVisibleRegion.range.x, sortedColumnWidthOverrides, defaultColumnWidth) -
      resolveColumnOffset(renderViewport.colStart, sortedColumnWidthOverrides, defaultColumnWidth) +
      nextVisibleRegion.tx,
    renderTy:
      resolveRowOffset(nextVisibleRegion.range.y, sortedRowHeightOverrides, defaultRowHeight) -
      resolveRowOffset(renderViewport.rowStart, sortedRowHeightOverrides, defaultRowHeight) +
      nextVisibleRegion.ty,
  }
}
