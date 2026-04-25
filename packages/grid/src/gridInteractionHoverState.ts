import type { GridMetrics } from './gridMetrics.js'
import type { GridHoverState } from './gridHover.js'
import type { HeaderSelection, PointerGeometry, VisibleRegionState } from './gridPointer.js'
import type { Item } from './gridTypes.js'

export function resolveGridInteractionHoverState(input: {
  readonly clientX: number
  readonly clientY: number
  readonly visibleRegion: VisibleRegionState
  readonly geometry: PointerGeometry
  readonly columnWidths: Readonly<Record<number, number>>
  readonly rowHeights: Readonly<Record<number, number>>
  readonly gridMetrics: GridMetrics
  readonly resolveColumnResizeTargetAtPointer: (
    clientX: number,
    clientY: number,
    region: VisibleRegionState,
    geometry?: PointerGeometry | null,
    columnWidths?: Readonly<Record<number, number>>,
    defaultWidth?: number,
  ) => number | null
  readonly resolveRowResizeTargetAtPointer: (
    clientX: number,
    clientY: number,
    region: VisibleRegionState,
    geometry?: PointerGeometry | null,
    rowHeights?: Readonly<Record<number, number>>,
    defaultHeight?: number,
  ) => number | null
  readonly resolveHeaderSelectionAtPointer: (
    clientX: number,
    clientY: number,
    region?: VisibleRegionState,
    geometry?: PointerGeometry | null,
  ) => HeaderSelection | null
  readonly resolvePointerCell: (
    clientX: number,
    clientY: number,
    region?: VisibleRegionState,
    geometry?: PointerGeometry | null,
  ) => Item | null
}): GridHoverState {
  const resizeTarget = input.resolveColumnResizeTargetAtPointer(
    input.clientX,
    input.clientY,
    input.visibleRegion,
    input.geometry,
    input.columnWidths,
    input.gridMetrics.columnWidth,
  )
  if (resizeTarget !== null) {
    return { cell: null, header: { kind: 'column', index: resizeTarget }, cursor: 'col-resize' }
  }

  const rowResizeTarget = input.resolveRowResizeTargetAtPointer(
    input.clientX,
    input.clientY,
    input.visibleRegion,
    input.geometry,
    input.rowHeights,
    input.gridMetrics.rowHeight,
  )
  if (rowResizeTarget !== null) {
    return { cell: null, header: { kind: 'row', index: rowResizeTarget }, cursor: 'row-resize' }
  }

  const header = input.resolveHeaderSelectionAtPointer(input.clientX, input.clientY, input.visibleRegion, input.geometry)
  if (header) {
    return { cell: null, header, cursor: 'pointer' }
  }

  const cell = input.resolvePointerCell(input.clientX, input.clientY, input.visibleRegion, input.geometry)
  return cell ? { cell, header: null, cursor: 'cell' } : { cell: null, header: null, cursor: 'default' }
}
