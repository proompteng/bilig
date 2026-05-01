import type { GridGeometrySnapshot } from './gridGeometry.js'
import type { GridMetrics } from './gridMetrics.js'
import type { GridHoverState } from './gridHover.js'
import type { HeaderSelection, VisibleRegionState } from './gridPointer.js'
import { resolveSelectionContentMoveCandidateCell, resolveSelectionMoveAnchorCell } from './gridRangeMove.js'
import type { Item, Rectangle } from './gridTypes.js'

const DEFAULT_HOVER_STATE: GridHoverState = { cell: null, header: null, cursor: 'default' }
const RANGE_MOVE_HOVER_STATE: GridHoverState = { cell: null, header: null, cursor: 'grab' }
const RANGE_MOVE_DRAG_HOVER_STATE: GridHoverState = { cell: null, header: null, cursor: 'grabbing' }

export function resolveGridInteractionHoverState(input: {
  readonly clientX: number
  readonly clientY: number
  readonly visibleRegion: VisibleRegionState
  readonly geometry: GridGeometrySnapshot
  readonly columnWidths: Readonly<Record<number, number>>
  readonly rowHeights: Readonly<Record<number, number>>
  readonly gridMetrics: GridMetrics
  readonly resolveColumnResizeTargetAtPointer: (
    clientX: number,
    clientY: number,
    region: VisibleRegionState,
    geometry?: GridGeometrySnapshot | null,
    columnWidths?: Readonly<Record<number, number>>,
    defaultWidth?: number,
  ) => number | null
  readonly resolveRowResizeTargetAtPointer: (
    clientX: number,
    clientY: number,
    region: VisibleRegionState,
    geometry?: GridGeometrySnapshot | null,
    rowHeights?: Readonly<Record<number, number>>,
    defaultHeight?: number,
  ) => number | null
  readonly resolveHeaderSelectionAtPointer: (
    clientX: number,
    clientY: number,
    region?: VisibleRegionState,
    geometry?: GridGeometrySnapshot | null,
  ) => HeaderSelection | null
  readonly resolvePointerCell: (
    clientX: number,
    clientY: number,
    region?: VisibleRegionState,
    geometry?: GridGeometrySnapshot | null,
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
  return cell ? { cell, header: null, cursor: 'cell' } : DEFAULT_HOVER_STATE
}

export function resolveWorkbookGridHoverState(input: {
  readonly clientX: number
  readonly clientY: number
  readonly buttons: number
  readonly isFillHandleDragging: boolean
  readonly isRangeMoveDragging: boolean
  readonly hasFillPreviewRange: boolean
  readonly allowsRangeMove: boolean
  readonly selectionRange: Rectangle | null
  readonly getCellScreenBounds: (col: number, row: number) => Rectangle | undefined
  readonly getVisibleRegion: () => VisibleRegionState
  readonly resolvePointerGeometry: (visibleRegion: VisibleRegionState) => GridGeometrySnapshot | null
  readonly columnWidths: Readonly<Record<number, number>>
  readonly rowHeights: Readonly<Record<number, number>>
  readonly gridMetrics: GridMetrics
  readonly resolveColumnResizeTargetAtPointer: (
    clientX: number,
    clientY: number,
    region: VisibleRegionState,
    geometry?: GridGeometrySnapshot | null,
    columnWidths?: Readonly<Record<number, number>>,
    defaultWidth?: number,
  ) => number | null
  readonly resolveRowResizeTargetAtPointer: (
    clientX: number,
    clientY: number,
    region: VisibleRegionState,
    geometry?: GridGeometrySnapshot | null,
    rowHeights?: Readonly<Record<number, number>>,
    defaultHeight?: number,
  ) => number | null
  readonly resolveHeaderSelectionAtPointer: (
    clientX: number,
    clientY: number,
    region?: VisibleRegionState,
    geometry?: GridGeometrySnapshot | null,
  ) => HeaderSelection | null
  readonly resolvePointerCell: (
    clientX: number,
    clientY: number,
    region?: VisibleRegionState,
    geometry?: GridGeometrySnapshot | null,
  ) => Item | null
}): GridHoverState {
  if (input.isFillHandleDragging) {
    return DEFAULT_HOVER_STATE
  }
  if (input.isRangeMoveDragging) {
    return RANGE_MOVE_DRAG_HOVER_STATE
  }
  if (input.buttons !== 0 || input.hasFillPreviewRange) {
    return DEFAULT_HOVER_STATE
  }

  const visibleRegion = input.getVisibleRegion()
  const geometry = input.resolvePointerGeometry(visibleRegion)
  if (!geometry) {
    return DEFAULT_HOVER_STATE
  }

  const rangeMoveAnchorCell =
    input.allowsRangeMove && input.selectionRange
      ? resolveSelectionMoveAnchorCell(input.clientX, input.clientY, input.selectionRange, input.getCellScreenBounds)
      : null
  if (rangeMoveAnchorCell) {
    return RANGE_MOVE_HOVER_STATE
  }
  const rangeMoveInteriorCell =
    input.allowsRangeMove && input.selectionRange
      ? resolveSelectionContentMoveCandidateCell(input.clientX, input.clientY, input.selectionRange, input.getCellScreenBounds)
      : null
  if (rangeMoveInteriorCell) {
    return RANGE_MOVE_HOVER_STATE
  }

  return resolveGridInteractionHoverState({
    clientX: input.clientX,
    clientY: input.clientY,
    columnWidths: input.columnWidths,
    geometry,
    gridMetrics: input.gridMetrics,
    resolveColumnResizeTargetAtPointer: input.resolveColumnResizeTargetAtPointer,
    resolveHeaderSelectionAtPointer: input.resolveHeaderSelectionAtPointer,
    resolvePointerCell: input.resolvePointerCell,
    resolveRowResizeTargetAtPointer: input.resolveRowResizeTargetAtPointer,
    rowHeights: input.rowHeights,
    visibleRegion,
  })
}
