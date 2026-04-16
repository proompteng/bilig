import { MAX_COLS, MAX_ROWS } from '@bilig/protocol'
import {
  COLUMN_RESIZE_HANDLE_THRESHOLD,
  SCROLLBAR_GUTTER,
  type GridMetrics,
  type GridRect,
  getResolvedColumnWidth,
  getResolvedRowHeight,
  getVisibleColumnBounds,
  getVisibleRowBounds,
  resolveColumnAtClientX,
  resolveRowAtClientY,
} from './gridMetrics.js'
import type { Item, Rectangle } from './gridTypes.js'

export interface VisibleRegionState {
  range: Rectangle
  tx: number
  ty: number
  freezeRows?: number
  freezeCols?: number
}

export interface PointerGeometry {
  hostBounds: GridRect
  cellWidth: number
  cellHeight: number
  dataLeft: number
  dataTop: number
  dataRight: number
  dataBottom: number
  frozenLeftWidth: number
  frozenTopHeight: number
  mainDataLeft: number
  mainDataTop: number
}

export interface SelectedCellBounds {
  x: number
  y: number
  width: number
  height: number
}

export type HeaderSelection = { kind: 'column'; index: number } | { kind: 'row'; index: number }

export interface PointerCellResolutionInput {
  clientX: number
  clientY: number
  region: VisibleRegionState
  geometry: PointerGeometry
  columnWidths: Readonly<Record<number, number>>
  rowHeights: Readonly<Record<number, number>>
  gridMetrics: GridMetrics
  selectedCell: Item
  selectedCellBounds?: SelectedCellBounds | null
  selectionRange?: Rectangle | null
  hasColumnSelection: boolean
  hasRowSelection: boolean
}

export function createPointerGeometry(
  hostBounds: GridRect,
  region: VisibleRegionState,
  columnWidths: Readonly<Record<number, number>>,
  rowHeights: Readonly<Record<number, number>>,
  gridMetrics: GridMetrics,
): PointerGeometry {
  const freezeCols = Math.max(0, Math.min(MAX_COLS, region.freezeCols ?? 0))
  const freezeRows = Math.max(0, Math.min(MAX_ROWS, region.freezeRows ?? 0))
  const cellWidth = gridMetrics.columnWidth
  const cellHeight = gridMetrics.rowHeight
  const dataLeft = hostBounds.left + gridMetrics.rowMarkerWidth
  const dataTop = hostBounds.top + gridMetrics.headerHeight
  const frozenLeftWidth = resolveFrozenColumnWidth(freezeCols, columnWidths, gridMetrics.columnWidth)
  const frozenTopHeight = resolveFrozenRowHeight(freezeRows, rowHeights, gridMetrics.rowHeight)
  const mainDataLeft = dataLeft + frozenLeftWidth
  const mainDataTop = dataTop + frozenTopHeight
  const visibleColumnBounds = getVisibleColumnBounds(region.range, mainDataLeft, MAX_COLS, columnWidths, gridMetrics.columnWidth)
  const mainDataWidth = visibleColumnBounds.length === 0 ? region.range.width * cellWidth : visibleColumnBounds.at(-1)!.right - mainDataLeft
  const visibleRowBounds = getVisibleRowBounds(region.range, mainDataTop, MAX_ROWS, rowHeights, gridMetrics.rowHeight)
  const mainDataHeight = visibleRowBounds.length === 0 ? region.range.height * cellHeight : visibleRowBounds.at(-1)!.bottom - mainDataTop
  return {
    hostBounds,
    cellWidth,
    cellHeight,
    dataLeft,
    dataTop,
    dataRight: Math.min(hostBounds.right - SCROLLBAR_GUTTER, dataLeft + frozenLeftWidth + mainDataWidth),
    dataBottom: Math.min(hostBounds.bottom - SCROLLBAR_GUTTER, dataTop + frozenTopHeight + mainDataHeight),
    frozenLeftWidth,
    frozenTopHeight,
    mainDataLeft,
    mainDataTop,
  }
}

export function resolveColumnResizeTarget(
  clientX: number,
  clientY: number,
  region: VisibleRegionState,
  geometry: PointerGeometry,
  columnWidths: Readonly<Record<number, number>>,
  defaultWidth: number,
): number | null {
  if (clientY < geometry.hostBounds.top || clientY >= geometry.dataTop) {
    return null
  }
  for (const column of getPointerVisibleColumns(region, geometry, columnWidths, defaultWidth)) {
    if (column.index >= MAX_COLS - 1) {
      continue
    }
    if (clientX >= column.right - COLUMN_RESIZE_HANDLE_THRESHOLD && clientX <= column.right + COLUMN_RESIZE_HANDLE_THRESHOLD) {
      return column.index
    }
  }
  return null
}

export function resolveRowResizeTarget(
  clientX: number,
  clientY: number,
  region: VisibleRegionState,
  geometry: PointerGeometry,
  rowHeights: Readonly<Record<number, number>>,
  defaultHeight: number,
): number | null {
  if (clientX < geometry.hostBounds.left || clientX >= geometry.dataLeft) {
    return null
  }
  for (const row of getPointerVisibleRows(region, geometry, rowHeights, defaultHeight)) {
    if (row.index >= MAX_ROWS - 1) {
      continue
    }
    if (clientY >= row.bottom - COLUMN_RESIZE_HANDLE_THRESHOLD && clientY <= row.bottom + COLUMN_RESIZE_HANDLE_THRESHOLD) {
      return row.index
    }
  }
  return null
}

export function resolvePointerCell(input: PointerCellResolutionInput): Item | null {
  const {
    clientX,
    clientY,
    region,
    geometry,
    columnWidths,
    rowHeights,
    gridMetrics,
    selectedCell,
    selectedCellBounds,
    selectionRange,
    hasColumnSelection,
    hasRowSelection,
  } = input
  const { hostBounds, dataLeft, dataTop, dataRight, dataBottom } = geometry

  if (clientX >= hostBounds.right - SCROLLBAR_GUTTER || clientY >= hostBounds.bottom - SCROLLBAR_GUTTER) {
    return null
  }

  if (clientX < dataLeft || clientX >= dataRight || clientY < dataTop || clientY >= dataBottom) {
    return null
  }

  if (
    selectedCellBounds &&
    !hasColumnSelection &&
    !hasRowSelection &&
    selectionRange?.width === 1 &&
    selectionRange?.height === 1 &&
    clientX >= selectedCellBounds.x - 1 &&
    clientX < selectedCellBounds.x + selectedCellBounds.width &&
    clientY >= selectedCellBounds.y - 1 &&
    clientY < selectedCellBounds.y + selectedCellBounds.height
  ) {
    return selectedCell
  }

  const col = resolvePointerColumnIndex(clientX, region, geometry, columnWidths, gridMetrics.columnWidth)
  const row = resolvePointerRowIndex(clientY, region, geometry, rowHeights, gridMetrics.rowHeight)
  if (col === null || col < 0 || col >= MAX_COLS || row === null || row < 0 || row >= MAX_ROWS) {
    return null
  }

  return [col, row]
}

export function resolveHeaderSelection(
  clientX: number,
  clientY: number,
  region: VisibleRegionState,
  geometry: PointerGeometry,
  columnWidths: Readonly<Record<number, number>>,
  rowHeights: Readonly<Record<number, number>>,
  gridMetrics: GridMetrics,
): HeaderSelection | null {
  const { hostBounds, dataLeft, dataTop, dataRight, dataBottom } = geometry
  const headerBottom = dataTop
  const rowAreaRight = dataLeft

  if (clientY >= hostBounds.top && clientY < headerBottom && clientX >= dataLeft && clientX < dataRight) {
    const col = resolvePointerColumnIndex(clientX, region, geometry, columnWidths, gridMetrics.columnWidth)
    if (col !== null && col >= 0 && col < MAX_COLS) {
      return { kind: 'column', index: col }
    }
  }

  if (clientX >= hostBounds.left && clientX < rowAreaRight && clientY >= dataTop && clientY < dataBottom) {
    const row = resolvePointerRowIndex(clientY, region, geometry, rowHeights, gridMetrics.rowHeight)
    if (row !== null && row >= 0 && row < MAX_ROWS) {
      return { kind: 'row', index: row }
    }
  }

  return null
}

export function resolveHeaderSelectionForDrag(
  kind: HeaderSelection['kind'],
  clientX: number,
  clientY: number,
  region: VisibleRegionState,
  geometry: PointerGeometry,
  columnWidths: Readonly<Record<number, number>>,
  rowHeights: Readonly<Record<number, number>>,
  gridMetrics: GridMetrics,
): HeaderSelection | null {
  const { hostBounds, dataLeft, dataTop, dataRight, dataBottom } = geometry

  if (kind === 'column') {
    if (clientX < dataLeft || clientX >= dataRight || clientY < hostBounds.top || clientY >= dataBottom) {
      return null
    }
    const col = resolvePointerColumnIndex(clientX, region, geometry, columnWidths, gridMetrics.columnWidth)
    if (col === null || col < 0 || col >= MAX_COLS) {
      return null
    }
    return { kind: 'column', index: col }
  }

  if (clientY < dataTop || clientY >= dataBottom || clientX < hostBounds.left || clientX >= dataRight) {
    return null
  }
  const row = resolvePointerRowIndex(clientY, region, geometry, rowHeights, gridMetrics.rowHeight)
  if (row === null || row < 0 || row >= MAX_ROWS) {
    return null
  }
  return { kind: 'row', index: row }
}

function resolveFrozenColumnWidth(freezeCols: number, columnWidths: Readonly<Record<number, number>>, defaultWidth: number): number {
  let width = 0
  for (let col = 0; col < freezeCols; col += 1) {
    width += getResolvedColumnWidth(columnWidths, col, defaultWidth)
  }
  return width
}

function resolveFrozenRowHeight(freezeRows: number, rowHeights: Readonly<Record<number, number>>, defaultHeight: number): number {
  let height = 0
  for (let row = 0; row < freezeRows; row += 1) {
    height += getResolvedRowHeight(rowHeights, row, defaultHeight)
  }
  return height
}

function getPointerVisibleColumns(
  region: VisibleRegionState,
  geometry: PointerGeometry,
  columnWidths: Readonly<Record<number, number>>,
  defaultWidth: number,
) {
  const freezeCols = Math.max(0, Math.min(MAX_COLS, region.freezeCols ?? 0))
  const frozenColumns =
    freezeCols === 0 ? [] : getVisibleColumnBounds({ x: 0, width: freezeCols }, geometry.dataLeft, MAX_COLS, columnWidths, defaultWidth)
  const mainColumns = getVisibleColumnBounds(region.range, geometry.mainDataLeft, MAX_COLS, columnWidths, defaultWidth)
  return [...frozenColumns, ...mainColumns]
}

function getPointerVisibleRows(
  region: VisibleRegionState,
  geometry: PointerGeometry,
  rowHeights: Readonly<Record<number, number>>,
  defaultHeight: number,
) {
  const freezeRows = Math.max(0, Math.min(MAX_ROWS, region.freezeRows ?? 0))
  const frozenRows =
    freezeRows === 0 ? [] : getVisibleRowBounds({ y: 0, height: freezeRows }, geometry.dataTop, MAX_ROWS, rowHeights, defaultHeight)
  const mainRows = getVisibleRowBounds(region.range, geometry.mainDataTop, MAX_ROWS, rowHeights, defaultHeight)
  return [...frozenRows, ...mainRows]
}

function resolvePointerColumnIndex(
  clientX: number,
  region: VisibleRegionState,
  geometry: PointerGeometry,
  columnWidths: Readonly<Record<number, number>>,
  defaultWidth: number,
): number | null {
  const freezeCols = Math.max(0, Math.min(MAX_COLS, region.freezeCols ?? 0))
  if (freezeCols > 0 && clientX < geometry.mainDataLeft) {
    return resolveColumnAtClientX(clientX, { x: 0, width: freezeCols }, geometry.dataLeft, MAX_COLS, columnWidths, defaultWidth)
  }
  return resolveColumnAtClientX(clientX, region.range, geometry.mainDataLeft, MAX_COLS, columnWidths, defaultWidth)
}

function resolvePointerRowIndex(
  clientY: number,
  region: VisibleRegionState,
  geometry: PointerGeometry,
  rowHeights: Readonly<Record<number, number>>,
  defaultHeight: number,
): number | null {
  const freezeRows = Math.max(0, Math.min(MAX_ROWS, region.freezeRows ?? 0))
  if (freezeRows > 0 && clientY < geometry.mainDataTop) {
    return resolveRowAtClientY(clientY, { y: 0, height: freezeRows }, geometry.dataTop, MAX_ROWS, rowHeights, defaultHeight)
  }
  return resolveRowAtClientY(clientY, region.range, geometry.mainDataTop, MAX_ROWS, rowHeights, defaultHeight)
}
