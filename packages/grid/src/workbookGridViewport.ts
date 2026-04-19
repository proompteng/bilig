import { MAX_COLS, MAX_ROWS, VIEWPORT_TILE_COLUMN_COUNT, VIEWPORT_TILE_ROW_COUNT, type Viewport } from '@bilig/protocol'
import type { getGridMetrics } from './gridMetrics.js'
import { MAX_COLUMN_WIDTH, MAX_ROW_HEIGHT, getResolvedColumnWidth, getResolvedRowHeight, resolveRowOffset } from './gridMetrics.js'
import type { Item } from './gridTypes.js'
import type { VisibleRegionState } from './gridPointer.js'

function resolveFrozenColumnOffset(freezeCols: number, columnWidths: Readonly<Record<number, number>>, defaultWidth: number): number {
  let width = 0
  for (let col = 0; col < freezeCols; col += 1) {
    width += getResolvedColumnWidth(columnWidths, col, defaultWidth)
  }
  return width
}

function resolveFrozenRowOffset(freezeRows: number, rowHeights: Readonly<Record<number, number>>, defaultHeight: number): number {
  let height = 0
  for (let row = 0; row < freezeRows; row += 1) {
    height += getResolvedRowHeight(rowHeights, row, defaultHeight)
  }
  return height
}

export function resolveFrozenColumnWidth(options: {
  freezeCols: number
  columnWidths: Readonly<Record<number, number>>
  gridMetrics: ReturnType<typeof getGridMetrics>
}): number {
  return resolveFrozenColumnOffset(Math.max(0, options.freezeCols), options.columnWidths, options.gridMetrics.columnWidth)
}

export function resolveFrozenRowHeight(options: {
  freezeRows: number
  rowHeights: Readonly<Record<number, number>>
  gridMetrics: ReturnType<typeof getGridMetrics>
}): number {
  return resolveFrozenRowOffset(Math.max(0, options.freezeRows), options.rowHeights, options.gridMetrics.rowHeight)
}

export function resolveVisibleRegionFromScroll(options: {
  scrollLeft: number
  scrollTop: number
  viewportWidth: number
  viewportHeight: number
  freezeRows?: number
  freezeCols?: number
  columnWidths: Readonly<Record<number, number>>
  rowHeights: Readonly<Record<number, number>>
  gridMetrics: ReturnType<typeof getGridMetrics>
}): VisibleRegionState {
  const {
    scrollLeft,
    scrollTop,
    viewportWidth,
    viewportHeight,
    freezeRows: requestedFreezeRows = 0,
    freezeCols: requestedFreezeCols = 0,
    columnWidths,
    rowHeights,
    gridMetrics,
  } = options
  const freezeRows = Math.max(0, Math.min(MAX_ROWS, requestedFreezeRows))
  const freezeCols = Math.max(0, Math.min(MAX_COLS, requestedFreezeCols))
  const frozenWidth = resolveFrozenColumnOffset(freezeCols, columnWidths, gridMetrics.columnWidth)
  const frozenHeight = resolveFrozenRowOffset(freezeRows, rowHeights, gridMetrics.rowHeight)
  const bodyWidth = Math.max(0, viewportWidth - gridMetrics.rowMarkerWidth - frozenWidth)
  const bodyHeight = Math.max(0, viewportHeight - gridMetrics.headerHeight - frozenHeight)
  const horizontalAnchor = resolveColumnAnchor(scrollLeft + frozenWidth, columnWidths, gridMetrics.columnWidth)
  const verticalAnchor = resolveRowAnchor(scrollTop + frozenHeight, rowHeights, gridMetrics.rowHeight)
  const rangeX = Math.max(freezeCols, horizontalAnchor.index)
  const rangeY = Math.max(freezeRows, verticalAnchor.index)

  return {
    range: {
      x: rangeX,
      y: rangeY,
      width: resolveVisibleColumnCount({
        startCol: rangeX,
        bodyWidth,
        columnWidths,
        defaultWidth: gridMetrics.columnWidth,
      }),
      height: resolveVisibleRowCount({
        startRow: rangeY,
        bodyHeight,
        rowHeights,
        defaultHeight: gridMetrics.rowHeight,
      }),
    },
    tx: horizontalAnchor.offset,
    ty: verticalAnchor.offset,
    freezeRows,
    freezeCols,
  }
}

export function resolveResidentViewport(viewport: Viewport): Viewport {
  const rowStart = clampAxisStart(viewport.rowStart, VIEWPORT_TILE_ROW_COUNT, MAX_ROWS)
  const rowEnd = clampAxisEnd(viewport.rowEnd, VIEWPORT_TILE_ROW_COUNT, MAX_ROWS)
  const colStart = clampAxisStart(viewport.colStart, VIEWPORT_TILE_COLUMN_COUNT, MAX_COLS)
  const colEnd = clampAxisEnd(viewport.colEnd, VIEWPORT_TILE_COLUMN_COUNT, MAX_COLS)
  return {
    rowStart,
    rowEnd,
    colStart,
    colEnd,
  }
}

export function resolveColumnOffset(
  targetColumn: number,
  sortedColumnWidthOverrides: readonly (readonly [number, number])[],
  defaultWidth: number,
): number {
  let offset = targetColumn * defaultWidth
  for (const [columnIndex, width] of sortedColumnWidthOverrides) {
    if (columnIndex >= targetColumn) {
      break
    }
    offset += width - defaultWidth
  }
  return offset
}

export function scrollCellIntoView(options: {
  cell: Item
  freezeRows?: number
  freezeCols?: number
  columnWidths: Readonly<Record<number, number>>
  rowHeights: Readonly<Record<number, number>>
  gridMetrics: ReturnType<typeof getGridMetrics>
  scrollViewport: HTMLDivElement
  sortedColumnWidthOverrides: readonly (readonly [number, number])[]
  sortedRowHeightOverrides: readonly (readonly [number, number])[]
}): void {
  const {
    cell,
    freezeRows: requestedFreezeRows = 0,
    freezeCols: requestedFreezeCols = 0,
    columnWidths,
    rowHeights,
    gridMetrics,
    scrollViewport,
    sortedColumnWidthOverrides,
    sortedRowHeightOverrides,
  } = options
  const freezeRows = Math.max(0, Math.min(MAX_ROWS, requestedFreezeRows))
  const freezeCols = Math.max(0, Math.min(MAX_COLS, requestedFreezeCols))
  const frozenWidth = resolveColumnOffset(freezeCols, sortedColumnWidthOverrides, gridMetrics.columnWidth)
  const frozenHeight = resolveRowOffset(freezeRows, sortedRowHeightOverrides, gridMetrics.rowHeight)
  const cellLeft = resolveColumnOffset(cell[0], sortedColumnWidthOverrides, gridMetrics.columnWidth)
  const cellWidth = getResolvedColumnWidth(columnWidths, cell[0], gridMetrics.columnWidth)
  const bodyWidth = Math.max(0, scrollViewport.clientWidth - gridMetrics.rowMarkerWidth - frozenWidth)
  if (cell[0] >= freezeCols) {
    const scrollCellLeft = cellLeft - frozenWidth
    if (scrollCellLeft < scrollViewport.scrollLeft) {
      scrollViewport.scrollLeft = scrollCellLeft
    } else if (scrollCellLeft + cellWidth > scrollViewport.scrollLeft + bodyWidth) {
      scrollViewport.scrollLeft = scrollCellLeft + cellWidth - bodyWidth
    }
  }

  const cellTop = resolveRowOffset(cell[1], sortedRowHeightOverrides, gridMetrics.rowHeight)
  const cellHeight = getResolvedRowHeight(rowHeights, cell[1], gridMetrics.rowHeight)
  const bodyHeight = Math.max(0, scrollViewport.clientHeight - gridMetrics.headerHeight - frozenHeight)
  if (cell[1] >= freezeRows) {
    const scrollCellTop = cellTop - frozenHeight
    if (scrollCellTop < scrollViewport.scrollTop) {
      scrollViewport.scrollTop = scrollCellTop
    } else if (scrollCellTop + cellHeight > scrollViewport.scrollTop + bodyHeight) {
      scrollViewport.scrollTop = scrollCellTop + cellHeight - bodyHeight
    }
  }
}

export function resolveViewportScrollPosition(options: {
  viewport: Pick<VisibleRegionState['range'], 'x' | 'y'> | { colStart: number; rowStart: number }
  freezeRows?: number
  freezeCols?: number
  sortedColumnWidthOverrides: readonly (readonly [number, number])[]
  sortedRowHeightOverrides: readonly (readonly [number, number])[]
  gridMetrics: ReturnType<typeof getGridMetrics>
}): { scrollLeft: number; scrollTop: number } {
  const colStart = 'colStart' in options.viewport ? options.viewport.colStart : options.viewport.x
  const rowStart = 'rowStart' in options.viewport ? options.viewport.rowStart : options.viewport.y
  const freezeCols = Math.max(0, Math.min(MAX_COLS, options.freezeCols ?? 0))
  const freezeRows = Math.max(0, Math.min(MAX_ROWS, options.freezeRows ?? 0))
  const frozenWidth = resolveColumnOffset(freezeCols, options.sortedColumnWidthOverrides, options.gridMetrics.columnWidth)
  const frozenHeight = resolveRowOffset(freezeRows, options.sortedRowHeightOverrides, options.gridMetrics.rowHeight)
  return {
    scrollLeft:
      colStart <= freezeCols
        ? 0
        : Math.max(0, resolveColumnOffset(colStart, options.sortedColumnWidthOverrides, options.gridMetrics.columnWidth) - frozenWidth),
    scrollTop:
      rowStart <= freezeRows
        ? 0
        : Math.max(0, resolveRowOffset(rowStart, options.sortedRowHeightOverrides, options.gridMetrics.rowHeight) - frozenHeight),
  }
}

export function hasSelectionTargetChanged(
  previousSelection: { sheetName: string; col: number; row: number } | null,
  nextSelection: { sheetName: string; col: number; row: number },
): boolean {
  return (
    previousSelection === null ||
    previousSelection.sheetName !== nextSelection.sheetName ||
    previousSelection.col !== nextSelection.col ||
    previousSelection.row !== nextSelection.row
  )
}

function resolveVisibleColumnCount(options: {
  startCol: number
  bodyWidth: number
  columnWidths: Readonly<Record<number, number>>
  defaultWidth: number
}): number {
  const { startCol, bodyWidth, columnWidths, defaultWidth } = options
  const targetWidth = bodyWidth + MAX_COLUMN_WIDTH
  let coveredWidth = 0
  let count = 0
  for (let col = startCol; col < MAX_COLS && coveredWidth < targetWidth; col += 1) {
    coveredWidth += getResolvedColumnWidth(columnWidths, col, defaultWidth)
    count += 1
  }
  return Math.max(1, count)
}

function resolveVisibleRowCount(options: {
  startRow: number
  bodyHeight: number
  rowHeights: Readonly<Record<number, number>>
  defaultHeight: number
}): number {
  const { startRow, bodyHeight, rowHeights, defaultHeight } = options
  const targetHeight = bodyHeight + MAX_ROW_HEIGHT
  let coveredHeight = 0
  let count = 0
  for (let row = startRow; row < MAX_ROWS && coveredHeight < targetHeight; row += 1) {
    coveredHeight += getResolvedRowHeight(rowHeights, row, defaultHeight)
    count += 1
  }
  return Math.max(1, count)
}

function resolveColumnAnchor(
  scrollLeft: number,
  columnWidths: Readonly<Record<number, number>>,
  defaultWidth: number,
): { index: number; offset: number } {
  let consumed = 0
  for (let col = 0; col < MAX_COLS; col += 1) {
    const width = getResolvedColumnWidth(columnWidths, col, defaultWidth)
    if (consumed + width > scrollLeft) {
      return { index: col, offset: scrollLeft - consumed }
    }
    consumed += width
  }
  return { index: MAX_COLS - 1, offset: 0 }
}

function clampAxisStart(value: number, tileSize: number, axisMax: number): number {
  const clamped = Math.max(0, Math.min(axisMax - 1, value))
  return Math.floor(clamped / tileSize) * tileSize
}

function clampAxisEnd(value: number, tileSize: number, axisMax: number): number {
  const clamped = Math.max(0, Math.min(axisMax - 1, value))
  return Math.min(axisMax - 1, Math.ceil((clamped + 1) / tileSize) * tileSize - 1)
}

function resolveRowAnchor(
  scrollTop: number,
  rowHeights: Readonly<Record<number, number>>,
  defaultHeight: number,
): { index: number; offset: number } {
  let consumed = 0
  for (let row = 0; row < MAX_ROWS; row += 1) {
    const height = getResolvedRowHeight(rowHeights, row, defaultHeight)
    if (consumed + height > scrollTop) {
      return { index: row, offset: scrollTop - consumed }
    }
    consumed += height
  }
  return { index: MAX_ROWS - 1, offset: 0 }
}
