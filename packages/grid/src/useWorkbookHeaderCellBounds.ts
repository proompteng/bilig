import { useCallback } from 'react'
import { MAX_COLS, MAX_ROWS, type Viewport } from '@bilig/protocol'
import { type GridMetrics, getResolvedColumnWidth, getResolvedRowHeight, resolveRowOffset } from './gridMetrics.js'
import type { Rectangle } from './gridTypes.js'
import { resolveColumnOffset } from './workbookGridViewport.js'

type SortedAxisOverrides = readonly (readonly [number, number])[]

export function resolveWorkbookHeaderCellBounds(input: {
  readonly col: number
  readonly row: number
  readonly columnWidths: Readonly<Record<number, number>>
  readonly freezeCols: number
  readonly freezeRows: number
  readonly frozenColumnWidth: number
  readonly frozenRowHeight: number
  readonly gridMetrics: GridMetrics
  readonly residentViewport: Viewport
  readonly rowHeights: Readonly<Record<number, number>>
  readonly sortedColumnWidthOverrides: SortedAxisOverrides
  readonly sortedRowHeightOverrides: SortedAxisOverrides
}): Rectangle | undefined {
  const {
    col,
    row,
    columnWidths,
    freezeCols,
    freezeRows,
    frozenColumnWidth,
    frozenRowHeight,
    gridMetrics,
    residentViewport,
    rowHeights,
    sortedColumnWidthOverrides,
    sortedRowHeightOverrides,
  } = input
  if (col < 0 || col >= MAX_COLS || row < 0 || row >= MAX_ROWS) {
    return undefined
  }
  return {
    x:
      col < freezeCols
        ? gridMetrics.rowMarkerWidth + resolveColumnOffset(col, sortedColumnWidthOverrides, gridMetrics.columnWidth)
        : gridMetrics.rowMarkerWidth +
          frozenColumnWidth +
          resolveColumnOffset(col, sortedColumnWidthOverrides, gridMetrics.columnWidth) -
          resolveColumnOffset(residentViewport.colStart, sortedColumnWidthOverrides, gridMetrics.columnWidth),
    y:
      row < freezeRows
        ? gridMetrics.headerHeight + resolveRowOffset(row, sortedRowHeightOverrides, gridMetrics.rowHeight)
        : gridMetrics.headerHeight +
          frozenRowHeight +
          resolveRowOffset(row, sortedRowHeightOverrides, gridMetrics.rowHeight) -
          resolveRowOffset(residentViewport.rowStart, sortedRowHeightOverrides, gridMetrics.rowHeight),
    width: getResolvedColumnWidth(columnWidths, col, gridMetrics.columnWidth),
    height: getResolvedRowHeight(rowHeights, row, gridMetrics.rowHeight),
  }
}

export function useWorkbookHeaderCellBounds(input: {
  readonly columnWidths: Readonly<Record<number, number>>
  readonly freezeCols: number
  readonly freezeRows: number
  readonly frozenColumnWidth: number
  readonly frozenRowHeight: number
  readonly gridMetrics: GridMetrics
  readonly residentViewport: Viewport
  readonly rowHeights: Readonly<Record<number, number>>
  readonly sortedColumnWidthOverrides: SortedAxisOverrides
  readonly sortedRowHeightOverrides: SortedAxisOverrides
}): (col: number, row: number) => Rectangle | undefined {
  const {
    columnWidths,
    freezeCols,
    freezeRows,
    frozenColumnWidth,
    frozenRowHeight,
    gridMetrics,
    residentViewport,
    rowHeights,
    sortedColumnWidthOverrides,
    sortedRowHeightOverrides,
  } = input
  return useCallback(
    (col: number, row: number): Rectangle | undefined =>
      resolveWorkbookHeaderCellBounds({
        col,
        row,
        columnWidths,
        freezeCols,
        freezeRows,
        frozenColumnWidth,
        frozenRowHeight,
        gridMetrics,
        residentViewport,
        rowHeights,
        sortedColumnWidthOverrides,
        sortedRowHeightOverrides,
      }),
    [
      columnWidths,
      freezeCols,
      freezeRows,
      frozenColumnWidth,
      frozenRowHeight,
      gridMetrics,
      residentViewport,
      rowHeights,
      sortedColumnWidthOverrides,
      sortedRowHeightOverrides,
    ],
  )
}
