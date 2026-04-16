import { getVisibleColumnBounds, getVisibleRowBounds, type GridMetrics } from './gridMetrics.js'
import type { HeaderSelection } from './gridPointer.js'
import type { GridSelection, Item, Rectangle } from './gridTypes.js'
import { collectVisibleColumnBounds, collectVisibleRowBounds } from './visibleGridAxes.js'
import type { GridGpuColor, GridGpuRect } from './gridGpuScene.js'

export interface GridGpuHeaderPalette {
  readonly gridLineColor: GridGpuColor
  readonly headerFillColor: GridGpuColor
  readonly headerSelectedFillColor: GridGpuColor
  readonly headerHoverFillColor: GridGpuColor
  readonly headerDragAnchorFillColor: GridGpuColor
  readonly selectionFillColor: GridGpuColor
  readonly resizeGuideColor: GridGpuColor
  readonly resizeGuideGlowColor: GridGpuColor
}

export function buildGridGpuHeaderScene(options: {
  palette: GridGpuHeaderPalette
  columnWidths: Readonly<Record<number, number>>
  gridMetrics: GridMetrics
  gridSelection: GridSelection
  rowHeights: Readonly<Record<number, number>>
  activeHeaderDrag: HeaderSelection | null
  hoveredHeader: HeaderSelection | null
  resizeGuideColumn: number | null
  resizeGuideRow: number | null
  selectedCell: Item
  selectionRange: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'> | null
  visibleRegion: {
    readonly range: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'>
    readonly tx: number
    readonly ty: number
    readonly freezeRows?: number
    readonly freezeCols?: number
  }
  visibleItems: readonly Item[]
  getCellBounds: (col: number, row: number) => Rectangle | undefined
}): { fillRects: GridGpuRect[]; borderRects: GridGpuRect[] } {
  const {
    palette,
    columnWidths,
    gridMetrics,
    gridSelection,
    rowHeights,
    activeHeaderDrag,
    hoveredHeader,
    resizeGuideColumn,
    resizeGuideRow,
    selectedCell,
    selectionRange,
    visibleRegion,
    visibleItems,
    getCellBounds,
  } = options

  const fillRects: GridGpuRect[] = []
  const borderRects: GridGpuRect[] = []

  const hasFrozenAxes = (visibleRegion.freezeRows ?? 0) > 0 || (visibleRegion.freezeCols ?? 0) > 0
  const visibleColumns = hasFrozenAxes
    ? collectVisibleColumnBounds(visibleItems, getCellBounds, gridMetrics)
    : getVisibleColumnBounds(
        visibleRegion.range,
        gridMetrics.rowMarkerWidth - visibleRegion.tx,
        Number.MAX_SAFE_INTEGER,
        columnWidths,
        gridMetrics.columnWidth,
      )
  const visibleRows = hasFrozenAxes
    ? collectVisibleRowBounds(visibleItems, getCellBounds, gridMetrics)
    : getVisibleRowBounds(
        visibleRegion.range,
        gridMetrics.headerHeight - visibleRegion.ty,
        Number.MAX_SAFE_INTEGER,
        rowHeights,
        gridMetrics.rowHeight,
      )
  const selectedColumns = resolveAxisSelectionRange(
    selectionRange?.x ?? selectedCell[0],
    selectionRange ? selectionRange.x + selectionRange.width - 1 : selectedCell[0],
    gridSelection.columns,
  )
  const selectedRows = resolveAxisSelectionRange(
    selectionRange?.y ?? selectedCell[1],
    selectionRange ? selectionRange.y + selectionRange.height - 1 : selectedCell[1],
    gridSelection.rows,
  )

  if (gridSelection.columns.length > 0) {
    pushColumnSelectionBodyRects({
      fillRects,
      gridMetrics,
      selectedColumns,
      selectionFillColor: palette.selectionFillColor,
      visibleColumns,
      visibleRows,
    })
  }

  if (gridSelection.rows.length > 0) {
    pushRowSelectionBodyRects({
      fillRects,
      gridMetrics,
      selectedRows,
      selectionFillColor: palette.selectionFillColor,
      visibleRows,
      visibleWidth: visibleColumns.length === 0 ? 0 : visibleColumns.at(-1)!.right - gridMetrics.rowMarkerWidth,
    })
  }

  if (activeHeaderDrag?.kind === 'column') {
    pushColumnHeaderDragGuideRects({
      activeHeaderDrag,
      borderRects,
      fillRects,
      gridMetrics,
      resizeGuideColor: palette.resizeGuideColor,
      selectedColumns,
      visibleColumns,
      visibleRows,
    })
  }

  if (activeHeaderDrag?.kind === 'row') {
    pushRowHeaderDragGuideRects({
      activeHeaderDrag,
      borderRects,
      fillRects,
      gridMetrics,
      resizeGuideColor: palette.resizeGuideColor,
      selectedRows,
      visibleRows,
      visibleWidth: visibleColumns.length === 0 ? 0 : visibleColumns.at(-1)!.right - gridMetrics.rowMarkerWidth,
    })
  }

  fillRects.push({
    x: 0,
    y: 0,
    width: gridMetrics.rowMarkerWidth,
    height: gridMetrics.headerHeight,
    color: palette.headerFillColor,
  })
  borderRects.push(
    {
      x: 0,
      y: gridMetrics.headerHeight - 1,
      width: gridMetrics.rowMarkerWidth,
      height: 1,
      color: palette.gridLineColor,
    },
    {
      x: gridMetrics.rowMarkerWidth - 1,
      y: 0,
      width: 1,
      height: gridMetrics.headerHeight,
      color: palette.gridLineColor,
    },
  )

  for (const column of visibleColumns) {
    fillRects.push({
      x: column.left,
      y: 0,
      width: column.width,
      height: gridMetrics.headerHeight,
      color:
        column.index >= selectedColumns.start && column.index <= selectedColumns.end
          ? activeHeaderDrag?.kind === 'column' && activeHeaderDrag.index === column.index
            ? palette.headerDragAnchorFillColor
            : palette.headerSelectedFillColor
          : hoveredHeader?.kind === 'column' && hoveredHeader.index === column.index
            ? palette.headerHoverFillColor
            : palette.headerFillColor,
    })
    borderRects.push(
      {
        x: column.left + column.width - 1,
        y: 0,
        width: 1,
        height: gridMetrics.headerHeight,
        color: palette.gridLineColor,
      },
      {
        x: column.left,
        y: gridMetrics.headerHeight - 1,
        width: column.width,
        height: 1,
        color: palette.gridLineColor,
      },
    )
  }

  for (const row of visibleRows) {
    fillRects.push({
      x: 0,
      y: row.top,
      width: gridMetrics.rowMarkerWidth,
      height: row.height,
      color:
        row.index >= selectedRows.start && row.index <= selectedRows.end
          ? activeHeaderDrag?.kind === 'row' && activeHeaderDrag.index === row.index
            ? palette.headerDragAnchorFillColor
            : palette.headerSelectedFillColor
          : hoveredHeader?.kind === 'row' && hoveredHeader.index === row.index
            ? palette.headerHoverFillColor
            : palette.headerFillColor,
    })
    borderRects.push(
      {
        x: gridMetrics.rowMarkerWidth - 1,
        y: row.top,
        width: 1,
        height: row.height,
        color: palette.gridLineColor,
      },
      {
        x: 0,
        y: row.bottom - 1,
        width: gridMetrics.rowMarkerWidth,
        height: 1,
        color: palette.gridLineColor,
      },
    )
  }

  if (resizeGuideColumn !== null) {
    pushResizeGuideRects({
      borderRects,
      fillRects,
      gridMetrics,
      resizeGuideColumn,
      resizeGuideColor: palette.resizeGuideColor,
      resizeGuideGlowColor: palette.resizeGuideGlowColor,
      visibleColumns,
      visibleRows,
    })
  }

  if (resizeGuideRow !== null) {
    pushRowResizeGuideRects({
      borderRects,
      fillRects,
      gridMetrics,
      resizeGuideRow,
      resizeGuideColor: palette.resizeGuideColor,
      resizeGuideGlowColor: palette.resizeGuideGlowColor,
      visibleColumns,
      visibleRows,
    })
  }

  return { fillRects, borderRects }
}

function resolveAxisSelectionRange(
  fallbackStart: number,
  fallbackEnd: number,
  selection: GridSelection['columns'],
): { start: number; end: number } {
  const start = selection.first()
  const end = selection.last()
  if (start === undefined || end === undefined) {
    return { start: fallbackStart, end: fallbackEnd }
  }
  return { start, end }
}

function pushColumnSelectionBodyRects(options: {
  fillRects: GridGpuRect[]
  gridMetrics: GridMetrics
  selectedColumns: { start: number; end: number }
  selectionFillColor: GridGpuColor
  visibleColumns: ReadonlyArray<{
    index: number
    left: number
    right: number
    width: number
  }>
  visibleRows: ReadonlyArray<{
    index: number
    top: number
    bottom: number
    height: number
  }>
}) {
  const { fillRects, gridMetrics, selectedColumns, selectionFillColor, visibleColumns, visibleRows } = options
  const visibleSelectionColumns = visibleColumns.filter(
    (column) => column.index >= selectedColumns.start && column.index <= selectedColumns.end,
  )
  if (visibleSelectionColumns.length === 0) {
    return
  }
  const left = visibleSelectionColumns[0]!.left
  const right = visibleSelectionColumns.at(-1)!.right
  const top = gridMetrics.headerHeight
  const height = visibleRows.length === 0 ? 0 : visibleRows.at(-1)!.bottom - gridMetrics.headerHeight
  fillRects.push({
    x: left + 1,
    y: top + 1,
    width: Math.max(0, right - left - 2),
    height: Math.max(0, height - 2),
    color: selectionFillColor,
  })
}

function pushRowSelectionBodyRects(options: {
  fillRects: GridGpuRect[]
  gridMetrics: GridMetrics
  selectedRows: { start: number; end: number }
  selectionFillColor: GridGpuColor
  visibleRows: ReadonlyArray<{
    index: number
    top: number
    bottom: number
    height: number
  }>
  visibleWidth: number
}) {
  const { fillRects, gridMetrics, selectedRows, selectionFillColor, visibleRows, visibleWidth } = options
  if (visibleWidth <= 0) {
    return
  }
  const bodyLeft = gridMetrics.rowMarkerWidth
  const visibleSelectionRows = visibleRows.filter((row) => row.index >= selectedRows.start && row.index <= selectedRows.end)
  if (visibleSelectionRows.length === 0) {
    return
  }
  const startRow = visibleSelectionRows[0]!
  const endRow = visibleSelectionRows.at(-1)!
  const top = startRow.top
  const height = endRow.bottom - startRow.top
  fillRects.push({
    x: bodyLeft + 1,
    y: top + 1,
    width: Math.max(0, visibleWidth - 2),
    height: Math.max(0, height - 2),
    color: selectionFillColor,
  })
}

function pushResizeGuideRects(options: {
  borderRects: GridGpuRect[]
  fillRects: GridGpuRect[]
  gridMetrics: GridMetrics
  resizeGuideColumn: number
  resizeGuideColor: GridGpuColor
  resizeGuideGlowColor: GridGpuColor
  visibleColumns: ReadonlyArray<{
    index: number
    left: number
    right: number
    width: number
  }>
  visibleRows: ReadonlyArray<{
    index: number
    top: number
    bottom: number
    height: number
  }>
}) {
  const { borderRects, fillRects, gridMetrics, resizeGuideColumn, resizeGuideColor, resizeGuideGlowColor, visibleColumns, visibleRows } =
    options
  const column = visibleColumns.find((entry) => entry.index === resizeGuideColumn)
  if (!column) {
    return
  }
  const lineX = column.right - 1
  const totalHeight = visibleRows.length === 0 ? gridMetrics.headerHeight : visibleRows.at(-1)!.bottom
  fillRects.push({
    x: lineX - 1,
    y: 0,
    width: 3,
    height: totalHeight,
    color: resizeGuideGlowColor,
  })
  borderRects.push({
    x: lineX,
    y: 0,
    width: 1,
    height: totalHeight,
    color: resizeGuideColor,
  })
}

function pushRowResizeGuideRects(options: {
  borderRects: GridGpuRect[]
  fillRects: GridGpuRect[]
  gridMetrics: GridMetrics
  resizeGuideRow: number
  resizeGuideColor: GridGpuColor
  resizeGuideGlowColor: GridGpuColor
  visibleColumns: ReadonlyArray<{
    index: number
    left: number
    right: number
    width: number
  }>
  visibleRows: ReadonlyArray<{
    index: number
    top: number
    bottom: number
    height: number
  }>
}) {
  const { borderRects, fillRects, gridMetrics, resizeGuideRow, resizeGuideColor, resizeGuideGlowColor, visibleColumns, visibleRows } =
    options
  const row = visibleRows.find((entry) => entry.index === resizeGuideRow)
  if (!row) {
    return
  }
  const lineY = row.bottom - 1
  const totalWidth = visibleColumns.length === 0 ? gridMetrics.rowMarkerWidth : visibleColumns.at(-1)!.right
  fillRects.push({
    x: 0,
    y: lineY - 1,
    width: totalWidth,
    height: 3,
    color: resizeGuideGlowColor,
  })
  borderRects.push({
    x: 0,
    y: lineY,
    width: totalWidth,
    height: 1,
    color: resizeGuideColor,
  })
}

function pushColumnHeaderDragGuideRects(options: {
  activeHeaderDrag: HeaderSelection
  borderRects: GridGpuRect[]
  fillRects: GridGpuRect[]
  gridMetrics: GridMetrics
  resizeGuideColor: GridGpuColor
  selectedColumns: { start: number; end: number }
  visibleColumns: ReadonlyArray<{
    index: number
    left: number
    right: number
    width: number
  }>
  visibleRows: ReadonlyArray<{
    index: number
    top: number
    bottom: number
    height: number
  }>
}) {
  const { activeHeaderDrag, borderRects, fillRects, gridMetrics, resizeGuideColor, selectedColumns, visibleColumns, visibleRows } = options
  const startColumn = visibleColumns.find((entry) => entry.index === selectedColumns.start)
  const endColumn = visibleColumns.find((entry) => entry.index === selectedColumns.end)
  if (!startColumn || !endColumn) {
    return
  }
  const left = startColumn.left
  const right = endColumn.right
  const totalHeight = visibleRows.length === 0 ? gridMetrics.headerHeight : visibleRows.at(-1)!.bottom
  borderRects.push(
    {
      x: left,
      y: 0,
      width: 1,
      height: totalHeight,
      color: resizeGuideColor,
    },
    {
      x: right - 1,
      y: 0,
      width: 1,
      height: totalHeight,
      color: resizeGuideColor,
    },
  )
  const anchorColumn = visibleColumns.find((entry) => entry.index === activeHeaderDrag.index)
  if (anchorColumn) {
    fillRects.push({
      x: anchorColumn.left,
      y: gridMetrics.headerHeight - 3,
      width: anchorColumn.width,
      height: 3,
      color: resizeGuideColor,
    })
  }
}

function pushRowHeaderDragGuideRects(options: {
  activeHeaderDrag: HeaderSelection
  borderRects: GridGpuRect[]
  fillRects: GridGpuRect[]
  gridMetrics: GridMetrics
  resizeGuideColor: GridGpuColor
  selectedRows: { start: number; end: number }
  visibleRows: ReadonlyArray<{
    index: number
    top: number
    bottom: number
    height: number
  }>
  visibleWidth: number
}) {
  const { activeHeaderDrag, borderRects, fillRects, gridMetrics, resizeGuideColor, selectedRows, visibleRows, visibleWidth } = options
  if (visibleWidth <= 0) {
    return
  }
  const visibleSelectionRows = visibleRows.filter((row) => row.index >= selectedRows.start && row.index <= selectedRows.end)
  if (visibleSelectionRows.length === 0) {
    return
  }
  const topEntry = visibleSelectionRows[0]!
  const bottomEntry = visibleSelectionRows.at(-1)!
  const topRow = topEntry.index
  const bottomRow = bottomEntry.index
  const top = topEntry.top
  const bottom = bottomEntry.bottom
  const totalWidth = gridMetrics.rowMarkerWidth + visibleWidth
  borderRects.push(
    {
      x: 0,
      y: top,
      width: totalWidth,
      height: 1,
      color: resizeGuideColor,
    },
    {
      x: 0,
      y: bottom - 1,
      width: totalWidth,
      height: 1,
      color: resizeGuideColor,
    },
  )
  if (activeHeaderDrag.index >= topRow && activeHeaderDrag.index <= bottomRow) {
    const anchorEntry = visibleRows.find((row) => row.index === activeHeaderDrag.index)
    if (!anchorEntry) {
      return
    }
    fillRects.push({
      x: gridMetrics.rowMarkerWidth - 3,
      y: anchorEntry.top,
      width: 3,
      height: anchorEntry.height,
      color: resizeGuideColor,
    })
  }
}
