import type { CellSnapshot, Viewport } from '@bilig/protocol'
import type { GridEngineLike } from './grid-engine.js'
import { buildGridGpuScene, type GridGpuScene } from './gridGpuScene.js'
import { getResolvedColumnWidth, getResolvedRowHeight, resolveRowOffset, type GridMetrics } from './gridMetrics.js'
import { collectViewportItems } from './gridViewportItems.js'
import { buildGridTextScene, type GridTextScene } from './gridTextScene.js'
import type { HeaderSelection } from './gridPointer.js'
import type { GridSelection, Item, Rectangle } from './gridTypes.js'
import { resolveColumnOffset } from './workbookGridViewport.js'

export type ResidentDataPaneId = 'corner' | 'top' | 'left' | 'body'

export interface ResidentDataPaneScene {
  readonly id: ResidentDataPaneId
  readonly frame: Rectangle
  readonly surfaceSize: {
    readonly width: number
    readonly height: number
  }
  readonly gpuScene: GridGpuScene
  readonly textScene: GridTextScene
}

export interface ResidentDataPaneRenderState extends ResidentDataPaneScene {
  readonly contentOffset: {
    readonly x: number
    readonly y: number
  }
}

function resolveViewportWidth(
  viewport: Viewport,
  gridMetrics: GridMetrics,
  sortedColumnWidthOverrides: readonly (readonly [number, number])[],
): number {
  return (
    resolveColumnOffset(viewport.colEnd + 1, sortedColumnWidthOverrides, gridMetrics.columnWidth) -
    resolveColumnOffset(viewport.colStart, sortedColumnWidthOverrides, gridMetrics.columnWidth)
  )
}

function resolveViewportHeight(
  viewport: Viewport,
  gridMetrics: GridMetrics,
  sortedRowHeightOverrides: readonly (readonly [number, number])[],
): number {
  return (
    resolveRowOffset(viewport.rowEnd + 1, sortedRowHeightOverrides, gridMetrics.rowHeight) -
    resolveRowOffset(viewport.rowStart, sortedRowHeightOverrides, gridMetrics.rowHeight)
  )
}

function createPaneCellBoundsResolver(input: {
  viewport: Viewport
  columnWidths: Readonly<Record<number, number>>
  rowHeights: Readonly<Record<number, number>>
  gridMetrics: GridMetrics
  sortedColumnWidthOverrides: readonly (readonly [number, number])[]
  sortedRowHeightOverrides: readonly (readonly [number, number])[]
}): (col: number, row: number) => Rectangle | undefined {
  const { viewport, columnWidths, rowHeights, gridMetrics, sortedColumnWidthOverrides, sortedRowHeightOverrides } = input
  const baseX = resolveColumnOffset(viewport.colStart, sortedColumnWidthOverrides, gridMetrics.columnWidth)
  const baseY = resolveRowOffset(viewport.rowStart, sortedRowHeightOverrides, gridMetrics.rowHeight)
  return (col: number, row: number) => {
    if (col < viewport.colStart || col > viewport.colEnd || row < viewport.rowStart || row > viewport.rowEnd) {
      return undefined
    }
    return {
      x: resolveColumnOffset(col, sortedColumnWidthOverrides, gridMetrics.columnWidth) - baseX,
      y: resolveRowOffset(row, sortedRowHeightOverrides, gridMetrics.rowHeight) - baseY,
      width: getResolvedColumnWidth(columnWidths, col, gridMetrics.columnWidth),
      height: getResolvedRowHeight(rowHeights, row, gridMetrics.rowHeight),
    }
  }
}

function buildPaneScene(input: {
  id: ResidentDataPaneId
  viewport: Viewport
  frame: Rectangle
  engine: GridEngineLike
  sheetName: string
  columnWidths: Readonly<Record<number, number>>
  rowHeights: Readonly<Record<number, number>>
  gridMetrics: GridMetrics
  sortedColumnWidthOverrides: readonly (readonly [number, number])[]
  sortedRowHeightOverrides: readonly (readonly [number, number])[]
  gridSelection: GridSelection
  selectedCell: Item
  selectedCellSnapshot: CellSnapshot
  selectionRange?: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'> | null
  editingCell?: Item | null
  hoveredCell?: Item | null
  hoveredHeader?: HeaderSelection | null
  resizeGuideColumn?: number | null
  resizeGuideRow?: number | null
  activeHeaderDrag?: HeaderSelection | null
}): ResidentDataPaneScene {
  const {
    id,
    viewport,
    frame,
    engine,
    sheetName,
    columnWidths,
    rowHeights,
    gridMetrics,
    sortedColumnWidthOverrides,
    sortedRowHeightOverrides,
    gridSelection,
    selectedCell,
    selectedCellSnapshot,
    selectionRange = null,
    editingCell = null,
    hoveredCell = null,
    hoveredHeader = null,
    resizeGuideColumn = null,
    resizeGuideRow = null,
    activeHeaderDrag = null,
  } = input
  const getCellBounds = createPaneCellBoundsResolver({
    viewport,
    columnWidths,
    rowHeights,
    gridMetrics,
    sortedColumnWidthOverrides,
    sortedRowHeightOverrides,
  })
  const visibleItems = collectViewportItems(viewport)
  const surfaceSize = {
    width: resolveViewportWidth(viewport, gridMetrics, sortedColumnWidthOverrides),
    height: resolveViewportHeight(viewport, gridMetrics, sortedRowHeightOverrides),
  }
  return {
    id,
    frame,
    surfaceSize,
    gpuScene: buildGridGpuScene({
      contentMode: 'data',
      engine,
      sheetName,
      visibleItems,
      visibleRegion: {
        range: {
          x: viewport.colStart,
          y: viewport.rowStart,
          width: viewport.colEnd - viewport.colStart + 1,
          height: viewport.rowEnd - viewport.rowStart + 1,
        },
        tx: 0,
        ty: 0,
        freezeRows: 0,
        freezeCols: 0,
      },
      gridMetrics,
      columnWidths,
      rowHeights,
      hostBounds: { left: 0, top: 0 },
      getCellBounds,
      gridSelection,
      selectedCell,
      selectionRange,
      hoveredCell,
      hoveredHeader,
      resizeGuideColumn,
      resizeGuideRow,
      activeHeaderDrag,
    }),
    textScene: buildGridTextScene({
      contentMode: 'data',
      engine,
      sheetName,
      visibleItems,
      visibleRegion: {
        range: {
          x: viewport.colStart,
          y: viewport.rowStart,
          width: viewport.colEnd - viewport.colStart + 1,
          height: viewport.rowEnd - viewport.rowStart + 1,
        },
        tx: 0,
        ty: 0,
        freezeRows: 0,
        freezeCols: 0,
      },
      gridMetrics,
      columnWidths,
      rowHeights,
      editingCell,
      selectedCell,
      selectedCellSnapshot,
      selectionRange,
      hoveredHeader,
      activeHeaderDrag,
      resizeGuideColumn,
      hostBounds: {
        left: 0,
        top: 0,
        width: surfaceSize.width,
        height: surfaceSize.height,
      },
      getCellBounds,
    }),
  }
}

export function buildResidentDataPaneScenes(input: {
  residentViewport: Viewport
  hostWidth: number
  hostHeight: number
  engine: GridEngineLike
  sheetName: string
  columnWidths: Readonly<Record<number, number>>
  rowHeights: Readonly<Record<number, number>>
  freezeRows: number
  freezeCols: number
  frozenColumnWidth: number
  frozenRowHeight: number
  gridMetrics: GridMetrics
  sortedColumnWidthOverrides: readonly (readonly [number, number])[]
  sortedRowHeightOverrides: readonly (readonly [number, number])[]
  gridSelection: GridSelection
  selectedCell: Item
  selectedCellSnapshot: CellSnapshot
  selectionRange?: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'> | null
  editingCell?: Item | null
  hoveredCell?: Item | null
  hoveredHeader?: HeaderSelection | null
  resizeGuideColumn?: number | null
  resizeGuideRow?: number | null
  activeHeaderDrag?: HeaderSelection | null
}): ResidentDataPaneScene[] {
  const {
    residentViewport,
    hostWidth,
    hostHeight,
    engine,
    sheetName,
    columnWidths,
    rowHeights,
    freezeRows,
    freezeCols,
    frozenColumnWidth,
    frozenRowHeight,
    gridMetrics,
    sortedColumnWidthOverrides,
    sortedRowHeightOverrides,
    gridSelection,
    selectedCell,
    selectedCellSnapshot,
    selectionRange = null,
    editingCell = null,
    hoveredCell = null,
    hoveredHeader = null,
    resizeGuideColumn = null,
    resizeGuideRow = null,
    activeHeaderDrag = null,
  } = input

  const bodyFrameWidth = Math.max(0, hostWidth - gridMetrics.rowMarkerWidth - frozenColumnWidth)
  const bodyFrameHeight = Math.max(0, hostHeight - gridMetrics.headerHeight - frozenRowHeight)
  const bodyViewportWidth = resolveViewportWidth(residentViewport, gridMetrics, sortedColumnWidthOverrides)
  const bodyViewportHeight = resolveViewportHeight(residentViewport, gridMetrics, sortedRowHeightOverrides)

  const panes: ResidentDataPaneScene[] = []

  if (bodyFrameWidth > 0 && bodyFrameHeight > 0 && bodyViewportWidth > 0 && bodyViewportHeight > 0) {
    panes.push(
      buildPaneScene({
        id: 'body',
        viewport: residentViewport,
        frame: {
          x: gridMetrics.rowMarkerWidth + frozenColumnWidth,
          y: gridMetrics.headerHeight + frozenRowHeight,
          width: bodyFrameWidth,
          height: bodyFrameHeight,
        },
        engine,
        sheetName,
        columnWidths,
        rowHeights,
        gridMetrics,
        sortedColumnWidthOverrides,
        sortedRowHeightOverrides,
        gridSelection,
        selectedCell,
        selectedCellSnapshot,
        selectionRange,
        editingCell,
        hoveredCell,
        hoveredHeader,
        resizeGuideColumn,
        resizeGuideRow,
        activeHeaderDrag,
      }),
    )
  }

  if (freezeRows > 0 && frozenRowHeight > 0 && bodyFrameWidth > 0 && bodyViewportWidth > 0) {
    panes.push(
      buildPaneScene({
        id: 'top',
        viewport: {
          rowStart: 0,
          rowEnd: Math.max(0, freezeRows - 1),
          colStart: residentViewport.colStart,
          colEnd: residentViewport.colEnd,
        },
        frame: {
          x: gridMetrics.rowMarkerWidth + frozenColumnWidth,
          y: gridMetrics.headerHeight,
          width: bodyFrameWidth,
          height: frozenRowHeight,
        },
        engine,
        sheetName,
        columnWidths,
        rowHeights,
        gridMetrics,
        sortedColumnWidthOverrides,
        sortedRowHeightOverrides,
        gridSelection,
        selectedCell,
        selectedCellSnapshot,
        selectionRange,
        editingCell,
        hoveredCell,
        hoveredHeader,
        resizeGuideColumn,
        resizeGuideRow,
        activeHeaderDrag,
      }),
    )
  }

  if (freezeCols > 0 && frozenColumnWidth > 0 && bodyFrameHeight > 0 && bodyViewportHeight > 0) {
    panes.push(
      buildPaneScene({
        id: 'left',
        viewport: {
          rowStart: residentViewport.rowStart,
          rowEnd: residentViewport.rowEnd,
          colStart: 0,
          colEnd: Math.max(0, freezeCols - 1),
        },
        frame: {
          x: gridMetrics.rowMarkerWidth,
          y: gridMetrics.headerHeight + frozenRowHeight,
          width: frozenColumnWidth,
          height: bodyFrameHeight,
        },
        engine,
        sheetName,
        columnWidths,
        rowHeights,
        gridMetrics,
        sortedColumnWidthOverrides,
        sortedRowHeightOverrides,
        gridSelection,
        selectedCell,
        selectedCellSnapshot,
        selectionRange,
        editingCell,
        hoveredCell,
        hoveredHeader,
        resizeGuideColumn,
        resizeGuideRow,
        activeHeaderDrag,
      }),
    )
  }

  if (freezeRows > 0 && freezeCols > 0 && frozenColumnWidth > 0 && frozenRowHeight > 0) {
    panes.push(
      buildPaneScene({
        id: 'corner',
        viewport: {
          rowStart: 0,
          rowEnd: Math.max(0, freezeRows - 1),
          colStart: 0,
          colEnd: Math.max(0, freezeCols - 1),
        },
        frame: {
          x: gridMetrics.rowMarkerWidth,
          y: gridMetrics.headerHeight,
          width: frozenColumnWidth,
          height: frozenRowHeight,
        },
        engine,
        sheetName,
        columnWidths,
        rowHeights,
        gridMetrics,
        sortedColumnWidthOverrides,
        sortedRowHeightOverrides,
        gridSelection,
        selectedCell,
        selectedCellSnapshot,
        selectionRange,
        editingCell,
        hoveredCell,
        hoveredHeader,
        resizeGuideColumn,
        resizeGuideRow,
        activeHeaderDrag,
      }),
    )
  }

  return panes
}

export function resolveResidentDataPaneRenderState(input: {
  panes: readonly ResidentDataPaneScene[]
  residentViewport: Viewport
  visibleViewport: Viewport
  visibleRegion: {
    readonly tx: number
    readonly ty: number
  }
  gridMetrics: GridMetrics
  sortedColumnWidthOverrides: readonly (readonly [number, number])[]
  sortedRowHeightOverrides: readonly (readonly [number, number])[]
}): ResidentDataPaneRenderState[] {
  const { panes, residentViewport, visibleViewport, visibleRegion, gridMetrics, sortedColumnWidthOverrides, sortedRowHeightOverrides } =
    input
  const bodyOffsetX = -(
    resolveColumnOffset(visibleViewport.colStart, sortedColumnWidthOverrides, gridMetrics.columnWidth) -
    resolveColumnOffset(residentViewport.colStart, sortedColumnWidthOverrides, gridMetrics.columnWidth) +
    visibleRegion.tx
  )
  const bodyOffsetY = -(
    resolveRowOffset(visibleViewport.rowStart, sortedRowHeightOverrides, gridMetrics.rowHeight) -
    resolveRowOffset(residentViewport.rowStart, sortedRowHeightOverrides, gridMetrics.rowHeight) +
    visibleRegion.ty
  )
  return panes.map((pane) => ({
    id: pane.id,
    frame: pane.frame,
    surfaceSize: pane.surfaceSize,
    gpuScene: pane.gpuScene,
    textScene: pane.textScene,
    contentOffset:
      pane.id === 'body'
        ? { x: bodyOffsetX, y: bodyOffsetY }
        : pane.id === 'top'
          ? { x: bodyOffsetX, y: 0 }
          : pane.id === 'left'
            ? { x: 0, y: bodyOffsetY }
            : { x: 0, y: 0 },
  }))
}
