import type { CellSnapshot, Viewport } from '@bilig/protocol'
import type { GridEngineLike } from './grid-engine.js'
import { buildGridGpuScene } from './gridGpuScene.js'
import { getResolvedColumnWidth, getResolvedRowHeight, resolveRowOffset, type GridMetrics } from './gridMetrics.js'
import { collectViewportItems } from './gridViewportItems.js'
import { buildGridTextScene } from './gridTextScene.js'
import type { HeaderSelection } from './gridPointer.js'
import type { GridSelection, Item, Rectangle } from './gridTypes.js'
import { resolveColumnOffset } from './workbookGridViewport.js'
import type { WorkbookPaneId, WorkbookPaneRenderState, WorkbookPaneScenePacket } from './renderer/pane-scene-types.js'
import { getPaneFrame, resolvePaneLayout } from './renderer/pane-layout.js'
import { packGridScenePacketV2 } from './renderer-v2/scene-packet-v2.js'

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
  id: WorkbookPaneId
  viewport: Viewport
  engine: GridEngineLike
  sheetName: string
  columnWidths: Readonly<Record<number, number>>
  rowHeights: Readonly<Record<number, number>>
  gridMetrics: GridMetrics
  sortedColumnWidthOverrides: readonly (readonly [number, number])[]
  sortedRowHeightOverrides: readonly (readonly [number, number])[]
  gridSelection: GridSelection
  selectedCell: Item
  selectedCellSnapshot: CellSnapshot | null
  selectionRange?: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'> | null
  editingCell?: Item | null
  hoveredCell?: Item | null
  hoveredHeader?: HeaderSelection | null
  resizeGuideColumn?: number | null
  resizeGuideRow?: number | null
  activeHeaderDrag?: HeaderSelection | null
}): WorkbookPaneScenePacket {
  const {
    id,
    viewport,
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
  const gpuScene = buildGridGpuScene({
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
  })
  const textScene = buildGridTextScene({
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
  })
  return {
    generation: 0,
    paneId: id,
    viewport,
    surfaceSize,
    gpuScene,
    packedScene: packGridScenePacketV2({
      generation: 0,
      gpuScene,
      paneId: id,
      sheetName,
      surfaceSize,
      textScene,
      viewport,
    }),
    textScene,
  }
}

export function buildResidentDataPaneScenes(input: {
  residentViewport: Viewport
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
  selectedCellSnapshot: CellSnapshot | null
  selectionRange?: Pick<Rectangle, 'x' | 'y' | 'width' | 'height'> | null
  editingCell?: Item | null
  hoveredCell?: Item | null
  hoveredHeader?: HeaderSelection | null
  resizeGuideColumn?: number | null
  resizeGuideRow?: number | null
  activeHeaderDrag?: HeaderSelection | null
}): WorkbookPaneScenePacket[] {
  const {
    residentViewport,
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

  const bodyViewportWidth = resolveViewportWidth(residentViewport, gridMetrics, sortedColumnWidthOverrides)
  const bodyViewportHeight = resolveViewportHeight(residentViewport, gridMetrics, sortedRowHeightOverrides)

  const panes: WorkbookPaneScenePacket[] = []
  if (bodyViewportWidth > 0 && bodyViewportHeight > 0) {
    panes.push(
      buildPaneScene({
        id: 'body',
        viewport: residentViewport,
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

  if (freezeRows > 0 && frozenRowHeight > 0 && bodyViewportWidth > 0) {
    panes.push(
      buildPaneScene({
        id: 'top',
        viewport: {
          rowStart: 0,
          rowEnd: Math.max(0, freezeRows - 1),
          colStart: residentViewport.colStart,
          colEnd: residentViewport.colEnd,
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

  if (freezeCols > 0 && frozenColumnWidth > 0 && bodyViewportHeight > 0) {
    panes.push(
      buildPaneScene({
        id: 'left',
        viewport: {
          rowStart: residentViewport.rowStart,
          rowEnd: residentViewport.rowEnd,
          colStart: 0,
          colEnd: Math.max(0, freezeCols - 1),
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
  panes: readonly WorkbookPaneScenePacket[]
  residentViewport: Viewport
  visibleViewport: Viewport
  visibleRegion: {
    readonly tx: number
    readonly ty: number
  }
  gridMetrics: GridMetrics
  sortedColumnWidthOverrides: readonly (readonly [number, number])[]
  sortedRowHeightOverrides: readonly (readonly [number, number])[]
  hostWidth: number
  hostHeight: number
  rowMarkerWidth: number
  headerHeight: number
  frozenColumnWidth: number
  frozenRowHeight: number
}): WorkbookPaneRenderState[] {
  const layout = resolvePaneLayout({
    hostWidth: input.hostWidth,
    hostHeight: input.hostHeight,
    rowMarkerWidth: input.rowMarkerWidth,
    headerHeight: input.headerHeight,
    frozenColumnWidth: input.frozenColumnWidth,
    frozenRowHeight: input.frozenRowHeight,
  })
  return input.panes.map((pane) => {
    const frame = getPaneFrame(layout, pane.paneId)
    const next: WorkbookPaneRenderState = {
      generation: pane.generation,
      paneId: pane.paneId,
      viewport: pane.viewport,
      frame,
      surfaceSize: pane.surfaceSize,
      gpuScene: pane.gpuScene,
      textScene: pane.textScene,
      contentOffset: { x: 0, y: 0 },
      packedScene: pane.packedScene,
    }
    return next
  })
}
