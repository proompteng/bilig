import { describe, expect, it } from 'vitest'
import type { CellSnapshot } from '@bilig/protocol'
import { ValueTag } from '@bilig/protocol'
import { createGridSelection } from '../gridSelection.js'
import { getGridMetrics } from '../gridMetrics.js'
import { buildResidentDataPaneScenes, resolveResidentDataPaneRenderState } from '../gridResidentDataLayer.js'

const emptyCell: CellSnapshot = {
  sheetName: 'Sheet1',
  address: 'A1',
  value: { tag: ValueTag.Empty },
  flags: 0,
  version: 0,
}

const engine = {
  workbook: {
    getSheet: () => undefined,
  },
  getCell: () => emptyCell,
  getCellStyle: () => undefined,
}

describe('gridResidentDataLayer', () => {
  it('builds quadrant pane scenes for frozen resident windows', () => {
    const gridMetrics = getGridMetrics()
    const panes = buildResidentDataPaneScenes({
      residentViewport: { rowStart: 10, rowEnd: 25, colStart: 8, colEnd: 23 },
      engine,
      sheetName: 'Sheet1',
      columnWidths: {},
      rowHeights: {},
      freezeRows: 2,
      freezeCols: 2,
      frozenColumnWidth: 144,
      frozenRowHeight: 44,
      gridMetrics,
      sortedColumnWidthOverrides: [],
      sortedRowHeightOverrides: [],
      gridSelection: createGridSelection(8, 10),
      selectedCell: [8, 10],
      selectedCellSnapshot: emptyCell,
      selectionRange: null,
      editingCell: null,
      hoveredCell: null,
      hoveredHeader: null,
      resizeGuideColumn: null,
      resizeGuideRow: null,
      activeHeaderDrag: null,
    })

    expect(panes.map((pane) => pane.paneId)).toEqual(['body', 'top', 'left', 'corner'])
    expect(panes.every((pane) => pane.packedScene?.paneId === pane.paneId)).toBe(true)
    expect(panes[0]?.viewport).toMatchObject({
      rowStart: 10,
      rowEnd: 25,
      colStart: 8,
      colEnd: 23,
    })
    expect(renderedFrame(panes[0])).toMatchObject({
      x: gridMetrics.rowMarkerWidth + 144,
      y: gridMetrics.headerHeight + 44,
      width: 960 - gridMetrics.rowMarkerWidth - 144,
      height: 640 - gridMetrics.headerHeight - 44,
    })
    expect(panes[1]?.surfaceSize.height).toBe(44)
    expect(panes[2]?.surfaceSize.width).toBe(208)
    expect(panes[3]?.surfaceSize).toEqual({ width: 208, height: 44 })
  })

  it('leaves retained resident panes at zero offset for the live camera uniform', () => {
    const gridMetrics = getGridMetrics()
    const panes = buildResidentDataPaneScenes({
      residentViewport: { rowStart: 10, rowEnd: 25, colStart: 8, colEnd: 23 },
      engine,
      sheetName: 'Sheet1',
      columnWidths: {},
      rowHeights: {},
      freezeRows: 2,
      freezeCols: 2,
      frozenColumnWidth: 144,
      frozenRowHeight: 44,
      gridMetrics,
      sortedColumnWidthOverrides: [],
      sortedRowHeightOverrides: [],
      gridSelection: createGridSelection(8, 10),
      selectedCell: [8, 10],
      selectedCellSnapshot: emptyCell,
      selectionRange: null,
      editingCell: null,
      hoveredCell: null,
      hoveredHeader: null,
      resizeGuideColumn: null,
      resizeGuideRow: null,
      activeHeaderDrag: null,
    })

    const rendered = resolveResidentDataPaneRenderState({
      panes,
      residentViewport: { rowStart: 10, rowEnd: 25, colStart: 8, colEnd: 23 },
      visibleViewport: { rowStart: 12, rowEnd: 21, colStart: 11, colEnd: 20 },
      visibleRegion: { tx: 17, ty: 9 },
      gridMetrics,
      sortedColumnWidthOverrides: [],
      sortedRowHeightOverrides: [],
      hostWidth: 960,
      hostHeight: 640,
      rowMarkerWidth: gridMetrics.rowMarkerWidth,
      headerHeight: gridMetrics.headerHeight,
      frozenColumnWidth: 144,
      frozenRowHeight: 44,
    })

    expect(rendered.find((pane) => pane.paneId === 'body')?.contentOffset).toEqual({ x: 0, y: 0 })
    expect(rendered.find((pane) => pane.paneId === 'top')?.contentOffset).toEqual({ x: 0, y: 0 })
    expect(rendered.find((pane) => pane.paneId === 'left')?.contentOffset).toEqual({ x: 0, y: 0 })
    expect(rendered.find((pane) => pane.paneId === 'corner')?.contentOffset).toEqual({ x: 0, y: 0 })
  })

  it('keeps body scene coordinates local while retaining resident content beyond the visible pane window', () => {
    const gridMetrics = getGridMetrics()
    const panes = buildResidentDataPaneScenes({
      residentViewport: { rowStart: 0, rowEnd: 31, colStart: 0, colEnd: 127 },
      engine,
      sheetName: 'Sheet1',
      columnWidths: {},
      rowHeights: {},
      freezeRows: 0,
      freezeCols: 0,
      frozenColumnWidth: 0,
      frozenRowHeight: 0,
      gridMetrics,
      sortedColumnWidthOverrides: [],
      sortedRowHeightOverrides: [],
      gridSelection: createGridSelection(1, 1),
      selectedCell: [1, 1],
      selectedCellSnapshot: emptyCell,
      selectionRange: { x: 1, y: 1, width: 2, height: 2 },
      editingCell: null,
      hoveredCell: null,
      hoveredHeader: null,
      resizeGuideColumn: null,
      resizeGuideRow: null,
      activeHeaderDrag: null,
    })

    const rendered = resolveResidentDataPaneRenderState({
      panes,
      residentViewport: { rowStart: 0, rowEnd: 31, colStart: 0, colEnd: 127 },
      visibleViewport: { rowStart: 0, rowEnd: 25, colStart: 0, colEnd: 5 },
      visibleRegion: { tx: 0, ty: 0 },
      gridMetrics,
      sortedColumnWidthOverrides: [],
      sortedRowHeightOverrides: [],
      hostWidth: 640,
      hostHeight: 586,
      rowMarkerWidth: gridMetrics.rowMarkerWidth,
      headerHeight: gridMetrics.headerHeight,
      frozenColumnWidth: 0,
      frozenRowHeight: 0,
    })

    const body = rendered.find((pane) => pane.paneId === 'body')
    expect(body).toBeDefined()
    expect(body?.gpuScene.fillRects).toContainEqual({
      x: 105,
      y: 23,
      width: 206,
      height: 42,
      color: {
        r: 0.12941176470588237,
        g: 0.33725490196078434,
        b: 0.22745098039215686,
        a: 0.08,
      },
    })
    expect(body?.gpuScene.borderRects.some((rect) => rect.x >= body.frame.width || rect.y >= body.frame.height)).toBe(true)
  })
})

function renderedFrame(pane: { paneId: string } | undefined) {
  if (!pane) {
    throw new Error('unexpected missing pane')
  }
  switch (pane.paneId) {
    case 'body':
      return { x: 190, y: 68, width: 770, height: 572 }
    case 'top':
      return { x: 190, y: 24, width: 770, height: 44 }
    case 'left':
      return { x: 46, y: 68, width: 144, height: 572 }
    case 'corner':
      return { x: 46, y: 24, width: 144, height: 44 }
  }
  throw new Error('unexpected pane')
}
