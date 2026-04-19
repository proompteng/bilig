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
      hostWidth: 960,
      hostHeight: 640,
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

    expect(panes.map((pane) => pane.id)).toEqual(['body', 'top', 'left', 'corner'])
    expect(panes[0]?.frame).toMatchObject({
      x: gridMetrics.rowMarkerWidth + 144,
      y: gridMetrics.headerHeight + 44,
      width: 960 - gridMetrics.rowMarkerWidth - 144,
      height: 640 - gridMetrics.headerHeight - 44,
    })
    expect(panes[1]?.surfaceSize.height).toBe(44)
    expect(panes[2]?.surfaceSize.width).toBe(208)
    expect(panes[3]?.surfaceSize).toEqual({ width: 208, height: 44 })
  })

  it('derives pane offsets from the visible viewport inside a resident window', () => {
    const gridMetrics = getGridMetrics()
    const panes = buildResidentDataPaneScenes({
      residentViewport: { rowStart: 10, rowEnd: 25, colStart: 8, colEnd: 23 },
      hostWidth: 960,
      hostHeight: 640,
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
    })

    expect(rendered.find((pane) => pane.id === 'body')?.contentOffset).toEqual({ x: -(3 * 104 + 17), y: -(2 * 22 + 9) })
    expect(rendered.find((pane) => pane.id === 'top')?.contentOffset).toEqual({ x: -(3 * 104 + 17), y: 0 })
    expect(rendered.find((pane) => pane.id === 'left')?.contentOffset).toEqual({ x: 0, y: -(2 * 22 + 9) })
    expect(rendered.find((pane) => pane.id === 'corner')?.contentOffset).toEqual({ x: 0, y: 0 })
  })
})
