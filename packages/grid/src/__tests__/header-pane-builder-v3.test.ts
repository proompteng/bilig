import { describe, expect, it } from 'vitest'
import { getGridMetrics } from '../gridMetrics.js'
import { buildWorkbookHeaderPaneStatesV3, buildWorkbookHeaderTextSceneV3 } from '../renderer-v3/header-pane-builder.js'

describe('renderer-v3 header pane builder', () => {
  it('builds fixed header labels and panes without engine-backed grid scenes', () => {
    const gridMetrics = getGridMetrics()
    const residentHeaderItems = [
      [0, 0],
      [1, 0],
      [0, 1],
      [1, 1],
    ] as const
    const input = {
      columnWidths: { 0: 80, 1: 120 },
      freezeCols: 1,
      freezeRows: 1,
      frozenColumnWidth: 80,
      frozenRowHeight: 30,
      getHeaderCellLocalBounds: (col: number, row: number) => ({
        x: gridMetrics.rowMarkerWidth + (col === 0 ? 0 : 80),
        y: gridMetrics.headerHeight + (row === 0 ? 0 : 30),
        width: col === 0 ? 80 : 120,
        height: row === 0 ? 30 : 26,
      }),
      gridMetrics,
      hostClientHeight: 320,
      hostClientWidth: 520,
      residentBodyHeight: 56,
      residentBodyWidth: 200,
      residentHeaderItems,
      residentHeaderRegion: {
        freezeCols: 1,
        freezeRows: 1,
        range: { x: 0, y: 0, width: 2, height: 2 },
        tx: 0,
        ty: 0,
      },
      residentViewport: { rowStart: 0, rowEnd: 1, colStart: 0, colEnd: 1 },
      rowHeights: { 0: 30, 1: 26 },
      sheetName: 'Sheet1',
    }

    expect(buildWorkbookHeaderTextSceneV3(input).items.map((item) => item.text)).toEqual(['A', 'B', '1', '2'])

    const panes = buildWorkbookHeaderPaneStatesV3(input)

    expect(panes.map((pane) => pane.paneId)).toEqual(['corner-header', 'top-frozen', 'top-body', 'left-frozen', 'left-body'])
    expect(panes.find((pane) => pane.paneId === 'top-frozen')?.textRuns).toEqual([expect.objectContaining({ text: 'A' })])
    expect(panes.find((pane) => pane.paneId === 'top-body')?.textRuns).toEqual([expect.objectContaining({ text: 'B' })])
    expect(panes.find((pane) => pane.paneId === 'left-frozen')?.textRuns).toEqual([expect.objectContaining({ text: '1' })])
    expect(panes.find((pane) => pane.paneId === 'left-body')?.textRuns).toEqual([expect.objectContaining({ text: '2' })])
    expect(panes.every((pane) => pane.rectCount > 0)).toBe(true)
  })
})
