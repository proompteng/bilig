import { describe, expect, it } from 'vitest'
import { buildHeaderPaneStates } from '../gridHeaderPanes.js'
import { getGridMetrics } from '../gridMetrics.js'

describe('buildHeaderPaneStates', () => {
  it('splits frozen and scrolling header panes with the correct scroll axes', () => {
    const gridMetrics = getGridMetrics()
    const panes = buildHeaderPaneStates({
      gpuScene: {
        fillRects: [
          { x: gridMetrics.rowMarkerWidth, y: 0, width: 80, height: gridMetrics.headerHeight, color: { r: 0, g: 0, b: 0, a: 1 } },
          { x: gridMetrics.rowMarkerWidth + 80, y: 0, width: 200, height: gridMetrics.headerHeight, color: { r: 0, g: 0, b: 0, a: 1 } },
          { x: 0, y: gridMetrics.headerHeight, width: gridMetrics.rowMarkerWidth, height: 44, color: { r: 0, g: 0, b: 0, a: 1 } },
          { x: 0, y: gridMetrics.headerHeight + 44, width: gridMetrics.rowMarkerWidth, height: 120, color: { r: 0, g: 0, b: 0, a: 1 } },
        ],
        borderRects: [],
      },
      textScene: {
        items: [
          {
            x: gridMetrics.rowMarkerWidth + 16,
            y: 0,
            width: 48,
            height: gridMetrics.headerHeight,
            clipInsetTop: 0,
            clipInsetRight: 0,
            clipInsetBottom: 0,
            clipInsetLeft: 0,
            text: 'A',
            align: 'center',
            wrap: false,
            color: '#000',
            font: '12px sans-serif',
            fontSize: 12,
            underline: false,
            strike: false,
          },
          {
            x: 8,
            y: gridMetrics.headerHeight + 60,
            width: 30,
            height: 20,
            clipInsetTop: 0,
            clipInsetRight: 0,
            clipInsetBottom: 0,
            clipInsetLeft: 0,
            text: '9',
            align: 'right',
            wrap: false,
            color: '#000',
            font: '12px sans-serif',
            fontSize: 12,
            underline: false,
            strike: false,
          },
        ],
      },
      hostWidth: 720,
      hostHeight: 480,
      gridMetrics,
      frozenColumnWidth: 80,
      frozenRowHeight: 44,
      residentBodyWidth: 200,
      residentBodyHeight: 120,
    })

    expect(panes.map((pane) => pane.paneId)).toEqual(['top-frozen', 'top-body', 'left-frozen', 'left-body'])
    expect(panes.map((pane) => pane.scrollAxes)).toEqual([
      { x: false, y: false },
      { x: true, y: false },
      { x: false, y: false },
      { x: false, y: true },
    ])
    expect(panes.find((pane) => pane.paneId === 'top-frozen')?.surfaceSize.width).toBe(80)
    expect(panes.find((pane) => pane.paneId === 'top-body')?.surfaceSize.width).toBe(200)
    expect(panes.find((pane) => pane.paneId === 'left-frozen')?.surfaceSize.height).toBe(44)
    expect(panes.find((pane) => pane.paneId === 'left-body')?.surfaceSize.height).toBe(120)
    expect(panes.find((pane) => pane.paneId === 'top-body')?.contentOffset).toEqual({ x: 0, y: 0 })
    expect(panes.find((pane) => pane.paneId === 'left-body')?.contentOffset).toEqual({ x: 0, y: 0 })
  })
})
