import { afterEach, describe, expect, it, vi } from 'vitest'
import { getGridMetrics } from '../gridMetrics.js'
import type { Rectangle } from '../gridTypes.js'
import { GridHeaderPaneRuntime, getGridHeaderPaneRuntime, type GridHeaderPaneRuntimeInput } from '../runtime/gridHeaderPaneRuntime.js'

function createInput(overrides: Partial<GridHeaderPaneRuntimeInput> = {}): GridHeaderPaneRuntimeInput {
  const gridMetrics = getGridMetrics()
  return {
    columnWidths: { 0: 80, 1: 120 },
    freezeCols: 1,
    freezeRows: 1,
    frozenColumnWidth: 80,
    frozenRowHeight: 30,
    getHeaderCellLocalBounds: (col: number, row: number): Rectangle => ({
      height: row === 0 ? 30 : 26,
      width: col === 0 ? 80 : 120,
      x: gridMetrics.rowMarkerWidth + (col === 0 ? 0 : 80),
      y: gridMetrics.headerHeight + (row === 0 ? 0 : 30),
    }),
    gridMetrics,
    hostClientHeight: 320,
    hostClientWidth: 520,
    hostReady: true,
    residentBodyPane: {
      contentOffset: { x: 33, y: 44 },
      surfaceSize: { width: 200, height: 56 },
    },
    residentHeaderItems: [
      [0, 0],
      [1, 0],
      [0, 1],
      [1, 1],
    ],
    residentHeaderRegion: {
      freezeCols: 1,
      freezeRows: 1,
      range: { height: 2, width: 2, x: 0, y: 0 },
      tx: 0,
      ty: 0,
    },
    residentViewport: { colEnd: 1, colStart: 0, rowEnd: 1, rowStart: 0 },
    rowHeights: { 0: 30, 1: 26 },
    sheetName: 'Sheet1',
    ...overrides,
  }
}

describe('GridHeaderPaneRuntime', () => {
  afterEach(() => {
    vi.unstubAllGlobals()
  })

  it('replaces stale runtime refs from live reloads', () => {
    const runtime = new GridHeaderPaneRuntime()

    expect(getGridHeaderPaneRuntime(runtime)).toBe(runtime)
    expect(getGridHeaderPaneRuntime({})).toBeInstanceOf(GridHeaderPaneRuntime)
    expect(getGridHeaderPaneRuntime(null)).toBeInstanceOf(GridHeaderPaneRuntime)
  })

  it('builds V3 header panes and applies body tile content offsets in runtime', () => {
    const panes = new GridHeaderPaneRuntime().resolve(createInput())

    expect(panes.map((pane) => pane.paneId)).toEqual(['corner-header', 'top-frozen', 'top-body', 'left-frozen', 'left-body'])
    expect(panes.find((pane) => pane.paneId === 'top-body')?.contentOffset).toEqual({ x: 33, y: 0 })
    expect(panes.find((pane) => pane.paneId === 'left-body')?.contentOffset).toEqual({ x: 0, y: 44 })
    expect(panes.find((pane) => pane.paneId === 'top-frozen')?.contentOffset).toEqual({ x: 0, y: 0 })
    expect(panes.find((pane) => pane.paneId === 'top-body')?.textRuns).toEqual([expect.objectContaining({ text: 'B' })])
    expect(panes.find((pane) => pane.paneId === 'left-body')?.textRuns).toEqual([expect.objectContaining({ text: '2' })])
  })

  it('does not build headers before the host is ready', () => {
    const panes = new GridHeaderPaneRuntime().resolve(
      createInput({
        getHeaderCellLocalBounds: () => {
          throw new Error('header bounds should not be read before host readiness')
        },
        hostReady: false,
      }),
    )

    expect(panes).toEqual([])
  })

  it('keeps header panes resident across semantically identical selection-only renders', () => {
    const noteHeaderPaneBuild = vi.fn()
    vi.stubGlobal('window', { __biligScrollPerf: { noteHeaderPaneBuild } })
    const runtime = new GridHeaderPaneRuntime()

    const first = runtime.resolve(createInput())
    const second = runtime.resolve(
      createInput({
        columnWidths: { 0: 80, 1: 120 },
        getHeaderCellLocalBounds: (col: number, row: number): Rectangle => ({
          height: row === 0 ? 30 : 26,
          width: col === 0 ? 80 : 120,
          x: getGridMetrics().rowMarkerWidth + (col === 0 ? 0 : 80),
          y: getGridMetrics().headerHeight + (row === 0 ? 0 : 30),
        }),
        residentHeaderItems: [
          [0, 0],
          [1, 0],
          [0, 1],
          [1, 1],
        ],
        rowHeights: { 0: 30, 1: 26 },
      }),
    )

    expect(second).toBe(first)
    expect(noteHeaderPaneBuild).toHaveBeenCalledTimes(1)
  })

  it('rebuilds resident header panes when rendered header geometry changes', () => {
    const noteHeaderPaneBuild = vi.fn()
    vi.stubGlobal('window', { __biligScrollPerf: { noteHeaderPaneBuild } })
    const runtime = new GridHeaderPaneRuntime()

    const first = runtime.resolve(createInput())
    const shifted = runtime.resolve(
      createInput({
        residentBodyPane: {
          contentOffset: { x: 41, y: 44 },
          surfaceSize: { height: 56, width: 200 },
        },
      }),
    )

    expect(shifted).not.toBe(first)
    expect(shifted.find((pane) => pane.paneId === 'top-body')?.contentOffset).toEqual({ x: 41, y: 0 })
    expect(noteHeaderPaneBuild).toHaveBeenCalledTimes(2)
  })
})
