// @vitest-environment jsdom
import { MAX_COLS, MAX_ROWS } from '@bilig/protocol'
import { describe, expect, it, vi } from 'vitest'
import { createGridAxisWorldIndexFromRecords } from '../gridAxisWorldIndex.js'
import { getGridMetrics } from '../gridMetrics.js'
import type { VisibleRegionState } from '../gridPointer.js'
import { GridCameraStore } from '../runtime/gridCameraStore.js'
import { GridRuntimeHost } from '../runtime/gridRuntimeHost.js'
import { shouldCommitWorkbookVisibleRegion } from '../useWorkbookViewportScrollRuntime.js'
import { WorkbookGridScrollStore } from '../workbookGridScrollStore.js'
import { WorkbookViewportScrollRuntime } from '../workbookViewportScrollRuntime.js'

function region(input: {
  readonly x: number
  readonly y: number
  readonly width?: number | undefined
  readonly height?: number | undefined
  readonly freezeRows?: number | undefined
  readonly freezeCols?: number | undefined
  readonly tx?: number | undefined
  readonly ty?: number | undefined
}): VisibleRegionState {
  return {
    freezeCols: input.freezeCols ?? 0,
    freezeRows: input.freezeRows ?? 0,
    range: {
      height: input.height ?? 12,
      width: input.width ?? 12,
      x: input.x,
      y: input.y,
    },
    tx: input.tx ?? 0,
    ty: input.ty ?? 0,
  }
}

describe('shouldCommitWorkbookVisibleRegion', () => {
  it('keeps steady scroll inside the same resident window out of React state', () => {
    const current = region({ x: 0, y: 0 })
    const next = region({ x: 8, y: 8, tx: 22, ty: 11 })

    expect(
      shouldCommitWorkbookVisibleRegion({
        current,
        next,
        requiresLiveViewportState: false,
      }),
    ).toBe(false)
  })

  it('commits when the resident render window changes', () => {
    const current = region({ x: 0, y: 0 })
    const next = region({ x: 260, y: 100 })

    expect(
      shouldCommitWorkbookVisibleRegion({
        current,
        next,
        requiresLiveViewportState: false,
      }),
    ).toBe(true)
  })

  it('commits every visible window movement while a live overlay needs viewport state', () => {
    const current = region({ x: 0, y: 0 })
    const next = region({ x: 8, y: 8, tx: 22, ty: 11 })

    expect(
      shouldCommitWorkbookVisibleRegion({
        current,
        next,
        requiresLiveViewportState: true,
      }),
    ).toBe(true)
  })
})

describe('WorkbookViewportScrollRuntime', () => {
  it('syncs camera, scroll transform, and visible viewport without React owning the scroll math', () => {
    const metrics = getGridMetrics()
    const columnAxis = createGridAxisWorldIndexFromRecords({
      axisLength: MAX_COLS,
      defaultSize: metrics.columnWidth,
    })
    const rowAxis = createGridAxisWorldIndexFromRecords({
      axisLength: MAX_ROWS,
      defaultSize: metrics.rowHeight,
    })
    const gridRuntimeHost = new GridRuntimeHost({
      columnCount: MAX_COLS,
      defaultColumnWidth: metrics.columnWidth,
      defaultRowHeight: metrics.rowHeight,
      gridMetrics: metrics,
      rowCount: MAX_ROWS,
      viewportHeight: 360,
      viewportWidth: 640,
    })
    const scrollViewport = document.createElement('div')
    Object.defineProperty(scrollViewport, 'clientWidth', { configurable: true, value: 640 })
    Object.defineProperty(scrollViewport, 'clientHeight', { configurable: true, value: 360 })
    scrollViewport.scrollLeft = 10 * metrics.columnWidth
    scrollViewport.scrollTop = 8 * metrics.rowHeight
    const scrollTransformStore = new WorkbookGridScrollStore()
    const scrollTransformRef = { current: scrollTransformStore.getSnapshot() }
    const liveVisibleRegionRef = { current: region({ x: 0, y: 0 }) }
    let committedRegion = liveVisibleRegionRef.current
    const setVisibleRegion = vi.fn((updater: VisibleRegionState | ((current: VisibleRegionState) => VisibleRegionState)) => {
      committedRegion = typeof updater === 'function' ? updater(committedRegion) : updater
    })
    const onVisibleViewportChange = vi.fn()
    const runtime = new WorkbookViewportScrollRuntime()

    runtime.updateInput({
      columnAxis,
      freezeCols: 0,
      freezeRows: 0,
      gridCameraStore: new GridCameraStore(),
      gridMetrics: metrics,
      gridRuntimeHost,
      hostElement: scrollViewport,
      liveVisibleRegionRef,
      onVisibleViewportChange,
      requiresLiveViewportState: true,
      rowAxis,
      scrollTransformRef,
      scrollTransformStore,
      scrollViewportRef: { current: scrollViewport },
      selectedCell: [0, 0],
      setVisibleRegion,
      sheetName: 'Sheet1',
      sortedColumnWidthOverrides: [],
      sortedRowHeightOverrides: [],
      syncRuntimeAxes: vi.fn(),
      viewport: { colEnd: 11, colStart: 0, rowEnd: 23, rowStart: 0 },
    })

    runtime.syncVisibleRegion()

    expect(scrollTransformStore.getSnapshot()).toMatchObject({
      scrollLeft: 10 * metrics.columnWidth,
      scrollTop: 8 * metrics.rowHeight,
    })
    expect(liveVisibleRegionRef.current.range).toMatchObject({
      x: 10,
      y: 8,
    })
    expect(onVisibleViewportChange).toHaveBeenCalledWith(
      expect.objectContaining({
        colStart: 10,
        rowStart: 8,
      }),
    )
    expect(setVisibleRegion).toHaveBeenCalled()
    expect(committedRegion.range).toMatchObject({
      x: 10,
      y: 8,
    })
  })
})
