import { describe, expect, it, vi } from 'vitest'
import { ValueTag, type RecalcMetrics } from '@bilig/protocol'
import { encodeViewportPatch, type ViewportPatch } from '@bilig/worker-transport'
import { ProjectedViewportAxisStore } from '../projected-viewport-axis-store.js'
import { ProjectedViewportCellCache } from '../projected-viewport-cell-cache.js'
import { ProjectedViewportPatchCoordinator } from '../projected-viewport-patch-coordinator.js'

const TEST_METRICS: RecalcMetrics = {
  batchId: 0,
  changedInputCount: 0,
  dirtyFormulaCount: 0,
  wasmFormulaCount: 0,
  jsFormulaCount: 0,
  rangeNodeVisits: 0,
  recalcMs: 0,
  compileMs: 0,
}

function createPatch(styleId = 'style-0'): ViewportPatch {
  return {
    version: 1,
    full: false,
    freezeRows: 1,
    freezeCols: 2,
    viewport: {
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 2,
      colStart: 0,
      colEnd: 2,
    },
    metrics: TEST_METRICS,
    styles: [],
    cells: [
      {
        row: 0,
        col: 0,
        snapshot: {
          sheetName: 'Sheet1',
          address: 'A1',
          value: { tag: ValueTag.Number, value: 42 },
          flags: 0,
          version: 1,
          styleId,
        },
        displayText: '42',
        copyText: '42',
        editorText: '42',
        formatId: 0,
        styleId,
      },
    ],
    columns: [{ index: 0, size: 93, hidden: false }],
    rows: [{ index: 0, size: 44, hidden: false }],
  }
}

describe('ProjectedViewportPatchCoordinator', () => {
  it('applies viewport patches against cell and axis state', () => {
    const cellCache = new ProjectedViewportCellCache()
    const axisStore = new ProjectedViewportAxisStore()
    const coordinator = new ProjectedViewportPatchCoordinator({
      cellCache,
      axisStore,
    })

    const damage = coordinator.applyViewportPatch(createPatch())

    expect(damage).toEqual([{ cell: [0, 0] }])
    expect(cellCache.getCell('Sheet1', 'A1').value).toEqual({
      tag: ValueTag.Number,
      value: 42,
    })
    expect(axisStore.getColumnWidths('Sheet1')[0]).toBe(93)
    expect(axisStore.getRowHeights('Sheet1')[0]).toBe(44)
    expect(axisStore.getFreezeRows('Sheet1')).toBe(1)
    expect(axisStore.getFreezeCols('Sheet1')).toBe(2)
  })

  it('tracks viewport subscriptions through the worker client', async () => {
    vi.useFakeTimers()
    const encodedPatch = encodeViewportPatch(createPatch())
    const subscribeViewportPatches = vi.fn((_viewport, listener: (bytes: Uint8Array) => void) => {
      listener(encodedPatch)
      return () => undefined
    })
    const cellCache = new ProjectedViewportCellCache()
    const axisStore = new ProjectedViewportAxisStore()
    const coordinator = new ProjectedViewportPatchCoordinator({
      client: {
        invoke: async () => undefined,
        ready: async () => undefined,
        subscribe: () => () => undefined,
        subscribeBatches: () => () => undefined,
        subscribeViewportPatches,
        dispose: () => undefined,
      },
      cellCache,
      axisStore,
    })
    const listener = vi.fn()

    const unsubscribe = coordinator.subscribeViewport('Sheet1', { rowStart: 0, rowEnd: 0, colStart: 0, colEnd: 0 }, listener)

    expect(subscribeViewportPatches).toHaveBeenCalledWith(
      { sheetName: 'Sheet1', rowStart: 0, rowEnd: 0, colStart: 0, colEnd: 0 },
      expect.any(Function),
    )
    await vi.runAllTimersAsync()
    expect(listener).toHaveBeenCalledWith([{ cell: [0, 0] }])

    unsubscribe()
    vi.useRealTimers()
  })

  it('ignores stale out-of-order viewport patches for a subscription', async () => {
    vi.useFakeTimers()
    const hiddenRowPatch: ViewportPatch = {
      ...createPatch(),
      version: 2,
      cells: [],
      columns: [],
      rows: [{ index: 1, size: 22, hidden: true }],
    }
    const staleVisibleRowPatch: ViewportPatch = {
      ...createPatch(),
      version: 1,
      cells: [],
      columns: [],
      rows: [{ index: 1, size: 22, hidden: false }],
    }
    const patches: ((bytes: Uint8Array) => void)[] = []
    const subscribeViewportPatches = vi.fn((_viewport, listener: (bytes: Uint8Array) => void) => {
      patches.push(listener)
      return () => undefined
    })
    const cellCache = new ProjectedViewportCellCache()
    const axisStore = new ProjectedViewportAxisStore()
    const coordinator = new ProjectedViewportPatchCoordinator({
      client: {
        invoke: async () => undefined,
        ready: async () => undefined,
        subscribe: () => () => undefined,
        subscribeBatches: () => () => undefined,
        subscribeViewportPatches,
        dispose: () => undefined,
      },
      cellCache,
      axisStore,
    })
    const listener = vi.fn()

    const unsubscribe = coordinator.subscribeViewport('Sheet1', { rowStart: 0, rowEnd: 2, colStart: 0, colEnd: 2 }, listener, {
      initialPatch: 'none',
    })

    patches[0]?.(encodeViewportPatch(hiddenRowPatch))
    patches[0]?.(encodeViewportPatch(staleVisibleRowPatch))
    await vi.runAllTimersAsync()

    expect(axisStore.getRowHeights('Sheet1')[1]).toBe(0)
    expect(axisStore.getHiddenRows('Sheet1')[1]).toBe(true)
    expect(listener).toHaveBeenCalledTimes(1)

    unsubscribe()
    vi.useRealTimers()
  })

  it('coalesces repeated patch notifications into a single frame callback', async () => {
    vi.useFakeTimers()
    const patches: ((bytes: Uint8Array) => void)[] = []
    const subscribeViewportPatches = vi.fn((_viewport, listener: (bytes: Uint8Array) => void) => {
      patches.push(listener)
      return () => undefined
    })
    const coordinator = new ProjectedViewportPatchCoordinator({
      client: {
        invoke: async () => undefined,
        ready: async () => undefined,
        subscribe: () => () => undefined,
        subscribeBatches: () => () => undefined,
        subscribeViewportPatches,
        dispose: () => undefined,
      },
      cellCache: new ProjectedViewportCellCache(),
      axisStore: new ProjectedViewportAxisStore(),
    })
    const listener = vi.fn()

    coordinator.subscribeViewport('Sheet1', { rowStart: 0, rowEnd: 0, colStart: 0, colEnd: 0 }, listener)
    const emit = patches[0]
    if (!emit) {
      throw new Error('expected viewport patch listener')
    }

    const firstPatch = createPatch('style-1')
    const secondPatch = { ...createPatch('style-2'), version: firstPatch.version + 1 }
    emit(encodeViewportPatch(firstPatch))
    emit(
      encodeViewportPatch({
        ...secondPatch,
        cells: [
          {
            ...secondPatch.cells[0],
            col: 1,
            snapshot: {
              ...secondPatch.cells[0].snapshot,
              address: 'B1',
            },
            displayText: '43',
            copyText: '43',
            editorText: '43',
          },
        ],
      }),
    )

    expect(listener).not.toHaveBeenCalled()
    await vi.runAllTimersAsync()
    expect(listener).toHaveBeenCalledTimes(1)
    expect(listener).toHaveBeenLastCalledWith([{ cell: [0, 0] }, { cell: [1, 0] }])
    vi.useRealTimers()
  })
})
