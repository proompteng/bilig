import { describe, expect, it, vi } from 'vitest'
import { ValueTag } from '@bilig/protocol'
import type { WorkerEngineClient } from '@bilig/worker-transport'
import { ProjectedViewportStore } from '../projected-viewport-store.js'
import { applyOptimisticClearRange } from '../workbook-optimistic-range.js'

function createStore() {
  const client: WorkerEngineClient = {
    dispose: vi.fn(),
    invoke: vi.fn(async () => undefined),
    ready: vi.fn(async () => undefined),
    subscribe: vi.fn(() => () => undefined),
    subscribeBatches: vi.fn(() => () => undefined),
    subscribeRenderTileDeltas: vi.fn(() => () => undefined),
    subscribeViewportPatches: vi.fn(() => () => undefined),
    subscribeWorkbookDeltas: vi.fn(() => () => undefined),
  }
  return new ProjectedViewportStore(client)
}

describe('whole-sheet clear projection', () => {
  it('publishes a batch once instead of notifying observers for every visible blank cell', () => {
    const cache = createStore()
    cache.setCellSnapshot({ sheetName: 'Sheet1', address: 'A1', value: { tag: ValueTag.Number, value: 42 }, flags: 0, version: 1 })
    const unsubscribe = cache.subscribeViewport(
      'Sheet1',
      { sheetName: 'Sheet1', rowStart: 0, rowEnd: 130, colStart: 0, colEnd: 144 },
      () => undefined,
      { initialPatch: 'none' },
    )
    let notifications = 0
    const listener = vi.fn(() => {
      if (++notifications > 20) throw new Error('Clear notified observers for each blank cell')
    })
    cache.subscribe(listener)
    const rollback = applyOptimisticClearRange(cache, { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'XFD1048576' })
    expect(listener).toHaveBeenCalledTimes(1)
    expect(cache.getCell('Sheet1', 'A1').value).toEqual({ tag: ValueTag.Empty })
    listener.mockClear()
    notifications = 0
    rollback?.()
    expect(cache.getCell('Sheet1', 'A1').value).toEqual({ tag: ValueTag.Number, value: 42 })
    unsubscribe()
  })

  it('retains formatting and clears later visible cells through the same overlay', () => {
    const cache = createStore()
    cache.setCellSnapshot({
      sheetName: 'Sheet1',
      address: 'B2',
      input: 'clear',
      styleId: 'style-fill',
      numberFormatId: 'fmt-test',
      value: { tag: ValueTag.String, value: 'clear', stringId: 1 },
      flags: 0,
      version: 3,
    })
    const rollback = applyOptimisticClearRange(cache, { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'XFD1048576' })
    expect(cache.getCell('Sheet1', 'B2')).toMatchObject({
      styleId: 'style-fill',
      numberFormatId: 'fmt-test',
      value: { tag: ValueTag.Empty },
      version: 4,
    })
    expect(cache.getCell('Sheet1', 'C12000').value).toEqual({ tag: ValueTag.Empty })
    rollback?.()
    expect(cache.getCell('Sheet1', 'B2')).toMatchObject({ input: 'clear', version: 3 })
  })
})
