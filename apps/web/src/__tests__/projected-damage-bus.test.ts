import { describe, expect, it } from 'vitest'
import { encodeWorkbookDeltaBatchV3 } from '@bilig/worker-transport'
import { DirtyMaskV3 } from '../../../../packages/grid/src/renderer-v3/tile-damage-index.js'
import { tileKeyFromCell } from '../../../../packages/grid/src/renderer-v3/tile-key.js'
import { ProjectedDamageBus } from '../projected-damage-bus.js'

function createEmptyDeltaBatch(sheetOrdinal: number, seq: number) {
  return {
    magic: 'bilig.workbook.delta.v3',
    version: 1,
    seq,
    source: 'remote',
    sheetId: sheetOrdinal,
    sheetOrdinal,
    valueSeq: seq,
    styleSeq: seq,
    axisSeqX: seq,
    axisSeqY: seq,
    freezeSeq: seq,
    calcSeq: seq,
    dirty: {
      axisX: new Uint32Array(),
      axisY: new Uint32Array(),
      cellRanges: new Uint32Array(),
    },
  } as const
}

describe('ProjectedDamageBus', () => {
  it('applies workbook delta batches once and exposes visible dirty tiles', () => {
    const bus = new ProjectedDamageBus()
    const key = tileKeyFromCell({ sheetOrdinal: 2, dprBucket: 1, row: 3, col: 4 })
    const batch = {
      magic: 'bilig.workbook.delta.v3',
      version: 1,
      seq: 10,
      source: 'workerAuthoritative',
      sheetId: 7,
      sheetOrdinal: 2,
      valueSeq: 11,
      styleSeq: 12,
      axisSeqX: 13,
      axisSeqY: 14,
      freezeSeq: 15,
      calcSeq: 16,
      dirty: {
        axisX: new Uint32Array(),
        axisY: new Uint32Array(),
        cellRanges: Uint32Array.from([3, 3, 4, 4, DirtyMaskV3.Value | DirtyMaskV3.Text]),
      },
    } as const

    expect(bus.applyWorkbookDelta(batch, { dprBucket: 1 })).toEqual({ applied: true, seq: 10 })
    expect(bus.applyWorkbookDelta(batch, { dprBucket: 1 })).toEqual({ applied: false, seq: 10 })
    expect(bus.peekWarm([key])).toEqual([key])
    expect(bus.consumeVisible([key])).toEqual([key])
    expect(bus.peekWarm([key])).toEqual([])
  })

  it('tracks delta ordering per sheet ordinal', () => {
    const bus = new ProjectedDamageBus()

    expect(bus.applyWorkbookDelta(createEmptyDeltaBatch(1, 5), { dprBucket: 1 }).applied).toBe(true)
    expect(bus.applyWorkbookDelta(createEmptyDeltaBatch(2, 5), { dprBucket: 1 }).applied).toBe(true)
    expect(bus.applyWorkbookDelta(createEmptyDeltaBatch(1, 4), { dprBucket: 1 }).applied).toBe(false)
  })

  it('subscribes to encoded workbook delta batches from the worker client', () => {
    const key = tileKeyFromCell({ sheetOrdinal: 1, dprBucket: 1, row: 0, col: 0 })
    const listeners: Array<(bytes: Uint8Array) => void> = []
    const bus = new ProjectedDamageBus({
      subscribeWorkbookDeltas(listener) {
        listeners.push(listener)
        return () => undefined
      },
    })
    const results: boolean[] = []

    bus.subscribeWorkbookDeltas({ dprBucket: 1 }, (result) => {
      results.push(result.applied)
    })

    listeners[0]?.(
      encodeWorkbookDeltaBatchV3({
        ...createEmptyDeltaBatch(1, 1),
        dirty: {
          axisX: new Uint32Array(),
          axisY: new Uint32Array(),
          cellRanges: Uint32Array.from([0, 0, 0, 0, DirtyMaskV3.Value]),
        },
      }),
    )

    expect(results).toEqual([true])
    expect(bus.consumeVisible([key])).toEqual([key])
  })
})
