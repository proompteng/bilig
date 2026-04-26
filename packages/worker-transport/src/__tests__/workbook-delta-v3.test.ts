import { describe, expect, it } from 'vitest'
import { decodeWorkbookDeltaBatchV3, encodeWorkbookDeltaBatchV3, type WorkbookDeltaBatchV3 } from '../index.js'

function createBatch(): WorkbookDeltaBatchV3 {
  return {
    magic: 'bilig.workbook.delta.v3',
    version: 1,
    seq: 99,
    source: 'workerAuthoritative',
    sheetId: 7,
    sheetOrdinal: 3,
    valueSeq: 11,
    styleSeq: 12,
    axisSeqX: 13,
    axisSeqY: 14,
    freezeSeq: 15,
    calcSeq: 16,
    dirty: {
      cellRanges: Uint32Array.from([1, 2, 3, 4, 7]),
      axisX: Uint32Array.from([3, 7]),
      axisY: Uint32Array.from([5, 9]),
      sheets: Uint32Array.from([7]),
    },
  }
}

describe('workbook delta v3 codec', () => {
  it('round-trips sheet-level dirty range batches through the binary codec', () => {
    const batch = createBatch()
    const bytes = encodeWorkbookDeltaBatchV3(batch)
    const decoded = decodeWorkbookDeltaBatchV3(bytes)

    expect(bytes[0]).not.toBe('{'.charCodeAt(0))
    expect(decoded).toEqual(batch)
  })
})
