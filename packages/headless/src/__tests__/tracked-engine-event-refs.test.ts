import { describe, expect, it } from 'vitest'
import { captureTrackedEngineEvent } from '../tracked-engine-event-refs.js'

describe('captureTrackedEngineEvent', () => {
  it('clones changed-cell indices and preserves invalidation flags', () => {
    const changedCellIndices = new Uint32Array([1, 4, 9])
    const tracked = captureTrackedEngineEvent({
      kind: 'batch',
      invalidation: 'cells',
      changedCellIndices,
      invalidatedRanges: [{ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' }],
      invalidatedRows: [{ sheetName: 'Sheet1', startIndex: 2, endIndex: 3 }],
      invalidatedColumns: [{ sheetName: 'Sheet1', startIndex: 1, endIndex: 1 }],
      metrics: {
        batchId: 1,
        changedInputCount: 3,
        dirtyFormulaCount: 0,
        wasmFormulaCount: 0,
        jsFormulaCount: 0,
        rangeNodeVisits: 0,
        recalcMs: 0,
        compileMs: 0,
      },
      explicitChangedCount: 2,
    })

    changedCellIndices[0] = 99

    expect(tracked.changedCellIndices).toEqual(new Uint32Array([1, 4, 9]))
    expect(tracked.changedInputCount).toBe(3)
    expect(tracked.explicitChangedCount).toBe(2)
    expect(tracked.hasInvalidatedRanges).toBe(true)
    expect(tracked.hasInvalidatedRows).toBe(true)
    expect(tracked.hasInvalidatedColumns).toBe(true)
  })
})
