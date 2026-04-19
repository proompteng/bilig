import { describe, expect, it } from 'vitest'
import { ValueTag } from '@bilig/protocol'
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

  it('copies tracked patch arrays for later lazy consumption', () => {
    const patches = [
      {
        kind: 'cell' as const,
        cellIndex: 1,
        address: { sheet: 7, row: 2, col: 3 },
        sheetName: 'Bench',
        a1: 'D3',
        newValue: { tag: ValueTag.Number, value: 42 },
      },
    ]
    const tracked = captureTrackedEngineEvent({
      kind: 'batch',
      invalidation: 'cells',
      changedCellIndices: new Uint32Array([1]),
      patches,
      invalidatedRanges: [],
      invalidatedRows: [],
      invalidatedColumns: [],
      metrics: {
        batchId: 2,
        changedInputCount: 1,
        dirtyFormulaCount: 0,
        wasmFormulaCount: 0,
        jsFormulaCount: 0,
        rangeNodeVisits: 0,
        recalcMs: 0,
        compileMs: 0,
      },
    })

    expect(tracked.patches).toEqual([
      {
        kind: 'cell',
        cellIndex: 1,
        address: { sheet: 7, row: 2, col: 3 },
        sheetName: 'Bench',
        a1: 'D3',
        newValue: { tag: ValueTag.Number, value: 42 },
      },
    ])

    patches[0] = {
      ...patches[0],
      address: { ...patches[0].address, row: 99 },
    }

    expect(tracked.patches).toEqual([
      {
        kind: 'cell',
        cellIndex: 1,
        address: { sheet: 7, row: 2, col: 3 },
        sheetName: 'Bench',
        a1: 'D3',
        newValue: { tag: ValueTag.Number, value: 42 },
      },
    ])
  })
})
