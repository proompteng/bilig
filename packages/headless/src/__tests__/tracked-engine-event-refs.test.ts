import { describe, expect, it } from 'vitest'
import { ValueTag } from '@bilig/protocol'
import { canUseSortedDisjointTrackedEngineEventChanges, captureTrackedEngineEvent } from '../tracked-engine-event-refs.js'

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
    expect(tracked.changedCellIndicesSortedDisjoint).toBe(true)
    expect(tracked.firstChangedCellIndex).toBe(1)
    expect(tracked.lastChangedCellIndex).toBe(9)
    expect(tracked.changedInputCount).toBe(3)
    expect(tracked.explicitChangedCount).toBe(2)
    expect(tracked.hasInvalidatedRanges).toBe(true)
    expect(tracked.hasInvalidatedRows).toBe(true)
    expect(tracked.hasInvalidatedColumns).toBe(true)
  })

  it('can retain changed-cell indices for owned internal events', () => {
    const changedCellIndices = new Uint32Array([2, 5])
    const tracked = captureTrackedEngineEvent(
      {
        kind: 'batch',
        invalidation: 'cells',
        changedCellIndices,
        invalidatedRanges: [],
        invalidatedRows: [],
        invalidatedColumns: [],
        metrics: {
          batchId: 1,
          changedInputCount: 2,
          dirtyFormulaCount: 0,
          wasmFormulaCount: 0,
          jsFormulaCount: 0,
          rangeNodeVisits: 0,
          recalcMs: 0,
          compileMs: 0,
        },
      },
      { cloneChangedCellIndices: false },
    )

    expect(tracked.changedCellIndices).toBe(changedCellIndices)
    expect(tracked.changedCellIndicesSortedDisjoint).toBe(true)
    expect(tracked.firstChangedCellIndex).toBe(2)
    expect(tracked.lastChangedCellIndex).toBe(5)
  })

  it('uses a tiny borrowed index fast path without trusting unsorted pairs', () => {
    const changedCellIndices = new Uint32Array([5, 2])
    const tracked = captureTrackedEngineEvent(
      {
        kind: 'batch',
        invalidation: 'cells',
        changedCellIndices,
        invalidatedRanges: [],
        invalidatedRows: [],
        invalidatedColumns: [],
        metrics: {
          batchId: 1,
          changedInputCount: 2,
          dirtyFormulaCount: 0,
          wasmFormulaCount: 0,
          jsFormulaCount: 0,
          rangeNodeVisits: 0,
          recalcMs: 0,
          compileMs: 0,
        },
      },
      { cloneChangedCellIndices: false, borrowChangedCellIndexViews: true },
    )

    expect(tracked.changedCellIndices).toBe(changedCellIndices)
    expect(tracked.changedCellIndicesSortedDisjoint).toBe(false)
    expect(tracked.firstChangedCellIndex).toBe(5)
    expect(tracked.lastChangedCellIndex).toBe(2)
  })

  it('trusts borrowed physical split metadata without scanning raw numeric order', () => {
    const changedCellIndices = new Uint32Array([2, 4, 1, 3])
    Reflect.set(changedCellIndices, '__biligTrackedPhysicalSheetId', 1)
    Reflect.set(changedCellIndices, '__biligTrackedPhysicalSortedSliceSplit', 2)

    const tracked = captureTrackedEngineEvent(
      {
        kind: 'batch',
        invalidation: 'cells',
        changedCellIndices,
        invalidatedRanges: [],
        invalidatedRows: [],
        invalidatedColumns: [],
        metrics: {
          batchId: 1,
          changedInputCount: 2,
          dirtyFormulaCount: 0,
          wasmFormulaCount: 0,
          jsFormulaCount: 0,
          rangeNodeVisits: 0,
          recalcMs: 0,
          compileMs: 0,
        },
        explicitChangedCount: 2,
      },
      { cloneChangedCellIndices: false, borrowChangedCellIndexViews: true },
    )

    expect(tracked.changedCellIndices).toBe(changedCellIndices)
    expect(tracked.changedCellIndicesSortedDisjoint).toBe(true)
    expect(tracked.firstChangedCellIndex).toBe(2)
    expect(tracked.lastChangedCellIndex).toBe(3)
  })

  it('omits empty patch arrays from tracked events', () => {
    const tracked = captureTrackedEngineEvent({
      kind: 'batch',
      invalidation: 'cells',
      changedCellIndices: new Uint32Array([2]),
      patches: [],
      invalidatedRanges: [],
      invalidatedRows: [],
      invalidatedColumns: [],
      metrics: {
        batchId: 1,
        changedInputCount: 1,
        dirtyFormulaCount: 0,
        wasmFormulaCount: 0,
        jsFormulaCount: 0,
        rangeNodeVisits: 0,
        recalcMs: 0,
        compileMs: 0,
      },
    })

    expect(tracked.patches).toBeUndefined()
  })

  it('copies changed-cell index views even when retention is requested', () => {
    const backing = new Uint32Array([2, 5, 8, 13])
    const changedCellIndices = backing.subarray(1, 3)
    const tracked = captureTrackedEngineEvent(
      {
        kind: 'batch',
        invalidation: 'cells',
        changedCellIndices,
        invalidatedRanges: [],
        invalidatedRows: [],
        invalidatedColumns: [],
        metrics: {
          batchId: 1,
          changedInputCount: 2,
          dirtyFormulaCount: 0,
          wasmFormulaCount: 0,
          jsFormulaCount: 0,
          rangeNodeVisits: 0,
          recalcMs: 0,
          compileMs: 0,
        },
      },
      { cloneChangedCellIndices: false },
    )

    backing[1] = 99

    expect(tracked.changedCellIndices).toEqual(new Uint32Array([5, 8]))
    expect(tracked.changedCellIndicesSortedDisjoint).toBe(true)
    expect(tracked.firstChangedCellIndex).toBe(5)
    expect(tracked.lastChangedCellIndex).toBe(8)
  })

  it('identifies sorted disjoint numeric changed-cell event runs', () => {
    const first = captureTrackedEngineEvent({
      kind: 'batch',
      invalidation: 'cells',
      changedCellIndices: new Uint32Array([1, 3, 5]),
      invalidatedRanges: [],
      invalidatedRows: [],
      invalidatedColumns: [],
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
    })
    const second = captureTrackedEngineEvent({
      kind: 'batch',
      invalidation: 'cells',
      changedCellIndices: new Uint32Array([8, 13]),
      invalidatedRanges: [],
      invalidatedRows: [],
      invalidatedColumns: [],
      metrics: {
        batchId: 2,
        changedInputCount: 2,
        dirtyFormulaCount: 0,
        wasmFormulaCount: 0,
        jsFormulaCount: 0,
        rangeNodeVisits: 0,
        recalcMs: 0,
        compileMs: 0,
      },
    })

    expect(first.changedCellIndicesSortedDisjoint).toBe(true)
    expect(second.changedCellIndicesSortedDisjoint).toBe(true)
    expect(canUseSortedDisjointTrackedEngineEventChanges([first, second])).toBe(true)
  })

  it('rejects overlapping, unsorted, and invalidated tracked event runs', () => {
    const sorted = captureTrackedEngineEvent({
      kind: 'batch',
      invalidation: 'cells',
      changedCellIndices: new Uint32Array([1, 4]),
      invalidatedRanges: [],
      invalidatedRows: [],
      invalidatedColumns: [],
      metrics: {
        batchId: 1,
        changedInputCount: 2,
        dirtyFormulaCount: 0,
        wasmFormulaCount: 0,
        jsFormulaCount: 0,
        rangeNodeVisits: 0,
        recalcMs: 0,
        compileMs: 0,
      },
    })
    const overlapping = captureTrackedEngineEvent({
      kind: 'batch',
      invalidation: 'cells',
      changedCellIndices: new Uint32Array([4, 5]),
      invalidatedRanges: [],
      invalidatedRows: [],
      invalidatedColumns: [],
      metrics: {
        batchId: 2,
        changedInputCount: 2,
        dirtyFormulaCount: 0,
        wasmFormulaCount: 0,
        jsFormulaCount: 0,
        rangeNodeVisits: 0,
        recalcMs: 0,
        compileMs: 0,
      },
    })
    const unsorted = captureTrackedEngineEvent({
      kind: 'batch',
      invalidation: 'cells',
      changedCellIndices: new Uint32Array([9, 7]),
      invalidatedRanges: [],
      invalidatedRows: [],
      invalidatedColumns: [],
      metrics: {
        batchId: 3,
        changedInputCount: 2,
        dirtyFormulaCount: 0,
        wasmFormulaCount: 0,
        jsFormulaCount: 0,
        rangeNodeVisits: 0,
        recalcMs: 0,
        compileMs: 0,
      },
    })
    const invalidated = captureTrackedEngineEvent({
      kind: 'batch',
      invalidation: 'cells',
      changedCellIndices: new Uint32Array([10, 11]),
      invalidatedRanges: [{ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A2' }],
      invalidatedRows: [],
      invalidatedColumns: [],
      metrics: {
        batchId: 4,
        changedInputCount: 2,
        dirtyFormulaCount: 0,
        wasmFormulaCount: 0,
        jsFormulaCount: 0,
        rangeNodeVisits: 0,
        recalcMs: 0,
        compileMs: 0,
      },
    })

    expect(unsorted.changedCellIndicesSortedDisjoint).toBe(false)
    expect(canUseSortedDisjointTrackedEngineEventChanges([sorted, overlapping])).toBe(false)
    expect(canUseSortedDisjointTrackedEngineEventChanges([unsorted])).toBe(false)
    expect(canUseSortedDisjointTrackedEngineEventChanges([invalidated])).toBe(false)
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

  it('derives invalidation flags from typed invalidation patches', () => {
    const tracked = captureTrackedEngineEvent({
      kind: 'batch',
      invalidation: 'cells',
      changedCellIndices: new Uint32Array(),
      patches: [
        {
          kind: 'row-invalidation',
          sheetName: 'Bench',
          startIndex: 3,
          endIndex: 4,
        },
        {
          kind: 'column-invalidation',
          sheetName: 'Bench',
          startIndex: 1,
          endIndex: 2,
        },
        {
          kind: 'range-invalidation',
          range: { sheetName: 'Bench', startAddress: 'A1', endAddress: 'C3' },
        },
      ],
      invalidatedRanges: [],
      invalidatedRows: [],
      invalidatedColumns: [],
      metrics: {
        batchId: 3,
        changedInputCount: 0,
        dirtyFormulaCount: 0,
        wasmFormulaCount: 0,
        jsFormulaCount: 0,
        rangeNodeVisits: 0,
        recalcMs: 0,
        compileMs: 0,
      },
    })

    expect(tracked.hasInvalidatedRows).toBe(true)
    expect(tracked.hasInvalidatedColumns).toBe(true)
    expect(tracked.hasInvalidatedRanges).toBe(true)
  })
})
