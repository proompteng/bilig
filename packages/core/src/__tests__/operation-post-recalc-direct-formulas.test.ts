import { describe, expect, it, vi } from 'vitest'
import { FormulaMode } from '@bilig/protocol'
import { createEngineCounters } from '../perf/engine-counters.js'
import type { RuntimeDirectScalarDescriptor, U32 } from '../engine/runtime-state.js'
import { DirectFormulaIndexCollection } from '../engine/services/direct-formula-index-collection.js'
import {
  applyPostRecalcDirectFormulaChanges,
  countOperationPostRecalcDirectFormulaMetric,
  tryApplySinglePostRecalcDirectFormula,
  type ApplyPostRecalcDirectFormulaChangesArgs,
  type DirectFormulaMetricCounts,
  type OperationPostRecalcFormula,
} from '../engine/services/operation-post-recalc-direct-formulas.js'

const directScalar: RuntimeDirectScalarDescriptor = {
  kind: 'abs',
  operand: { kind: 'literal-number', value: 1 },
}

function formula(overrides: Partial<OperationPostRecalcFormula> = {}): OperationPostRecalcFormula {
  return {
    compiled: { producesSpill: false },
    directAggregate: undefined,
    directCriteria: undefined,
    directScalar: undefined,
    ...overrides,
  }
}

function makeState(formulas: Map<number, OperationPostRecalcFormula> = new Map()): ApplyPostRecalcDirectFormulaChangesArgs['state'] {
  return {
    workbook: {
      cellStore: {
        flags: [],
      },
      withBatchedColumnVersionUpdates: (apply: () => void): void => apply(),
    },
    formulas: {
      get: (cellIndex: number) => formulas.get(cellIndex),
    },
    counters: createEngineCounters(),
  }
}

function makeArgs(overrides: Partial<ApplyPostRecalcDirectFormulaChangesArgs> = {}): ApplyPostRecalcDirectFormulaChangesArgs {
  const collection = new DirectFormulaIndexCollection()
  const metrics: DirectFormulaMetricCounts = { wasmFormulaCount: 0, jsFormulaCount: 0 }
  return {
    state: makeState(),
    collection,
    recalculated: new Uint32Array(),
    didRunRecalc: false,
    metrics,
    applyDirectFormulaCurrentResult: vi.fn(() => true),
    applyDirectFormulaNumericDelta: vi.fn(() => true),
    applyDirectScalarCurrentValue: vi.fn(() => true),
    tryApplyDirectScalarDeltas: vi.fn(() => undefined),
    tryApplyDirectFormulaDeltas: vi.fn(() => undefined),
    countPostRecalcDirectFormulaMetric: vi.fn(),
    evaluateDirectFormula: vi.fn(() => undefined),
    ...overrides,
  }
}

describe('operation post-recalc direct formula helpers', () => {
  it('applies a single current result while allowing callers to suppress captured changes', () => {
    const collection = new DirectFormulaIndexCollection()
    collection.addCurrentResult(7, { kind: 'number', value: 42 })
    const applyCurrent = vi.fn(() => true)
    const args = makeArgs({
      collection,
      applyDirectFormulaCurrentResult: applyCurrent,
    })

    const changed = tryApplySinglePostRecalcDirectFormula(args, false)

    expect(Array.from(changed ?? [])).toEqual([])
    expect(applyCurrent).toHaveBeenCalledWith(7, { kind: 'number', value: 42 })
  })

  it('uses complete delta application before falling back to per-formula evaluation', () => {
    const collection = new DirectFormulaIndexCollection()
    collection.addDelta(2, 1)
    collection.addDelta(3, 1)
    const tryApplyDirectFormulaDeltas = vi.fn((): U32 => Uint32Array.of(2, 3))
    const evaluateDirectFormula = vi.fn(() => [99])

    const changed = applyPostRecalcDirectFormulaChanges(
      makeArgs({
        collection,
        recalculated: Uint32Array.of(1),
        tryApplyDirectFormulaDeltas,
        evaluateDirectFormula,
      }),
    )

    expect(Array.from(changed)).toEqual([1, 2, 3])
    expect(tryApplyDirectFormulaDeltas).toHaveBeenCalledWith(collection, true)
    expect(evaluateDirectFormula).not.toHaveBeenCalled()
  })

  it('counts post-recalc direct formula metrics by formula mode', () => {
    const counts: DirectFormulaMetricCounts = { wasmFormulaCount: 0, jsFormulaCount: 0 }
    const formulas = new Map<number, OperationPostRecalcFormula>([
      [
        1,
        formula({
          compiled: { mode: FormulaMode.WasmFastPath, producesSpill: false },
          directScalar,
        }),
      ],
      [
        2,
        formula({
          compiled: { mode: FormulaMode.JsOnly, producesSpill: false },
          directAggregate: {},
        }),
      ],
      [
        3,
        formula({
          compiled: { mode: FormulaMode.JsOnly, producesSpill: false },
        }),
      ],
    ])

    ;[1, 2, 3, 4].forEach((cellIndex) => {
      countOperationPostRecalcDirectFormulaMetric({
        formulas: { get: (lookupCellIndex) => formulas.get(lookupCellIndex) },
        cellIndex,
        counts,
      })
    })

    expect(counts).toEqual({ wasmFormulaCount: 1, jsFormulaCount: 1 })
  })

  it('evaluates remaining direct formulas in a batched column-version update', () => {
    const collection = new DirectFormulaIndexCollection()
    collection.add(2)
    collection.add(3)
    const withBatchedColumnVersionUpdates = vi.fn((apply: () => void): void => apply())
    const formulas = new Map([
      [2, formula({ cellIndex: 2 })],
      [3, formula({ cellIndex: 3 })],
    ])
    const evaluateDirectFormula = vi.fn((cellIndex: number) => [cellIndex + 100])

    const changed = applyPostRecalcDirectFormulaChanges(
      makeArgs({
        state: {
          ...makeState(formulas),
          workbook: {
            cellStore: {
              flags: [],
            },
            withBatchedColumnVersionUpdates,
          },
        },
        collection,
        recalculated: Uint32Array.of(1),
        didRunRecalc: true,
        applyDirectScalarCurrentValue: vi.fn(() => false),
        evaluateDirectFormula,
      }),
    )

    expect(Array.from(changed)).toEqual([1, 2, 3, 102, 103])
    expect(withBatchedColumnVersionUpdates).toHaveBeenCalledOnce()
    expect(evaluateDirectFormula).toHaveBeenCalledWith(2)
    expect(evaluateDirectFormula).toHaveBeenCalledWith(3)
  })

  it('counts aggregate and scalar delta applications on the single-cell delta path', () => {
    const collection = new DirectFormulaIndexCollection()
    collection.addDelta(4, 2)
    const state = makeState(
      new Map([
        [
          4,
          formula({
            directAggregate: {},
            directScalar,
          }),
        ],
      ]),
    )

    const changed = tryApplySinglePostRecalcDirectFormula(
      makeArgs({
        state,
        collection,
      }),
    )

    expect(Array.from(changed ?? [])).toEqual([4])
    expect(state.counters.directAggregateDeltaApplications).toBe(1)
    expect(state.counters.directScalarDeltaApplications).toBe(1)
  })
})
