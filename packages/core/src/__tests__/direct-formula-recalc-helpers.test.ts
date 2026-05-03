import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import { createEngineCounters } from '../perf/engine-counters.js'
import type { RuntimeDirectCriteriaDescriptor } from '../engine/runtime-state.js'
import { DirectFormulaIndexCollection } from '../engine/services/direct-formula-index-collection.js'
import {
  aggregateColumnDependencyKey,
  canEvaluatePostRecalcDirectFormulasWithoutKernel,
  collectTrackedDependents,
  composeSingleDisjointExplicitEventChanges,
  countDirectFormulaDeltaSkip,
  directAggregateNumericContribution,
  directCriteriaTouchesPoint,
  directCriteriaValueString,
  directFormulaChangesAreDisjointFromInputs,
  hasCompleteDirectFormulaDeltas,
  lookupImpactCacheKey,
} from '../engine/services/direct-formula-recalc-helpers.js'

describe('direct formula recalc helpers', () => {
  it('normalizes aggregate and criteria values for direct formula paths', () => {
    expect(directAggregateNumericContribution({ tag: ValueTag.Number, value: 3 })).toBe(3)
    expect(directAggregateNumericContribution({ tag: ValueTag.Boolean, value: true })).toBe(1)
    expect(directAggregateNumericContribution({ tag: ValueTag.Boolean, value: false })).toBe(0)
    expect(directAggregateNumericContribution({ tag: ValueTag.Empty })).toBe(0)
    expect(directAggregateNumericContribution({ tag: ValueTag.String, value: 'x' })).toBe(0)
    expect(directAggregateNumericContribution({ tag: ValueTag.Error, code: ErrorCode.VALUE })).toBeUndefined()

    expect(directCriteriaValueString({ tag: ValueTag.Empty })).toBe('')
    expect(directCriteriaValueString({ tag: ValueTag.Number, value: -0 })).toBe('0')
    expect(directCriteriaValueString({ tag: ValueTag.Boolean, value: true })).toBe('TRUE')
    expect(directCriteriaValueString({ tag: ValueTag.Boolean, value: false })).toBe('FALSE')
    expect(directCriteriaValueString({ tag: ValueTag.String, value: 'needle' })).toBe('needle')
    expect(directCriteriaValueString({ tag: ValueTag.Error, code: ErrorCode.NA })).toBe(String(ErrorCode.NA))
  })

  it('detects direct criteria ranges and criterion cells touched by a write', () => {
    const criteria = {
      aggregateKind: 'sum',
      aggregateRange: { regionId: 1, sheetName: 'Sheet1', rowStart: 1, rowEnd: 3, col: 4, length: 3 },
      criteriaPairs: [
        {
          range: { regionId: 2, sheetName: 'Sheet1', rowStart: 1, rowEnd: 3, col: 0, length: 3 },
          criterion: { kind: 'literal', value: { tag: ValueTag.String, value: 'A' } },
        },
        {
          range: { regionId: 3, sheetName: 'Sheet1', rowStart: 1, rowEnd: 3, col: 2, length: 3 },
          criterion: { kind: 'cell', cellIndex: 42 },
        },
      ],
    } satisfies RuntimeDirectCriteriaDescriptor

    expect(directCriteriaTouchesPoint(criteria, { sheetName: 'Sheet1', row: 2, col: 4 })).toBe(true)
    expect(directCriteriaTouchesPoint(criteria, { sheetName: 'Sheet1', row: 2, col: 0 })).toBe(true)
    expect(directCriteriaTouchesPoint(criteria, { sheetName: 'Other', row: 2, col: 0, inputCellIndex: 42 })).toBe(true)
    expect(directCriteriaTouchesPoint(criteria, { sheetName: 'Sheet1', row: 4, col: 0 })).toBe(false)
    expect(directCriteriaTouchesPoint(criteria, { sheetName: 'Other', row: 2, col: 0 })).toBe(false)
  })

  it('collects dependent candidates and composes explicit event changes', () => {
    const registry = new Map<string | number, Set<number>>([
      ['a', new Set([1, 2])],
      ['b', new Set([2, 3])],
    ])
    expect(collectTrackedDependents(registry, ['a', 'b', 'missing'])).toEqual([1, 2, 3])
    expect(lookupImpactCacheKey(7, 4)).toBe('7:4')
    expect(aggregateColumnDependencyKey(2, 5)).toBe(2 * 16_384 + 5)
    expect(Array.from(composeSingleDisjointExplicitEventChanges(5, new Uint32Array()))).toEqual([5])
    expect(Array.from(composeSingleDisjointExplicitEventChanges(5, new Uint32Array([6, 7])))).toEqual([5, 6, 7])
  })

  it('decides when direct post-recalc formulas can avoid the kernel', () => {
    const collection = new DirectFormulaIndexCollection()
    collection.appendConstantDelta(new Uint32Array([1, 2]), 3, 'scalar')

    expect(hasCompleteDirectFormulaDeltas(collection)).toBe(true)
    expect(directFormulaChangesAreDisjointFromInputs(new Uint32Array([3, 4]), 2, collection)).toBe(true)
    expect(directFormulaChangesAreDisjointFromInputs(new Uint32Array([1, 4]), 2, collection)).toBe(false)
    expect(
      canEvaluatePostRecalcDirectFormulasWithoutKernel(
        new Map([
          [1, { directScalar: {} }],
          [2, { directAggregate: {} }],
        ]),
        collection,
      ),
    ).toBe(true)
    expect(
      canEvaluatePostRecalcDirectFormulasWithoutKernel(
        new Map([
          [1, { directScalar: {} }],
          [2, {}],
        ]),
        collection,
      ),
    ).toBe(false)
    expect(canEvaluatePostRecalcDirectFormulasWithoutKernel(new Map(), new DirectFormulaIndexCollection())).toBe(false)

    const counters = createEngineCounters()
    countDirectFormulaDeltaSkip(
      new Map([
        [1, { directCriteria: {} }],
        [2, { directScalar: {} }],
      ]),
      collection,
      counters,
    )
    expect(counters.directAggregateDeltaOnlyRecalcSkips).toBe(1)
    expect(counters.directScalarDeltaOnlyRecalcSkips).toBe(1)
  })
})
