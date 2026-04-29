import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import type { EngineOp } from '@bilig/workbook-domain'
import { operationServiceTestHooks } from '../engine/services/operation-service.js'
import type {
  RuntimeDirectCriteriaDescriptor,
  RuntimeDirectLookupDescriptor,
  RuntimeDirectScalarDescriptor,
} from '../engine/runtime-state.js'
import { SpreadsheetEngine } from '../engine.js'

const {
  DirectFormulaIndexCollection,
  PendingNumericCellValues,
  ROW_PAIR_LEFT_DIV_RIGHT,
  ROW_PAIR_LEFT_MINUS_RIGHT,
  ROW_PAIR_LEFT_PLUS_RIGHT,
  ROW_PAIR_LEFT_TIMES_RIGHT,
  ROW_PAIR_RIGHT_DIV_LEFT,
  ROW_PAIR_RIGHT_MINUS_LEFT,
  aggregateColumnDependencyKey,
  approximateUniformLookupCurrentResult,
  approximateUniformLookupNumericResult,
  canEvaluatePostRecalcDirectFormulasWithoutKernel,
  canSkipUniformApproximateNumericTailWrite,
  canSkipUniformApproximateNumericTailWriteFromCurrentResult,
  canSkipUniformExactNumericTailWriteFromCurrentResult,
  cellRange,
  collectTrackedDependents,
  composeSingleDisjointExplicitEventChanges,
  countDirectFormulaDeltaSkip,
  directAggregateNumericContribution,
  directCriteriaTouchesPoint,
  directFormulaChangesAreDisjointFromInputs,
  directLookupRowBounds,
  directScalarLiteralNumericValue,
  evaluateRowPairDirectScalarCode,
  exactLookupLiteralNumericValue,
  exactUniformLookupCurrentResult,
  exactUniformLookupNumericResult,
  lookupImpactCacheKey,
  makeCompactExistingNumericMutationResult,
  makeExistingNumericMutationResult,
  mergeChangedCellIndices,
  normalizeApproximateNumericValue,
  normalizeApproximateTextValue,
  normalizeExactLookupKey,
  normalizeExactNumericValue,
  rangesIntersect,
  rowPairDirectScalarCode,
  sameExactNumericValue,
  singleInputAffineDirectScalar,
  tagTrustedPhysicalTrackedChanges,
  throwProtectionBlocked,
  withOptionalLookupStringIds,
} = operationServiceTestHooks

function exactUniform(overrides: Partial<Extract<RuntimeDirectLookupDescriptor, { kind: 'exact-uniform-numeric' }>> = {}) {
  return {
    kind: 'exact-uniform-numeric',
    operandCellIndex: 1,
    sheetName: 'Sheet1',
    sheetId: 7,
    rowStart: 0,
    rowEnd: 4,
    col: 0,
    length: 5,
    columnVersion: 2,
    structureVersion: 3,
    sheetColumnVersions: new Uint32Array([2]),
    start: 1,
    step: 1,
    searchMode: 1,
    ...overrides,
  } satisfies Extract<RuntimeDirectLookupDescriptor, { kind: 'exact-uniform-numeric' }>
}

function approximateUniform(overrides: Partial<Extract<RuntimeDirectLookupDescriptor, { kind: 'approximate-uniform-numeric' }>> = {}) {
  return {
    kind: 'approximate-uniform-numeric',
    operandCellIndex: 1,
    sheetName: 'Sheet1',
    sheetId: 7,
    rowStart: 0,
    rowEnd: 4,
    col: 0,
    length: 5,
    columnVersion: 2,
    structureVersion: 3,
    sheetColumnVersions: new Uint32Array([2]),
    start: 1,
    step: 1,
    matchMode: 1,
    ...overrides,
  } satisfies Extract<RuntimeDirectLookupDescriptor, { kind: 'approximate-uniform-numeric' }>
}

function binaryScalar(operator: '+' | '-' | '*' | '/', leftCellIndex: number, rightCellIndex: number): RuntimeDirectScalarDescriptor {
  return {
    kind: 'binary',
    operator,
    left: { kind: 'cell', cellIndex: leftCellIndex },
    right: { kind: 'cell', cellIndex: rightCellIndex },
  }
}

const lookupString = (id: number) => (id === 1 ? 'alpha' : 'beta')

function runtimeHooks(engine: SpreadsheetEngine): object {
  const runtime = Reflect.get(engine, 'runtime')
  if (typeof runtime !== 'object' || runtime === null) {
    throw new TypeError('Expected engine runtime')
  }
  const operations = Reflect.get(runtime, 'operations')
  if (typeof operations !== 'object' || operations === null) {
    throw new TypeError('Expected operation service')
  }
  const hooks = Reflect.get(operations, '__testHooks')
  if (typeof hooks !== 'object' || hooks === null) {
    throw new TypeError('Expected operation service hooks')
  }
  return hooks
}

function hookFunction<Args extends readonly unknown[], Return>(hooks: object, name: string): (...args: Args) => Return {
  const fn = Reflect.get(hooks, name)
  if (typeof fn !== 'function') {
    throw new TypeError(`Expected hook ${name}`)
  }
  return (...args: Args): Return => Reflect.apply(fn, hooks, args)
}

describe('operation-service internals', () => {
  it('tracks pending numeric cells with sparse assigned slots', () => {
    const values = new PendingNumericCellValues()

    expect(values.has(4)).toBe(false)
    expect(values.get(4)).toBeUndefined()

    values.set(4, -0)

    expect(values.has(4)).toBe(true)
    expect(Object.is(values.get(4), -0)).toBe(true)
    expect(values.has(5)).toBe(false)
  })

  it('maintains direct formula indices, deltas, validation, and covered inputs through sparse and bulk paths', () => {
    const indices = new DirectFormulaIndexCollection()

    expect(indices.size).toBe(0)
    expect(indices.has(8)).toBe(false)
    expect(indices.hasCompleteDeltas()).toBe(false)
    expect(indices.hasCoveredDirectRangeInput(3)).toBe(false)
    expect(indices.hasCoveredDirectFormulaInput(3)).toBe(false)

    indices.add(8)
    indices.add(8)
    indices.addDelta(8, 2)
    indices.addScalarDelta(8, 3)
    indices.addCurrentResult(8, { kind: 'number', value: 5 })
    indices.markDirectRangeInputCovered(3)
    indices.markDirectRangeInputCovered(3)
    indices.markDirectFormulaInputCovered(4)
    indices.markDirectFormulaInputCovered(4)

    expect(indices.size).toBe(1)
    expect(indices.has(8)).toBe(true)
    expect(indices.getDelta(8)).toBe(5)
    expect(indices.getDelta(9)).toBeUndefined()
    expect(indices.getDeltaAt(0)).toBe(5)
    expect(indices.getCurrentResult(8)).toEqual({ kind: 'number', value: 5 })
    expect(indices.getCurrentResultAt(0)).toEqual({ kind: 'number', value: 5 })
    expect(indices.hasCoveredDirectRangeInput(3)).toBe(true)
    expect(indices.hasCoveredDirectRangeInput(4)).toBe(false)
    expect(indices.hasCoveredDirectFormulaInput(4)).toBe(true)
    expect(indices.hasCoveredDirectFormulaInput(5)).toBe(false)
    expect(indices.getConstantScalarDelta()).toBeUndefined()
    expect(indices.hasCompleteScalarDeltas()).toBe(false)

    const visited: number[] = []
    indices.forEach((cellIndex) => visited.push(cellIndex))
    expect(visited).toEqual([8])

    const indexed: [number, number][] = []
    indices.forEachIndexed((cellIndex, index) => indexed.push([cellIndex, index]))
    expect(indexed).toEqual([[8, 0]])

    indices.appendDeltas([10, 11], [1, 2], 'scalar')
    expect(indices.getDelta(10)).toBe(1)
    expect(indices.getScalarDeltaAt(1)).toBe(1)
    expect(indices.getScalarDeltaAt(2)).toBe(2)

    indices.appendDeltas(new Uint32Array([10, 12]), [4, 6])
    expect(indices.getDelta(10)).toBe(5)
    expect(indices.getDelta(12)).toBe(6)
    expect(indices.getScalarDeltaAt(1)).toBeUndefined()

    const constant = new DirectFormulaIndexCollection()
    constant.appendConstantDelta(new Uint32Array([20, 21, 22]), 7, 'scalar')
    expect(constant.size).toBe(3)
    expect(constant.getDelta(20)).toBe(7)
    expect(constant.getDeltaAt(2)).toBe(7)
    expect(constant.getScalarDeltaAt(1)).toBe(7)
    expect(constant.getConstantScalarDelta()).toBe(7)
    expect(constant.hasCompleteScalarDeltas()).toBe(true)
    expect(constant.hasValidatedScalarDeltaCells()).toBe(false)
    constant.markScalarDeltaCellsValidated()
    expect(constant.hasValidatedScalarDeltaCells()).toBe(true)

    constant.addScalarDelta(23, 7)
    expect(constant.getDelta(23)).toBe(7)
    expect(constant.getConstantScalarDelta()).toBe(7)
    constant.addDelta(24, 7)
    expect(constant.getConstantScalarDelta()).toBeUndefined()

    const bulk = new DirectFormulaIndexCollection()
    bulk.appendConstantDelta(
      Array.from({ length: 18 }, (_, index) => index + 100),
      1,
      'scalar',
    )
    bulk.appendDeltas([110, 118, 119], [2, 3, 4], 'scalar')
    bulk.appendConstantDelta(new Uint32Array([119, 120]), 5)

    expect(bulk.has(110)).toBe(true)
    expect(bulk.getCellIndexAt(0)).toBe(100)
    expect(bulk.getDelta(110)).toBe(3)
    expect(bulk.getDelta(118)).toBe(3)
    expect(bulk.getDelta(119)).toBe(9)
    expect(bulk.getDelta(120)).toBe(5)
    expect(bulk.getScalarDeltaAt(19)).toBeUndefined()
  })

  it('normalizes lookup values and computes lookup row results across exact and approximate uniform paths', () => {
    expect(directAggregateNumericContribution({ tag: ValueTag.Number, value: 3 })).toBe(3)
    expect(directAggregateNumericContribution({ tag: ValueTag.Boolean, value: true })).toBe(1)
    expect(directAggregateNumericContribution({ tag: ValueTag.Boolean, value: false })).toBe(0)
    expect(directAggregateNumericContribution({ tag: ValueTag.Empty })).toBe(0)
    expect(directAggregateNumericContribution({ tag: ValueTag.String, value: 'x' })).toBe(0)
    expect(directAggregateNumericContribution({ tag: ValueTag.Error, code: ErrorCode.VALUE })).toBeUndefined()

    expect(normalizeExactLookupKey({ tag: ValueTag.Empty }, lookupString)).toBe('e:')
    expect(normalizeExactLookupKey({ tag: ValueTag.Number, value: -0 }, lookupString)).toBe('n:0')
    expect(normalizeExactLookupKey({ tag: ValueTag.Boolean, value: true }, lookupString)).toBe('b:1')
    expect(normalizeExactLookupKey({ tag: ValueTag.Boolean, value: false }, lookupString)).toBe('b:0')
    expect(normalizeExactLookupKey({ tag: ValueTag.String, value: 'local' }, lookupString)).toBe('s:LOCAL')
    expect(normalizeExactLookupKey({ tag: ValueTag.String, value: 'ignored' }, lookupString, 1)).toBe('s:ALPHA')
    expect(normalizeExactLookupKey({ tag: ValueTag.Error, code: ErrorCode.NA }, lookupString)).toBeUndefined()

    expect(normalizeExactNumericValue({ tag: ValueTag.Number, value: -0 })).toBe(0)
    expect(normalizeExactNumericValue({ tag: ValueTag.Empty })).toBeUndefined()
    expect(sameExactNumericValue(-0, 0)).toBe(true)
    expect(exactLookupLiteralNumericValue(-0)).toBe(0)
    expect(exactLookupLiteralNumericValue('1')).toBeUndefined()

    expect(normalizeApproximateNumericValue({ tag: ValueTag.Empty })).toBe(0)
    expect(normalizeApproximateNumericValue({ tag: ValueTag.Number, value: -0 })).toBe(0)
    expect(normalizeApproximateNumericValue({ tag: ValueTag.Boolean, value: true })).toBe(1)
    expect(normalizeApproximateNumericValue({ tag: ValueTag.String, value: '1' })).toBeUndefined()
    expect(normalizeApproximateNumericValue({ tag: ValueTag.Error, code: ErrorCode.NA })).toBeUndefined()

    expect(normalizeApproximateTextValue({ tag: ValueTag.Empty }, lookupString)).toBe('')
    expect(normalizeApproximateTextValue({ tag: ValueTag.String, value: 'local' }, lookupString)).toBe('LOCAL')
    expect(normalizeApproximateTextValue({ tag: ValueTag.String, value: 'ignored' }, lookupString, 2)).toBe('BETA')
    expect(normalizeApproximateTextValue({ tag: ValueTag.Number, value: 1 }, lookupString)).toBeUndefined()

    expect(directScalarLiteralNumericValue(null)).toBe(0)
    expect(directScalarLiteralNumericValue(-0)).toBe(0)
    expect(directScalarLiteralNumericValue(true)).toBe(1)
    expect(directScalarLiteralNumericValue(false)).toBe(0)
    expect(directScalarLiteralNumericValue('1')).toBeUndefined()
    expect(directScalarLiteralNumericValue(undefined)).toBeUndefined()

    expect(exactUniformLookupNumericResult(exactUniform(), 3)).toBe(3)
    expect(exactUniformLookupNumericResult(exactUniform(), 3.5)).toBeUndefined()
    expect(exactUniformLookupNumericResult(exactUniform({ step: -1, start: 5 }), 3)).toBe(3)
    expect(exactUniformLookupNumericResult(exactUniform({ step: 2, start: 2 }), 8)).toBe(4)
    expect(exactUniformLookupNumericResult(exactUniform({ step: 2, start: 2 }), 9)).toBeUndefined()
    expect(
      exactUniformLookupNumericResult(exactUniform({ tailPatch: { row: 4, oldNumeric: 5, newNumeric: 9, columnVersion: 4 } }), 9),
    ).toBe(5)
    expect(
      exactUniformLookupNumericResult(exactUniform({ tailPatch: { row: 4, oldNumeric: 5, newNumeric: 9, columnVersion: 4 } }), 5),
    ).toBeUndefined()
    expect(exactUniformLookupCurrentResult(exactUniform(), 8)).toEqual({ kind: 'error', code: ErrorCode.NA })

    expect(approximateUniformLookupNumericResult(approximateUniform(), 3.2)).toBe(3)
    expect(approximateUniformLookupNumericResult(approximateUniform(), 0)).toBeUndefined()
    expect(approximateUniformLookupNumericResult(approximateUniform({ start: 10, step: -1, matchMode: -1 }), 8.8)).toBe(2)
    expect(approximateUniformLookupNumericResult(approximateUniform({ start: 10, step: -1, matchMode: -1 }), 11)).toBeUndefined()
    expect(approximateUniformLookupNumericResult(approximateUniform({ step: 2 }), 4.9)).toBe(2)
    expect(
      approximateUniformLookupNumericResult(
        approximateUniform({ tailPatch: { row: 4, oldNumeric: 5, newNumeric: 9, columnVersion: 4 } }),
        6,
      ),
    ).toBe(4)
    expect(
      approximateUniformLookupNumericResult(
        approximateUniform({ tailPatch: { row: 4, oldNumeric: 5, newNumeric: 9, columnVersion: 4 } }),
        9,
      ),
    ).toBe(5)
    expect(
      approximateUniformLookupNumericResult(
        approximateUniform({ start: 10, step: -1, matchMode: -1, tailPatch: { row: 4, oldNumeric: 6, newNumeric: 2, columnVersion: 4 } }),
        4,
      ),
    ).toBe(4)
    expect(
      approximateUniformLookupNumericResult(
        approximateUniform({ start: 10, step: -1, matchMode: -1, tailPatch: { row: 4, oldNumeric: 6, newNumeric: 2, columnVersion: 4 } }),
        1,
      ),
    ).toBe(5)
    expect(approximateUniformLookupCurrentResult(approximateUniform(), 0)).toEqual({ kind: 'error', code: ErrorCode.NA })
    expect(approximateUniformLookupCurrentResult(approximateUniform({ start: 10, step: -1, matchMode: -1 }), 11)).toEqual({
      kind: 'error',
      code: ErrorCode.NA,
    })
  })

  it('covers lookup skip guards, range utilities, changed-cell merges, and result shaping', () => {
    const exact = exactUniform()
    const approximate = approximateUniform()
    const cellStore = { tags: [ValueTag.Empty, ValueTag.Number], numbers: [0, 2] }

    expect(canSkipUniformApproximateNumericTailWrite(approximate, 4, 2, 5, 6)).toBe(true)
    expect(canSkipUniformApproximateNumericTailWrite(approximate, 3, 2, 5, 6)).toBe(false)
    expect(canSkipUniformApproximateNumericTailWrite(approximateUniform({ start: 5, step: -1, matchMode: -1 }), 4, 3, 1, 0)).toBe(true)
    expect(canSkipUniformApproximateNumericTailWriteFromCurrentResult(cellStore, 1, approximate, 4, 5, 6)).toBe(true)
    expect(canSkipUniformApproximateNumericTailWriteFromCurrentResult(cellStore, 1, approximate, 4, 6, 5)).toBe(false)
    expect(canSkipUniformExactNumericTailWriteFromCurrentResult(cellStore, 1, exact, 4, 5, 6)).toBe(true)
    expect(canSkipUniformExactNumericTailWriteFromCurrentResult(cellStore, 1, exact, 3, 5, 6)).toBe(false)
    expect(
      canSkipUniformExactNumericTailWriteFromCurrentResult(
        cellStore,
        1,
        exactUniform({ tailPatch: { row: 4, oldNumeric: 5, newNumeric: 6, columnVersion: 4 } }),
        4,
        5,
        6,
      ),
    ).toBe(false)

    expect(directLookupRowBounds(exact)).toEqual({ rowStart: 0, rowEnd: 4 })
    expect(
      directLookupRowBounds({
        kind: 'exact',
        operandCellIndex: 1,
        searchMode: 1,
        prepared: {
          sheetName: 'Sheet1',
          rowStart: 2,
          rowEnd: 8,
          col: 3,
          length: 7,
          columnVersion: 1,
          structureVersion: 1,
          sheetColumnVersions: new Uint32Array([1]),
          comparableKind: 'numeric',
          uniformStart: undefined,
          uniformStep: undefined,
          firstPositions: new Map(),
          lastPositions: new Map(),
          firstNumericPositions: undefined,
          lastNumericPositions: undefined,
          firstTextPositions: undefined,
          lastTextPositions: undefined,
        },
      }),
    ).toEqual({ rowStart: 2, rowEnd: 8 })

    expect(cellRange('Sheet1', 'B2')).toEqual({ sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'B2' })
    expect(
      rangesIntersect(
        { sheetName: 'Sheet1', startAddress: 'C3', endAddress: 'A1' },
        { sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'D4' },
      ),
    ).toBe(true)
    expect(
      rangesIntersect(
        { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A2' },
        { sheetName: 'Other', startAddress: 'A1', endAddress: 'A2' },
      ),
    ).toBe(false)
    expect(
      withOptionalLookupStringIds({
        sheetName: 'Sheet1',
        row: 1,
        col: 2,
        oldValue: { tag: ValueTag.String, value: 'a' },
        newValue: { tag: ValueTag.String, value: 'b' },
        oldStringId: undefined,
        newStringId: 4,
        inputCellIndex: 9,
      }),
    ).toEqual({
      sheetName: 'Sheet1',
      row: 1,
      col: 2,
      oldValue: { tag: ValueTag.String, value: 'a' },
      newValue: { tag: ValueTag.String, value: 'b' },
      newStringId: 4,
      inputCellIndex: 9,
    })

    const registry = new Map<string | number, Set<number>>([
      ['a', new Set([1, 2])],
      ['b', new Set([2, 3])],
    ])
    expect(collectTrackedDependents(registry, ['a', 'b', 'missing'])).toEqual([1, 2, 3])
    expect(lookupImpactCacheKey(7, 4)).toBe('7:4')
    expect(aggregateColumnDependencyKey(2, 5)).toBe(2 * 16_384 + 5)

    expect(Array.from(mergeChangedCellIndices([], [1, 2]))).toEqual([1, 2])
    expect(Array.from(mergeChangedCellIndices([1, 2], []))).toEqual([1, 2])
    expect(Array.from(mergeChangedCellIndices([1], [1]))).toEqual([1])
    expect(Array.from(mergeChangedCellIndices([1], [2]))).toEqual([1, 2])
    expect(Array.from(mergeChangedCellIndices([1, 2], [2, 3]))).toEqual([1, 2, 3])
    expect(Array.from(composeSingleDisjointExplicitEventChanges(5, new Uint32Array()))).toEqual([5])
    expect(Array.from(composeSingleDisjointExplicitEventChanges(5, new Uint32Array([6, 7])))).toEqual([5, 6, 7])

    const tagged = new Uint32Array([1, 2])
    tagTrustedPhysicalTrackedChanges(tagged, 9, 1)
    expect(Reflect.get(tagged, '__biligTrackedPhysicalSheetId')).toBe(9)
    expect(Reflect.get(tagged, '__biligTrackedPhysicalSortedSliceSplit')).toBe(1)
    expect(makeExistingNumericMutationResult(tagged, 1)).toEqual({ changedCellIndices: tagged, explicitChangedCount: 1 })
    expect(makeCompactExistingNumericMutationResult(1, undefined, 1)).toEqual({
      firstChangedCellIndex: 1,
      changedCellCount: 1,
      explicitChangedCount: 1,
    })
    expect(makeCompactExistingNumericMutationResult(1, 2, 1, 12, { row: 3, col: 4 })).toEqual({
      firstChangedCellIndex: 1,
      secondChangedCellIndex: 2,
      changedCellCount: 2,
      explicitChangedCount: 1,
      secondChangedNumericValue: 12,
      secondChangedRow: 3,
      secondChangedCol: 4,
    })

    expect(() => throwProtectionBlocked('locked')).toThrow('Workbook protection blocks this change: locked')
  })

  it('evaluates direct criteria and scalar helper decisions', () => {
    const criteria = {
      aggregateKind: 'sum',
      aggregateRange: { regionId: 1, sheetName: 'Sheet1', rowStart: 1, rowEnd: 3, col: 4, length: 3 },
      criteriaPairs: [
        {
          range: { regionId: 2, sheetName: 'Sheet1', rowStart: 1, rowEnd: 3, col: 0, length: 3 },
          criterion: { kind: 'literal', value: { tag: ValueTag.String, value: 'A' } },
        },
      ],
    } satisfies RuntimeDirectCriteriaDescriptor

    expect(directCriteriaTouchesPoint(criteria, { sheetName: 'Sheet1', row: 2, col: 4 })).toBe(true)
    expect(directCriteriaTouchesPoint(criteria, { sheetName: 'Sheet1', row: 2, col: 0 })).toBe(true)
    expect(directCriteriaTouchesPoint(criteria, { sheetName: 'Sheet1', row: 4, col: 0 })).toBe(false)
    expect(directCriteriaTouchesPoint(criteria, { sheetName: 'Other', row: 2, col: 0 })).toBe(false)

    const indexCollection = new DirectFormulaIndexCollection()
    indexCollection.appendConstantDelta(new Uint32Array([1, 2]), 3, 'scalar')
    expect(directFormulaChangesAreDisjointFromInputs(new Uint32Array([3, 4]), 2, indexCollection)).toBe(true)
    expect(directFormulaChangesAreDisjointFromInputs(new Uint32Array([1, 4]), 2, indexCollection)).toBe(false)
    expect(
      canEvaluatePostRecalcDirectFormulasWithoutKernel(
        new Map([
          [1, { directScalar: {} }],
          [2, { directAggregate: {} }],
        ]),
        indexCollection,
      ),
    ).toBe(true)
    expect(
      canEvaluatePostRecalcDirectFormulasWithoutKernel(
        new Map([
          [1, { directScalar: {} }],
          [2, {}],
        ]),
        indexCollection,
      ),
    ).toBe(false)
    expect(canEvaluatePostRecalcDirectFormulasWithoutKernel(new Map(), new DirectFormulaIndexCollection())).toBe(false)

    const counters = {
      directAggregateDeltaOnlyRecalcSkips: 0,
      directScalarDeltaOnlyRecalcSkips: 0,
    }
    countDirectFormulaDeltaSkip(
      new Map([
        [1, { directCriteria: {} }],
        [2, { directScalar: {} }],
      ]),
      indexCollection,
      counters,
    )
    expect(counters).toEqual({
      directAggregateDeltaOnlyRecalcSkips: 1,
      directScalarDeltaOnlyRecalcSkips: 1,
    })

    expect(singleInputAffineDirectScalar({ kind: 'abs', operand: { kind: 'cell', cellIndex: 1 } }, 1)).toBeNull()
    expect(
      singleInputAffineDirectScalar(
        { kind: 'binary', operator: '+', left: { kind: 'cell', cellIndex: 1 }, right: { kind: 'literal-number', value: 2 } },
        1,
      ),
    ).toEqual({ scale: 1, offset: 2 })
    expect(
      singleInputAffineDirectScalar(
        { kind: 'binary', operator: '-', left: { kind: 'cell', cellIndex: 1 }, right: { kind: 'literal-number', value: 2 } },
        1,
      ),
    ).toEqual({ scale: 1, offset: -2 })
    expect(
      singleInputAffineDirectScalar(
        { kind: 'binary', operator: '*', left: { kind: 'cell', cellIndex: 1 }, right: { kind: 'literal-number', value: 2 } },
        1,
      ),
    ).toEqual({ scale: 2, offset: 0 })
    expect(
      singleInputAffineDirectScalar(
        { kind: 'binary', operator: '/', left: { kind: 'cell', cellIndex: 1 }, right: { kind: 'literal-number', value: 2 } },
        1,
      ),
    ).toEqual({ scale: 0.5, offset: 0 })
    expect(
      singleInputAffineDirectScalar(
        { kind: 'binary', operator: '/', left: { kind: 'cell', cellIndex: 1 }, right: { kind: 'literal-number', value: 0 } },
        1,
      ),
    ).toBeNull()
    expect(
      singleInputAffineDirectScalar(
        { kind: 'binary', operator: '+', left: { kind: 'literal-number', value: 2 }, right: { kind: 'cell', cellIndex: 1 } },
        1,
      ),
    ).toEqual({ scale: 1, offset: 2 })
    expect(
      singleInputAffineDirectScalar(
        { kind: 'binary', operator: '-', left: { kind: 'literal-number', value: 2 }, right: { kind: 'cell', cellIndex: 1 } },
        1,
      ),
    ).toEqual({ scale: -1, offset: 2 })
    expect(
      singleInputAffineDirectScalar(
        { kind: 'binary', operator: '*', left: { kind: 'literal-number', value: 2 }, right: { kind: 'cell', cellIndex: 1 } },
        1,
      ),
    ).toEqual({ scale: 2, offset: 0 })
    expect(
      singleInputAffineDirectScalar(
        { kind: 'binary', operator: '/', left: { kind: 'literal-number', value: 2 }, right: { kind: 'cell', cellIndex: 1 } },
        1,
      ),
    ).toBeNull()

    expect(rowPairDirectScalarCode(binaryScalar('+', 1, 2), 1, 2)).toBe(ROW_PAIR_LEFT_PLUS_RIGHT)
    expect(rowPairDirectScalarCode(binaryScalar('-', 1, 2), 1, 2)).toBe(ROW_PAIR_LEFT_MINUS_RIGHT)
    expect(rowPairDirectScalarCode(binaryScalar('*', 1, 2), 1, 2)).toBe(ROW_PAIR_LEFT_TIMES_RIGHT)
    expect(rowPairDirectScalarCode(binaryScalar('/', 1, 2), 1, 2)).toBe(ROW_PAIR_LEFT_DIV_RIGHT)
    expect(rowPairDirectScalarCode(binaryScalar('-', 2, 1), 1, 2)).toBe(ROW_PAIR_RIGHT_MINUS_LEFT)
    expect(rowPairDirectScalarCode(binaryScalar('/', 2, 1), 1, 2)).toBe(ROW_PAIR_RIGHT_DIV_LEFT)
    expect(rowPairDirectScalarCode({ kind: 'abs', operand: { kind: 'cell', cellIndex: 1 } }, 1, 2)).toBe(0)
    expect(rowPairDirectScalarCode(binaryScalar('+', 1, 3), 1, 2)).toBe(0)
    expect(evaluateRowPairDirectScalarCode(ROW_PAIR_LEFT_PLUS_RIGHT, 8, 2)).toBe(10)
    expect(evaluateRowPairDirectScalarCode(ROW_PAIR_LEFT_MINUS_RIGHT, 8, 2)).toBe(6)
    expect(evaluateRowPairDirectScalarCode(ROW_PAIR_RIGHT_MINUS_LEFT, 8, 2)).toBe(-6)
    expect(evaluateRowPairDirectScalarCode(ROW_PAIR_LEFT_TIMES_RIGHT, 8, 2)).toBe(16)
    expect(evaluateRowPairDirectScalarCode(ROW_PAIR_LEFT_DIV_RIGHT, 8, 2)).toBe(4)
    expect(evaluateRowPairDirectScalarCode(ROW_PAIR_LEFT_DIV_RIGHT, 8, 0)).toBeUndefined()
    expect(evaluateRowPairDirectScalarCode(ROW_PAIR_RIGHT_DIV_LEFT, 8, 2)).toBe(0.25)
    expect(evaluateRowPairDirectScalarCode(ROW_PAIR_RIGHT_DIV_LEFT, 0, 2)).toBeUndefined()
    expect(evaluateRowPairDirectScalarCode(0, 8, 2)).toBeUndefined()
  })

  it('exercises operation-service runtime lookup and protection closures over real engine state', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'operation-runtime-hook-coverage', useColumnIndex: true })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A5' }, [[1], [2], [3], [4], [5]])
    engine.setRangeValues({ sheetName: 'Sheet1', startAddress: 'B1', endAddress: 'B5' }, [['A'], ['B'], ['C'], ['D'], ['E']])
    engine.setCellValue('Sheet1', 'D1', 3)
    engine.setCellValue('Sheet1', 'D2', 3.5)
    engine.setCellValue('Sheet1', 'D3', 'C')
    engine.setCellFormula('Sheet1', 'E1', 'MATCH(D1,A1:A5,0)')
    engine.setCellFormula('Sheet1', 'E2', 'MATCH(D2,A1:A5,1)')
    engine.setCellFormula('Sheet1', 'E3', 'MATCH(D3,B1:B5,1)')
    engine.setCellFormula('Sheet1', 'F1', 'SUM(A1:A5)')
    engine.setCellFormula('Sheet1', 'F2', 'SUMIF(B1:B5,"C",A1:A5)')

    const hooks = runtimeHooks(engine)
    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const a1 = engine.workbook.getCellIndex('Sheet1', 'A1')!
    const b1 = engine.workbook.getCellIndex('Sheet1', 'B1')!
    const d1 = engine.workbook.getCellIndex('Sheet1', 'D1')!
    const e1 = engine.workbook.getCellIndex('Sheet1', 'E1')!
    const e2 = engine.workbook.getCellIndex('Sheet1', 'E2')!
    const e3 = engine.workbook.getCellIndex('Sheet1', 'E3')!
    const f1 = engine.workbook.getCellIndex('Sheet1', 'F1')!
    const f2 = engine.workbook.getCellIndex('Sheet1', 'F2')!

    const readCellValueForLookup = hookFunction<[number | undefined], { value: unknown; stringId: number | undefined }>(
      hooks,
      'readCellValueForLookup',
    )
    const readCellValueAtForLookup = hookFunction<[string, number, number], { value: unknown; stringId: number | undefined }>(
      hooks,
      'readCellValueAtForLookup',
    )
    const readApproximateNumericValueForLookup = hookFunction<[number | undefined], number | undefined>(
      hooks,
      'readApproximateNumericValueForLookup',
    )
    const readApproximateNumericValueAtForLookup = hookFunction<[string, number, number], number | undefined>(
      hooks,
      'readApproximateNumericValueAtForLookup',
    )
    const readExactNumericValueForLookup = hookFunction<[number | undefined], number | undefined>(hooks, 'readExactNumericValueForLookup')

    expect(readCellValueForLookup(undefined).value).toEqual({ tag: ValueTag.Empty })
    expect(readCellValueForLookup(a1).value).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(readCellValueForLookup(b1).value).toMatchObject({ tag: ValueTag.String, value: 'A' })
    expect(readCellValueAtForLookup('Sheet1', 0, 0).value).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(readCellValueAtForLookup('Missing', 0, 0).value).toEqual({ tag: ValueTag.Empty })
    expect(readApproximateNumericValueForLookup(undefined)).toBe(0)
    expect(readApproximateNumericValueForLookup(a1)).toBe(1)
    expect(readApproximateNumericValueForLookup(b1)).toBeUndefined()
    expect(readApproximateNumericValueAtForLookup('Sheet1', 2, 0)).toBe(3)
    expect(readApproximateNumericValueAtForLookup('Missing', 2, 0)).toBe(0)
    expect(readExactNumericValueForLookup(undefined)).toBeUndefined()
    expect(readExactNumericValueForLookup(d1)).toBe(3)
    expect(readExactNumericValueForLookup(b1)).toBeUndefined()

    const isLocallySortedNumericWrite = hookFunction<[string, number, number, number, number, 1 | -1, number], boolean>(
      hooks,
      'isLocallySortedNumericWrite',
    )
    const isLocallySortedTextWrite = hookFunction<[string, number, number, number, number, 1 | -1, string], boolean>(
      hooks,
      'isLocallySortedTextWrite',
    )
    expect(isLocallySortedNumericWrite('Sheet1', 2, 0, 0, 4, 1, 3.25)).toBe(true)
    expect(isLocallySortedNumericWrite('Sheet1', 2, 0, 0, 4, 1, 6)).toBe(false)
    expect(isLocallySortedTextWrite('Sheet1', 2, 1, 0, 4, 1, 'CC')).toBe(true)
    expect(isLocallySortedTextWrite('Sheet1', 2, 1, 0, 4, 1, 'Z')).toBe(false)

    const planSingleExactLookupNumericColumnWrite = hookFunction<[number, number, number, number], { handled: boolean }>(
      hooks,
      'planSingleExactLookupNumericColumnWrite',
    )
    const planSingleApproximateLookupNumericColumnWrite = hookFunction<
      [number, string, number, number, number, number],
      { handled: boolean }
    >(hooks, 'planSingleApproximateLookupNumericColumnWrite')
    const planExactLookupNumericColumnWrite = hookFunction<[number, number, number, number, number], { handled: boolean }>(
      hooks,
      'planExactLookupNumericColumnWrite',
    )
    const planApproximateLookupNumericColumnWrite = hookFunction<[number, string, number, number, number, number], { handled: boolean }>(
      hooks,
      'planApproximateLookupNumericColumnWrite',
    )
    const canSkipApproximateLookupNewNumericColumnWrite = hookFunction<[number, number, number], boolean>(
      hooks,
      'canSkipApproximateLookupNewNumericColumnWrite',
    )

    expect(planSingleExactLookupNumericColumnWrite(e1, 4, 5, 99).handled).toBe(true)
    expect(planSingleExactLookupNumericColumnWrite(e1, 2, 3, 99).handled).toBe(false)
    expect(planSingleExactLookupNumericColumnWrite(e2, 4, 5, 99).handled).toBe(false)
    expect(planSingleApproximateLookupNumericColumnWrite(e2, 'Sheet1', 4, 0, 5, 6).handled).toBe(true)
    expect(planSingleApproximateLookupNumericColumnWrite(e2, 'Sheet1', 2, 0, 3, 0).handled).toBe(false)
    expect(planSingleApproximateLookupNumericColumnWrite(e1, 'Sheet1', 4, 0, 5, 6).handled).toBe(false)
    expect(planExactLookupNumericColumnWrite(sheetId, 0, 4, 5, 99).handled).toBe(true)
    expect(planApproximateLookupNumericColumnWrite(sheetId, 'Sheet1', 0, 4, 5, 6).handled).toBe(true)
    expect(canSkipApproximateLookupNewNumericColumnWrite(sheetId, 0, 7)).toBe(true)
    expect(canSkipApproximateLookupNewNumericColumnWrite(sheetId, 0, 2)).toBe(false)

    const canSkipApproximateLookupDirtyMark = hookFunction<
      [
        unknown,
        {
          sheetName: string
          row: number
          col: number
          oldValue: unknown
          newValue: unknown
        },
      ],
      boolean
    >(hooks, 'canSkipApproximateLookupDirtyMark')
    expect(
      canSkipApproximateLookupDirtyMark(engine.state.formulas.get(e2)?.directLookup, {
        sheetName: 'Sheet1',
        row: 4,
        col: 0,
        oldValue: { tag: ValueTag.Number, value: 5 },
        newValue: { tag: ValueTag.Number, value: 6 },
      }),
    ).toBe(true)
    expect(
      canSkipApproximateLookupDirtyMark(engine.state.formulas.get(e3)?.directLookup, {
        sheetName: 'Sheet1',
        row: 4,
        col: 1,
        oldValue: { tag: ValueTag.String, value: 'E' },
        newValue: { tag: ValueTag.String, value: 'F' },
      }),
    ).toBe(true)

    const collectAffectedDirectRangeDependents = hookFunction<[{ sheetName: string; row: number; col: number }], number[]>(
      hooks,
      'collectAffectedDirectRangeDependents',
    )
    const collectSingleAffectedDirectRangeDependent = hookFunction<
      [{ sheetName: string; row: number; col: number; sheetId?: number }],
      number
    >(hooks, 'collectSingleAffectedDirectRangeDependent')
    const tryDirectCriteriaSumDelta = hookFunction<
      [
        unknown,
        {
          sheetName: string
          row: number
          col: number
          oldValue?: unknown
          newValue?: unknown
        },
      ],
      number | undefined
    >(hooks, 'tryDirectCriteriaSumDelta')

    expect(collectAffectedDirectRangeDependents({ sheetName: 'Sheet1', row: 2, col: 0 })).toContain(f1)
    expect(collectSingleAffectedDirectRangeDependent({ sheetName: 'Missing', row: 2, col: 0 })).toBe(-1)
    expect(
      tryDirectCriteriaSumDelta(engine.state.formulas.get(f2)?.directCriteria, {
        sheetName: 'Sheet1',
        row: 2,
        col: 0,
        oldValue: { tag: ValueTag.Number, value: 3 },
        newValue: { tag: ValueTag.Number, value: 30 },
      }),
    ).toBe(27)
    expect(
      tryDirectCriteriaSumDelta(engine.state.formulas.get(f2)?.directCriteria, {
        sheetName: 'Sheet1',
        row: 1,
        col: 0,
        oldValue: { tag: ValueTag.Number, value: 2 },
        newValue: { tag: ValueTag.Number, value: 20 },
      }),
    ).toBe(0)

    const rangeIsProtected = hookFunction<[{ sheetName: string; startAddress: string; endAddress: string }], boolean>(
      hooks,
      'rangeIsProtected',
    )
    const sheetHasProtection = hookFunction<[string], boolean>(hooks, 'sheetHasProtection')
    const assertProtectionAllowsOp = hookFunction<[unknown], void>(hooks, 'assertProtectionAllowsOp')

    expect(sheetHasProtection('Sheet1')).toBe(false)
    expect(rangeIsProtected({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' })).toBe(false)
    engine.setRangeProtection({ id: 'protect-a1', range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' } })
    expect(rangeIsProtected({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B1' })).toBe(true)
    expect(() => assertProtectionAllowsOp({ kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 1 })).toThrow(
      /Workbook protection blocks this change/,
    )
    engine.setSheetProtection({ sheetName: 'Sheet1' })
    expect(sheetHasProtection('Sheet1')).toBe(true)
    expect(() => assertProtectionAllowsOp({ kind: 'deleteSheet', name: 'Sheet1' })).toThrow(/Workbook protection blocks this change/)
    expect(() => assertProtectionAllowsOp({ kind: 'setWorkbookMetadata', key: 'owner', value: 'ops' })).not.toThrow()
  })

  it('maps operation entity keys and stale-order barriers for every operation family', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'operation-order-hook-coverage' })
    await engine.ready()
    const hooks = runtimeHooks(engine)
    const entityKeyForOp = hookFunction<[EngineOp], string>(hooks, 'entityKeyForOp')
    const sheetDeleteBarrierForOp = hookFunction<[EngineOp], undefined>(hooks, 'sheetDeleteBarrierForOp')
    const shouldApplyOp = hookFunction<[EngineOp, { counter: number; replicaId: string; batchId: string; opIndex: number }], boolean>(
      hooks,
      'shouldApplyOp',
    )
    const range = { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B2' }
    const opsWithKeys: Array<[EngineOp, string]> = [
      [{ kind: 'upsertWorkbook', name: 'Book' }, 'workbook'],
      [{ kind: 'setWorkbookMetadata', key: 'owner', value: 'ops' }, 'workbook-meta:owner'],
      [{ kind: 'setCalculationSettings', settings: { mode: 'manual' } }, 'workbook-calc'],
      [{ kind: 'setVolatileContext', context: { recalcEpoch: 1 } }, 'workbook-volatile'],
      [{ kind: 'upsertSheet', name: 'Sheet1', order: 0 }, 'sheet:Sheet1'],
      [{ kind: 'renameSheet', oldName: 'Sheet1', newName: 'Renamed' }, 'sheet:Sheet1'],
      [{ kind: 'deleteSheet', name: 'Sheet1' }, 'sheet:Sheet1'],
      [{ kind: 'insertRows', sheetName: 'Sheet1', start: 0, count: 1 }, 'row-structure:Sheet1'],
      [{ kind: 'deleteRows', sheetName: 'Sheet1', start: 0, count: 1 }, 'row-structure:Sheet1'],
      [{ kind: 'moveRows', sheetName: 'Sheet1', start: 0, count: 1, target: 2 }, 'row-structure:Sheet1'],
      [{ kind: 'insertColumns', sheetName: 'Sheet1', start: 0, count: 1 }, 'column-structure:Sheet1'],
      [{ kind: 'deleteColumns', sheetName: 'Sheet1', start: 0, count: 1 }, 'column-structure:Sheet1'],
      [{ kind: 'moveColumns', sheetName: 'Sheet1', start: 0, count: 1, target: 2 }, 'column-structure:Sheet1'],
      [{ kind: 'updateRowMetadata', sheetName: 'Sheet1', start: 1, count: 2, size: 20, hidden: false }, 'row-meta:Sheet1:1:2'],
      [{ kind: 'updateColumnMetadata', sheetName: 'Sheet1', start: 1, count: 2, size: 80, hidden: false }, 'column-meta:Sheet1:1:2'],
      [{ kind: 'setFreezePane', sheetName: 'Sheet1', rows: 1, cols: 1 }, 'freeze:Sheet1'],
      [{ kind: 'clearFreezePane', sheetName: 'Sheet1' }, 'freeze:Sheet1'],
      [{ kind: 'setSheetProtection', protection: { sheetName: 'Sheet1' } }, 'sheet-protection:Sheet1'],
      [{ kind: 'clearSheetProtection', sheetName: 'Sheet1' }, 'sheet-protection:Sheet1'],
      [{ kind: 'setFilter', sheetName: 'Sheet1', range }, 'filter:Sheet1:A1:B2'],
      [{ kind: 'clearFilter', sheetName: 'Sheet1', range }, 'filter:Sheet1:A1:B2'],
      [{ kind: 'setSort', sheetName: 'Sheet1', range, keys: [{ keyAddress: 'A1', direction: 'asc' }] }, 'sort:Sheet1:A1:B2'],
      [{ kind: 'clearSort', sheetName: 'Sheet1', range }, 'sort:Sheet1:A1:B2'],
      [{ kind: 'setDataValidation', validation: { id: 'v1', range, rule: { kind: 'list', values: ['A'] } } }, 'validation:Sheet1:A1:B2'],
      [{ kind: 'clearDataValidation', sheetName: 'Sheet1', range }, 'validation:Sheet1:A1:B2'],
      [
        { kind: 'upsertConditionalFormat', format: { id: 'cf1', range, rule: { kind: 'textContains', text: 'A' }, style: {} } },
        'conditional-format:cf1',
      ],
      [{ kind: 'deleteConditionalFormat', id: 'cf1', sheetName: 'Sheet1' }, 'conditional-format:cf1'],
      [{ kind: 'upsertRangeProtection', protection: { id: 'rp1', range } }, 'range-protection:rp1'],
      [{ kind: 'deleteRangeProtection', id: 'rp1', sheetName: 'Sheet1' }, 'range-protection:rp1'],
      [{ kind: 'upsertCommentThread', thread: { threadId: 't1', sheetName: 'Sheet1', address: 'A1', comments: [] } }, 'comment:Sheet1!A1'],
      [{ kind: 'deleteCommentThread', sheetName: 'Sheet1', address: 'A1' }, 'comment:Sheet1!A1'],
      [{ kind: 'upsertNote', note: { sheetName: 'Sheet1', address: 'A1', text: 'note' } }, 'note:Sheet1!A1'],
      [{ kind: 'deleteNote', sheetName: 'Sheet1', address: 'A1' }, 'note:Sheet1!A1'],
      [{ kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1', value: 1 }, 'cell:Sheet1!A1'],
      [{ kind: 'setCellFormula', sheetName: 'Sheet1', address: 'A1', formula: '1+1' }, 'cell:Sheet1!A1'],
      [{ kind: 'setCellFormat', sheetName: 'Sheet1', address: 'A1', format: '0.00' }, 'format:Sheet1!A1'],
      [{ kind: 'clearCell', sheetName: 'Sheet1', address: 'A1' }, 'cell:Sheet1!A1'],
      [{ kind: 'upsertCellStyle', style: { id: 'style-1' } }, 'style:style-1'],
      [{ kind: 'upsertCellNumberFormat', format: { id: 'fmt-1', code: '0.00', kind: 'number' } }, 'number-format:fmt-1'],
      [{ kind: 'setStyleRange', range, styleId: 'style-1' }, 'style-range:Sheet1:A1:B2'],
      [{ kind: 'setFormatRange', range, formatId: 'fmt-1' }, 'format-range:Sheet1:A1:B2'],
      [{ kind: 'upsertDefinedName', name: ' Revenue ', value: { kind: 'formula', formula: 'Sheet1!A1' } }, 'defined-name:REVENUE'],
      [{ kind: 'deleteDefinedName', name: ' Revenue ' }, 'defined-name:REVENUE'],
      [
        {
          kind: 'upsertTable',
          table: {
            name: ' Table1 ',
            sheetName: 'Sheet1',
            startAddress: 'A1',
            endAddress: 'B2',
            columnNames: ['A', 'B'],
            headerRow: true,
            totalsRow: false,
          },
        },
        'table:TABLE1',
      ],
      [{ kind: 'deleteTable', name: ' Table1 ' }, 'table:TABLE1'],
      [{ kind: 'upsertSpillRange', sheetName: 'Sheet1', address: 'C1', rows: 2, cols: 2 }, 'spill:Sheet1!C1'],
      [{ kind: 'deleteSpillRange', sheetName: 'Sheet1', address: 'C1' }, 'spill:Sheet1!C1'],
      [
        {
          kind: 'upsertPivotTable',
          name: 'Pivot1',
          sheetName: 'Sheet1',
          address: 'D1',
          source: range,
          groupBy: ['A'],
          values: [{ sourceColumn: 'B', summarizeBy: 'sum', outputLabel: 'B' }],
          rows: 3,
          cols: 2,
        },
        'pivot:Sheet1!D1',
      ],
      [{ kind: 'deletePivotTable', sheetName: 'Sheet1', address: 'D1' }, 'pivot:Sheet1!D1'],
      [
        {
          kind: 'upsertChart',
          chart: { id: ' chart-1 ', sheetName: 'Sheet1', address: 'F1', source: range, chartType: 'line', rows: 3, cols: 3 },
        },
        'chart:CHART-1',
      ],
      [{ kind: 'deleteChart', id: ' chart-1 ' }, 'chart:CHART-1'],
      [
        {
          kind: 'upsertImage',
          image: { id: ' image-1 ', sheetName: 'Sheet1', address: 'F5', sourceUrl: 'https://example.com/i.png', rows: 2, cols: 2 },
        },
        'image:IMAGE-1',
      ],
      [{ kind: 'deleteImage', id: ' image-1 ' }, 'image:IMAGE-1'],
      [
        { kind: 'upsertShape', shape: { id: ' shape-1 ', sheetName: 'Sheet1', address: 'F8', shapeType: 'rectangle', rows: 2, cols: 2 } },
        'shape:SHAPE-1',
      ],
      [{ kind: 'deleteShape', id: ' shape-1 ' }, 'shape:SHAPE-1'],
    ]

    expect(opsWithKeys.map(([op]) => entityKeyForOp(op))).toEqual(opsWithKeys.map(([, key]) => key))
    expect(opsWithKeys.map(([op]) => sheetDeleteBarrierForOp(op))).toHaveLength(opsWithKeys.length)

    const lowOrder = { counter: 0, replicaId: 'r', batchId: 'r:0', opIndex: 0 }
    const highOrder = { counter: 999, replicaId: 'r', batchId: 'r:999', opIndex: 0 }
    expect(shouldApplyOp({ kind: 'upsertWorkbook', name: 'Book' }, lowOrder)).toBe(true)
    engine.applyOps(
      [
        { kind: 'upsertWorkbook', name: 'Applied' },
        { kind: 'upsertSheet', name: 'Deleted', order: 0 },
        { kind: 'deleteSheet', name: 'Deleted' },
      ],
      {
        source: 'local',
      },
    )
    expect(shouldApplyOp({ kind: 'upsertWorkbook', name: 'Stale' }, lowOrder)).toBe(false)
    expect(shouldApplyOp({ kind: 'upsertWorkbook', name: 'Fresh' }, highOrder)).toBe(true)
    expect(sheetDeleteBarrierForOp({ kind: 'setCellValue', sheetName: 'Deleted', address: 'A1', value: 1 })).toBeDefined()
    expect(shouldApplyOp({ kind: 'setCellValue', sheetName: 'Deleted', address: 'A1', value: 1 }, lowOrder)).toBe(false)
    expect(shouldApplyOp({ kind: 'setCellValue', sheetName: 'Deleted', address: 'A1', value: 1 }, highOrder)).toBe(true)
  })
})
