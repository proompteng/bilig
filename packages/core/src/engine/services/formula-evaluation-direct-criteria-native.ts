import { compileCriteriaMatcher } from '@bilig/formula'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { addEngineCounter } from '../../perf/engine-counters.js'
import type { EngineRuntimeState, RuntimeDirectCriteriaDescriptor } from '../runtime-state.js'
import type { CriterionRangeDescriptor, CriterionRangeMatch, CriterionRangePair } from './criterion-range-cache-service.js'
import type { EngineRuntimeColumnStoreService } from './runtime-column-store-service.js'

const MIN_NATIVE_DIRECT_CRITERIA_MATCHED_AGGREGATE_ROWS = 64
const MIN_NATIVE_DIRECT_CRITERIA_PREDICATE_AGGREGATE_ROWS = 65_536
const MAX_NATIVE_DIRECT_CRITERIA_PREDICATE_PAIRS = 4
const NATIVE_DIRECT_AGGREGATE_OP_SUM = 1
const NATIVE_DIRECT_AGGREGATE_OP_AVERAGE = 2
const NATIVE_DIRECT_AGGREGATE_OP_COUNT = 3
const NATIVE_DIRECT_AGGREGATE_OP_MIN = 4
const NATIVE_DIRECT_AGGREGATE_OP_MAX = 5
const NATIVE_DIRECT_CRITERIA_OP_EQ = 0
const NATIVE_DIRECT_CRITERIA_OP_NE = 1
const NATIVE_DIRECT_CRITERIA_OP_GT = 2
const NATIVE_DIRECT_CRITERIA_OP_GTE = 3
const NATIVE_DIRECT_CRITERIA_OP_LT = 4
const NATIVE_DIRECT_CRITERIA_OP_LTE = 5

export function tryEvaluateNativeDirectCriteriaPredicateAggregate(
  args: {
    readonly state: Pick<EngineRuntimeState, 'wasm' | 'counters'>
    readonly runtimeColumnStore: EngineRuntimeColumnStoreService
  },
  input: {
    readonly aggregateKind: RuntimeDirectCriteriaDescriptor['aggregateKind']
    readonly aggregateRange: CriterionRangeDescriptor | undefined
    readonly criteriaPairs: readonly CriterionRangePair[]
    readonly shouldUseSharedCriteriaCache?: () => boolean
  },
): CellValue | undefined {
  const aggregateKind = nativeDirectCriteriaAggregateKind(input.aggregateKind)
  if (
    aggregateKind === undefined ||
    input.criteriaPairs.length === 0 ||
    input.criteriaPairs.length > MAX_NATIVE_DIRECT_CRITERIA_PREDICATE_PAIRS ||
    !args.state.wasm.initSyncIfPossible()
  ) {
    return undefined
  }

  const rowCount = input.criteriaPairs[0]?.range.length ?? 0
  if (
    rowCount < MIN_NATIVE_DIRECT_CRITERIA_PREDICATE_AGGREGATE_ROWS ||
    input.criteriaPairs.some((pair) => pair.range.length !== rowCount) ||
    (input.aggregateKind !== 'count' && input.aggregateRange === undefined) ||
    (input.aggregateRange !== undefined && input.aggregateRange.length !== rowCount)
  ) {
    return undefined
  }

  const loweredCriteria = input.criteriaPairs.map((pair) => lowerNativeNumericCriteria(pair.criteria))
  if (loweredCriteria.some((criterion) => criterion === undefined)) {
    return undefined
  }
  if (input.shouldUseSharedCriteriaCache?.()) {
    return undefined
  }

  const criteriaOps = new Uint8Array(loweredCriteria.length)
  const criteriaValues = new Float64Array(loweredCriteria.length)
  const criteriaTags = new Uint8Array(rowCount * loweredCriteria.length)
  const criteriaNumbers = new Float64Array(rowCount * loweredCriteria.length)
  for (let pairIndex = 0; pairIndex < input.criteriaPairs.length; pairIndex += 1) {
    const lowered = loweredCriteria[pairIndex]
    if (lowered === undefined) {
      return undefined
    }
    criteriaOps[pairIndex] = lowered.operator
    criteriaValues[pairIndex] = lowered.value
    const criteriaSlice = args.runtimeColumnStore.getColumnSlice(input.criteriaPairs[pairIndex]!.range)
    const offset = pairIndex * rowCount
    if (!copyNativeNumericCriteriaSlice(criteriaSlice.tags, criteriaSlice.numbers, criteriaTags, criteriaNumbers, offset)) {
      return undefined
    }
  }

  const aggregateSlice = input.aggregateRange === undefined ? undefined : args.runtimeColumnStore.getColumnSlice(input.aggregateRange)
  const outTags = new Uint8Array(1)
  const outNumbers = new Float64Array(1)
  const outErrors = new Uint16Array(1)
  if (
    !args.state.wasm.evalDirectCriteriaPredicateAggregateBatch({
      aggregateKind,
      rowCount,
      criteriaOps,
      criteriaValues,
      criteriaTags,
      criteriaNumbers,
      aggregateTags: aggregateSlice?.tags ?? new Uint8Array(0),
      aggregateNumbers: aggregateSlice?.numbers ?? new Float64Array(0),
      aggregateErrors: aggregateSlice?.errors ?? new Uint16Array(0),
      outTags,
      outNumbers,
      outErrors,
    })
  ) {
    return undefined
  }

  addEngineCounter(args.state.counters, 'nativeDirectCriteriaPredicateAggregateEvaluations')
  return decodeNativeDirectCriteriaAggregateResult(outTags, outNumbers, outErrors)
}

export function tryEvaluateNativeDirectCriteriaMatchedAggregate(
  args: {
    readonly state: Pick<EngineRuntimeState, 'wasm' | 'counters'>
    readonly runtimeColumnStore: EngineRuntimeColumnStoreService
  },
  input: {
    readonly aggregateKind: RuntimeDirectCriteriaDescriptor['aggregateKind']
    readonly aggregateRange: CriterionRangeDescriptor
    readonly matches: CriterionRangeMatch
  },
): CellValue | undefined {
  const aggregateKind = nativeDirectCriteriaAggregateKind(input.aggregateKind)
  if (
    aggregateKind === undefined ||
    input.matches.length < MIN_NATIVE_DIRECT_CRITERIA_MATCHED_AGGREGATE_ROWS ||
    !args.state.wasm.initSyncIfPossible()
  ) {
    return undefined
  }

  const aggregateSlice = args.runtimeColumnStore.getColumnSlice({
    sheetName: input.aggregateRange.sheetName,
    rowStart: input.aggregateRange.rowStart,
    rowEnd: input.aggregateRange.rowEnd,
    col: input.aggregateRange.col,
  })
  const outTags = new Uint8Array(1)
  const outNumbers = new Float64Array(1)
  const outErrors = new Uint16Array(1)
  const matchedRows =
    input.matches.rows.length === input.matches.length ? input.matches.rows : input.matches.rows.subarray(0, input.matches.length)

  if (
    !args.state.wasm.evalDirectCriteriaMatchedAggregateBatch({
      aggregateKinds: Uint8Array.of(aggregateKind),
      matchStarts: Uint32Array.of(0),
      matchLengths: Uint32Array.of(input.matches.length),
      matchedRows,
      aggregateTags: aggregateSlice.tags,
      aggregateNumbers: aggregateSlice.numbers,
      aggregateErrors: aggregateSlice.errors,
      outTags,
      outNumbers,
      outErrors,
    })
  ) {
    return undefined
  }

  addEngineCounter(args.state.counters, 'nativeDirectCriteriaAggregateEvaluations')
  return decodeNativeDirectCriteriaAggregateResult(outTags, outNumbers, outErrors)
}

function decodeNativeDirectCriteriaAggregateResult(
  outTags: Uint8Array,
  outNumbers: Float64Array,
  outErrors: Uint16Array,
): CellValue | undefined {
  const tag = (outTags[0] as ValueTag | undefined) ?? ValueTag.Empty
  if (tag === ValueTag.Number) {
    return { tag: ValueTag.Number, value: outNumbers[0] ?? 0 }
  }
  if (tag === ValueTag.Error) {
    return { tag: ValueTag.Error, code: (outErrors[0] as ErrorCode | undefined) ?? ErrorCode.None }
  }
  return undefined
}

function lowerNativeNumericCriteria(criteria: CellValue): { operator: number; value: number } | undefined {
  const compiled = compileCriteriaMatcher(criteria)
  if (compiled.wildcardPattern !== undefined) {
    return undefined
  }
  const operator = nativeDirectCriteriaOperator(compiled.operator)
  if (operator === undefined) {
    return undefined
  }
  if (compiled.operand.tag === ValueTag.Number) {
    if (!isNativeRawCriteriaNumberSafe(compiled.operand.value)) {
      return undefined
    }
    return { operator, value: compiled.operand.value }
  }
  if (compiled.operand.tag === ValueTag.Boolean) {
    return { operator, value: compiled.operand.value ? 1 : 0 }
  }
  return undefined
}

function copyNativeNumericCriteriaSlice(
  sourceTags: Uint8Array,
  sourceNumbers: Float64Array,
  targetTags: Uint8Array,
  targetNumbers: Float64Array,
  offset: number,
): boolean {
  for (let index = 0; index < sourceTags.length; index += 1) {
    const tag = (sourceTags[index] ?? ValueTag.Empty) as ValueTag
    const value = sourceNumbers[index] ?? 0
    if (tag === ValueTag.Number && !isNativeRawCriteriaNumberSafe(value)) {
      return false
    }
    targetTags[offset + index] = tag
    targetNumbers[offset + index] = value
  }
  return true
}

function isNativeRawCriteriaNumberSafe(value: number): boolean {
  return Number.isFinite(value) && Number.isInteger(value) && Math.abs(value) < 1_000_000_000_000_000
}

function nativeDirectCriteriaOperator(operator: ReturnType<typeof compileCriteriaMatcher>['operator']): number | undefined {
  switch (operator) {
    case '=':
      return NATIVE_DIRECT_CRITERIA_OP_EQ
    case '<>':
      return NATIVE_DIRECT_CRITERIA_OP_NE
    case '>':
      return NATIVE_DIRECT_CRITERIA_OP_GT
    case '>=':
      return NATIVE_DIRECT_CRITERIA_OP_GTE
    case '<':
      return NATIVE_DIRECT_CRITERIA_OP_LT
    case '<=':
      return NATIVE_DIRECT_CRITERIA_OP_LTE
  }
}

function nativeDirectCriteriaAggregateKind(kind: RuntimeDirectCriteriaDescriptor['aggregateKind']): number | undefined {
  switch (kind) {
    case 'sum':
      return NATIVE_DIRECT_AGGREGATE_OP_SUM
    case 'average':
      return NATIVE_DIRECT_AGGREGATE_OP_AVERAGE
    case 'count':
      return NATIVE_DIRECT_AGGREGATE_OP_COUNT
    case 'min':
      return NATIVE_DIRECT_AGGREGATE_OP_MIN
    case 'max':
      return NATIVE_DIRECT_AGGREGATE_OP_MAX
    case 'first':
      return undefined
  }
}
