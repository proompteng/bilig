import { compileCriteriaMatcher } from '@bilig/formula'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { addEngineCounter } from '../../perf/engine-counters.js'
import type { EngineRuntimeState, RuntimeDirectCriteriaDescriptor } from '../runtime-state.js'
import type { CriterionRangeDescriptor, CriterionRangeMatch, CriterionRangePair } from './criterion-range-cache-service.js'
import type { EngineRuntimeColumnStoreService, RuntimeColumnSlice } from './runtime-column-store-service.js'

const MIN_NATIVE_DIRECT_CRITERIA_MATCHED_AGGREGATE_ROWS = 64
const MIN_NATIVE_DIRECT_CRITERIA_PREDICATE_AGGREGATE_ROWS = 65_536
const MIN_NATIVE_DIRECT_CRITERIA_STRING_PREDICATE_AGGREGATE_ROWS = 512
const NATIVE_DIRECT_CRITERIA_PREDICATE_LAYOUT_CACHE_LIMIT = 32
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
const NATIVE_DIRECT_CRITERIA_KIND_NUMBER = 0
const NATIVE_DIRECT_CRITERIA_KIND_STRING_ID = 1

interface LoweredNativeCriteria {
  readonly operator: number
  readonly kind: number
  readonly value: number
  readonly stringId: number
}

export interface NativeDirectCriteriaPredicateLayout {
  readonly rowCount: number
  readonly hasStringCriteria: boolean
  readonly criteriaOps: Uint8Array
  readonly criteriaKinds: Uint8Array
  readonly criteriaValues: Float64Array
  readonly criteriaStringIds: Uint32Array
  readonly criteriaTags: Uint8Array
  readonly criteriaNumbers: Float64Array
  readonly criteriaStringIdsByRow: Uint32Array
}

export type NativeDirectCriteriaPredicateLayoutCache = Map<string, NativeDirectCriteriaPredicateLayout>

export function tryEvaluateNativeDirectCriteriaPredicateAggregate(
  args: {
    readonly state: Pick<EngineRuntimeState, 'wasm' | 'counters'>
    readonly runtimeColumnStore: EngineRuntimeColumnStoreService
  },
  input: {
    readonly aggregateKind: RuntimeDirectCriteriaDescriptor['aggregateKind']
    readonly aggregateRange: CriterionRangeDescriptor | undefined
    readonly criteriaPairs: readonly CriterionRangePair[]
    readonly criteriaLayoutCache?: NativeDirectCriteriaPredicateLayoutCache
    readonly criteriaLayoutCacheKey?: string
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
    input.criteriaPairs.some((pair) => pair.range.length !== rowCount) ||
    (input.aggregateKind !== 'count' && input.aggregateRange === undefined) ||
    (input.aggregateRange !== undefined && input.aggregateRange.length !== rowCount)
  ) {
    return undefined
  }

  const criteriaLayout = getOrBuildNativePredicateCriteriaLayout(
    {
      runtimeColumnStore: args.runtimeColumnStore,
    },
    {
      cache: input.criteriaLayoutCache,
      cacheKey: input.criteriaLayoutCacheKey,
      criteriaPairs: input.criteriaPairs,
      rowCount,
      useSharedCriteriaCache: input.shouldUseSharedCriteriaCache?.() ?? false,
    },
  )
  if (criteriaLayout === undefined) {
    return undefined
  }

  const aggregateSlice = input.aggregateRange === undefined ? undefined : args.runtimeColumnStore.getColumnSlice(input.aggregateRange)
  const outTags = new Uint8Array(1)
  const outNumbers = new Float64Array(1)
  const outErrors = new Uint16Array(1)
  if (
    !args.state.wasm.evalDirectCriteriaPredicateAggregateBatch({
      aggregateKind,
      rowCount,
      criteriaOps: criteriaLayout.criteriaOps,
      criteriaKinds: criteriaLayout.criteriaKinds,
      criteriaValues: criteriaLayout.criteriaValues,
      criteriaStringIds: criteriaLayout.criteriaStringIds,
      criteriaTags: criteriaLayout.criteriaTags,
      criteriaNumbers: criteriaLayout.criteriaNumbers,
      criteriaStringIdsByRow: criteriaLayout.criteriaStringIdsByRow,
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

function getOrBuildNativePredicateCriteriaLayout(
  args: {
    readonly runtimeColumnStore: EngineRuntimeColumnStoreService
  },
  input: {
    readonly cache: NativeDirectCriteriaPredicateLayoutCache | undefined
    readonly cacheKey: string | undefined
    readonly criteriaPairs: readonly CriterionRangePair[]
    readonly rowCount: number
    readonly useSharedCriteriaCache: boolean
  },
): NativeDirectCriteriaPredicateLayout | undefined {
  const cachedLayout = input.cacheKey === undefined ? undefined : input.cache?.get(input.cacheKey)
  if (
    cachedLayout !== undefined &&
    cachedLayout.rowCount === input.rowCount &&
    cachedLayout.criteriaOps.length === input.criteriaPairs.length
  ) {
    if (!shouldUseNativePredicateCriteriaLayout(cachedLayout.hasStringCriteria, input.rowCount, input.useSharedCriteriaCache)) {
      return undefined
    }
    return cachedLayout
  }

  const criteriaSlices: Array<RuntimeColumnSlice | undefined> = Array.from({ length: input.criteriaPairs.length }, () => undefined)
  const loweredCriteria = input.criteriaPairs.map((pair, index) =>
    lowerNativeCriteria(
      pair.criteria,
      () => {
        const slice = args.runtimeColumnStore.getColumnSlice(pair.range)
        criteriaSlices[index] = slice
        return slice
      },
      args.runtimeColumnStore,
    ),
  )
  if (loweredCriteria.some((criterion) => criterion === undefined)) {
    return undefined
  }
  const hasStringCriteria = loweredCriteria.some((criterion) => criterion?.kind === NATIVE_DIRECT_CRITERIA_KIND_STRING_ID)
  if (!shouldUseNativePredicateCriteriaLayout(hasStringCriteria, input.rowCount, input.useSharedCriteriaCache)) {
    return undefined
  }

  const criteriaOps = new Uint8Array(loweredCriteria.length)
  const criteriaKinds = new Uint8Array(loweredCriteria.length)
  const criteriaValues = new Float64Array(loweredCriteria.length)
  const criteriaStringIds = new Uint32Array(loweredCriteria.length)
  const criteriaTags = new Uint8Array(input.rowCount * loweredCriteria.length)
  const criteriaNumbers = new Float64Array(input.rowCount * loweredCriteria.length)
  const criteriaStringIdsByRow = new Uint32Array(input.rowCount * loweredCriteria.length)
  for (let pairIndex = 0; pairIndex < input.criteriaPairs.length; pairIndex += 1) {
    const lowered = loweredCriteria[pairIndex]
    if (lowered === undefined) {
      return undefined
    }
    criteriaOps[pairIndex] = lowered.operator
    criteriaKinds[pairIndex] = lowered.kind
    criteriaValues[pairIndex] = lowered.value
    criteriaStringIds[pairIndex] = lowered.stringId
    const criteriaSlice = criteriaSlices[pairIndex] ?? args.runtimeColumnStore.getColumnSlice(input.criteriaPairs[pairIndex]!.range)
    const offset = pairIndex * input.rowCount
    if (
      !copyNativeCriteriaSlice(
        criteriaSlice.tags,
        criteriaSlice.numbers,
        criteriaSlice.stringIds,
        criteriaTags,
        criteriaNumbers,
        criteriaStringIdsByRow,
        offset,
      )
    ) {
      return undefined
    }
  }

  const layout: NativeDirectCriteriaPredicateLayout = {
    rowCount: input.rowCount,
    hasStringCriteria,
    criteriaOps,
    criteriaKinds,
    criteriaValues,
    criteriaStringIds,
    criteriaTags,
    criteriaNumbers,
    criteriaStringIdsByRow,
  }
  if (hasStringCriteria && input.cacheKey !== undefined && input.cache !== undefined) {
    rememberNativePredicateCriteriaLayout(input.cache, input.cacheKey, layout)
  }
  return layout
}

function shouldUseNativePredicateCriteriaLayout(hasStringCriteria: boolean, rowCount: number, useSharedCriteriaCache: boolean): boolean {
  const minRows = hasStringCriteria
    ? MIN_NATIVE_DIRECT_CRITERIA_STRING_PREDICATE_AGGREGATE_ROWS
    : MIN_NATIVE_DIRECT_CRITERIA_PREDICATE_AGGREGATE_ROWS
  if (rowCount < minRows) {
    return false
  }
  return hasStringCriteria || !useSharedCriteriaCache
}

function rememberNativePredicateCriteriaLayout(
  cache: NativeDirectCriteriaPredicateLayoutCache,
  key: string,
  layout: NativeDirectCriteriaPredicateLayout,
): NativeDirectCriteriaPredicateLayout {
  if (cache.size >= NATIVE_DIRECT_CRITERIA_PREDICATE_LAYOUT_CACHE_LIMIT) {
    const firstKey = cache.keys().next().value
    if (firstKey !== undefined) {
      cache.delete(firstKey)
    }
  }
  cache.set(key, layout)
  return layout
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

function lowerNativeCriteria(
  criteria: CellValue,
  getCriteriaSlice: () => RuntimeColumnSlice,
  runtimeColumnStore: EngineRuntimeColumnStoreService,
): LoweredNativeCriteria | undefined {
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
    return { operator, kind: NATIVE_DIRECT_CRITERIA_KIND_NUMBER, value: compiled.operand.value, stringId: 0 }
  }
  if (compiled.operand.tag === ValueTag.Boolean) {
    return { operator, kind: NATIVE_DIRECT_CRITERIA_KIND_NUMBER, value: compiled.operand.value ? 1 : 0, stringId: 0 }
  }
  if (compiled.operand.tag === ValueTag.String && operator === NATIVE_DIRECT_CRITERIA_OP_EQ) {
    const stringId = findUniqueNormalizedStringIdInSlice(compiled.operand.value, getCriteriaSlice(), runtimeColumnStore)
    if (stringId !== undefined) {
      return { operator, kind: NATIVE_DIRECT_CRITERIA_KIND_STRING_ID, value: 0, stringId }
    }
  }
  return undefined
}

function copyNativeCriteriaSlice(
  sourceTags: Uint8Array,
  sourceNumbers: Float64Array,
  sourceStringIds: Uint32Array,
  targetTags: Uint8Array,
  targetNumbers: Float64Array,
  targetStringIds: Uint32Array,
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
    targetStringIds[offset + index] = sourceStringIds[index] ?? 0
  }
  return true
}

function findUniqueNormalizedStringIdInSlice(
  value: string,
  slice: RuntimeColumnSlice,
  runtimeColumnStore: EngineRuntimeColumnStoreService,
): number | undefined {
  const normalized = value.toUpperCase()
  if (normalized === '') {
    return undefined
  }
  let foundStringId = 0
  for (let index = 0; index < slice.length; index += 1) {
    if (slice.tags[index] !== ValueTag.String) {
      continue
    }
    const stringId = slice.stringIds[index] ?? 0
    if (stringId === 0 || runtimeColumnStore.normalizeStringId(stringId) !== normalized) {
      continue
    }
    if (foundStringId !== 0 && foundStringId !== stringId) {
      return undefined
    }
    foundStringId = stringId
  }
  return foundStringId === 0 ? undefined : foundStringId
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
