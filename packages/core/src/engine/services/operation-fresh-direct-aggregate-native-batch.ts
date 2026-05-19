import { addEngineCounter } from '../../perf/engine-counters.js'
import type { EngineRuntimeState, RuntimeDirectAggregateDescriptor } from '../runtime-state.js'

const NATIVE_FRESH_DIRECT_AGGREGATE_MATRIX_MIN_ROWS = 64
const NATIVE_DIRECT_AGGREGATE_OP_SUM = 1
const NATIVE_DIRECT_AGGREGATE_OP_AVERAGE = 2
const NATIVE_DIRECT_AGGREGATE_OP_COUNT = 3
const NATIVE_DIRECT_AGGREGATE_OP_MIN = 4
const NATIVE_DIRECT_AGGREGATE_OP_MAX = 5

export interface FreshDirectAggregateNativeMatrixSeed {
  readonly aggregateKind: RuntimeDirectAggregateDescriptor['aggregateKind']
  readonly aggregateColStart: number
  readonly aggregateColEnd: number
  readonly resultOffset: number | undefined
}

export function tryEvaluateNativeFreshDirectAggregateMatrixResults(
  args: {
    readonly state: Pick<EngineRuntimeState, 'wasm' | 'counters'>
  },
  input: {
    readonly inputColCount: number
    readonly matrixColStart: number
    readonly seeds: readonly FreshDirectAggregateNativeMatrixSeed[]
    readonly values: Float64Array
  },
): Float64Array | undefined {
  const first = input.seeds[0]
  if (
    first === undefined ||
    input.seeds.length < NATIVE_FRESH_DIRECT_AGGREGATE_MATRIX_MIN_ROWS ||
    input.inputColCount <= 0 ||
    !args.state.wasm.initSyncIfPossible()
  ) {
    return undefined
  }
  const aggregateKind = nativeDirectAggregateKind(first.aggregateKind)
  const startColOffset = first.aggregateColStart - input.matrixColStart
  const aggregateColCount = first.aggregateColEnd - first.aggregateColStart + 1
  if (startColOffset < 0 || aggregateColCount <= 0 || startColOffset + aggregateColCount > input.inputColCount) {
    return undefined
  }
  const resultOffset = first.resultOffset ?? 0
  for (let index = 1; index < input.seeds.length; index += 1) {
    const seed = input.seeds[index]!
    if (
      seed.aggregateKind !== first.aggregateKind ||
      seed.aggregateColStart !== first.aggregateColStart ||
      seed.aggregateColEnd !== first.aggregateColEnd ||
      seed.resultOffset !== first.resultOffset
    ) {
      return undefined
    }
  }
  const outNumbers = new Float64Array(input.seeds.length)
  if (
    !args.state.wasm.evalDenseNumericRowAggregateBatch({
      aggregateKind,
      values: input.values,
      rowCount: input.seeds.length,
      prefixColCount: input.inputColCount,
      startColOffset,
      aggregateColCount,
      resultOffset,
      outNumbers,
    })
  ) {
    return undefined
  }
  addEngineCounter(args.state.counters, 'nativeDirectAggregatePrefixEvaluations', input.seeds.length)
  return outNumbers
}

function nativeDirectAggregateKind(kind: RuntimeDirectAggregateDescriptor['aggregateKind']): number {
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
  }
}
