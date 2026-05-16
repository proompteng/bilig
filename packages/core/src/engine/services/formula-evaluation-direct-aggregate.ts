import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import type { EngineCounters } from '../../perf/engine-counters.js'
import { addEngineCounter } from '../../perf/engine-counters.js'
import type { WorkbookStore } from '../../workbook-store.js'
import type { RuntimeFormula } from '../runtime-state.js'
import type { RangeAggregateCacheService } from './range-aggregate-cache-service.js'
import { decodeErrorCode, directErrorResult, directNumberResult, offsetDirectAggregateResult } from './formula-evaluation-helpers.js'

const DIRECT_AGGREGATE_SCAN_MAX_LENGTH = 64
const DIRECT_AGGREGATE_PREFIX_MIN_LENGTH = 16

export function tryEvaluateDirectAggregate(input: {
  readonly formula: RuntimeFormula
  readonly workbook: Pick<WorkbookStore, 'getSheet'>
  readonly counters: EngineCounters
  readonly aggregateCache: RangeAggregateCacheService
  readonly readCellValueByIndex: (cellIndex: number | undefined) => CellValue
}): CellValue | undefined {
  const directAggregate = input.formula.directAggregate
  if (!directAggregate) {
    return undefined
  }
  const columnCount = directAggregate.colEnd - directAggregate.col + 1
  const canUseSlidingPrefix =
    input.formula.dependencyIndices.length === 0 &&
    (directAggregate.aggregateKind === 'sum' || directAggregate.aggregateKind === 'average' || directAggregate.aggregateKind === 'count') &&
    directAggregate.length > DIRECT_AGGREGATE_PREFIX_MIN_LENGTH
  const canUseExistingLargePrefix =
    input.formula.dependencyIndices.length === 0 && directAggregate.length > DIRECT_AGGREGATE_SCAN_MAX_LENGTH
  if (!canUseSlidingPrefix && !canUseExistingLargePrefix) {
    addEngineCounter(input.counters, 'directAggregateScanEvaluations')
    addEngineCounter(input.counters, 'directAggregateScanCells', directAggregate.length)
    const aggregateSheet = input.workbook.getSheet(directAggregate.sheetName)
    if (!aggregateSheet) {
      return undefined
    }
    let sum = 0
    let count = 0
    let averageCount = 0
    let minimum = Number.POSITIVE_INFINITY
    let maximum = Number.NEGATIVE_INFINITY
    for (let col = directAggregate.col; col <= directAggregate.colEnd; col += 1) {
      for (let row = directAggregate.rowStart; row <= directAggregate.rowEnd; row += 1) {
        const memberCellIndex =
          aggregateSheet.structureVersion === 1 ? aggregateSheet.grid.getPhysical(row, col) : aggregateSheet.grid.get(row, col)
        const value: CellValue = memberCellIndex === -1 ? { tag: ValueTag.Empty } : input.readCellValueByIndex(memberCellIndex)
        switch (value.tag) {
          case ValueTag.Number:
            sum += value.value
            count += 1
            averageCount += 1
            minimum = Math.min(minimum, value.value)
            maximum = Math.max(maximum, value.value)
            break
          case ValueTag.Boolean: {
            const booleanNumber = value.value ? 1 : 0
            sum += booleanNumber
            count += 1
            averageCount += 1
            minimum = Math.min(minimum, booleanNumber)
            maximum = Math.max(maximum, booleanNumber)
            break
          }
          case ValueTag.Empty:
            break
          case ValueTag.Error:
            if (directAggregate.aggregateKind === 'sum' || directAggregate.aggregateKind === 'average') {
              return directErrorResult(value.code)
            }
            break
          case ValueTag.String:
            break
        }
      }
    }
    if (directAggregate.aggregateKind === 'sum') {
      return offsetDirectAggregateResult(directAggregate, directNumberResult(sum))
    }
    if (directAggregate.aggregateKind === 'count') {
      return offsetDirectAggregateResult(directAggregate, directNumberResult(count))
    }
    if (directAggregate.aggregateKind === 'average') {
      return averageCount === 0
        ? directErrorResult(ErrorCode.Div0)
        : offsetDirectAggregateResult(directAggregate, directNumberResult(sum / averageCount))
    }
    if (directAggregate.aggregateKind === 'min') {
      return offsetDirectAggregateResult(directAggregate, directNumberResult(minimum === Number.POSITIVE_INFINITY ? 0 : minimum))
    }
    return offsetDirectAggregateResult(directAggregate, directNumberResult(maximum === Number.NEGATIVE_INFINITY ? 0 : maximum))
  }

  addEngineCounter(input.counters, 'directAggregatePrefixEvaluations')
  // SUM/AVERAGE ranges should reuse any compatible lower-start prefix to avoid
  // rescanning shifted windows, while still allowing narrower anchors when no
  // compatible reusable prefix exists.
  const sharedPrefixStart =
    directAggregate.aggregateKind === 'sum' || directAggregate.aggregateKind === 'average' || directAggregate.aggregateKind === 'count'
      ? 0
      : directAggregate.rowStart
  let errorCode = ErrorCode.None
  let sum = 0
  let count = 0
  let averageCount = 0
  let errorCount = 0
  let minimum = Number.POSITIVE_INFINITY
  let maximum = Number.NEGATIVE_INFINITY
  let hasShiftedPrefixStart = false
  for (let colOffset = 0; colOffset < columnCount; colOffset += 1) {
    const prefix = input.aggregateCache.getOrBuildColumnPrefix(
      {
        sheetName: directAggregate.sheetName,
        rowStart: sharedPrefixStart,
        rowEnd: directAggregate.rowEnd,
        col: directAggregate.col + colOffset,
      },
      directAggregate.aggregateKind,
    )
    const endOffset = directAggregate.rowEnd - prefix.rowStart
    const startOffset = directAggregate.rowStart - prefix.rowStart - 1
    hasShiftedPrefixStart ||= startOffset >= 0
    const prefixSum = prefix.prefixSums[endOffset] ?? 0
    const prefixCount = prefix.prefixCount[endOffset] ?? 0
    const prefixAverageCount = prefix.prefixAverageCount[endOffset] ?? 0
    const prefixErrorCount = prefix.prefixErrorCounts[endOffset] ?? 0
    sum += startOffset >= 0 ? prefixSum - (prefix.prefixSums[startOffset] ?? 0) : prefixSum
    count += startOffset >= 0 ? prefixCount - (prefix.prefixCount[startOffset] ?? 0) : prefixCount
    averageCount += startOffset >= 0 ? prefixAverageCount - (prefix.prefixAverageCount[startOffset] ?? 0) : prefixAverageCount
    errorCount += startOffset >= 0 ? prefixErrorCount - (prefix.prefixErrorCounts[startOffset] ?? 0) : prefixErrorCount
    const nextErrorCode = prefix.prefixErrorCodes[endOffset]
    if (errorCode === ErrorCode.None && nextErrorCode !== undefined && nextErrorCode !== Number(ErrorCode.None)) {
      errorCode = decodeErrorCode(nextErrorCode)
    }
    minimum = Math.min(minimum, prefix.prefixMinimums[endOffset] ?? Number.POSITIVE_INFINITY)
    maximum = Math.max(maximum, prefix.prefixMaximums[endOffset] ?? Number.NEGATIVE_INFINITY)
  }
  if (
    errorCode !== ErrorCode.None &&
    errorCount > 0 &&
    (directAggregate.aggregateKind === 'sum' || directAggregate.aggregateKind === 'average')
  ) {
    return hasShiftedPrefixStart ? undefined : directErrorResult(decodeErrorCode(errorCode))
  }
  if (directAggregate.aggregateKind === 'sum') {
    return offsetDirectAggregateResult(directAggregate, directNumberResult(sum))
  }
  if (directAggregate.aggregateKind === 'count') {
    return offsetDirectAggregateResult(directAggregate, directNumberResult(count))
  }
  if (directAggregate.aggregateKind === 'min') {
    return offsetDirectAggregateResult(directAggregate, directNumberResult(minimum === Number.POSITIVE_INFINITY ? 0 : minimum))
  }
  if (directAggregate.aggregateKind === 'max') {
    return offsetDirectAggregateResult(directAggregate, directNumberResult(maximum === Number.NEGATIVE_INFINITY ? 0 : maximum))
  }
  return averageCount === 0
    ? directErrorResult(ErrorCode.Div0)
    : offsetDirectAggregateResult(directAggregate, directNumberResult(sum / averageCount))
}
