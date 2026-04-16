import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import type { LookupBuiltin, LookupBuiltinArgument, RangeBuiltinArgument } from './lookup.js'

interface LookupCriteriaBuiltinDeps {
  errorValue: (code: ErrorCode) => CellValue
  numberResult: (value: number) => CellValue
  isError: (value: LookupBuiltinArgument | undefined) => value is Extract<CellValue, { tag: ValueTag.Error }>
  isRangeArg: (value: LookupBuiltinArgument | undefined) => value is RangeBuiltinArgument
  toNumber: (value: CellValue) => number | undefined
  requireCellRange: (arg: LookupBuiltinArgument) => RangeBuiltinArgument | CellValue
  matchesCriteria: (value: CellValue, criteria: CellValue) => boolean
  numericAggregateCandidate: (value: CellValue) => number | undefined
}

type RangeCriteriaPair = {
  range: RangeBuiltinArgument
  criteria: CellValue
}

function validateCriteriaPairs(
  criteriaArgs: readonly LookupBuiltinArgument[],
  deps: LookupCriteriaBuiltinDeps,
): RangeCriteriaPair[] | CellValue {
  if (criteriaArgs.length === 0 || criteriaArgs.length % 2 !== 0) {
    return deps.errorValue(ErrorCode.Value)
  }
  const rangeCriteriaPairs: RangeCriteriaPair[] = []
  for (let index = 0; index < criteriaArgs.length; index += 2) {
    const range = deps.requireCellRange(criteriaArgs[index]!)
    if (!deps.isRangeArg(range)) {
      return range
    }
    const criteria = criteriaArgs[index + 1]!
    if (deps.isRangeArg(criteria)) {
      return deps.errorValue(ErrorCode.Value)
    }
    if (deps.isError(criteria)) {
      return criteria
    }
    rangeCriteriaPairs.push({ range, criteria })
  }
  return rangeCriteriaPairs
}

function findMatchingRowIndexes(
  targetRange: RangeBuiltinArgument,
  criteriaArgs: readonly LookupBuiltinArgument[],
  deps: LookupCriteriaBuiltinDeps,
): number[] | CellValue {
  const rangeCriteriaPairs = validateCriteriaPairs(criteriaArgs, deps)
  if (!Array.isArray(rangeCriteriaPairs)) {
    return rangeCriteriaPairs
  }
  if (rangeCriteriaPairs.some((pair) => pair.range.values.length !== targetRange.values.length)) {
    return deps.errorValue(ErrorCode.Value)
  }

  const matchingRows: number[] = []
  for (let row = 0; row < targetRange.values.length; row += 1) {
    if (rangeCriteriaPairs.every((pair) => deps.matchesCriteria(pair.range.values[row]!, pair.criteria))) {
      matchingRows.push(row)
    }
  }
  return matchingRows
}

function sumMatchingRows(range: RangeBuiltinArgument, rows: readonly number[], deps: LookupCriteriaBuiltinDeps): number {
  let sum = 0
  for (const row of rows) {
    sum += deps.toNumber(range.values[row]!) ?? 0
  }
  return sum
}

export function createLookupCriteriaBuiltins(deps: LookupCriteriaBuiltinDeps): Record<string, LookupBuiltin> {
  return {
    COUNTIF: (rangeArg, criteriaArg) => {
      const range = deps.requireCellRange(rangeArg)
      if (!deps.isRangeArg(range)) {
        return range
      }
      if (deps.isRangeArg(criteriaArg)) {
        return deps.errorValue(ErrorCode.Value)
      }
      if (deps.isError(criteriaArg)) {
        return criteriaArg
      }

      let count = 0
      for (const value of range.values) {
        if (deps.matchesCriteria(value, criteriaArg)) {
          count += 1
        }
      }
      return deps.numberResult(count)
    },
    COUNTIFS: (...args) => {
      const rangeCriteriaPairs = validateCriteriaPairs(args, deps)
      if (!Array.isArray(rangeCriteriaPairs)) {
        return rangeCriteriaPairs
      }

      const expectedLength = rangeCriteriaPairs[0]!.range.values.length
      if (rangeCriteriaPairs.some((pair) => pair.range.values.length !== expectedLength)) {
        return deps.errorValue(ErrorCode.Value)
      }

      let count = 0
      for (let row = 0; row < expectedLength; row += 1) {
        if (rangeCriteriaPairs.every((pair) => deps.matchesCriteria(pair.range.values[row]!, pair.criteria))) {
          count += 1
        }
      }
      return deps.numberResult(count)
    },
    SUMIF: (rangeArg, criteriaArg, sumRangeArg = rangeArg) => {
      const range = deps.requireCellRange(rangeArg)
      const sumRange = deps.requireCellRange(sumRangeArg)
      if (!deps.isRangeArg(range)) {
        return range
      }
      if (!deps.isRangeArg(sumRange)) {
        return sumRange
      }
      if (range.values.length !== sumRange.values.length) {
        return deps.errorValue(ErrorCode.Value)
      }
      if (deps.isRangeArg(criteriaArg)) {
        return deps.errorValue(ErrorCode.Value)
      }
      if (deps.isError(criteriaArg)) {
        return criteriaArg
      }

      let sum = 0
      for (let index = 0; index < range.values.length; index += 1) {
        if (!deps.matchesCriteria(range.values[index]!, criteriaArg)) {
          continue
        }
        sum += deps.toNumber(sumRange.values[index]!) ?? 0
      }
      return deps.numberResult(sum)
    },
    SUMIFS: (sumRangeArg, ...criteriaArgs) => {
      const sumRange = deps.requireCellRange(sumRangeArg)
      if (!deps.isRangeArg(sumRange)) {
        return sumRange
      }
      const matchingRows = findMatchingRowIndexes(sumRange, criteriaArgs, deps)
      if (!Array.isArray(matchingRows)) {
        return matchingRows
      }
      return deps.numberResult(sumMatchingRows(sumRange, matchingRows, deps))
    },
    AVERAGEIF: (rangeArg, criteriaArg, averageRangeArg = rangeArg) => {
      const range = deps.requireCellRange(rangeArg)
      const averageRange = deps.requireCellRange(averageRangeArg)
      if (!deps.isRangeArg(range)) {
        return range
      }
      if (!deps.isRangeArg(averageRange)) {
        return averageRange
      }
      if (range.values.length !== averageRange.values.length) {
        return deps.errorValue(ErrorCode.Value)
      }
      if (deps.isRangeArg(criteriaArg)) {
        return deps.errorValue(ErrorCode.Value)
      }
      if (deps.isError(criteriaArg)) {
        return criteriaArg
      }

      let count = 0
      let sum = 0
      for (let index = 0; index < range.values.length; index += 1) {
        if (!deps.matchesCriteria(range.values[index]!, criteriaArg)) {
          continue
        }
        const numeric = deps.toNumber(averageRange.values[index]!)
        if (numeric === undefined) {
          continue
        }
        count += 1
        sum += numeric
      }
      return count === 0 ? deps.errorValue(ErrorCode.Div0) : deps.numberResult(sum / count)
    },
    AVERAGEIFS: (averageRangeArg, ...criteriaArgs) => {
      const averageRange = deps.requireCellRange(averageRangeArg)
      if (!deps.isRangeArg(averageRange)) {
        return averageRange
      }
      const matchingRows = findMatchingRowIndexes(averageRange, criteriaArgs, deps)
      if (!Array.isArray(matchingRows)) {
        return matchingRows
      }

      let count = 0
      let sum = 0
      for (const row of matchingRows) {
        const numeric = deps.toNumber(averageRange.values[row]!)
        if (numeric === undefined) {
          continue
        }
        count += 1
        sum += numeric
      }
      return count === 0 ? deps.errorValue(ErrorCode.Div0) : deps.numberResult(sum / count)
    },
    MINIFS: (minRangeArg, ...criteriaArgs) => {
      const minRange = deps.requireCellRange(minRangeArg)
      if (!deps.isRangeArg(minRange)) {
        return minRange
      }
      const matchingRows = findMatchingRowIndexes(minRange, criteriaArgs, deps)
      if (!Array.isArray(matchingRows)) {
        return matchingRows
      }

      let minimum = Number.POSITIVE_INFINITY
      for (const row of matchingRows) {
        const numeric = deps.numericAggregateCandidate(minRange.values[row]!)
        if (numeric === undefined) {
          continue
        }
        minimum = Math.min(minimum, numeric)
      }
      return deps.numberResult(minimum === Number.POSITIVE_INFINITY ? 0 : minimum)
    },
    MAXIFS: (maxRangeArg, ...criteriaArgs) => {
      const maxRange = deps.requireCellRange(maxRangeArg)
      if (!deps.isRangeArg(maxRange)) {
        return maxRange
      }
      const matchingRows = findMatchingRowIndexes(maxRange, criteriaArgs, deps)
      if (!Array.isArray(matchingRows)) {
        return matchingRows
      }

      let maximum = Number.NEGATIVE_INFINITY
      for (const row of matchingRows) {
        const numeric = deps.numericAggregateCandidate(maxRange.values[row]!)
        if (numeric === undefined) {
          continue
        }
        maximum = Math.max(maximum, numeric)
      }
      return deps.numberResult(maximum === Number.NEGATIVE_INFINITY ? 0 : maximum)
    },
  }
}
