import { ErrorCode, type CellValue, type ValueTag } from '@bilig/protocol'
import type { LookupBuiltin, LookupBuiltinArgument, RangeBuiltinArgument } from './lookup.js'

interface LookupCriteriaBuiltinDeps {
  errorValue: (code: ErrorCode) => CellValue
  numberResult: (value: number) => CellValue
  isError: (value: LookupBuiltinArgument | undefined) => value is Extract<CellValue, { tag: ValueTag.Error }>
  isRangeArg: (value: LookupBuiltinArgument | undefined) => value is RangeBuiltinArgument
  toNumber: (value: CellValue | undefined) => number | undefined
  requireCellRange: (arg: LookupBuiltinArgument) => RangeBuiltinArgument | CellValue
  matchesCriteria: (value: CellValue, criteria: CellValue) => boolean
  numericAggregateCandidate: (value: CellValue | undefined) => number | undefined
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
    const range = requireCriteriaRange(criteriaArgs[index], deps)
    if (!deps.isRangeArg(range)) {
      return range
    }
    const criteria = criteriaArgs[index + 1]
    if (criteria === undefined || deps.isRangeArg(criteria)) {
      return deps.errorValue(ErrorCode.Value)
    }
    if (deps.isError(criteria)) {
      return criteria
    }
    rangeCriteriaPairs.push({ range, criteria })
  }
  return rangeCriteriaPairs
}

function requireCriteriaRange(arg: LookupBuiltinArgument | undefined, deps: LookupCriteriaBuiltinDeps): RangeBuiltinArgument | CellValue {
  if (deps.isError(arg)) {
    return arg
  }
  return arg === undefined ? deps.errorValue(ErrorCode.Value) : deps.requireCellRange(arg)
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

function sumMatchingRows(range: RangeBuiltinArgument, rows: readonly number[], deps: LookupCriteriaBuiltinDeps): number | CellValue {
  let sum = 0
  for (const row of rows) {
    const value = range.values[row]
    if (deps.isError(value)) {
      return value
    }
    sum += deps.toNumber(value) ?? 0
  }
  return sum
}

export function createLookupCriteriaBuiltins(deps: LookupCriteriaBuiltinDeps): Record<string, LookupBuiltin> {
  return {
    COUNTIF: (rangeArg, criteriaArg) => {
      const range = requireCriteriaRange(rangeArg, deps)
      if (!deps.isRangeArg(range)) {
        return range
      }
      if (criteriaArg === undefined || deps.isRangeArg(criteriaArg)) {
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
      const range = requireCriteriaRange(rangeArg, deps)
      const sumRange = requireCriteriaRange(sumRangeArg, deps)
      if (!deps.isRangeArg(range)) {
        return range
      }
      if (!deps.isRangeArg(sumRange)) {
        return sumRange
      }
      if (range.values.length !== sumRange.values.length) {
        return deps.errorValue(ErrorCode.Value)
      }
      if (criteriaArg === undefined || deps.isRangeArg(criteriaArg)) {
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
        const value = sumRange.values[index]
        if (deps.isError(value)) {
          return value
        }
        sum += deps.toNumber(value) ?? 0
      }
      return deps.numberResult(sum)
    },
    SUMIFS: (sumRangeArg, ...criteriaArgs) => {
      const sumRange = requireCriteriaRange(sumRangeArg, deps)
      if (!deps.isRangeArg(sumRange)) {
        return sumRange
      }
      const matchingRows = findMatchingRowIndexes(sumRange, criteriaArgs, deps)
      if (!Array.isArray(matchingRows)) {
        return matchingRows
      }
      const sum = sumMatchingRows(sumRange, matchingRows, deps)
      return typeof sum === 'number' ? deps.numberResult(sum) : sum
    },
    AVERAGEIF: (rangeArg, criteriaArg, averageRangeArg = rangeArg) => {
      const range = requireCriteriaRange(rangeArg, deps)
      const averageRange = requireCriteriaRange(averageRangeArg, deps)
      if (!deps.isRangeArg(range)) {
        return range
      }
      if (!deps.isRangeArg(averageRange)) {
        return averageRange
      }
      if (range.values.length !== averageRange.values.length) {
        return deps.errorValue(ErrorCode.Value)
      }
      if (criteriaArg === undefined || deps.isRangeArg(criteriaArg)) {
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
        const value = averageRange.values[index]
        if (deps.isError(value)) {
          return value
        }
        const numeric = deps.toNumber(value)
        if (numeric === undefined) {
          continue
        }
        count += 1
        sum += numeric
      }
      return count === 0 ? deps.errorValue(ErrorCode.Div0) : deps.numberResult(sum / count)
    },
    AVERAGEIFS: (averageRangeArg, ...criteriaArgs) => {
      const averageRange = requireCriteriaRange(averageRangeArg, deps)
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
        const value = averageRange.values[row]
        if (deps.isError(value)) {
          return value
        }
        const numeric = deps.toNumber(value)
        if (numeric === undefined) {
          continue
        }
        count += 1
        sum += numeric
      }
      return count === 0 ? deps.errorValue(ErrorCode.Div0) : deps.numberResult(sum / count)
    },
    MINIFS: (minRangeArg, ...criteriaArgs) => {
      const minRange = requireCriteriaRange(minRangeArg, deps)
      if (!deps.isRangeArg(minRange)) {
        return minRange
      }
      const matchingRows = findMatchingRowIndexes(minRange, criteriaArgs, deps)
      if (!Array.isArray(matchingRows)) {
        return matchingRows
      }

      let minimum = Number.POSITIVE_INFINITY
      for (const row of matchingRows) {
        const value = minRange.values[row]
        if (deps.isError(value)) {
          return value
        }
        const numeric = deps.numericAggregateCandidate(value)
        if (numeric === undefined) {
          continue
        }
        minimum = Math.min(minimum, numeric)
      }
      return deps.numberResult(minimum === Number.POSITIVE_INFINITY ? 0 : minimum)
    },
    MAXIFS: (maxRangeArg, ...criteriaArgs) => {
      const maxRange = requireCriteriaRange(maxRangeArg, deps)
      if (!deps.isRangeArg(maxRange)) {
        return maxRange
      }
      const matchingRows = findMatchingRowIndexes(maxRange, criteriaArgs, deps)
      if (!Array.isArray(matchingRows)) {
        return matchingRows
      }

      let maximum = Number.NEGATIVE_INFINITY
      for (const row of matchingRows) {
        const value = maxRange.values[row]
        if (deps.isError(value)) {
          return value
        }
        const numeric = deps.numericAggregateCandidate(value)
        if (numeric === undefined) {
          continue
        }
        maximum = Math.max(maximum, numeric)
      }
      return deps.numberResult(maximum === Number.NEGATIVE_INFINITY ? 0 : maximum)
    },
  }
}
