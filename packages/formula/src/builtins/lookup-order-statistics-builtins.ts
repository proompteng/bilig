import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import type { ArrayValue } from '../runtime-values.js'
import type { LookupBuiltin, LookupBuiltinArgument, RangeBuiltinArgument } from './lookup.js'

interface LookupOrderStatisticsBuiltinDeps {
  errorValue: (code: ErrorCode) => CellValue
  numberResult: (value: number) => CellValue
  arrayResult: (values: CellValue[], rows: number, cols: number) => ArrayValue
  requireCellRange: (arg: LookupBuiltinArgument) => RangeBuiltinArgument | CellValue
  isError: (value: LookupBuiltinArgument | undefined) => value is Extract<CellValue, { tag: ValueTag.Error }>
  isRangeArg: (value: LookupBuiltinArgument | undefined) => value is RangeBuiltinArgument
  toNumber: (value: CellValue) => number | undefined
  toInteger: (value: CellValue) => number | undefined
  flattenNumbers: (arg: LookupBuiltinArgument) => number[] | CellValue
}

function flattenNumbersOrValueError(
  arg: LookupBuiltinArgument | undefined,
  { errorValue, flattenNumbers }: LookupOrderStatisticsBuiltinDeps,
): number[] | CellValue {
  return arg === undefined ? errorValue(ErrorCode.Value) : flattenNumbers(arg)
}

function flattenNumericArguments(args: readonly LookupBuiltinArgument[], deps: LookupOrderStatisticsBuiltinDeps): number[] | CellValue {
  const values: number[] = []
  for (const arg of args) {
    const flattened = flattenNumbersOrValueError(arg, deps)
    if (!Array.isArray(flattened)) {
      return flattened
    }
    values.push(...flattened)
  }
  return values
}

function rankFromValues(
  numberArg: LookupBuiltinArgument | undefined,
  arrayArg: LookupBuiltinArgument | undefined,
  orderArg: LookupBuiltinArgument | undefined,
  useAverage: boolean,
  deps: LookupOrderStatisticsBuiltinDeps,
): CellValue {
  if (numberArg === undefined || arrayArg === undefined) {
    return deps.errorValue(ErrorCode.Value)
  }
  if (deps.isRangeArg(numberArg) || deps.isRangeArg(orderArg)) {
    return deps.errorValue(ErrorCode.Value)
  }
  if (deps.isError(numberArg)) {
    return numberArg
  }
  if (deps.isError(arrayArg)) {
    return arrayArg
  }
  if (deps.isError(orderArg)) {
    return orderArg
  }

  const target = deps.toNumber(numberArg)
  if (target === undefined) {
    return deps.errorValue(ErrorCode.Value)
  }
  const order = orderArg === undefined ? 0 : deps.toInteger(orderArg)
  if (order === undefined || ![0, 1].includes(order)) {
    return deps.errorValue(ErrorCode.Value)
  }

  const values = flattenNumbersOrValueError(arrayArg, deps)
  if (!Array.isArray(values)) {
    return values
  }
  if (values.length === 0) {
    return deps.errorValue(ErrorCode.NA)
  }

  let preceding = 0
  let ties = 0
  for (const value of values) {
    if (value === target) {
      ties += 1
      continue
    }
    if (order === 0 ? value > target : value < target) {
      preceding += 1
    }
  }

  if (ties === 0) {
    return deps.errorValue(ErrorCode.NA)
  }

  return deps.numberResult(useAverage ? preceding + (ties + 1) / 2 : preceding + 1)
}

function nthValue(
  arg: LookupBuiltinArgument | undefined,
  positionArg: LookupBuiltinArgument | undefined,
  ascending: boolean,
  deps: LookupOrderStatisticsBuiltinDeps,
): CellValue {
  if (arg === undefined || positionArg === undefined) {
    return deps.errorValue(ErrorCode.Value)
  }
  if (deps.isRangeArg(positionArg)) {
    return deps.errorValue(ErrorCode.Value)
  }
  if (deps.isError(positionArg)) {
    return positionArg
  }
  const values = flattenNumbersOrValueError(arg, deps)
  if (!Array.isArray(values)) {
    return values
  }
  if (values.length === 0) {
    return deps.errorValue(ErrorCode.Value)
  }

  const position = deps.toInteger(positionArg)
  if (position === undefined || position < 1 || position > values.length) {
    return deps.errorValue(ErrorCode.Value)
  }

  const sortedValues = values.toSorted(ascending ? (a, b) => a - b : (a, b) => b - a)
  return deps.numberResult(sortedValues[position - 1] ?? 0)
}

function interpolatePercentile(sortedValues: readonly number[], percentile: number, exclusive: boolean): number | undefined {
  const count = sortedValues.length
  if (count === 0 || !Number.isFinite(percentile)) {
    return undefined
  }

  if (exclusive) {
    if (!(percentile > 0 && percentile < 1)) {
      return undefined
    }
    const rank = percentile * (count + 1)
    if (rank < 1 || rank > count) {
      return undefined
    }
    const lowerRank = Math.floor(rank)
    const upperRank = Math.ceil(rank)
    if (lowerRank === upperRank) {
      return sortedValues[lowerRank - 1]
    }
    const lower = sortedValues[lowerRank - 1]
    const upper = sortedValues[upperRank - 1]
    return lower !== undefined && upper !== undefined ? lower + (rank - lowerRank) * (upper - lower) : undefined
  }

  if (percentile < 0 || percentile > 1) {
    return undefined
  }
  if (count === 1) {
    return sortedValues[0]
  }
  const rank = percentile * (count - 1) + 1
  const lowerRank = Math.floor(rank)
  const upperRank = Math.ceil(rank)
  if (lowerRank === upperRank) {
    return sortedValues[lowerRank - 1]
  }
  const lower = sortedValues[lowerRank - 1]
  const upper = sortedValues[upperRank - 1]
  return lower !== undefined && upper !== undefined ? lower + (rank - lowerRank) * (upper - lower) : undefined
}

function percentileFromValues(
  arrayArg: LookupBuiltinArgument | undefined,
  percentileArg: LookupBuiltinArgument | undefined,
  exclusive: boolean,
  deps: LookupOrderStatisticsBuiltinDeps,
): CellValue {
  if (arrayArg === undefined || percentileArg === undefined) {
    return deps.errorValue(ErrorCode.Value)
  }
  if (deps.isRangeArg(percentileArg)) {
    return deps.errorValue(ErrorCode.Value)
  }
  if (deps.isError(arrayArg)) {
    return arrayArg
  }
  if (deps.isError(percentileArg)) {
    return percentileArg
  }

  const values = flattenNumbersOrValueError(arrayArg, deps)
  if (!Array.isArray(values)) {
    return values
  }
  if (values.length === 0) {
    return deps.errorValue(ErrorCode.Value)
  }

  const percentile = deps.toNumber(percentileArg)
  if (percentile === undefined) {
    return deps.errorValue(ErrorCode.Value)
  }

  const sortedValues = values.toSorted((left, right) => left - right)
  const interpolated = interpolatePercentile(sortedValues, percentile, exclusive)
  return interpolated === undefined ? deps.errorValue(ErrorCode.Value) : deps.numberResult(interpolated)
}

function truncateToSignificance(value: number, significance: number): number {
  const scale = 10 ** significance
  return Math.trunc(value * scale) / scale
}

function interpolatePercentRank(sortedValues: readonly number[], target: number, exclusive: boolean): number | undefined {
  const count = sortedValues.length
  if (count < 2 || !Number.isFinite(target)) {
    return undefined
  }

  let exactFirst = -1
  let exactLast = -1
  for (let index = 0; index < count; index += 1) {
    const value = sortedValues[index]!
    if (value !== target) {
      continue
    }
    if (exactFirst === -1) {
      exactFirst = index
    }
    exactLast = index
  }

  if (exactFirst !== -1) {
    const averageIndex = (exactFirst + exactLast) / 2
    return exclusive ? (averageIndex + 1) / (count + 1) : averageIndex / (count - 1)
  }

  if (target < sortedValues[0]! || target > sortedValues[count - 1]!) {
    return undefined
  }

  let lowerIndex = -1
  for (let index = 0; index < count; index += 1) {
    if (sortedValues[index]! < target) {
      lowerIndex = index
      continue
    }
    break
  }
  const upperIndex = lowerIndex + 1
  if (lowerIndex < 0 || upperIndex >= count) {
    return undefined
  }

  const lower = sortedValues[lowerIndex]!
  const upper = sortedValues[upperIndex]!
  if (upper === lower) {
    return undefined
  }

  const lowerRank = exclusive ? (lowerIndex + 1) / (count + 1) : lowerIndex / (count - 1)
  const upperRank = exclusive ? (upperIndex + 1) / (count + 1) : upperIndex / (count - 1)
  const fraction = (target - lower) / (upper - lower)
  return lowerRank + fraction * (upperRank - lowerRank)
}

function percentRankFromValues(
  arrayArg: LookupBuiltinArgument | undefined,
  targetArg: LookupBuiltinArgument | undefined,
  significanceArg: LookupBuiltinArgument | undefined,
  exclusive: boolean,
  deps: LookupOrderStatisticsBuiltinDeps,
): CellValue {
  if (arrayArg === undefined || targetArg === undefined) {
    return deps.errorValue(ErrorCode.Value)
  }
  if (deps.isRangeArg(targetArg) || deps.isRangeArg(significanceArg)) {
    return deps.errorValue(ErrorCode.Value)
  }
  if (deps.isError(arrayArg)) {
    return arrayArg
  }
  if (deps.isError(targetArg)) {
    return targetArg
  }
  if (deps.isError(significanceArg)) {
    return significanceArg
  }

  const values = flattenNumbersOrValueError(arrayArg, deps)
  if (!Array.isArray(values)) {
    return values
  }
  if (values.length < 2) {
    return deps.errorValue(ErrorCode.Value)
  }

  const target = deps.toNumber(targetArg)
  if (target === undefined) {
    return deps.errorValue(ErrorCode.Value)
  }

  const significance = significanceArg === undefined ? 3 : deps.toInteger(significanceArg)
  if (significance === undefined || significance < 1) {
    return deps.errorValue(ErrorCode.Value)
  }

  const sortedValues = values.toSorted((left, right) => left - right)
  const rank = interpolatePercentRank(sortedValues, target, exclusive)
  return rank === undefined ? deps.errorValue(ErrorCode.NA) : deps.numberResult(truncateToSignificance(rank, significance))
}

function quartileFromValues(
  arrayArg: LookupBuiltinArgument | undefined,
  quartArg: LookupBuiltinArgument | undefined,
  exclusive: boolean,
  deps: LookupOrderStatisticsBuiltinDeps,
): CellValue {
  if (arrayArg === undefined || quartArg === undefined) {
    return deps.errorValue(ErrorCode.Value)
  }
  if (deps.isRangeArg(quartArg)) {
    return deps.errorValue(ErrorCode.Value)
  }
  if (deps.isError(arrayArg)) {
    return arrayArg
  }
  if (deps.isError(quartArg)) {
    return quartArg
  }

  const values = flattenNumbersOrValueError(arrayArg, deps)
  if (!Array.isArray(values)) {
    return values
  }
  if (values.length === 0) {
    return deps.errorValue(ErrorCode.Value)
  }

  const quartile = deps.toInteger(quartArg)
  if (quartile === undefined) {
    return deps.errorValue(ErrorCode.Value)
  }

  const sortedValues = values.toSorted((left, right) => left - right)
  if (!exclusive && (quartile === 0 || quartile === 4)) {
    return deps.numberResult(sortedValues[quartile === 0 ? 0 : sortedValues.length - 1] ?? 0)
  }
  if (exclusive && (quartile <= 0 || quartile >= 4)) {
    return deps.errorValue(ErrorCode.Value)
  }
  if (!exclusive && (quartile < 0 || quartile > 4)) {
    return deps.errorValue(ErrorCode.Value)
  }

  const interpolated = interpolatePercentile(sortedValues, quartile / 4, exclusive)
  return interpolated === undefined ? deps.errorValue(ErrorCode.Value) : deps.numberResult(interpolated)
}

function collectPlainNumericValues(
  arg: LookupBuiltinArgument | undefined,
  ignoreNonNumbersInRange: boolean,
  deps: LookupOrderStatisticsBuiltinDeps,
): number[] | CellValue {
  if (arg === undefined) {
    return deps.errorValue(ErrorCode.Value)
  }
  if (!deps.isRangeArg(arg)) {
    if (deps.isError(arg)) {
      return arg
    }
    return arg.tag === ValueTag.Number ? [arg.value] : deps.errorValue(ErrorCode.Value)
  }

  if (arg.refKind !== 'cells') {
    return deps.errorValue(ErrorCode.Value)
  }

  const values: number[] = []
  for (const value of arg.values) {
    if (value.tag === ValueTag.Error) {
      return value
    }
    if (value.tag === ValueTag.Number) {
      values.push(value.value)
      continue
    }
    if (!ignoreNonNumbersInRange && value.tag !== ValueTag.Empty) {
      return deps.errorValue(ErrorCode.Value)
    }
  }
  return values
}

function validateProbabilityInputs(
  xArg: LookupBuiltinArgument | undefined,
  probabilitiesArg: LookupBuiltinArgument | undefined,
  deps: LookupOrderStatisticsBuiltinDeps,
): { xValues: number[]; probabilityValues: number[] } | CellValue {
  if (xArg === undefined || probabilitiesArg === undefined) {
    return deps.errorValue(ErrorCode.Value)
  }
  const xRange = deps.requireCellRange(xArg)
  if (!deps.isRangeArg(xRange)) {
    return xRange
  }
  const probabilityRange = deps.requireCellRange(probabilitiesArg)
  if (!deps.isRangeArg(probabilityRange)) {
    return probabilityRange
  }
  if (xRange.rows !== probabilityRange.rows || xRange.cols !== probabilityRange.cols || xRange.values.length === 0) {
    return deps.errorValue(ErrorCode.Value)
  }

  const xValues: number[] = []
  const probabilityValues: number[] = []
  let probabilitySum = 0
  for (let index = 0; index < xRange.values.length; index += 1) {
    const xValue = xRange.values[index]!
    const probability = probabilityRange.values[index]!
    if (xValue.tag === ValueTag.Error) {
      return xValue
    }
    if (probability.tag === ValueTag.Error) {
      return probability
    }
    if (xValue.tag !== ValueTag.Number || probability.tag !== ValueTag.Number) {
      return deps.errorValue(ErrorCode.Value)
    }
    if (!Number.isFinite(probability.value) || probability.value < 0 || probability.value > 1) {
      return deps.errorValue(ErrorCode.Value)
    }
    xValues.push(xValue.value)
    probabilityValues.push(probability.value)
    probabilitySum += probability.value
  }
  if (Math.abs(probabilitySum - 1) > 1e-9) {
    return deps.errorValue(ErrorCode.Value)
  }
  return { xValues, probabilityValues }
}

function probFromValues(
  xArg: LookupBuiltinArgument | undefined,
  probabilitiesArg: LookupBuiltinArgument | undefined,
  lowerArg: LookupBuiltinArgument | undefined,
  upperArg: LookupBuiltinArgument | undefined,
  deps: LookupOrderStatisticsBuiltinDeps,
): CellValue {
  if (lowerArg === undefined) {
    return deps.errorValue(ErrorCode.Value)
  }
  if (deps.isRangeArg(lowerArg) || deps.isRangeArg(upperArg)) {
    return deps.errorValue(ErrorCode.Value)
  }
  if (deps.isError(lowerArg)) {
    return lowerArg
  }
  if (deps.isError(upperArg)) {
    return upperArg
  }

  const lower = deps.toNumber(lowerArg)
  const upper = upperArg === undefined ? lower : deps.toNumber(upperArg)
  if (lower === undefined || upper === undefined || upper < lower) {
    return deps.errorValue(ErrorCode.Value)
  }

  const values = validateProbabilityInputs(xArg, probabilitiesArg, deps)
  if (!('xValues' in values)) {
    return values
  }

  let total = 0
  for (let index = 0; index < values.xValues.length; index += 1) {
    const value = values.xValues[index]!
    if (value >= lower && value <= upper) {
      total += values.probabilityValues[index]!
    }
  }
  return deps.numberResult(total)
}

function trimMeanFromValues(
  arrayArg: LookupBuiltinArgument | undefined,
  percentArg: LookupBuiltinArgument | undefined,
  deps: LookupOrderStatisticsBuiltinDeps,
): CellValue {
  if (percentArg === undefined) {
    return deps.errorValue(ErrorCode.Value)
  }
  if (deps.isRangeArg(percentArg)) {
    return deps.errorValue(ErrorCode.Value)
  }
  if (deps.isError(arrayArg)) {
    return arrayArg
  }
  if (deps.isError(percentArg)) {
    return percentArg
  }

  const values = collectPlainNumericValues(arrayArg, true, deps)
  if (!Array.isArray(values) || values.length === 0) {
    return Array.isArray(values) ? deps.errorValue(ErrorCode.Value) : values
  }

  const percent = deps.toNumber(percentArg)
  if (percent === undefined || percent < 0 || percent >= 1) {
    return deps.errorValue(ErrorCode.Value)
  }

  let excluded = Math.floor(values.length * percent)
  if (excluded % 2 === 1) {
    excluded -= 1
  }
  const retainedCount = values.length - excluded
  if (retainedCount <= 0) {
    return deps.errorValue(ErrorCode.Value)
  }

  const sorted = values.toSorted((left, right) => left - right)
  const trimEachSide = excluded / 2
  const retained = sorted.slice(trimEachSide, sorted.length - trimEachSide)
  if (retained.length === 0) {
    return deps.errorValue(ErrorCode.Value)
  }
  const sum = retained.reduce((total, value) => total + value, 0)
  return deps.numberResult(sum / retained.length)
}

function flattenNumbersIgnoringNonNumeric(
  arg: LookupBuiltinArgument | undefined,
  deps: LookupOrderStatisticsBuiltinDeps,
): number[] | CellValue {
  if (arg === undefined) {
    return deps.errorValue(ErrorCode.Value)
  }
  if (!deps.isRangeArg(arg)) {
    if (deps.isError(arg)) {
      return arg
    }
    const numeric = deps.toNumber(arg)
    return numeric === undefined ? deps.errorValue(ErrorCode.Value) : [numeric]
  }

  const values: number[] = []
  for (const value of arg.values) {
    if (value.tag === ValueTag.Error) {
      return value
    }
    if (value.tag === ValueTag.Number) {
      values.push(value.value)
    }
  }
  return values
}

function flattenNumbersIgnoringNonNumericArgs(
  args: readonly LookupBuiltinArgument[],
  deps: LookupOrderStatisticsBuiltinDeps,
): number[] | CellValue {
  const values: number[] = []
  for (const arg of args) {
    const flattened = flattenNumbersIgnoringNonNumeric(arg, deps)
    if (!Array.isArray(flattened)) {
      return flattened
    }
    values.push(...flattened)
  }
  return values
}

function modeMultiFromValues(args: readonly LookupBuiltinArgument[], deps: LookupOrderStatisticsBuiltinDeps): ArrayValue | CellValue {
  const values = flattenNumbersIgnoringNonNumericArgs(args, deps)
  if (!Array.isArray(values)) {
    return values
  }
  if (values.length === 0) {
    return deps.errorValue(ErrorCode.NA)
  }

  const counts = new Map<number, number>()
  let maxCount = 0
  for (const value of values) {
    const nextCount = (counts.get(value) ?? 0) + 1
    counts.set(value, nextCount)
    if (nextCount > maxCount) {
      maxCount = nextCount
    }
  }
  if (maxCount < 2) {
    return deps.errorValue(ErrorCode.NA)
  }

  const modes = Array.from(counts.entries())
    .filter(([, count]) => count === maxCount)
    .map(([value]) => value)
    .toSorted((left, right) => left - right)
  return deps.arrayResult(
    modes.map((value) => deps.numberResult(value)),
    modes.length,
    1,
  )
}

function frequencyFromValues(
  dataArg: LookupBuiltinArgument | undefined,
  binsArg: LookupBuiltinArgument | undefined,
  deps: LookupOrderStatisticsBuiltinDeps,
): ArrayValue | CellValue {
  const dataValues = flattenNumbersIgnoringNonNumeric(dataArg, deps)
  if (!Array.isArray(dataValues)) {
    return dataValues
  }
  const binValues = flattenNumbersIgnoringNonNumeric(binsArg, deps)
  if (!Array.isArray(binValues)) {
    return binValues
  }

  const sortedBins = binValues.toSorted((left, right) => left - right)
  const counts = Array.from({ length: sortedBins.length + 1 }, () => 0)
  for (const value of dataValues) {
    let bucket = sortedBins.length
    for (let index = 0; index < sortedBins.length; index += 1) {
      if (value <= sortedBins[index]!) {
        bucket = index
        break
      }
    }
    counts[bucket]! += 1
  }

  return deps.arrayResult(
    counts.map((value) => deps.numberResult(value)),
    counts.length,
    1,
  )
}

export function createLookupOrderStatisticsBuiltins(deps: LookupOrderStatisticsBuiltinDeps): Record<string, LookupBuiltin> {
  return {
    AVEDEV: (...args) => {
      const values = flattenNumericArguments(args, deps)
      if (!Array.isArray(values)) {
        return values
      }
      if (values.length === 0) {
        return deps.errorValue(ErrorCode.Value)
      }
      const mean = values.reduce((sum, value) => sum + value, 0) / values.length
      const totalAbsoluteDeviation = values.reduce((sum, value) => sum + Math.abs(value - mean), 0)
      return deps.numberResult(totalAbsoluteDeviation / values.length)
    },
    DEVSQ: (...args) => {
      const values = flattenNumericArguments(args, deps)
      if (!Array.isArray(values)) {
        return values
      }
      if (values.length === 0) {
        return deps.errorValue(ErrorCode.Value)
      }
      const mean = values.reduce((sum, value) => sum + value, 0) / values.length
      const total = values.reduce((sum, value) => {
        const deviation = value - mean
        return sum + deviation * deviation
      }, 0)
      return deps.numberResult(total)
    },
    MEDIAN: (...args) => {
      const values = flattenNumericArguments(args, deps)
      if (!Array.isArray(values)) {
        return values
      }
      if (values.length === 0) {
        return deps.errorValue(ErrorCode.Value)
      }

      const sortedValues = values.toSorted((left, right) => left - right)
      const center = Math.floor(sortedValues.length / 2)
      const isEven = sortedValues.length % 2 === 0
      if (!isEven) {
        return deps.numberResult(sortedValues[center] ?? 0)
      }

      const lower = sortedValues[center - 1]
      const upper = sortedValues[center]
      return deps.numberResult(((lower ?? 0) + (upper ?? 0)) / 2)
    },
    SMALL: (arg, positionArg) => nthValue(arg, positionArg, true, deps),
    LARGE: (arg, positionArg) => nthValue(arg, positionArg, false, deps),
    PERCENTILE: (arrayArg, percentileArg) => percentileFromValues(arrayArg, percentileArg, false, deps),
    'PERCENTILE.INC': (arrayArg, percentileArg) => percentileFromValues(arrayArg, percentileArg, false, deps),
    'PERCENTILE.EXC': (arrayArg, percentileArg) => percentileFromValues(arrayArg, percentileArg, true, deps),
    PERCENTRANK: (arrayArg, targetArg, significanceArg = { tag: ValueTag.Number, value: 3 }) =>
      percentRankFromValues(arrayArg, targetArg, significanceArg, false, deps),
    'PERCENTRANK.INC': (arrayArg, targetArg, significanceArg = { tag: ValueTag.Number, value: 3 }) =>
      percentRankFromValues(arrayArg, targetArg, significanceArg, false, deps),
    'PERCENTRANK.EXC': (arrayArg, targetArg, significanceArg = { tag: ValueTag.Number, value: 3 }) =>
      percentRankFromValues(arrayArg, targetArg, significanceArg, true, deps),
    QUARTILE: (arrayArg, quartArg) => quartileFromValues(arrayArg, quartArg, false, deps),
    'QUARTILE.INC': (arrayArg, quartArg) => quartileFromValues(arrayArg, quartArg, false, deps),
    'QUARTILE.EXC': (arrayArg, quartArg) => quartileFromValues(arrayArg, quartArg, true, deps),
    'MODE.MULT': (...args) => modeMultiFromValues(args, deps),
    FREQUENCY: (dataArg, binsArg) => frequencyFromValues(dataArg, binsArg, deps),
    PROB: (xArg, probabilitiesArg, lowerArg, upperArg) => probFromValues(xArg, probabilitiesArg, lowerArg, upperArg, deps),
    TRIMMEAN: (arrayArg, percentArg) => trimMeanFromValues(arrayArg, percentArg, deps),
    RANK: (numberArg, arrayArg, orderArg = { tag: ValueTag.Number, value: 0 }) =>
      rankFromValues(numberArg, arrayArg, orderArg, false, deps),
    'RANK.EQ': (numberArg, arrayArg, orderArg = { tag: ValueTag.Number, value: 0 }) =>
      rankFromValues(numberArg, arrayArg, orderArg, false, deps),
    'RANK.AVG': (numberArg, arrayArg, orderArg = { tag: ValueTag.Number, value: 0 }) =>
      rankFromValues(numberArg, arrayArg, orderArg, true, deps),
  }
}
