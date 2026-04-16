import { ErrorCode, type CellValue } from '@bilig/protocol'
import { excelSerialToDateParts } from './datetime.js'
import type { LookupBuiltin, LookupBuiltinArgument, RangeBuiltinArgument } from './lookup.js'

interface LookupFinancialBuiltinDeps {
  errorValue: (code: ErrorCode) => CellValue
  numberResult: (value: number) => CellValue
  isRangeArg: (value: LookupBuiltinArgument | undefined) => value is RangeBuiltinArgument
  toNumber: (value: CellValue) => number | undefined
  collectNumericSeries: (arg: LookupBuiltinArgument, mode: 'lenient' | 'strict') => number[] | CellValue
}

function isCellError(
  value: LookupBuiltinArgument | undefined,
  isRangeArg: LookupFinancialBuiltinDeps['isRangeArg'],
): value is Extract<CellValue, { code: ErrorCode }> {
  return value !== undefined && !isRangeArg(value) && 'code' in value
}

function collectDateSerialSeries(
  arg: LookupBuiltinArgument | undefined,
  { errorValue, collectNumericSeries }: LookupFinancialBuiltinDeps,
): number[] | CellValue {
  if (arg === undefined) {
    return errorValue(ErrorCode.Value)
  }
  const numericValues = collectNumericSeries(arg, 'strict')
  if (!Array.isArray(numericValues)) {
    return numericValues
  }
  const serials: number[] = []
  for (const numeric of numericValues) {
    const serial = Math.trunc(numeric)
    if (!Number.isFinite(serial) || excelSerialToDateParts(serial) === undefined) {
      return errorValue(ErrorCode.Value)
    }
    serials.push(serial)
  }
  return serials
}

function hasPositiveAndNegative(values: readonly number[]): boolean {
  let hasPositive = false
  let hasNegative = false
  for (const value of values) {
    if (value > 0) {
      hasPositive = true
    } else if (value < 0) {
      hasNegative = true
    }
  }
  return hasPositive && hasNegative
}

function periodicCashflowNetPresentValue(rate: number, values: readonly number[]): number | undefined {
  if (!Number.isFinite(rate) || rate <= -0.999999999) {
    return undefined
  }
  const base = 1 + rate
  let total = 0
  for (let index = 0; index < values.length; index += 1) {
    total += values[index]! / base ** index
  }
  return Number.isFinite(total) ? total : undefined
}

function periodicCashflowNetPresentValueDerivative(rate: number, values: readonly number[]): number | undefined {
  if (!Number.isFinite(rate) || rate <= -0.999999999) {
    return undefined
  }
  const base = 1 + rate
  let total = 0
  for (let index = 1; index < values.length; index += 1) {
    total -= (index * values[index]!) / base ** (index + 1)
  }
  return Number.isFinite(total) ? total : undefined
}

function xnpvValue(rate: number, values: readonly number[], dates: readonly number[]): number | undefined {
  if (!Number.isFinite(rate) || rate <= -0.999999999 || values.length !== dates.length) {
    return undefined
  }
  const base = 1 + rate
  const start = dates[0]
  if (start === undefined) {
    return undefined
  }
  let total = 0
  for (let index = 0; index < values.length; index += 1) {
    const elapsed = (dates[index]! - start) / 365
    total += values[index]! / base ** elapsed
  }
  return Number.isFinite(total) ? total : undefined
}

function xnpvDerivative(rate: number, values: readonly number[], dates: readonly number[]): number | undefined {
  if (!Number.isFinite(rate) || rate <= -0.999999999 || values.length !== dates.length) {
    return undefined
  }
  const base = 1 + rate
  const start = dates[0]
  if (start === undefined) {
    return undefined
  }
  let total = 0
  for (let index = 0; index < values.length; index += 1) {
    const elapsed = (dates[index]! - start) / 365
    total -= (elapsed * values[index]!) / base ** (elapsed + 1)
  }
  return Number.isFinite(total) ? total : undefined
}

function solveDiscountRate(
  guess: number,
  evaluate: (rate: number) => number | undefined,
  derivative: (rate: number) => number | undefined,
): number | undefined {
  const zeroError = evaluate(0)
  if (zeroError !== undefined && Math.abs(zeroError) < 1e-7) {
    return 0
  }

  let previousRate = Number.isFinite(guess) ? guess : 0.1
  if (previousRate <= -0.999999999) {
    previousRate = -0.9
  }
  let currentRate = previousRate === 0 ? 0.1 : previousRate * 1.1
  if (currentRate <= -0.999999999) {
    currentRate = -0.8
  }

  let previousError = evaluate(previousRate)
  let currentError = evaluate(currentRate)
  if (previousError === undefined || currentError === undefined) {
    return undefined
  }

  for (let iteration = 0; iteration < 100; iteration += 1) {
    if (!Number.isFinite(currentError)) {
      return undefined
    }
    if (Math.abs(currentError) < 1e-10) {
      return currentRate
    }

    let nextRate: number | undefined
    if (Number.isFinite(previousError) && currentError !== previousError) {
      nextRate = currentRate - (currentError * (currentRate - previousRate)) / (currentError - previousError)
    }
    if (nextRate === undefined || !Number.isFinite(nextRate) || nextRate <= -0.999999999) {
      const slope = derivative(currentRate)
      if (slope === undefined || !Number.isFinite(slope) || slope === 0) {
        return undefined
      }
      nextRate = currentRate - currentError / slope
    }
    if (!Number.isFinite(nextRate) || nextRate <= -0.999999999) {
      return undefined
    }

    previousRate = currentRate
    previousError = currentError
    currentRate = nextRate
    const nextError = evaluate(currentRate)
    if (nextError === undefined) {
      return undefined
    }
    currentError = nextError
  }

  return Math.abs(currentError) < 1e-7 ? currentRate : undefined
}

function modifiedInternalRateOfReturn(values: readonly number[], financeRate: number, reinvestRate: number): number | undefined {
  if (values.length < 2 || financeRate <= -1 || reinvestRate <= -1 || !Number.isFinite(financeRate) || !Number.isFinite(reinvestRate)) {
    return undefined
  }
  let positiveFutureValue = 0
  let negativePresentValue = 0
  const lastIndex = values.length - 1
  for (let index = 0; index < values.length; index += 1) {
    const value = values[index]!
    if (value > 0) {
      positiveFutureValue += value * (1 + reinvestRate) ** (lastIndex - index)
    } else if (value < 0) {
      negativePresentValue += value / (1 + financeRate) ** index
    }
  }
  if (positiveFutureValue <= 0 || negativePresentValue >= 0 || lastIndex <= 0) {
    return undefined
  }
  const result = (-positiveFutureValue / negativePresentValue) ** (1 / lastIndex) - 1
  return Number.isFinite(result) ? result : undefined
}

export function createLookupFinancialBuiltins(deps: LookupFinancialBuiltinDeps): Record<string, LookupBuiltin> {
  return {
    IRR: (valuesArg, guessArg = deps.numberResult(0.1)) => {
      if (valuesArg === undefined || guessArg === undefined || deps.isRangeArg(guessArg)) {
        return deps.errorValue(ErrorCode.Value)
      }
      if (isCellError(guessArg, deps.isRangeArg)) {
        return guessArg
      }
      const values = deps.collectNumericSeries(valuesArg, 'lenient')
      if (!Array.isArray(values)) {
        return values
      }
      if (!hasPositiveAndNegative(values)) {
        return deps.errorValue(ErrorCode.Value)
      }
      const guess = deps.toNumber(guessArg)
      if (guess === undefined) {
        return deps.errorValue(ErrorCode.Value)
      }
      const rate = solveDiscountRate(
        guess,
        (candidate) => periodicCashflowNetPresentValue(candidate, values),
        (candidate) => periodicCashflowNetPresentValueDerivative(candidate, values),
      )
      return rate === undefined ? deps.errorValue(ErrorCode.Value) : deps.numberResult(rate)
    },
    MIRR: (valuesArg, financeRateArg, reinvestRateArg) => {
      if (
        valuesArg === undefined ||
        financeRateArg === undefined ||
        reinvestRateArg === undefined ||
        deps.isRangeArg(financeRateArg) ||
        deps.isRangeArg(reinvestRateArg)
      ) {
        return deps.errorValue(ErrorCode.Value)
      }
      if (isCellError(financeRateArg, deps.isRangeArg)) {
        return financeRateArg
      }
      if (isCellError(reinvestRateArg, deps.isRangeArg)) {
        return reinvestRateArg
      }
      const values = deps.collectNumericSeries(valuesArg, 'lenient')
      if (!Array.isArray(values)) {
        return values
      }
      if (!hasPositiveAndNegative(values)) {
        return deps.errorValue(ErrorCode.Div0)
      }
      const financeRate = deps.toNumber(financeRateArg)
      const reinvestRate = deps.toNumber(reinvestRateArg)
      if (financeRate === undefined || reinvestRate === undefined) {
        return deps.errorValue(ErrorCode.Value)
      }
      const result = modifiedInternalRateOfReturn(values, financeRate, reinvestRate)
      return result === undefined ? deps.errorValue(ErrorCode.Div0) : deps.numberResult(result)
    },
    XNPV: (rateArg, valuesArg, datesArg) => {
      if (rateArg === undefined || valuesArg === undefined || datesArg === undefined || deps.isRangeArg(rateArg)) {
        return deps.errorValue(ErrorCode.Value)
      }
      if (isCellError(rateArg, deps.isRangeArg)) {
        return rateArg
      }
      const rate = deps.toNumber(rateArg)
      if (rate === undefined) {
        return deps.errorValue(ErrorCode.Value)
      }
      const values = deps.collectNumericSeries(valuesArg, 'strict')
      if (!Array.isArray(values)) {
        return values
      }
      const dates = collectDateSerialSeries(datesArg, deps)
      if (!Array.isArray(dates)) {
        return dates
      }
      if (values.length !== dates.length || values.length === 0 || !hasPositiveAndNegative(values)) {
        return deps.errorValue(ErrorCode.Value)
      }
      const start = dates[0]!
      if (dates.some((date) => date < start)) {
        return deps.errorValue(ErrorCode.Value)
      }
      const result = xnpvValue(rate, values, dates)
      return result === undefined ? deps.errorValue(ErrorCode.Value) : deps.numberResult(result)
    },
    XIRR: (valuesArg, datesArg, guessArg = deps.numberResult(0.1)) => {
      if (valuesArg === undefined || datesArg === undefined || guessArg === undefined || deps.isRangeArg(guessArg)) {
        return deps.errorValue(ErrorCode.Value)
      }
      if (isCellError(guessArg, deps.isRangeArg)) {
        return guessArg
      }
      const values = deps.collectNumericSeries(valuesArg, 'strict')
      if (!Array.isArray(values)) {
        return values
      }
      const dates = collectDateSerialSeries(datesArg, deps)
      if (!Array.isArray(dates)) {
        return dates
      }
      if (values.length !== dates.length || values.length === 0 || !hasPositiveAndNegative(values)) {
        return deps.errorValue(ErrorCode.Value)
      }
      const start = dates[0]!
      if (dates.some((date) => date < start)) {
        return deps.errorValue(ErrorCode.Value)
      }
      const guess = deps.toNumber(guessArg)
      if (guess === undefined) {
        return deps.errorValue(ErrorCode.Value)
      }
      const result = solveDiscountRate(
        guess,
        (candidate) => xnpvValue(candidate, values, dates),
        (candidate) => xnpvDerivative(candidate, values, dates),
      )
      return result === undefined ? deps.errorValue(ErrorCode.Value) : deps.numberResult(result)
    },
  }
}
