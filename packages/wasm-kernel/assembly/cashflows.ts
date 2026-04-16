export function futureValueCalc(rate: f64, periods: f64, payment: f64, present: f64, paymentTypeValue: i32): f64 {
  if (rate == 0.0) {
    return -(present + payment * periods)
  }
  const growth = Math.pow(1.0 + rate, periods)
  return -(present * growth + payment * (1.0 + rate * <f64>paymentTypeValue) * ((growth - 1.0) / rate))
}

export function presentValueCalc(rate: f64, periods: f64, payment: f64, future: f64, paymentTypeValue: i32): f64 {
  if (rate == 0.0) {
    return -(future + payment * periods)
  }
  const growth = Math.pow(1.0 + rate, periods)
  return -(future + payment * (1.0 + rate * <f64>paymentTypeValue) * ((growth - 1.0) / rate)) / growth
}

export function periodicPaymentCalc(rate: f64, periods: f64, present: f64, future: f64, paymentTypeValue: i32): f64 {
  if (periods <= 0.0) {
    return NaN
  }
  if (rate == 0.0) {
    return -(future + present) / periods
  }
  const growth = Math.pow(1.0 + rate, periods)
  const denominator = (1.0 + rate * <f64>paymentTypeValue) * (growth - 1.0)
  if (denominator == 0.0) {
    return NaN
  }
  return (-rate * (future + present * growth)) / denominator
}

export function totalPeriodsCalc(rate: f64, payment: f64, present: f64, future: f64, paymentTypeValue: i32): f64 {
  if (payment == 0.0 && rate == 0.0) {
    return NaN
  }
  if (rate == 0.0) {
    return payment == 0.0 ? NaN : -(future + present) / payment
  }
  const adjustedPayment = payment * (1.0 + rate * <f64>paymentTypeValue)
  const numerator = adjustedPayment - future * rate
  const denominator = adjustedPayment + present * rate
  if (numerator == 0.0 || denominator == 0.0 || numerator / denominator <= 0.0) {
    return NaN
  }
  return Math.log(numerator / denominator) / Math.log(1.0 + rate)
}

export function interestPaymentCalc(rate: f64, period: f64, periods: f64, present: f64, future: f64, paymentTypeValue: i32): f64 {
  if (period < 1.0 || period > periods) {
    return NaN
  }
  const payment = periodicPaymentCalc(rate, periods, present, future, paymentTypeValue)
  if (isNaN(payment)) {
    return NaN
  }
  if (paymentTypeValue == 1 && period == 1.0) {
    return 0.0
  }
  const balance = futureValueCalc(rate, paymentTypeValue == 1 ? period - 2.0 : period - 1.0, payment, present, paymentTypeValue)
  return balance * rate
}

export function principalPaymentCalc(rate: f64, period: f64, periods: f64, present: f64, future: f64, paymentTypeValue: i32): f64 {
  const payment = periodicPaymentCalc(rate, periods, present, future, paymentTypeValue)
  const interest = interestPaymentCalc(rate, period, periods, present, future, paymentTypeValue)
  return isNaN(payment) || isNaN(interest) ? NaN : payment - interest
}

function annuityRateEquationCalc(rate: f64, periods: f64, payment: f64, present: f64, future: f64, paymentTypeValue: i32): f64 {
  if (Math.abs(rate) < 1e-12) {
    return future + present + payment * periods
  }
  const growth = Math.pow(1.0 + rate, periods)
  return future + present * growth + payment * (1.0 + rate * <f64>paymentTypeValue) * ((growth - 1.0) / rate)
}

export function solveRateCalc(periods: f64, payment: f64, present: f64, future: f64, paymentTypeValue: i32, guess: f64): f64 {
  if (!isFinite(periods) || periods <= 0.0) {
    return NaN
  }

  if (Math.abs(annuityRateEquationCalc(0.0, periods, payment, present, future, paymentTypeValue)) < 1e-7) {
    return 0.0
  }

  let previousRate = isFinite(guess) ? guess : 0.1
  if (previousRate <= -0.999999999) {
    previousRate = -0.9
  }
  let currentRate = previousRate == 0.0 ? 0.1 : previousRate * 1.1
  if (currentRate <= -0.999999999) {
    currentRate = -0.8
  }

  let previousError = annuityRateEquationCalc(previousRate, periods, payment, present, future, paymentTypeValue)
  let currentError = annuityRateEquationCalc(currentRate, periods, payment, present, future, paymentTypeValue)

  for (let iteration = 0; iteration < 50; iteration += 1) {
    if (!isFinite(currentError)) {
      return NaN
    }
    if (Math.abs(currentError) < 1e-7) {
      return currentRate
    }

    let nextRate: f64
    if (!isFinite(previousError) || currentError == previousError) {
      const epsilon = Math.max(1e-7, Math.abs(currentRate) * 1e-7)
      const forward = annuityRateEquationCalc(currentRate + epsilon, periods, payment, present, future, paymentTypeValue)
      const backward = annuityRateEquationCalc(currentRate - epsilon, periods, payment, present, future, paymentTypeValue)
      const derivative = (forward - backward) / (2.0 * epsilon)
      if (!isFinite(derivative) || derivative == 0.0) {
        return NaN
      }
      nextRate = currentRate - currentError / derivative
    } else {
      nextRate = currentRate - (currentError * (currentRate - previousRate)) / (currentError - previousError)
    }

    if (!isFinite(nextRate) || nextRate <= -0.999999999) {
      return NaN
    }

    previousRate = currentRate
    previousError = currentError
    currentRate = nextRate
    currentError = annuityRateEquationCalc(currentRate, periods, payment, present, future, paymentTypeValue)
  }

  return Math.abs(currentError) < 1e-6 ? currentRate : NaN
}

export function hasPositiveAndNegativeSeries(values: Array<f64>): bool {
  let hasPositive = false
  let hasNegative = false
  for (let index = 0; index < values.length; index += 1) {
    const value = unchecked(values[index])
    if (value > 0.0) {
      hasPositive = true
    } else if (value < 0.0) {
      hasNegative = true
    }
  }
  return hasPositive && hasNegative
}

export function periodicCashflowNetPresentValueCalc(rate: f64, values: Array<f64>): f64 {
  if (!isFinite(rate) || rate <= -0.999999999) {
    return NaN
  }
  const base = 1.0 + rate
  let total = 0.0
  for (let index = 0; index < values.length; index += 1) {
    total += unchecked(values[index]) / Math.pow(base, <f64>index)
  }
  return isFinite(total) ? total : NaN
}

function periodicCashflowNetPresentValueDerivativeCalc(rate: f64, values: Array<f64>): f64 {
  if (!isFinite(rate) || rate <= -0.999999999) {
    return NaN
  }
  const base = 1.0 + rate
  let total = 0.0
  for (let index = 1; index < values.length; index += 1) {
    total -= (<f64>index * unchecked(values[index])) / Math.pow(base, <f64>(index + 1))
  }
  return isFinite(total) ? total : NaN
}

export function xnpvCalc(rate: f64, values: Array<f64>, dates: Array<i32>): f64 {
  if (!isFinite(rate) || rate <= -0.999999999 || values.length != dates.length || values.length == 0) {
    return NaN
  }
  const base = 1.0 + rate
  const start = unchecked(dates[0])
  let total = 0.0
  for (let index = 0; index < values.length; index += 1) {
    const elapsed = <f64>(unchecked(dates[index]) - start) / 365.0
    total += unchecked(values[index]) / Math.pow(base, elapsed)
  }
  return isFinite(total) ? total : NaN
}

function xnpvDerivativeCalc(rate: f64, values: Array<f64>, dates: Array<i32>): f64 {
  if (!isFinite(rate) || rate <= -0.999999999 || values.length != dates.length || values.length == 0) {
    return NaN
  }
  const base = 1.0 + rate
  const start = unchecked(dates[0])
  let total = 0.0
  for (let index = 0; index < values.length; index += 1) {
    const elapsed = <f64>(unchecked(dates[index]) - start) / 365.0
    total -= (elapsed * unchecked(values[index])) / Math.pow(base, elapsed + 1.0)
  }
  return isFinite(total) ? total : NaN
}

export function solvePeriodicCashflowRateCalc(values: Array<f64>, guess: f64): f64 {
  const zeroError = periodicCashflowNetPresentValueCalc(0.0, values)
  if (isFinite(zeroError) && Math.abs(zeroError) < 1e-7) {
    return 0.0
  }

  let previousRate = isFinite(guess) ? guess : 0.1
  if (previousRate <= -0.999999999) {
    previousRate = -0.9
  }
  let currentRate = previousRate == 0.0 ? 0.1 : previousRate * 1.1
  if (currentRate <= -0.999999999) {
    currentRate = -0.8
  }

  let previousError = periodicCashflowNetPresentValueCalc(previousRate, values)
  let currentError = periodicCashflowNetPresentValueCalc(currentRate, values)
  if (isNaN(previousError) || isNaN(currentError)) {
    return NaN
  }

  for (let iteration = 0; iteration < 100; iteration += 1) {
    if (!isFinite(currentError)) {
      return NaN
    }
    if (Math.abs(currentError) < 1e-10) {
      return currentRate
    }

    let nextRate: f64
    if (isFinite(previousError) && currentError != previousError) {
      nextRate = currentRate - (currentError * (currentRate - previousRate)) / (currentError - previousError)
    } else {
      nextRate = NaN
    }
    if (!isFinite(nextRate) || nextRate <= -0.999999999) {
      const derivative = periodicCashflowNetPresentValueDerivativeCalc(currentRate, values)
      if (!isFinite(derivative) || derivative == 0.0) {
        return NaN
      }
      nextRate = currentRate - currentError / derivative
    }
    if (!isFinite(nextRate) || nextRate <= -0.999999999) {
      return NaN
    }
    previousRate = currentRate
    previousError = currentError
    currentRate = nextRate
    currentError = periodicCashflowNetPresentValueCalc(currentRate, values)
  }

  return Math.abs(currentError) < 1e-7 ? currentRate : NaN
}

export function solveXirrCalc(values: Array<f64>, dates: Array<i32>, guess: f64): f64 {
  const zeroError = xnpvCalc(0.0, values, dates)
  if (isFinite(zeroError) && Math.abs(zeroError) < 1e-7) {
    return 0.0
  }

  let previousRate = isFinite(guess) ? guess : 0.1
  if (previousRate <= -0.999999999) {
    previousRate = -0.9
  }
  let currentRate = previousRate == 0.0 ? 0.1 : previousRate * 1.1
  if (currentRate <= -0.999999999) {
    currentRate = -0.8
  }

  let previousError = xnpvCalc(previousRate, values, dates)
  let currentError = xnpvCalc(currentRate, values, dates)
  if (isNaN(previousError) || isNaN(currentError)) {
    return NaN
  }

  for (let iteration = 0; iteration < 100; iteration += 1) {
    if (!isFinite(currentError)) {
      return NaN
    }
    if (Math.abs(currentError) < 1e-10) {
      return currentRate
    }

    let nextRate: f64
    if (isFinite(previousError) && currentError != previousError) {
      nextRate = currentRate - (currentError * (currentRate - previousRate)) / (currentError - previousError)
    } else {
      nextRate = NaN
    }
    if (!isFinite(nextRate) || nextRate <= -0.999999999) {
      const derivative = xnpvDerivativeCalc(currentRate, values, dates)
      if (!isFinite(derivative) || derivative == 0.0) {
        return NaN
      }
      nextRate = currentRate - currentError / derivative
    }
    if (!isFinite(nextRate) || nextRate <= -0.999999999) {
      return NaN
    }
    previousRate = currentRate
    previousError = currentError
    currentRate = nextRate
    currentError = xnpvCalc(currentRate, values, dates)
  }

  return Math.abs(currentError) < 1e-7 ? currentRate : NaN
}

export function mirrCalc(values: Array<f64>, financeRate: f64, reinvestRate: f64): f64 {
  if (values.length < 2 || !isFinite(financeRate) || !isFinite(reinvestRate) || financeRate <= -1.0 || reinvestRate <= -1.0) {
    return NaN
  }
  let positiveFutureValue = 0.0
  let negativePresentValue = 0.0
  const lastIndex = values.length - 1
  for (let index = 0; index < values.length; index += 1) {
    const value = unchecked(values[index])
    if (value > 0.0) {
      positiveFutureValue += value * Math.pow(1.0 + reinvestRate, <f64>(lastIndex - index))
    } else if (value < 0.0) {
      negativePresentValue += value / Math.pow(1.0 + financeRate, <f64>index)
    }
  }
  if (positiveFutureValue <= 0.0 || negativePresentValue >= 0.0 || lastIndex <= 0) {
    return NaN
  }
  const result = Math.pow(-positiveFutureValue / negativePresentValue, 1.0 / <f64>lastIndex) - 1.0
  return isFinite(result) ? result : NaN
}

export function cumulativePeriodicPaymentCalc(
  rate: f64,
  periods: f64,
  present: f64,
  startPeriod: i32,
  endPeriod: i32,
  paymentTypeValue: i32,
  principalOnly: bool,
): f64 {
  if (
    !isFinite(rate) ||
    !isFinite(periods) ||
    !isFinite(present) ||
    rate <= 0.0 ||
    periods <= 0.0 ||
    present <= 0.0 ||
    startPeriod < 1 ||
    endPeriod < startPeriod ||
    <f64>endPeriod > periods
  ) {
    return NaN
  }

  let total = 0.0
  for (let period = startPeriod; period <= endPeriod; period += 1) {
    const value = principalOnly
      ? principalPaymentCalc(rate, <f64>period, periods, present, 0.0, paymentTypeValue)
      : interestPaymentCalc(rate, <f64>period, periods, present, 0.0, paymentTypeValue)
    if (isNaN(value)) {
      return NaN
    }
    total += value
  }
  return total
}
