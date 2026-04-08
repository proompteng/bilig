export function futureValue(
  rate: number,
  periods: number,
  payment: number,
  present: number,
  type: number,
): number {
  if (rate === 0) {
    return -(present + payment * periods);
  }
  const growth = (1 + rate) ** periods;
  return -(present * growth + payment * (1 + rate * type) * ((growth - 1) / rate));
}

export function presentValue(
  rate: number,
  periods: number,
  payment: number,
  future: number,
  type: number,
): number {
  if (rate === 0) {
    return -(future + payment * periods);
  }
  const growth = (1 + rate) ** periods;
  return -(future + payment * (1 + rate * type) * ((growth - 1) / rate)) / growth;
}

export function periodicPayment(
  rate: number,
  periods: number,
  present: number,
  future: number,
  type: number,
): number | undefined {
  if (periods <= 0) {
    return undefined;
  }
  if (rate === 0) {
    return -(future + present) / periods;
  }
  const growth = (1 + rate) ** periods;
  const denominator = (1 + rate * type) * (growth - 1);
  if (denominator === 0) {
    return undefined;
  }
  return (-rate * (future + present * growth)) / denominator;
}

export function totalPeriods(
  rate: number,
  payment: number,
  present: number,
  future: number,
  type: number,
): number | undefined {
  if (payment === 0 && rate === 0) {
    return undefined;
  }
  if (rate === 0) {
    return payment === 0 ? undefined : -(future + present) / payment;
  }
  const adjustedPayment = payment * (1 + rate * type);
  const numerator = adjustedPayment - future * rate;
  const denominator = adjustedPayment + present * rate;
  if (numerator === 0 || denominator === 0 || numerator / denominator <= 0) {
    return undefined;
  }
  return Math.log(numerator / denominator) / Math.log(1 + rate);
}

function annuityRateEquation(
  rate: number,
  periods: number,
  payment: number,
  present: number,
  future: number,
  type: number,
): number {
  if (Math.abs(rate) < 1e-12) {
    return future + present + payment * periods;
  }
  const growth = (1 + rate) ** periods;
  return future + present * growth + payment * (1 + rate * type) * ((growth - 1) / rate);
}

export function solveRate(
  periods: number,
  payment: number,
  present: number,
  future: number,
  type: number,
  guess: number,
): number | undefined {
  if (!Number.isFinite(periods) || periods <= 0) {
    return undefined;
  }

  if (Math.abs(annuityRateEquation(0, periods, payment, present, future, type)) < 1e-7) {
    return 0;
  }

  let previousRate = Number.isFinite(guess) ? guess : 0.1;
  if (previousRate <= -0.999999999) {
    previousRate = -0.9;
  }
  let currentRate = previousRate === 0 ? 0.1 : previousRate * 1.1;
  if (currentRate <= -0.999999999) {
    currentRate = -0.8;
  }

  let previousError = annuityRateEquation(previousRate, periods, payment, present, future, type);
  let currentError = annuityRateEquation(currentRate, periods, payment, present, future, type);

  for (let iteration = 0; iteration < 50; iteration += 1) {
    if (!Number.isFinite(currentError)) {
      return undefined;
    }
    if (Math.abs(currentError) < 1e-7) {
      return currentRate;
    }

    let nextRate: number;
    if (
      !Number.isFinite(previousError) ||
      !Number.isFinite(currentError) ||
      currentError === previousError
    ) {
      const epsilon = Math.max(1e-7, Math.abs(currentRate) * 1e-7);
      const forward = annuityRateEquation(
        currentRate + epsilon,
        periods,
        payment,
        present,
        future,
        type,
      );
      const backward = annuityRateEquation(
        currentRate - epsilon,
        periods,
        payment,
        present,
        future,
        type,
      );
      const derivative = (forward - backward) / (2 * epsilon);
      if (!Number.isFinite(derivative) || derivative === 0) {
        return undefined;
      }
      nextRate = currentRate - currentError / derivative;
    } else {
      nextRate =
        currentRate -
        (currentError * (currentRate - previousRate)) / (currentError - previousError);
    }

    if (!Number.isFinite(nextRate) || nextRate <= -0.999999999) {
      return undefined;
    }

    previousRate = currentRate;
    previousError = currentError;
    currentRate = nextRate;
    currentError = annuityRateEquation(currentRate, periods, payment, present, future, type);
  }

  return Math.abs(currentError) < 1e-6 ? currentRate : undefined;
}

function fixedDecliningBalanceRate(
  cost: number,
  salvage: number,
  life: number,
): number | undefined {
  if (
    !Number.isFinite(cost) ||
    !Number.isFinite(salvage) ||
    !Number.isFinite(life) ||
    cost <= 0 ||
    salvage < 0 ||
    life <= 0
  ) {
    return undefined;
  }
  const ratio = salvage / cost;
  if (ratio < 0) {
    return undefined;
  }
  return Math.round((1 - ratio ** (1 / life)) * 1000) / 1000;
}

export function dbDepreciation(
  cost: number,
  salvage: number,
  life: number,
  period: number,
  month: number,
): number | undefined {
  const rate = fixedDecliningBalanceRate(cost, salvage, life);
  if (rate === undefined || month < 1 || month > 12 || period < 1 || period > life + 1) {
    return undefined;
  }

  let bookValue = cost;
  let depreciation = 0;
  for (let currentPeriod = 1; currentPeriod <= period; currentPeriod += 1) {
    const raw =
      currentPeriod === 1
        ? bookValue * rate * (month / 12)
        : currentPeriod === Math.floor(life) + 1
          ? bookValue * rate * ((12 - month) / 12)
          : bookValue * rate;
    depreciation = Math.min(Math.max(raw, 0), Math.max(0, bookValue - salvage));
    bookValue -= depreciation;
  }
  return depreciation;
}

function ddbPeriodDepreciation(
  bookValue: number,
  salvage: number,
  life: number,
  factor: number,
  remainingLife: number,
  noSwitch: boolean,
): number {
  const declining = (bookValue * factor) / life;
  const straightLine = remainingLife <= 0 ? 0 : (bookValue - salvage) / remainingLife;
  const base = noSwitch ? declining : Math.max(declining, straightLine);
  return Math.min(Math.max(base, 0), Math.max(0, bookValue - salvage));
}

export function ddbDepreciation(
  cost: number,
  salvage: number,
  life: number,
  period: number,
  factor: number,
): number | undefined {
  if (
    !Number.isFinite(cost) ||
    !Number.isFinite(salvage) ||
    !Number.isFinite(life) ||
    !Number.isFinite(period) ||
    !Number.isFinite(factor) ||
    cost <= 0 ||
    salvage < 0 ||
    life <= 0 ||
    period <= 0 ||
    factor <= 0
  ) {
    return undefined;
  }
  let bookValue = cost;
  let current = 0;
  let depreciation = 0;
  while (current < period && bookValue > salvage) {
    const segment = Math.min(1, period - current);
    const full = Math.min(
      Math.max((bookValue * factor) / life, 0),
      Math.max(0, bookValue - salvage),
    );
    depreciation = Math.min(full * segment, Math.max(0, bookValue - salvage));
    bookValue -= depreciation;
    current += segment;
  }
  return depreciation;
}

export function vdbDepreciation(
  cost: number,
  salvage: number,
  life: number,
  startPeriod: number,
  endPeriod: number,
  factor: number,
  noSwitch: boolean,
): number | undefined {
  if (
    !Number.isFinite(cost) ||
    !Number.isFinite(salvage) ||
    !Number.isFinite(life) ||
    !Number.isFinite(startPeriod) ||
    !Number.isFinite(endPeriod) ||
    !Number.isFinite(factor) ||
    cost <= 0 ||
    salvage < 0 ||
    life <= 0 ||
    startPeriod < 0 ||
    endPeriod < startPeriod ||
    factor <= 0
  ) {
    return undefined;
  }

  let bookValue = cost;
  let total = 0;
  for (let current = 0; current < endPeriod && bookValue > salvage; current += 1) {
    const overlap = Math.max(0, Math.min(endPeriod, current + 1) - Math.max(startPeriod, current));
    if (overlap <= 0) {
      const full = ddbPeriodDepreciation(
        bookValue,
        salvage,
        life,
        factor,
        life - current,
        noSwitch,
      );
      bookValue -= full;
      continue;
    }
    const full = ddbPeriodDepreciation(bookValue, salvage, life, factor, life - current, noSwitch);
    const applied = Math.min(full * overlap, Math.max(0, bookValue - salvage));
    total += applied;
    bookValue -= full;
  }
  return total;
}

export function interestPayment(
  rate: number,
  period: number,
  periods: number,
  present: number,
  future: number,
  type: number,
): number | undefined {
  if (period < 1 || period > periods) {
    return undefined;
  }
  const payment = periodicPayment(rate, periods, present, future, type);
  if (payment === undefined) {
    return undefined;
  }
  if (type === 1 && period === 1) {
    return 0;
  }
  const balance = futureValue(rate, type === 1 ? period - 2 : period - 1, payment, present, type);
  return balance * rate;
}

export function principalPayment(
  rate: number,
  period: number,
  periods: number,
  present: number,
  future: number,
  type: number,
): number | undefined {
  const payment = periodicPayment(rate, periods, present, future, type);
  const interest = interestPayment(rate, period, periods, present, future, type);
  if (payment === undefined || interest === undefined) {
    return undefined;
  }
  return payment - interest;
}

export function cumulativePeriodicPayment(
  rate: number,
  periods: number,
  present: number,
  startPeriod: number,
  endPeriod: number,
  type: number,
  principalOnly: boolean,
): number | undefined {
  if (
    rate <= 0 ||
    periods <= 0 ||
    present <= 0 ||
    startPeriod < 1 ||
    endPeriod < startPeriod ||
    endPeriod > periods
  ) {
    return undefined;
  }

  let total = 0;
  for (let period = startPeriod; period <= endPeriod; period += 1) {
    const value = principalOnly
      ? principalPayment(rate, period, periods, present, 0, type)
      : interestPayment(rate, period, periods, present, 0, type);
    if (value === undefined) {
      return undefined;
    }
    total += value;
  }
  return total;
}
