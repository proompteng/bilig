import { addMonthsToExcelDate, excelSerialToDateParts } from './datetime.js'

export interface MaturityAtIssueFractions {
  issueToMaturity: number
  settlementToMaturity: number
  issueToSettlement: number
}

export interface OddLastCouponFractions {
  accruedFraction: number
  remainingFraction: number
  totalFraction: number
}

export interface OddFirstCouponMetrics {
  accruedFraction: number
  remainingFraction: number
  totalFraction: number
  regularPeriodsAfterFirst: number
}

export interface CouponSchedule {
  previousCoupon: number
  nextCoupon: number
  periodsRemaining: number
}

export interface CouponMetrics extends CouponSchedule {
  accruedDays: number
  daysToNextCoupon: number
  daysInPeriod: number
}

function isLeapYear(year: number): boolean {
  return year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0)
}

export function isValidBasis(basis: number): boolean {
  return basis === 0 || basis === 1 || basis === 2 || basis === 3 || basis === 4
}

export function isValidFrequency(frequency: number): boolean {
  return frequency === 1 || frequency === 2 || frequency === 4
}

function yearsDaysInYear(year: number): number {
  return isLeapYear(year) ? 366 : 365
}

export function yearFracByBasis(startSerial: number, endSerial: number, basis: number): number | undefined {
  if (!isValidBasis(basis)) {
    return undefined
  }

  let start = startSerial
  let end = endSerial
  if (start > end) {
    ;[start, end] = [end, start]
  }

  const startParts = excelSerialToDateParts(start)
  const endParts = excelSerialToDateParts(end)
  if (startParts === undefined || endParts === undefined) {
    return undefined
  }

  let startDay = startParts.day
  let startMonth = startParts.month
  let startYear = startParts.year
  let endDay = endParts.day
  let endMonth = endParts.month
  let endYear = endParts.year

  let totalDays: number
  switch (basis) {
    case 0:
      if (startDay === 31) {
        startDay -= 1
      }
      if (startDay === 30 && endDay === 31) {
        endDay -= 1
      } else if (startMonth === 2 && startDay === (isLeapYear(startYear) ? 29 : 28)) {
        startDay = 30
        if (endMonth === 2 && endDay === (isLeapYear(endYear) ? 29 : 28)) {
          endDay = 30
        }
      }
      totalDays = (endYear - startYear) * 360 + (endMonth - startMonth) * 30 + (endDay - startDay)
      break
    case 1:
    case 2:
    case 3:
      totalDays = end - start
      break
    case 4:
      if (startDay === 31) {
        startDay -= 1
      }
      if (endDay === 31) {
        endDay -= 1
      }
      totalDays = (endYear - startYear) * 360 + (endMonth - startMonth) * 30 + (endDay - startDay)
      break
    default:
      return undefined
  }

  let daysInYear: number
  switch (basis) {
    case 1: {
      const isYearDifferent = startYear !== endYear
      if (isYearDifferent && (endYear !== startYear + 1 || endMonth < startMonth || (endMonth === startMonth && endDay > startDay))) {
        let dayCount = 0
        for (let year = startYear; year <= endYear; year += 1) {
          dayCount += yearsDaysInYear(year)
        }
        daysInYear = dayCount / (endYear - startYear + 1)
      } else if (isYearDifferent) {
        const crossesLeap =
          (isLeapYear(startYear) && (startMonth < 2 || (startMonth === 2 && startDay <= 29))) ||
          (isLeapYear(endYear) && (endMonth > 2 || (endMonth === 2 && endDay === 29)))
        daysInYear = crossesLeap ? 366 : 365
      } else {
        daysInYear = yearsDaysInYear(startYear)
      }
      break
    }
    case 3:
      daysInYear = 365
      break
    case 0:
    case 2:
    case 4:
      daysInYear = 360
      break
    default:
      return undefined
  }

  return totalDays / daysInYear
}

export function securityAnnualizedYearFraction(settlement: number, maturity: number, basis: number): number | undefined {
  if (settlement >= maturity || !isValidBasis(basis)) {
    return undefined
  }
  const years = yearFracByBasis(settlement, maturity, basis)
  return years !== undefined && years > 0 ? years : undefined
}

export function treasuryBillDays(settlement: number, maturity: number): number | undefined {
  if (settlement >= maturity) {
    return undefined
  }
  const days = maturity - settlement
  return days > 0 && days <= 365 ? days : undefined
}

export function maturityAtIssueFractions(
  settlement: number,
  maturity: number,
  issue: number,
  basis: number,
): MaturityAtIssueFractions | undefined {
  if (issue >= settlement || issue >= maturity || settlement >= maturity || !isValidBasis(basis)) {
    return undefined
  }
  const issueToMaturity = yearFracByBasis(issue, maturity, basis)
  const settlementToMaturity = yearFracByBasis(settlement, maturity, basis)
  const issueToSettlement = yearFracByBasis(issue, settlement, basis)
  if (
    issueToMaturity === undefined ||
    settlementToMaturity === undefined ||
    issueToSettlement === undefined ||
    issueToMaturity <= 0 ||
    settlementToMaturity <= 0
  ) {
    return undefined
  }
  return {
    issueToMaturity,
    settlementToMaturity,
    issueToSettlement,
  }
}

export function oddLastCouponFractions(
  settlement: number,
  maturity: number,
  lastInterest: number,
  frequency: number,
  basis: number,
): OddLastCouponFractions | undefined {
  if (lastInterest >= settlement || settlement >= maturity || !isValidFrequency(frequency) || !isValidBasis(basis)) {
    return undefined
  }

  const stepMonths = 12 / frequency
  let periodStart = lastInterest
  let accruedFraction = 0
  let remainingFraction = 0
  let totalFraction = 0
  let iterations = 0

  while (periodStart < maturity && iterations < 32) {
    const normalEnd = addMonthsToExcelDate(periodStart, stepMonths)
    if (normalEnd === undefined || normalEnd <= periodStart) {
      return undefined
    }
    const actualEnd = Math.min(normalEnd, maturity)
    const normalDays = couponDaysByBasis(periodStart, normalEnd, basis)
    const countedDays = couponDaysByBasis(periodStart, actualEnd, basis)
    if (normalDays === undefined || countedDays === undefined || normalDays <= 0 || countedDays < 0) {
      return undefined
    }

    totalFraction += countedDays / normalDays

    if (settlement > periodStart) {
      const accruedEnd = Math.min(settlement, actualEnd)
      const accruedDays = couponDaysByBasis(periodStart, accruedEnd, basis)
      if (accruedDays === undefined || accruedDays < 0) {
        return undefined
      }
      accruedFraction += accruedDays / normalDays
    }

    if (settlement < actualEnd) {
      const remainingStart = Math.max(settlement, periodStart)
      const remainingDays = couponDaysByBasis(remainingStart, actualEnd, basis)
      if (remainingDays === undefined || remainingDays < 0) {
        return undefined
      }
      remainingFraction += remainingDays / normalDays
    }

    periodStart = actualEnd
    iterations += 1
  }

  if (periodStart !== maturity || iterations >= 32 || remainingFraction <= 0 || totalFraction <= 0) {
    return undefined
  }

  return {
    accruedFraction,
    remainingFraction,
    totalFraction,
  }
}

function oddFirstCouponMetrics(
  settlement: number,
  maturity: number,
  issue: number,
  firstCoupon: number,
  frequency: number,
  basis: number,
): OddFirstCouponMetrics | undefined {
  if (issue >= settlement || settlement >= firstCoupon || firstCoupon >= maturity || !isValidFrequency(frequency) || !isValidBasis(basis)) {
    return undefined
  }

  const stepMonths = 12 / frequency
  const segments: Array<{
    actualStart: number
    periodEnd: number
    normalDays: number
    countedDays: number
  }> = []
  let periodEnd = firstCoupon
  let iterations = 0
  while (periodEnd > issue && iterations < 64) {
    const normalStart = addMonthsToExcelDate(periodEnd, -stepMonths)
    if (normalStart === undefined || normalStart >= periodEnd) {
      return undefined
    }
    const actualStart = Math.max(normalStart, issue)
    const normalDays = couponDaysByBasis(normalStart, periodEnd, basis)
    const countedDays = couponDaysByBasis(actualStart, periodEnd, basis)
    if (normalDays === undefined || countedDays === undefined || normalDays <= 0 || countedDays < 0) {
      return undefined
    }
    segments.unshift({ actualStart, periodEnd, normalDays, countedDays })
    periodEnd = actualStart
    iterations += 1
  }
  if (periodEnd !== issue || iterations >= 64 || segments.length === 0) {
    return undefined
  }

  let accruedFraction = 0
  let remainingFraction = 0
  let totalFraction = 0
  for (const segment of segments) {
    totalFraction += segment.countedDays / segment.normalDays

    if (settlement > segment.actualStart) {
      const accruedEnd = Math.min(settlement, segment.periodEnd)
      const accruedDays = couponDaysByBasis(segment.actualStart, accruedEnd, basis)
      if (accruedDays === undefined || accruedDays < 0) {
        return undefined
      }
      accruedFraction += accruedDays / segment.normalDays
    }

    if (settlement < segment.periodEnd) {
      const remainingStart = Math.max(settlement, segment.actualStart)
      const remainingDays = couponDaysByBasis(remainingStart, segment.periodEnd, basis)
      if (remainingDays === undefined || remainingDays < 0) {
        return undefined
      }
      remainingFraction += remainingDays / segment.normalDays
    }
  }

  let regularPeriodsAfterFirst = 0
  let couponDate = firstCoupon
  while (couponDate < maturity && regularPeriodsAfterFirst < 256) {
    const nextCouponDate = addMonthsToExcelDate(couponDate, stepMonths)
    if (nextCouponDate === undefined || nextCouponDate <= couponDate) {
      return undefined
    }
    couponDate = nextCouponDate
    regularPeriodsAfterFirst += 1
  }
  if (
    couponDate !== maturity ||
    regularPeriodsAfterFirst <= 0 ||
    regularPeriodsAfterFirst >= 256 ||
    remainingFraction <= 0 ||
    totalFraction <= 0
  ) {
    return undefined
  }

  return {
    accruedFraction,
    remainingFraction,
    totalFraction,
    regularPeriodsAfterFirst,
  }
}

export function oddFirstPriceValue(
  settlement: number,
  maturity: number,
  issue: number,
  firstCoupon: number,
  rate: number,
  yieldRate: number,
  redemption: number,
  frequency: number,
  basis: number,
): number | undefined {
  if (!Number.isFinite(rate) || !Number.isFinite(yieldRate) || !Number.isFinite(redemption) || rate < 0 || redemption <= 0) {
    return undefined
  }
  const metrics = oddFirstCouponMetrics(settlement, maturity, issue, firstCoupon, frequency, basis)
  if (metrics === undefined) {
    return undefined
  }

  const discountBase = 1 + yieldRate / frequency
  if (discountBase <= 0) {
    return undefined
  }
  const coupon = (100 * rate) / frequency
  let price = (coupon * metrics.totalFraction) / Math.pow(discountBase, metrics.remainingFraction) - coupon * metrics.accruedFraction

  for (let period = 1; period <= metrics.regularPeriodsAfterFirst; period += 1) {
    const exponent = metrics.remainingFraction + period
    const cashflow = period === metrics.regularPeriodsAfterFirst ? redemption + coupon : coupon
    price += cashflow / Math.pow(discountBase, exponent)
  }
  return price
}

export function solveOddFirstYield(
  settlement: number,
  maturity: number,
  issue: number,
  firstCoupon: number,
  rate: number,
  price: number,
  redemption: number,
  frequency: number,
  basis: number,
): number | undefined {
  if (!Number.isFinite(rate) || !Number.isFinite(price) || !Number.isFinite(redemption) || rate < 0 || price <= 0 || redemption <= 0) {
    return undefined
  }

  const metrics = oddFirstCouponMetrics(settlement, maturity, issue, firstCoupon, frequency, basis)
  if (metrics === undefined) {
    return undefined
  }

  const priceAtYield = (yieldRate: number): number | undefined =>
    oddFirstPriceValue(settlement, maturity, issue, firstCoupon, rate, yieldRate, redemption, frequency, basis)

  let lower = -frequency + 1e-10
  let upper = Math.max(1, rate * 2 + 0.1)
  let lowerPrice = priceAtYield(lower)
  let upperPrice = priceAtYield(upper)
  for (
    let iteration = 0;
    iteration < 200 && (lowerPrice === undefined || upperPrice === undefined || lowerPrice < price || upperPrice > price);
    iteration += 1
  ) {
    if (upperPrice === undefined || upperPrice > price) {
      upper = upper * 2 + 1
      upperPrice = priceAtYield(upper)
      continue
    }
    lower = (lower - frequency) / 2
    lowerPrice = priceAtYield(lower)
  }
  if (lowerPrice === undefined || upperPrice === undefined || lowerPrice < price || upperPrice > price) {
    return undefined
  }

  let guess = Math.min(Math.max(rate, lower + 1e-8), upper - 1e-8)
  for (let iteration = 0; iteration < 200; iteration += 1) {
    const estimatedPrice = priceAtYield(guess)
    if (estimatedPrice === undefined) {
      return undefined
    }
    const error = estimatedPrice - price
    if (Math.abs(error) < 1e-14) {
      return guess
    }

    const epsilon = Math.max(1e-7, Math.abs(guess) * 1e-6)
    const shiftedPrice = priceAtYield(guess + epsilon)
    const derivative = shiftedPrice === undefined ? undefined : (shiftedPrice - estimatedPrice) / epsilon
    let nextGuess =
      derivative === undefined || !Number.isFinite(derivative) || derivative === 0 ? (lower + upper) / 2 : guess - error / derivative
    if (!Number.isFinite(nextGuess) || nextGuess <= lower || nextGuess >= upper) {
      nextGuess = (lower + upper) / 2
    }

    const boundedPrice = priceAtYield(nextGuess)
    if (boundedPrice === undefined) {
      return undefined
    }
    if (boundedPrice > price) {
      lower = nextGuess
    } else {
      upper = nextGuess
    }
    guess = nextGuess
    if (Math.abs(upper - lower) < 1e-14) {
      return (lower + upper) / 2
    }
  }

  return (lower + upper) / 2
}

function days360Us(startSerial: number, endSerial: number): number | undefined {
  const startParts = excelSerialToDateParts(startSerial)
  const endParts = excelSerialToDateParts(endSerial)
  if (startParts === undefined || endParts === undefined) {
    return undefined
  }

  let startDay = startParts.day
  let endDay = endParts.day
  if (startDay === 31) {
    startDay = 30
  }
  if (startDay === 30 && endDay === 31) {
    endDay = 30
  } else if (startParts.month === 2 && startDay === (isLeapYear(startParts.year) ? 29 : 28)) {
    startDay = 30
    if (endParts.month === 2 && endDay === (isLeapYear(endParts.year) ? 29 : 28)) {
      endDay = 30
    }
  }

  return (endParts.year - startParts.year) * 360 + (endParts.month - startParts.month) * 30 + (endDay - startDay)
}

function days360European(startSerial: number, endSerial: number): number | undefined {
  const startParts = excelSerialToDateParts(startSerial)
  const endParts = excelSerialToDateParts(endSerial)
  if (startParts === undefined || endParts === undefined) {
    return undefined
  }

  const startDay = startParts.day === 31 ? 30 : startParts.day
  const endDay = endParts.day === 31 ? 30 : endParts.day
  return (endParts.year - startParts.year) * 360 + (endParts.month - startParts.month) * 30 + (endDay - startDay)
}

function couponDaysByBasis(startSerial: number, endSerial: number, basis: number): number | undefined {
  if (!isValidBasis(basis) || startSerial > endSerial) {
    return undefined
  }
  switch (basis) {
    case 0:
      return days360Us(startSerial, endSerial)
    case 4:
      return days360European(startSerial, endSerial)
    default:
      return endSerial - startSerial
  }
}

export function couponSchedule(settlement: number, maturity: number, frequency: number): CouponSchedule | undefined {
  if (settlement >= maturity || !isValidFrequency(frequency)) {
    return undefined
  }

  const stepMonths = 12 / frequency
  let periodsRemaining = 1
  let previousCoupon = addMonthsToExcelDate(maturity, -stepMonths)
  while (previousCoupon !== undefined && previousCoupon > settlement) {
    periodsRemaining += 1
    previousCoupon = addMonthsToExcelDate(maturity, -periodsRemaining * stepMonths)
  }

  if (previousCoupon === undefined) {
    return undefined
  }
  const nextCoupon = addMonthsToExcelDate(maturity, -(periodsRemaining - 1) * stepMonths)
  if (nextCoupon === undefined || nextCoupon <= settlement) {
    return undefined
  }

  return {
    previousCoupon,
    nextCoupon,
    periodsRemaining,
  }
}

export function couponMetrics(settlement: number, maturity: number, frequency: number, basis: number): CouponMetrics | undefined {
  const schedule = couponSchedule(settlement, maturity, frequency)
  if (schedule === undefined || !isValidBasis(basis)) {
    return undefined
  }

  const accruedDays = couponDaysByBasis(schedule.previousCoupon, settlement, basis)
  const daysToNextCoupon = couponDaysByBasis(settlement, schedule.nextCoupon, basis)
  const daysInPeriod =
    basis === 1 ? couponDaysByBasis(schedule.previousCoupon, schedule.nextCoupon, basis) : basis === 3 ? 365 / frequency : 360 / frequency
  if (accruedDays === undefined || daysToNextCoupon === undefined || daysInPeriod === undefined || daysInPeriod <= 0) {
    return undefined
  }

  return {
    ...schedule,
    accruedDays,
    daysToNextCoupon,
    daysInPeriod,
  }
}

export function pricePeriodicSecurity(
  metrics: CouponMetrics,
  rate: number,
  yieldRate: number,
  redemption: number,
  frequency: number,
): number | undefined {
  const coupon = (100 * rate) / frequency
  const periodsToNextCoupon = metrics.daysToNextCoupon / metrics.daysInPeriod
  if (metrics.periodsRemaining === 1) {
    const denominator = 1 + (yieldRate / frequency) * periodsToNextCoupon
    return denominator <= 0 ? undefined : (redemption + coupon) / denominator - coupon * (metrics.accruedDays / metrics.daysInPeriod)
  }

  const discountBase = 1 + yieldRate / frequency
  if (discountBase <= 0) {
    return undefined
  }

  let price = 0
  for (let period = 1; period <= metrics.periodsRemaining; period += 1) {
    const periodsToCashflow = period - 1 + periodsToNextCoupon
    price += coupon / Math.pow(discountBase, periodsToCashflow)
  }
  price += redemption / Math.pow(discountBase, metrics.periodsRemaining - 1 + periodsToNextCoupon)
  return price - coupon * (metrics.accruedDays / metrics.daysInPeriod)
}

export function solvePeriodicSecurityYield(
  metrics: CouponMetrics,
  rate: number,
  price: number,
  redemption: number,
  frequency: number,
): number | undefined {
  const coupon = (100 * rate) / frequency
  if (metrics.periodsRemaining === 1) {
    const dirtyPrice = price + coupon * (metrics.accruedDays / metrics.daysInPeriod)
    if (dirtyPrice <= 0 || metrics.daysToNextCoupon <= 0) {
      return undefined
    }
    return ((redemption + coupon) / dirtyPrice - 1) * frequency * (metrics.daysInPeriod / metrics.daysToNextCoupon)
  }

  const targetPrice = price
  const priceAtYield = (yieldRate: number): number | undefined => pricePeriodicSecurity(metrics, rate, yieldRate, redemption, frequency)
  let lower = -frequency + 1e-10
  let upper = Math.max(1, rate * 2 + 0.1)
  let lowerPrice = priceAtYield(lower)
  let upperPrice = priceAtYield(upper)
  for (
    let iteration = 0;
    iteration < 100 && (lowerPrice === undefined || upperPrice === undefined || lowerPrice < targetPrice || upperPrice > targetPrice);
    iteration += 1
  ) {
    if (upperPrice === undefined || upperPrice > targetPrice) {
      upper = upper * 2 + 1
      upperPrice = priceAtYield(upper)
      continue
    }
    lower = (lower - frequency) / 2
    lowerPrice = priceAtYield(lower)
  }
  if (lowerPrice === undefined || upperPrice === undefined || lowerPrice < targetPrice || upperPrice > targetPrice) {
    return undefined
  }

  let guess = Math.min(Math.max(rate, lower + 1e-8), upper - 1e-8)
  for (let iteration = 0; iteration < 100; iteration += 1) {
    const estimatedPrice = priceAtYield(guess)
    if (estimatedPrice === undefined) {
      return undefined
    }
    const error = estimatedPrice - targetPrice
    if (Math.abs(error) < 1e-12) {
      return guess
    }

    const epsilon = Math.max(1e-7, Math.abs(guess) * 1e-6)
    const shiftedPrice = priceAtYield(guess + epsilon)
    const derivative = shiftedPrice === undefined ? undefined : (shiftedPrice - estimatedPrice) / epsilon
    let nextGuess =
      derivative === undefined || !Number.isFinite(derivative) || derivative === 0 ? (lower + upper) / 2 : guess - error / derivative
    if (!Number.isFinite(nextGuess) || nextGuess <= lower || nextGuess >= upper) {
      nextGuess = (lower + upper) / 2
    }

    const boundedPrice = priceAtYield(nextGuess)
    if (boundedPrice === undefined) {
      return undefined
    }
    if (boundedPrice > targetPrice) {
      lower = nextGuess
    } else {
      upper = nextGuess
    }
    guess = nextGuess
    if (Math.abs(upper - lower) < 1e-12) {
      return guess
    }
  }
  return guess
}

export function macaulayDuration(metrics: CouponMetrics, couponRate: number, yieldRate: number, frequency: number): number | undefined {
  const price = pricePeriodicSecurity(metrics, couponRate, yieldRate, 100, frequency)
  if (price === undefined || price <= 0) {
    return undefined
  }

  const coupon = (100 * couponRate) / frequency
  const periodsToNextCoupon = metrics.daysToNextCoupon / metrics.daysInPeriod
  const discountBase = 1 + yieldRate / frequency
  if (discountBase <= 0) {
    return undefined
  }

  let weightedPresentValue = 0
  for (let period = 1; period <= metrics.periodsRemaining; period += 1) {
    const periodsToCashflow = period - 1 + periodsToNextCoupon
    const timeInYears = periodsToCashflow / frequency
    const cashflow = period === metrics.periodsRemaining ? 100 + coupon : coupon
    weightedPresentValue += (timeInYears * cashflow) / Math.pow(discountBase, periodsToCashflow)
  }
  return weightedPresentValue / price
}

export function getAmordegrc(
  cost: number,
  datePurchased: number,
  firstPeriod: number,
  salvage: number,
  period: number,
  rate: number,
  basis: number,
): number | undefined {
  if (datePurchased > firstPeriod || rate <= 0 || salvage > cost || cost <= 0 || salvage < 0 || period < 0) {
    return undefined
  }

  const serialPer = Math.trunc(period)
  if (!Number.isFinite(serialPer) || serialPer < 0) {
    return undefined
  }

  const fractionalRate = yearFracByBasis(datePurchased, firstPeriod, basis)
  if (fractionalRate === undefined) {
    return undefined
  }

  const useRate = 1 / rate
  let amortizationCoefficient = 1.0
  if (useRate < 3.0) {
    amortizationCoefficient = 1.0
  } else if (useRate < 5.0) {
    amortizationCoefficient = 1.5
  } else if (useRate <= 6.0) {
    amortizationCoefficient = 2.0
  } else {
    amortizationCoefficient = 2.5
  }

  const adjustedRate = rate * amortizationCoefficient
  let currentRate = Math.round(fractionalRate * adjustedRate * cost)
  let currentCost = cost - currentRate
  let remaining = currentCost - salvage

  for (let step = 0; step < serialPer; step += 1) {
    currentRate = Math.round(adjustedRate * currentCost)
    remaining -= currentRate
    if (remaining < 0.0) {
      if (serialPer - step === 0 || serialPer - step === 1) {
        return Math.round(currentCost * 0.5)
      }
      return 0.0
    }
    currentCost -= currentRate
  }

  return currentRate
}

export function getAmorlinc(
  cost: number,
  datePurchased: number,
  firstPeriod: number,
  salvage: number,
  period: number,
  rate: number,
  basis: number,
): number | undefined {
  if (datePurchased > firstPeriod || rate <= 0 || salvage > cost || cost <= 0 || salvage < 0 || period < 0) {
    return undefined
  }

  const serialPer = Math.trunc(period)
  if (!Number.isFinite(serialPer) || serialPer < 0) {
    return undefined
  }

  const fractionalRate = yearFracByBasis(datePurchased, firstPeriod, basis)
  if (fractionalRate === undefined) {
    return undefined
  }

  const fullRate = cost * rate
  const remainingCost = cost - salvage
  const firstRate = fractionalRate * rate * cost
  const fullPeriods = Math.trunc((cost - salvage - firstRate) / fullRate)

  let result = 0.0
  if (serialPer === 0) {
    result = firstRate
  } else if (serialPer <= fullPeriods) {
    result = fullRate
  } else if (serialPer === fullPeriods + 1) {
    result = remainingCost - fullRate * fullPeriods - firstRate
  }

  return result > 0.0 ? result : 0.0
}
