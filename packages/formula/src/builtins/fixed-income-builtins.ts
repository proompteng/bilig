import type { CellValue } from '@bilig/protocol'
import {
  couponMetrics,
  couponSchedule,
  getAmordegrc,
  getAmorlinc,
  isValidBasis,
  isValidFrequency,
  macaulayDuration,
  maturityAtIssueFractions,
  oddFirstPriceValue,
  oddLastCouponFractions,
  pricePeriodicSecurity,
  securityAnnualizedYearFraction,
  solveOddFirstYield,
  solvePeriodicSecurityYield,
  treasuryBillDays,
  yearFracByBasis,
} from './fixed-income.js'
import type { EvaluationResult } from '../runtime-values.js'

type Builtin = (...args: CellValue[]) => EvaluationResult

interface FixedIncomeBuiltinDeps {
  toNumber: (value: CellValue) => number | undefined
  coerceBoolean: (value: CellValue | undefined, fallback: boolean) => boolean | undefined
  coerceDateSerial: (value: CellValue | undefined) => number | undefined
  coerceNumber: (value: CellValue | undefined, fallback: number) => number | undefined
  integerValue: (value: CellValue | undefined, fallback?: number) => number | undefined
  numberResult: (value: number) => EvaluationResult
  valueError: () => EvaluationResult
}

function isValidBasisValue(basis: number | undefined): basis is number {
  return basis !== undefined && isValidBasis(basis)
}

function isValidFrequencyValue(frequency: number | undefined): frequency is number {
  return frequency !== undefined && isValidFrequency(frequency)
}

export function createFixedIncomeBuiltins({
  toNumber,
  coerceBoolean,
  coerceDateSerial,
  coerceNumber,
  integerValue,
  numberResult,
  valueError,
}: FixedIncomeBuiltinDeps): Record<string, Builtin> {
  return {
    ACCRINT: (issueArg, firstInterestArg, settlementArg, rateArg, parArg, frequencyArg, basisArg, calcMethodArg) => {
      const issue = coerceDateSerial(issueArg)
      const firstInterest = coerceDateSerial(firstInterestArg)
      const settlement = coerceDateSerial(settlementArg)
      const rate = toNumber(rateArg)
      const par = coerceNumber(parArg, 1000)
      const frequency = integerValue(frequencyArg)
      const basis = integerValue(basisArg, 0)
      const calcMethod = coerceBoolean(calcMethodArg, true)
      if (
        issue === undefined ||
        firstInterest === undefined ||
        settlement === undefined ||
        rate === undefined ||
        calcMethod === undefined ||
        par === undefined ||
        rate <= 0 ||
        par <= 0 ||
        issue >= settlement ||
        firstInterest <= issue ||
        firstInterest >= settlement ||
        !isValidFrequencyValue(frequency) ||
        !isValidBasisValue(basis)
      ) {
        return valueError()
      }
      const accrualStart = settlement > firstInterest && !calcMethod ? firstInterest : issue
      const years = yearFracByBasis(accrualStart, settlement, basis)
      return years === undefined ? valueError() : numberResult(par * rate * years)
    },
    ACCRINTM: (issueArg, settlementArg, rateArg, parArg, basisArg) => {
      const issue = coerceDateSerial(issueArg)
      const settlement = coerceDateSerial(settlementArg)
      const rate = toNumber(rateArg)
      const par = coerceNumber(parArg, 1000)
      const basis = integerValue(basisArg, 0)
      if (
        issue === undefined ||
        settlement === undefined ||
        rate === undefined ||
        par === undefined ||
        rate <= 0 ||
        par <= 0 ||
        issue >= settlement ||
        !isValidBasisValue(basis)
      ) {
        return valueError()
      }
      const years = yearFracByBasis(issue, settlement, basis)
      return years === undefined ? valueError() : numberResult(par * rate * years)
    },
    COUPDAYBS: (settlementArg, maturityArg, frequencyArg, basisArg) => {
      const settlement = coerceDateSerial(settlementArg)
      const maturity = coerceDateSerial(maturityArg)
      const frequency = integerValue(frequencyArg)
      const basis = integerValue(basisArg, 0)
      if (settlement === undefined || maturity === undefined || !isValidFrequencyValue(frequency) || !isValidBasisValue(basis)) {
        return valueError()
      }
      const metrics = couponMetrics(settlement, maturity, frequency, basis)
      return metrics === undefined ? valueError() : numberResult(metrics.accruedDays)
    },
    COUPDAYS: (settlementArg, maturityArg, frequencyArg, basisArg) => {
      const settlement = coerceDateSerial(settlementArg)
      const maturity = coerceDateSerial(maturityArg)
      const frequency = integerValue(frequencyArg)
      const basis = integerValue(basisArg, 0)
      if (settlement === undefined || maturity === undefined || !isValidFrequencyValue(frequency) || !isValidBasisValue(basis)) {
        return valueError()
      }
      const metrics = couponMetrics(settlement, maturity, frequency, basis)
      return metrics === undefined ? valueError() : numberResult(metrics.daysInPeriod)
    },
    COUPDAYSNC: (settlementArg, maturityArg, frequencyArg, basisArg) => {
      const settlement = coerceDateSerial(settlementArg)
      const maturity = coerceDateSerial(maturityArg)
      const frequency = integerValue(frequencyArg)
      const basis = integerValue(basisArg, 0)
      if (settlement === undefined || maturity === undefined || !isValidFrequencyValue(frequency) || !isValidBasisValue(basis)) {
        return valueError()
      }
      const metrics = couponMetrics(settlement, maturity, frequency, basis)
      return metrics === undefined ? valueError() : numberResult(metrics.daysToNextCoupon)
    },
    COUPNCD: (settlementArg, maturityArg, frequencyArg, basisArg) => {
      const settlement = coerceDateSerial(settlementArg)
      const maturity = coerceDateSerial(maturityArg)
      const frequency = integerValue(frequencyArg)
      const basis = integerValue(basisArg, 0)
      if (settlement === undefined || maturity === undefined || !isValidFrequencyValue(frequency) || !isValidBasisValue(basis)) {
        return valueError()
      }
      const schedule = couponSchedule(settlement, maturity, frequency)
      return schedule === undefined ? valueError() : numberResult(schedule.nextCoupon)
    },
    COUPNUM: (settlementArg, maturityArg, frequencyArg, basisArg) => {
      const settlement = coerceDateSerial(settlementArg)
      const maturity = coerceDateSerial(maturityArg)
      const frequency = integerValue(frequencyArg)
      const basis = integerValue(basisArg, 0)
      if (settlement === undefined || maturity === undefined || !isValidFrequencyValue(frequency) || !isValidBasisValue(basis)) {
        return valueError()
      }
      const schedule = couponSchedule(settlement, maturity, frequency)
      return schedule === undefined ? valueError() : numberResult(schedule.periodsRemaining)
    },
    COUPPCD: (settlementArg, maturityArg, frequencyArg, basisArg) => {
      const settlement = coerceDateSerial(settlementArg)
      const maturity = coerceDateSerial(maturityArg)
      const frequency = integerValue(frequencyArg)
      const basis = integerValue(basisArg, 0)
      if (settlement === undefined || maturity === undefined || !isValidFrequencyValue(frequency) || !isValidBasisValue(basis)) {
        return valueError()
      }
      const schedule = couponSchedule(settlement, maturity, frequency)
      return schedule === undefined ? valueError() : numberResult(schedule.previousCoupon)
    },
    DISC: (settlementArg, maturityArg, priceArg, redemptionArg, basisArg) => {
      const settlement = coerceDateSerial(settlementArg)
      const maturity = coerceDateSerial(maturityArg)
      const price = toNumber(priceArg)
      const redemption = toNumber(redemptionArg)
      const basis = integerValue(basisArg, 0)
      if (
        settlement === undefined ||
        maturity === undefined ||
        price === undefined ||
        redemption === undefined ||
        price <= 0 ||
        redemption <= 0 ||
        !isValidBasisValue(basis)
      ) {
        return valueError()
      }
      const years = securityAnnualizedYearFraction(settlement, maturity, basis)
      return years === undefined ? valueError() : numberResult((redemption - price) / redemption / years)
    },
    INTRATE: (settlementArg, maturityArg, investmentArg, redemptionArg, basisArg) => {
      const settlement = coerceDateSerial(settlementArg)
      const maturity = coerceDateSerial(maturityArg)
      const investment = toNumber(investmentArg)
      const redemption = toNumber(redemptionArg)
      const basis = integerValue(basisArg, 0)
      if (
        settlement === undefined ||
        maturity === undefined ||
        investment === undefined ||
        redemption === undefined ||
        investment <= 0 ||
        redemption <= 0 ||
        !isValidBasisValue(basis)
      ) {
        return valueError()
      }
      const years = securityAnnualizedYearFraction(settlement, maturity, basis)
      return years === undefined ? valueError() : numberResult((redemption - investment) / investment / years)
    },
    RECEIVED: (settlementArg, maturityArg, investmentArg, discountArg, basisArg) => {
      const settlement = coerceDateSerial(settlementArg)
      const maturity = coerceDateSerial(maturityArg)
      const investment = toNumber(investmentArg)
      const discount = toNumber(discountArg)
      const basis = integerValue(basisArg, 0)
      if (
        settlement === undefined ||
        maturity === undefined ||
        investment === undefined ||
        discount === undefined ||
        investment <= 0 ||
        discount <= 0 ||
        !isValidBasisValue(basis)
      ) {
        return valueError()
      }
      const years = securityAnnualizedYearFraction(settlement, maturity, basis)
      if (years === undefined) {
        return valueError()
      }
      const denominator = 1 - discount * years
      return denominator <= 0 ? valueError() : numberResult(investment / denominator)
    },
    PRICEDISC: (settlementArg, maturityArg, discountArg, redemptionArg, basisArg) => {
      const settlement = coerceDateSerial(settlementArg)
      const maturity = coerceDateSerial(maturityArg)
      const discount = toNumber(discountArg)
      const redemption = toNumber(redemptionArg)
      const basis = integerValue(basisArg, 0)
      if (
        settlement === undefined ||
        maturity === undefined ||
        discount === undefined ||
        redemption === undefined ||
        discount <= 0 ||
        redemption <= 0 ||
        !isValidBasisValue(basis)
      ) {
        return valueError()
      }
      const years = securityAnnualizedYearFraction(settlement, maturity, basis)
      return years === undefined ? valueError() : numberResult(redemption * (1 - discount * years))
    },
    YIELDDISC: (settlementArg, maturityArg, priceArg, redemptionArg, basisArg) => {
      const settlement = coerceDateSerial(settlementArg)
      const maturity = coerceDateSerial(maturityArg)
      const price = toNumber(priceArg)
      const redemption = toNumber(redemptionArg)
      const basis = integerValue(basisArg, 0)
      if (
        settlement === undefined ||
        maturity === undefined ||
        price === undefined ||
        redemption === undefined ||
        price <= 0 ||
        redemption <= 0 ||
        !isValidBasisValue(basis)
      ) {
        return valueError()
      }
      const years = securityAnnualizedYearFraction(settlement, maturity, basis)
      return years === undefined ? valueError() : numberResult((redemption - price) / price / years)
    },
    TBILLPRICE: (settlementArg, maturityArg, discountArg) => {
      const settlement = coerceDateSerial(settlementArg)
      const maturity = coerceDateSerial(maturityArg)
      const discount = toNumber(discountArg)
      if (settlement === undefined || maturity === undefined || discount === undefined || discount <= 0) {
        return valueError()
      }
      const days = treasuryBillDays(settlement, maturity)
      return days === undefined ? valueError() : numberResult(100 * (1 - (discount * days) / 360))
    },
    TBILLYIELD: (settlementArg, maturityArg, priceArg) => {
      const settlement = coerceDateSerial(settlementArg)
      const maturity = coerceDateSerial(maturityArg)
      const price = toNumber(priceArg)
      if (settlement === undefined || maturity === undefined || price === undefined || price <= 0) {
        return valueError()
      }
      const days = treasuryBillDays(settlement, maturity)
      return days === undefined ? valueError() : numberResult(((100 - price) * 360) / (price * days))
    },
    TBILLEQ: (settlementArg, maturityArg, discountArg) => {
      const settlement = coerceDateSerial(settlementArg)
      const maturity = coerceDateSerial(maturityArg)
      const discount = toNumber(discountArg)
      if (settlement === undefined || maturity === undefined || discount === undefined || discount <= 0) {
        return valueError()
      }
      const days = treasuryBillDays(settlement, maturity)
      if (days === undefined) {
        return valueError()
      }
      const denominator = 360 - discount * days
      return denominator === 0 ? valueError() : numberResult((365 * discount) / denominator)
    },
    PRICEMAT: (settlementArg, maturityArg, issueArg, rateArg, yieldArg, basisArg) => {
      const settlement = coerceDateSerial(settlementArg)
      const maturity = coerceDateSerial(maturityArg)
      const issue = coerceDateSerial(issueArg)
      const rate = toNumber(rateArg)
      const yieldRate = toNumber(yieldArg)
      const basis = integerValue(basisArg, 0)
      if (
        settlement === undefined ||
        maturity === undefined ||
        issue === undefined ||
        rate === undefined ||
        yieldRate === undefined ||
        rate < 0 ||
        yieldRate < 0 ||
        !isValidBasisValue(basis)
      ) {
        return valueError()
      }
      const fractions = maturityAtIssueFractions(settlement, maturity, issue, basis)
      if (fractions === undefined) {
        return valueError()
      }
      const maturityValue = 100 * (1 + rate * fractions.issueToMaturity)
      const accruedInterest = 100 * rate * fractions.issueToSettlement
      const denominator = 1 + yieldRate * fractions.settlementToMaturity
      return numberResult(maturityValue / denominator - accruedInterest)
    },
    YIELDMAT: (settlementArg, maturityArg, issueArg, rateArg, priceArg, basisArg) => {
      const settlement = coerceDateSerial(settlementArg)
      const maturity = coerceDateSerial(maturityArg)
      const issue = coerceDateSerial(issueArg)
      const rate = toNumber(rateArg)
      const price = toNumber(priceArg)
      const basis = integerValue(basisArg, 0)
      if (
        settlement === undefined ||
        maturity === undefined ||
        issue === undefined ||
        rate === undefined ||
        price === undefined ||
        rate < 0 ||
        price <= 0 ||
        !isValidBasisValue(basis)
      ) {
        return valueError()
      }
      const fractions = maturityAtIssueFractions(settlement, maturity, issue, basis)
      if (fractions === undefined) {
        return valueError()
      }
      const settlementValue = price + 100 * rate * fractions.issueToSettlement
      const maturityValue = 100 * (1 + rate * fractions.issueToMaturity)
      return numberResult((maturityValue / settlementValue - 1) / fractions.settlementToMaturity)
    },
    ODDFPRICE: (settlementArg, maturityArg, issueArg, firstCouponArg, rateArg, yieldArg, redemptionArg, frequencyArg, basisArg) => {
      const settlement = coerceDateSerial(settlementArg)
      const maturity = coerceDateSerial(maturityArg)
      const issue = coerceDateSerial(issueArg)
      const firstCoupon = coerceDateSerial(firstCouponArg)
      const rate = toNumber(rateArg)
      const yieldRate = toNumber(yieldArg)
      const redemption = toNumber(redemptionArg)
      const frequency = integerValue(frequencyArg)
      const basis = integerValue(basisArg, 0)
      if (
        settlement === undefined ||
        maturity === undefined ||
        issue === undefined ||
        firstCoupon === undefined ||
        rate === undefined ||
        yieldRate === undefined ||
        redemption === undefined ||
        rate < 0 ||
        yieldRate < 0 ||
        redemption <= 0 ||
        !isValidFrequencyValue(frequency) ||
        !isValidBasisValue(basis)
      ) {
        return valueError()
      }
      const price = oddFirstPriceValue(settlement, maturity, issue, firstCoupon, rate, yieldRate, redemption, frequency, basis)
      return price === undefined ? valueError() : numberResult(price)
    },
    ODDFYIELD: (settlementArg, maturityArg, issueArg, firstCouponArg, rateArg, priceArg, redemptionArg, frequencyArg, basisArg) => {
      const settlement = coerceDateSerial(settlementArg)
      const maturity = coerceDateSerial(maturityArg)
      const issue = coerceDateSerial(issueArg)
      const firstCoupon = coerceDateSerial(firstCouponArg)
      const rate = toNumber(rateArg)
      const price = toNumber(priceArg)
      const redemption = toNumber(redemptionArg)
      const frequency = integerValue(frequencyArg)
      const basis = integerValue(basisArg, 0)
      if (
        settlement === undefined ||
        maturity === undefined ||
        issue === undefined ||
        firstCoupon === undefined ||
        rate === undefined ||
        price === undefined ||
        redemption === undefined ||
        rate < 0 ||
        price <= 0 ||
        redemption <= 0 ||
        !isValidFrequencyValue(frequency) ||
        !isValidBasisValue(basis)
      ) {
        return valueError()
      }
      const yieldRate = solveOddFirstYield(settlement, maturity, issue, firstCoupon, rate, price, redemption, frequency, basis)
      return yieldRate === undefined ? valueError() : numberResult(yieldRate)
    },
    ODDLPRICE: (settlementArg, maturityArg, lastInterestArg, rateArg, yieldArg, redemptionArg, frequencyArg, basisArg) => {
      const settlement = coerceDateSerial(settlementArg)
      const maturity = coerceDateSerial(maturityArg)
      const lastInterest = coerceDateSerial(lastInterestArg)
      const rate = toNumber(rateArg)
      const yieldRate = toNumber(yieldArg)
      const redemption = toNumber(redemptionArg)
      const frequency = integerValue(frequencyArg)
      const basis = integerValue(basisArg, 0)
      if (
        settlement === undefined ||
        maturity === undefined ||
        lastInterest === undefined ||
        rate === undefined ||
        yieldRate === undefined ||
        redemption === undefined ||
        rate < 0 ||
        yieldRate < 0 ||
        redemption <= 0 ||
        !isValidFrequencyValue(frequency) ||
        !isValidBasisValue(basis)
      ) {
        return valueError()
      }

      const fractions = oddLastCouponFractions(settlement, maturity, lastInterest, frequency, basis)
      if (fractions === undefined) {
        return valueError()
      }

      const coupon = (100 * rate) / frequency
      const maturityValue = redemption + coupon * fractions.totalFraction
      const denominator = 1 + (yieldRate * fractions.remainingFraction) / frequency
      if (denominator <= 0) {
        return valueError()
      }
      return numberResult(maturityValue / denominator - coupon * fractions.accruedFraction)
    },
    ODDLYIELD: (settlementArg, maturityArg, lastInterestArg, rateArg, priceArg, redemptionArg, frequencyArg, basisArg) => {
      const settlement = coerceDateSerial(settlementArg)
      const maturity = coerceDateSerial(maturityArg)
      const lastInterest = coerceDateSerial(lastInterestArg)
      const rate = toNumber(rateArg)
      const price = toNumber(priceArg)
      const redemption = toNumber(redemptionArg)
      const frequency = integerValue(frequencyArg)
      const basis = integerValue(basisArg, 0)
      if (
        settlement === undefined ||
        maturity === undefined ||
        lastInterest === undefined ||
        rate === undefined ||
        price === undefined ||
        redemption === undefined ||
        rate < 0 ||
        price <= 0 ||
        redemption <= 0 ||
        !isValidFrequencyValue(frequency) ||
        !isValidBasisValue(basis)
      ) {
        return valueError()
      }

      const fractions = oddLastCouponFractions(settlement, maturity, lastInterest, frequency, basis)
      if (fractions === undefined) {
        return valueError()
      }

      const coupon = (100 * rate) / frequency
      const dirtyPrice = price + coupon * fractions.accruedFraction
      if (dirtyPrice <= 0 || fractions.remainingFraction <= 0) {
        return valueError()
      }
      const maturityValue = redemption + coupon * fractions.totalFraction
      return numberResult(((maturityValue - dirtyPrice) / dirtyPrice) * (frequency / fractions.remainingFraction))
    },
    PRICE: (settlementArg, maturityArg, rateArg, yieldArg, redemptionArg, frequencyArg, basisArg) => {
      const settlement = coerceDateSerial(settlementArg)
      const maturity = coerceDateSerial(maturityArg)
      const rate = toNumber(rateArg)
      const yieldRate = toNumber(yieldArg)
      const redemption = toNumber(redemptionArg)
      const frequency = integerValue(frequencyArg)
      const basis = integerValue(basisArg, 0)
      if (
        settlement === undefined ||
        maturity === undefined ||
        rate === undefined ||
        yieldRate === undefined ||
        redemption === undefined ||
        rate < 0 ||
        yieldRate < 0 ||
        redemption <= 0 ||
        !isValidFrequencyValue(frequency) ||
        !isValidBasisValue(basis)
      ) {
        return valueError()
      }
      const metrics = couponMetrics(settlement, maturity, frequency, basis)
      const price = metrics ? pricePeriodicSecurity(metrics, rate, yieldRate, redemption, frequency) : undefined
      return price === undefined ? valueError() : numberResult(price)
    },
    YIELD: (settlementArg, maturityArg, rateArg, priceArg, redemptionArg, frequencyArg, basisArg) => {
      const settlement = coerceDateSerial(settlementArg)
      const maturity = coerceDateSerial(maturityArg)
      const rate = toNumber(rateArg)
      const price = toNumber(priceArg)
      const redemption = toNumber(redemptionArg)
      const frequency = integerValue(frequencyArg)
      const basis = integerValue(basisArg, 0)
      if (
        settlement === undefined ||
        maturity === undefined ||
        rate === undefined ||
        price === undefined ||
        redemption === undefined ||
        rate < 0 ||
        price <= 0 ||
        redemption <= 0 ||
        !isValidFrequencyValue(frequency) ||
        !isValidBasisValue(basis)
      ) {
        return valueError()
      }
      const metrics = couponMetrics(settlement, maturity, frequency, basis)
      const yieldRate = metrics ? solvePeriodicSecurityYield(metrics, rate, price, redemption, frequency) : undefined
      return yieldRate === undefined ? valueError() : numberResult(yieldRate)
    },
    DURATION: (settlementArg, maturityArg, couponArg, yieldArg, frequencyArg, basisArg) => {
      const settlement = coerceDateSerial(settlementArg)
      const maturity = coerceDateSerial(maturityArg)
      const coupon = toNumber(couponArg)
      const yieldRate = toNumber(yieldArg)
      const frequency = integerValue(frequencyArg)
      const basis = integerValue(basisArg, 0)
      if (
        settlement === undefined ||
        maturity === undefined ||
        coupon === undefined ||
        yieldRate === undefined ||
        coupon < 0 ||
        yieldRate < 0 ||
        !isValidFrequencyValue(frequency) ||
        !isValidBasisValue(basis)
      ) {
        return valueError()
      }
      const metrics = couponMetrics(settlement, maturity, frequency, basis)
      const duration = metrics ? macaulayDuration(metrics, coupon, yieldRate, frequency) : undefined
      return duration === undefined ? valueError() : numberResult(duration)
    },
    MDURATION: (settlementArg, maturityArg, couponArg, yieldArg, frequencyArg, basisArg) => {
      const settlement = coerceDateSerial(settlementArg)
      const maturity = coerceDateSerial(maturityArg)
      const coupon = toNumber(couponArg)
      const yieldRate = toNumber(yieldArg)
      const frequency = integerValue(frequencyArg)
      const basis = integerValue(basisArg, 0)
      if (
        settlement === undefined ||
        maturity === undefined ||
        coupon === undefined ||
        yieldRate === undefined ||
        coupon < 0 ||
        yieldRate < 0 ||
        !isValidFrequencyValue(frequency) ||
        !isValidBasisValue(basis)
      ) {
        return valueError()
      }
      const metrics = couponMetrics(settlement, maturity, frequency, basis)
      const duration = metrics ? macaulayDuration(metrics, coupon, yieldRate, frequency) : undefined
      return duration === undefined ? valueError() : numberResult(duration / (1 + yieldRate / frequency))
    },
    AMORDEGRC: (costArg, datePurchasedArg, firstPeriodArg, salvageArg, periodArg, rateArg, basisArg) => {
      const cost = toNumber(costArg)
      const datePurchased = coerceDateSerial(datePurchasedArg)
      const firstPeriod = coerceDateSerial(firstPeriodArg)
      const salvage = toNumber(salvageArg)
      const period = toNumber(periodArg)
      const rate = toNumber(rateArg)
      const basis = integerValue(basisArg, 0)
      if (
        cost === undefined ||
        datePurchased === undefined ||
        firstPeriod === undefined ||
        salvage === undefined ||
        period === undefined ||
        rate === undefined ||
        !isValidBasisValue(basis)
      ) {
        return valueError()
      }
      const depreciation = getAmordegrc(cost, datePurchased, firstPeriod, salvage, period, rate, basis)
      return depreciation === undefined ? valueError() : numberResult(depreciation)
    },
    AMORLINC: (costArg, datePurchasedArg, firstPeriodArg, salvageArg, periodArg, rateArg, basisArg) => {
      const cost = toNumber(costArg)
      const datePurchased = coerceDateSerial(datePurchasedArg)
      const firstPeriod = coerceDateSerial(firstPeriodArg)
      const salvage = toNumber(salvageArg)
      const period = toNumber(periodArg)
      const rate = toNumber(rateArg)
      const basis = integerValue(basisArg, 0)
      if (
        cost === undefined ||
        datePurchased === undefined ||
        firstPeriod === undefined ||
        salvage === undefined ||
        period === undefined ||
        rate === undefined ||
        !isValidBasisValue(basis)
      ) {
        return valueError()
      }
      const depreciation = getAmorlinc(cost, datePurchased, firstPeriod, salvage, period, rate, basis)
      return depreciation === undefined ? valueError() : numberResult(depreciation)
    },
  }
}
