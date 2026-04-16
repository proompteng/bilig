import { BuiltinId, ErrorCode, ValueTag } from './protocol'
import {
  accruedIssueYearfracValue,
  couponDateFromMaturityValue,
  couponDaysByBasisValue,
  couponPeriodDaysValue,
  couponPeriodsRemainingValue,
  couponPriceFromMetricsValue,
  excelSerialWhole,
  macaulayDurationValue,
  maturityIssueYearfracValue,
  oddFirstPriceValue,
  oddFirstYieldValue,
  oddLastPriceValue,
  oddLastYieldValue,
  securityAnnualizedYearfracValue,
  solveCouponYieldValue,
  treasuryBillDaysValue,
} from './date-finance'
import { truncToInt } from './numeric-core'
import { toNumberExact } from './operands'
import { scalarErrorAt } from './builtin-args'
import { STACK_KIND_SCALAR, writeResult } from './result-io'

export function tryApplyFinanceSecuritiesBuiltin(
  builtinId: i32,
  argc: i32,
  base: i32,
  rangeIndexStack: Uint32Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  kindStack: Uint8Array,
): i32 {
  if (builtinId == BuiltinId.Disc && (argc == 4 || argc == 5)) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const settlement = excelSerialWhole(tagStack[base], valueStack[base])
    const maturity = excelSerialWhole(tagStack[base + 1], valueStack[base + 1])
    const price = toNumberExact(tagStack[base + 2], valueStack[base + 2])
    const redemption = toNumberExact(tagStack[base + 3], valueStack[base + 3])
    const basis = argc == 5 ? truncToInt(tagStack[base + 4], valueStack[base + 4]) : 0
    const years =
      settlement == i32.MIN_VALUE || maturity == i32.MIN_VALUE ? NaN : securityAnnualizedYearfracValue(settlement, maturity, basis)
    const value =
      isNaN(price) || isNaN(redemption) || redemption <= 0.0 || price <= 0.0 || isNaN(years)
        ? NaN
        : (redemption - price) / redemption / years
    return isNaN(value)
      ? writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, value, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (
    (builtinId == BuiltinId.Coupdaybs ||
      builtinId == BuiltinId.Coupdays ||
      builtinId == BuiltinId.Coupdaysnc ||
      builtinId == BuiltinId.Coupncd ||
      builtinId == BuiltinId.Coupnum ||
      builtinId == BuiltinId.Couppcd) &&
    (argc == 3 || argc == 4)
  ) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const settlement = excelSerialWhole(tagStack[base], valueStack[base])
    const maturity = excelSerialWhole(tagStack[base + 1], valueStack[base + 1])
    const frequency = truncToInt(tagStack[base + 2], valueStack[base + 2])
    const basis = argc == 4 ? truncToInt(tagStack[base + 3], valueStack[base + 3]) : 0
    const periodsRemaining =
      settlement == i32.MIN_VALUE || maturity == i32.MIN_VALUE
        ? i32.MIN_VALUE
        : couponPeriodsRemainingValue(settlement, maturity, frequency)
    const previousCoupon =
      periodsRemaining == i32.MIN_VALUE ? i32.MIN_VALUE : couponDateFromMaturityValue(maturity, periodsRemaining, frequency)
    const nextCoupon =
      periodsRemaining == i32.MIN_VALUE ? i32.MIN_VALUE : couponDateFromMaturityValue(maturity, periodsRemaining - 1, frequency)
    const accruedDays =
      previousCoupon == i32.MIN_VALUE || settlement == i32.MIN_VALUE ? NaN : couponDaysByBasisValue(previousCoupon, settlement, basis)
    const daysToNextCoupon =
      nextCoupon == i32.MIN_VALUE || settlement == i32.MIN_VALUE ? NaN : couponDaysByBasisValue(settlement, nextCoupon, basis)
    const daysInPeriod =
      previousCoupon == i32.MIN_VALUE || nextCoupon == i32.MIN_VALUE
        ? NaN
        : couponPeriodDaysValue(previousCoupon, nextCoupon, basis, frequency)
    let value = NaN
    if (builtinId == BuiltinId.Coupdaybs) {
      value = accruedDays
    } else if (builtinId == BuiltinId.Coupdays) {
      value = daysInPeriod
    } else if (builtinId == BuiltinId.Coupdaysnc) {
      value = daysToNextCoupon
    } else if (builtinId == BuiltinId.Coupncd) {
      value = nextCoupon == i32.MIN_VALUE ? NaN : <f64>nextCoupon
    } else if (builtinId == BuiltinId.Coupnum) {
      value = periodsRemaining == i32.MIN_VALUE ? NaN : <f64>periodsRemaining
    } else {
      value = previousCoupon == i32.MIN_VALUE ? NaN : <f64>previousCoupon
    }
    return isNaN(value)
      ? writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, value, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Intrate && (argc == 4 || argc == 5)) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const settlement = excelSerialWhole(tagStack[base], valueStack[base])
    const maturity = excelSerialWhole(tagStack[base + 1], valueStack[base + 1])
    const investment = toNumberExact(tagStack[base + 2], valueStack[base + 2])
    const redemption = toNumberExact(tagStack[base + 3], valueStack[base + 3])
    const basis = argc == 5 ? truncToInt(tagStack[base + 4], valueStack[base + 4]) : 0
    const years =
      settlement == i32.MIN_VALUE || maturity == i32.MIN_VALUE ? NaN : securityAnnualizedYearfracValue(settlement, maturity, basis)
    const value =
      isNaN(investment) || isNaN(redemption) || investment <= 0.0 || redemption <= 0.0 || isNaN(years)
        ? NaN
        : (redemption - investment) / investment / years
    return isNaN(value)
      ? writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, value, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Received && (argc == 4 || argc == 5)) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const settlement = excelSerialWhole(tagStack[base], valueStack[base])
    const maturity = excelSerialWhole(tagStack[base + 1], valueStack[base + 1])
    const investment = toNumberExact(tagStack[base + 2], valueStack[base + 2])
    const discount = toNumberExact(tagStack[base + 3], valueStack[base + 3])
    const basis = argc == 5 ? truncToInt(tagStack[base + 4], valueStack[base + 4]) : 0
    const years =
      settlement == i32.MIN_VALUE || maturity == i32.MIN_VALUE ? NaN : securityAnnualizedYearfracValue(settlement, maturity, basis)
    const denominator = isNaN(years) ? NaN : 1.0 - discount * years
    const value =
      isNaN(investment) || isNaN(discount) || investment <= 0.0 || discount <= 0.0 || !isFinite(denominator) || denominator <= 0.0
        ? NaN
        : investment / denominator
    return isNaN(value)
      ? writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, value, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Pricedisc && (argc == 4 || argc == 5)) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const settlement = excelSerialWhole(tagStack[base], valueStack[base])
    const maturity = excelSerialWhole(tagStack[base + 1], valueStack[base + 1])
    const discount = toNumberExact(tagStack[base + 2], valueStack[base + 2])
    const redemption = toNumberExact(tagStack[base + 3], valueStack[base + 3])
    const basis = argc == 5 ? truncToInt(tagStack[base + 4], valueStack[base + 4]) : 0
    const years =
      settlement == i32.MIN_VALUE || maturity == i32.MIN_VALUE ? NaN : securityAnnualizedYearfracValue(settlement, maturity, basis)
    const value =
      isNaN(discount) || isNaN(redemption) || discount <= 0.0 || redemption <= 0.0 || isNaN(years)
        ? NaN
        : redemption * (1.0 - discount * years)
    return isNaN(value)
      ? writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, value, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Yielddisc && (argc == 4 || argc == 5)) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const settlement = excelSerialWhole(tagStack[base], valueStack[base])
    const maturity = excelSerialWhole(tagStack[base + 1], valueStack[base + 1])
    const price = toNumberExact(tagStack[base + 2], valueStack[base + 2])
    const redemption = toNumberExact(tagStack[base + 3], valueStack[base + 3])
    const basis = argc == 5 ? truncToInt(tagStack[base + 4], valueStack[base + 4]) : 0
    const years =
      settlement == i32.MIN_VALUE || maturity == i32.MIN_VALUE ? NaN : securityAnnualizedYearfracValue(settlement, maturity, basis)
    const value =
      isNaN(price) || isNaN(redemption) || price <= 0.0 || redemption <= 0.0 || isNaN(years) ? NaN : (redemption - price) / price / years
    return isNaN(value)
      ? writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, value, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Tbillprice && argc == 3) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const settlement = excelSerialWhole(tagStack[base], valueStack[base])
    const maturity = excelSerialWhole(tagStack[base + 1], valueStack[base + 1])
    const discount = toNumberExact(tagStack[base + 2], valueStack[base + 2])
    const days = settlement == i32.MIN_VALUE || maturity == i32.MIN_VALUE ? NaN : treasuryBillDaysValue(settlement, maturity)
    const value = isNaN(discount) || discount <= 0.0 || isNaN(days) ? NaN : 100.0 * (1.0 - (discount * days) / 360.0)
    return isNaN(value)
      ? writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, value, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Tbillyield && argc == 3) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const settlement = excelSerialWhole(tagStack[base], valueStack[base])
    const maturity = excelSerialWhole(tagStack[base + 1], valueStack[base + 1])
    const price = toNumberExact(tagStack[base + 2], valueStack[base + 2])
    const days = settlement == i32.MIN_VALUE || maturity == i32.MIN_VALUE ? NaN : treasuryBillDaysValue(settlement, maturity)
    const value = isNaN(price) || price <= 0.0 || isNaN(days) ? NaN : ((100.0 - price) * 360.0) / (price * days)
    return isNaN(value)
      ? writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, value, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Tbilleq && argc == 3) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const settlement = excelSerialWhole(tagStack[base], valueStack[base])
    const maturity = excelSerialWhole(tagStack[base + 1], valueStack[base + 1])
    const discount = toNumberExact(tagStack[base + 2], valueStack[base + 2])
    const days = settlement == i32.MIN_VALUE || maturity == i32.MIN_VALUE ? NaN : treasuryBillDaysValue(settlement, maturity)
    const denominator = isNaN(days) ? NaN : 360.0 - discount * days
    const value =
      isNaN(discount) || discount <= 0.0 || !isFinite(denominator) || denominator == 0.0 ? NaN : (365.0 * discount) / denominator
    return isNaN(value)
      ? writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, value, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Pricemat && (argc == 5 || argc == 6)) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const settlement = excelSerialWhole(tagStack[base], valueStack[base])
    const maturity = excelSerialWhole(tagStack[base + 1], valueStack[base + 1])
    const issue = excelSerialWhole(tagStack[base + 2], valueStack[base + 2])
    const rate = toNumberExact(tagStack[base + 3], valueStack[base + 3])
    const yieldRate = toNumberExact(tagStack[base + 4], valueStack[base + 4])
    const basis = argc == 6 ? truncToInt(tagStack[base + 5], valueStack[base + 5]) : 0
    const issueToMaturity =
      settlement == i32.MIN_VALUE || maturity == i32.MIN_VALUE || issue == i32.MIN_VALUE
        ? NaN
        : maturityIssueYearfracValue(issue, settlement, maturity, basis)
    const settlementToMaturity =
      settlement == i32.MIN_VALUE || maturity == i32.MIN_VALUE ? NaN : securityAnnualizedYearfracValue(settlement, maturity, basis)
    const issueToSettlement =
      settlement == i32.MIN_VALUE || issue == i32.MIN_VALUE || maturity == i32.MIN_VALUE
        ? NaN
        : accruedIssueYearfracValue(issue, settlement, maturity, basis)
    const denominator = isNaN(settlementToMaturity) ? NaN : 1.0 + yieldRate * settlementToMaturity
    const value =
      isNaN(rate) ||
      isNaN(yieldRate) ||
      rate < 0.0 ||
      yieldRate < 0.0 ||
      isNaN(issueToMaturity) ||
      isNaN(issueToSettlement) ||
      !isFinite(denominator) ||
      denominator <= 0.0
        ? NaN
        : (100.0 * (1.0 + rate * issueToMaturity)) / denominator - 100.0 * rate * issueToSettlement
    return isNaN(value)
      ? writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, value, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Yieldmat && (argc == 5 || argc == 6)) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const settlement = excelSerialWhole(tagStack[base], valueStack[base])
    const maturity = excelSerialWhole(tagStack[base + 1], valueStack[base + 1])
    const issue = excelSerialWhole(tagStack[base + 2], valueStack[base + 2])
    const rate = toNumberExact(tagStack[base + 3], valueStack[base + 3])
    const price = toNumberExact(tagStack[base + 4], valueStack[base + 4])
    const basis = argc == 6 ? truncToInt(tagStack[base + 5], valueStack[base + 5]) : 0
    const issueToMaturity =
      settlement == i32.MIN_VALUE || maturity == i32.MIN_VALUE || issue == i32.MIN_VALUE
        ? NaN
        : maturityIssueYearfracValue(issue, settlement, maturity, basis)
    const settlementToMaturity =
      settlement == i32.MIN_VALUE || maturity == i32.MIN_VALUE ? NaN : securityAnnualizedYearfracValue(settlement, maturity, basis)
    const issueToSettlement =
      settlement == i32.MIN_VALUE || issue == i32.MIN_VALUE || maturity == i32.MIN_VALUE
        ? NaN
        : accruedIssueYearfracValue(issue, settlement, maturity, basis)
    const settlementValue = isNaN(price) || isNaN(rate) || isNaN(issueToSettlement) ? NaN : price + 100.0 * rate * issueToSettlement
    const value =
      isNaN(rate) ||
      isNaN(price) ||
      rate < 0.0 ||
      price <= 0.0 ||
      isNaN(issueToMaturity) ||
      isNaN(settlementToMaturity) ||
      !isFinite(settlementValue) ||
      settlementValue <= 0.0
        ? NaN
        : ((100.0 * (1.0 + rate * issueToMaturity)) / settlementValue - 1.0) / settlementToMaturity
    return isNaN(value)
      ? writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, value, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Oddlprice && (argc == 7 || argc == 8)) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const settlement = excelSerialWhole(tagStack[base], valueStack[base])
    const maturity = excelSerialWhole(tagStack[base + 1], valueStack[base + 1])
    const lastInterest = excelSerialWhole(tagStack[base + 2], valueStack[base + 2])
    const rate = toNumberExact(tagStack[base + 3], valueStack[base + 3])
    const yieldRate = toNumberExact(tagStack[base + 4], valueStack[base + 4])
    const redemption = toNumberExact(tagStack[base + 5], valueStack[base + 5])
    const frequency = truncToInt(tagStack[base + 6], valueStack[base + 6])
    const basis = argc == 8 ? truncToInt(tagStack[base + 7], valueStack[base + 7]) : 0
    const value =
      settlement == i32.MIN_VALUE || maturity == i32.MIN_VALUE || lastInterest == i32.MIN_VALUE
        ? NaN
        : oddLastPriceValue(settlement, maturity, lastInterest, rate, yieldRate, redemption, frequency, basis)
    return isNaN(value)
      ? writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, value, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Oddlyield && (argc == 7 || argc == 8)) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const settlement = excelSerialWhole(tagStack[base], valueStack[base])
    const maturity = excelSerialWhole(tagStack[base + 1], valueStack[base + 1])
    const lastInterest = excelSerialWhole(tagStack[base + 2], valueStack[base + 2])
    const rate = toNumberExact(tagStack[base + 3], valueStack[base + 3])
    const price = toNumberExact(tagStack[base + 4], valueStack[base + 4])
    const redemption = toNumberExact(tagStack[base + 5], valueStack[base + 5])
    const frequency = truncToInt(tagStack[base + 6], valueStack[base + 6])
    const basis = argc == 8 ? truncToInt(tagStack[base + 7], valueStack[base + 7]) : 0
    const value =
      settlement == i32.MIN_VALUE || maturity == i32.MIN_VALUE || lastInterest == i32.MIN_VALUE
        ? NaN
        : oddLastYieldValue(settlement, maturity, lastInterest, rate, price, redemption, frequency, basis)
    return isNaN(value)
      ? writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, value, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Oddfprice && (argc == 8 || argc == 9)) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const settlement = excelSerialWhole(tagStack[base], valueStack[base])
    const maturity = excelSerialWhole(tagStack[base + 1], valueStack[base + 1])
    const issue = excelSerialWhole(tagStack[base + 2], valueStack[base + 2])
    const firstCoupon = excelSerialWhole(tagStack[base + 3], valueStack[base + 3])
    const rate = toNumberExact(tagStack[base + 4], valueStack[base + 4])
    const yieldRate = toNumberExact(tagStack[base + 5], valueStack[base + 5])
    const redemption = toNumberExact(tagStack[base + 6], valueStack[base + 6])
    const frequency = truncToInt(tagStack[base + 7], valueStack[base + 7])
    const basis = argc == 9 ? truncToInt(tagStack[base + 8], valueStack[base + 8]) : 0
    const value =
      settlement == i32.MIN_VALUE ||
      maturity == i32.MIN_VALUE ||
      issue == i32.MIN_VALUE ||
      firstCoupon == i32.MIN_VALUE ||
      rate < 0.0 ||
      yieldRate < 0.0 ||
      redemption <= 0.0
        ? NaN
        : oddFirstPriceValue(settlement, maturity, issue, firstCoupon, rate, yieldRate, redemption, frequency, basis)
    return isNaN(value)
      ? writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, value, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Oddfyield && (argc == 8 || argc == 9)) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const settlement = excelSerialWhole(tagStack[base], valueStack[base])
    const maturity = excelSerialWhole(tagStack[base + 1], valueStack[base + 1])
    const issue = excelSerialWhole(tagStack[base + 2], valueStack[base + 2])
    const firstCoupon = excelSerialWhole(tagStack[base + 3], valueStack[base + 3])
    const rate = toNumberExact(tagStack[base + 4], valueStack[base + 4])
    const price = toNumberExact(tagStack[base + 5], valueStack[base + 5])
    const redemption = toNumberExact(tagStack[base + 6], valueStack[base + 6])
    const frequency = truncToInt(tagStack[base + 7], valueStack[base + 7])
    const basis = argc == 9 ? truncToInt(tagStack[base + 8], valueStack[base + 8]) : 0
    const value =
      settlement == i32.MIN_VALUE || maturity == i32.MIN_VALUE || issue == i32.MIN_VALUE || firstCoupon == i32.MIN_VALUE
        ? NaN
        : oddFirstYieldValue(settlement, maturity, issue, firstCoupon, rate, price, redemption, frequency, basis)
    return isNaN(value)
      ? writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, value, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Price && (argc == 6 || argc == 7)) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const settlement = excelSerialWhole(tagStack[base], valueStack[base])
    const maturity = excelSerialWhole(tagStack[base + 1], valueStack[base + 1])
    const rate = toNumberExact(tagStack[base + 2], valueStack[base + 2])
    const yieldRate = toNumberExact(tagStack[base + 3], valueStack[base + 3])
    const redemption = toNumberExact(tagStack[base + 4], valueStack[base + 4])
    const frequency = truncToInt(tagStack[base + 5], valueStack[base + 5])
    const basis = argc == 7 ? truncToInt(tagStack[base + 6], valueStack[base + 6]) : 0
    const periodsRemaining =
      settlement == i32.MIN_VALUE || maturity == i32.MIN_VALUE
        ? i32.MIN_VALUE
        : couponPeriodsRemainingValue(settlement, maturity, frequency)
    const previousCoupon =
      periodsRemaining == i32.MIN_VALUE ? i32.MIN_VALUE : couponDateFromMaturityValue(maturity, periodsRemaining, frequency)
    const nextCoupon =
      periodsRemaining == i32.MIN_VALUE ? i32.MIN_VALUE : couponDateFromMaturityValue(maturity, periodsRemaining - 1, frequency)
    const accruedDays = previousCoupon == i32.MIN_VALUE ? NaN : couponDaysByBasisValue(previousCoupon, settlement, basis)
    const daysToNextCoupon = nextCoupon == i32.MIN_VALUE ? NaN : couponDaysByBasisValue(settlement, nextCoupon, basis)
    const daysInPeriod =
      previousCoupon == i32.MIN_VALUE || nextCoupon == i32.MIN_VALUE
        ? NaN
        : couponPeriodDaysValue(previousCoupon, nextCoupon, basis, frequency)
    const value =
      isNaN(rate) ||
      isNaN(yieldRate) ||
      isNaN(redemption) ||
      rate < 0.0 ||
      yieldRate < 0.0 ||
      redemption <= 0.0 ||
      periodsRemaining == i32.MIN_VALUE
        ? NaN
        : couponPriceFromMetricsValue(periodsRemaining, accruedDays, daysToNextCoupon, daysInPeriod, rate, yieldRate, redemption, frequency)
    return isNaN(value)
      ? writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, value, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Yield && (argc == 6 || argc == 7)) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const settlement = excelSerialWhole(tagStack[base], valueStack[base])
    const maturity = excelSerialWhole(tagStack[base + 1], valueStack[base + 1])
    const rate = toNumberExact(tagStack[base + 2], valueStack[base + 2])
    const price = toNumberExact(tagStack[base + 3], valueStack[base + 3])
    const redemption = toNumberExact(tagStack[base + 4], valueStack[base + 4])
    const frequency = truncToInt(tagStack[base + 5], valueStack[base + 5])
    const basis = argc == 7 ? truncToInt(tagStack[base + 6], valueStack[base + 6]) : 0
    const periodsRemaining =
      settlement == i32.MIN_VALUE || maturity == i32.MIN_VALUE
        ? i32.MIN_VALUE
        : couponPeriodsRemainingValue(settlement, maturity, frequency)
    const previousCoupon =
      periodsRemaining == i32.MIN_VALUE ? i32.MIN_VALUE : couponDateFromMaturityValue(maturity, periodsRemaining, frequency)
    const nextCoupon =
      periodsRemaining == i32.MIN_VALUE ? i32.MIN_VALUE : couponDateFromMaturityValue(maturity, periodsRemaining - 1, frequency)
    const accruedDays = previousCoupon == i32.MIN_VALUE ? NaN : couponDaysByBasisValue(previousCoupon, settlement, basis)
    const daysToNextCoupon = nextCoupon == i32.MIN_VALUE ? NaN : couponDaysByBasisValue(settlement, nextCoupon, basis)
    const daysInPeriod =
      previousCoupon == i32.MIN_VALUE || nextCoupon == i32.MIN_VALUE
        ? NaN
        : couponPeriodDaysValue(previousCoupon, nextCoupon, basis, frequency)
    const value =
      isNaN(rate) ||
      isNaN(price) ||
      isNaN(redemption) ||
      rate < 0.0 ||
      price <= 0.0 ||
      redemption <= 0.0 ||
      periodsRemaining == i32.MIN_VALUE
        ? NaN
        : solveCouponYieldValue(periodsRemaining, accruedDays, daysToNextCoupon, daysInPeriod, rate, price, redemption, frequency)
    return isNaN(value)
      ? writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, value, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if ((builtinId == BuiltinId.Duration || builtinId == BuiltinId.Mduration) && (argc == 5 || argc == 6)) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const settlement = excelSerialWhole(tagStack[base], valueStack[base])
    const maturity = excelSerialWhole(tagStack[base + 1], valueStack[base + 1])
    const couponRate = toNumberExact(tagStack[base + 2], valueStack[base + 2])
    const yieldRate = toNumberExact(tagStack[base + 3], valueStack[base + 3])
    const frequency = truncToInt(tagStack[base + 4], valueStack[base + 4])
    const basis = argc == 6 ? truncToInt(tagStack[base + 5], valueStack[base + 5]) : 0
    const periodsRemaining =
      settlement == i32.MIN_VALUE || maturity == i32.MIN_VALUE
        ? i32.MIN_VALUE
        : couponPeriodsRemainingValue(settlement, maturity, frequency)
    const previousCoupon =
      periodsRemaining == i32.MIN_VALUE ? i32.MIN_VALUE : couponDateFromMaturityValue(maturity, periodsRemaining, frequency)
    const nextCoupon =
      periodsRemaining == i32.MIN_VALUE ? i32.MIN_VALUE : couponDateFromMaturityValue(maturity, periodsRemaining - 1, frequency)
    const accruedDays = previousCoupon == i32.MIN_VALUE ? NaN : couponDaysByBasisValue(previousCoupon, settlement, basis)
    const daysToNextCoupon = nextCoupon == i32.MIN_VALUE ? NaN : couponDaysByBasisValue(settlement, nextCoupon, basis)
    const daysInPeriod =
      previousCoupon == i32.MIN_VALUE || nextCoupon == i32.MIN_VALUE
        ? NaN
        : couponPeriodDaysValue(previousCoupon, nextCoupon, basis, frequency)
    const duration =
      isNaN(couponRate) || isNaN(yieldRate) || couponRate < 0.0 || yieldRate < 0.0 || periodsRemaining == i32.MIN_VALUE
        ? NaN
        : macaulayDurationValue(periodsRemaining, accruedDays, daysToNextCoupon, daysInPeriod, couponRate, yieldRate, frequency)
    const value = builtinId == BuiltinId.Mduration && !isNaN(duration) ? duration / (1.0 + yieldRate / <f64>frequency) : duration
    return isNaN(value)
      ? writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, value, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  return -1
}
