import { BUILTINS, BuiltinId, ErrorCode, ValueTag } from "@bilig/protocol";
import type { CellValue } from "@bilig/protocol";
import { builtinJsSpecialNames } from "./builtin-capabilities.js";
import { createComplexBuiltins } from "./builtins/complex.js";
import { createRadixBuiltins } from "./builtins/radix.js";
import {
  collectAStyleNumericArgs,
  erfApprox,
  modeSingle,
  populationStandardDeviation,
  populationVariance,
  sampleStandardDeviation,
  sampleVariance,
} from "./builtins/statistics.js";
import {
  addMonthsToExcelDate,
  datetimeBuiltins,
  excelSerialToDateParts,
} from "./builtins/datetime.js";
import { convertBuiltin, euroconvertBuiltin } from "./builtins/convert.js";
import { logicalBuiltins } from "./builtins/logical.js";
import { lookupBuiltins } from "./builtins/lookup.js";
import { createBlockedBuiltinMap, scalarPlaceholderBuiltinNames } from "./builtins/placeholder.js";
import { getExternalScalarFunction, hasExternalFunction } from "./external-function-adapter.js";
import type { ArrayValue, EvaluationResult } from "./runtime-values.js";
import { textBuiltins } from "./builtins/text.js";

type Builtin = (...args: CellValue[]) => EvaluationResult;

function toNumber(value: CellValue): number | undefined {
  switch (value.tag) {
    case ValueTag.Number:
      return value.value;
    case ValueTag.Boolean:
      return value.value ? 1 : 0;
    case ValueTag.Empty:
      return 0;
    case ValueTag.String:
    case ValueTag.Error:
      return undefined;
    default:
      return undefined;
  }
}

function toBitwiseUnsigned(value: CellValue | undefined): number | undefined {
  if (value === undefined) {
    return undefined;
  }
  const numeric = toNumber(value);
  if (numeric === undefined || !Number.isFinite(numeric)) {
    return undefined;
  }
  const truncated = Math.trunc(numeric);
  if (!Number.isSafeInteger(truncated)) {
    return undefined;
  }
  return truncated >>> 0;
}

function coerceShiftAmount(value: CellValue | undefined): number | undefined {
  if (value === undefined) {
    return undefined;
  }
  const numeric = toNumber(value);
  if (numeric === undefined || !Number.isFinite(numeric)) {
    return undefined;
  }
  const truncated = Math.trunc(numeric);
  return truncated >= 0 ? truncated : undefined;
}

function numberResult(value: number): CellValue {
  return { tag: ValueTag.Number, value };
}

function valueError(): CellValue {
  return { tag: ValueTag.Error, code: ErrorCode.Value };
}

function blockedError(): CellValue {
  return { tag: ValueTag.Error, code: ErrorCode.Blocked };
}

function numError(): CellValue {
  return { tag: ValueTag.Error, code: ErrorCode.Value };
}

function numericResultOrError(value: number): CellValue {
  return Number.isFinite(value) ? numberResult(value) : valueError();
}

function firstError(args: CellValue[]): CellValue | undefined {
  return args.find((arg) => arg.tag === ValueTag.Error);
}

function coercePositiveInteger(value: CellValue | undefined, fallback: number): number | undefined {
  if (value === undefined) {
    return fallback;
  }
  const numeric = toNumber(value);
  if (numeric === undefined || !Number.isFinite(numeric)) {
    return undefined;
  }
  const truncated = Math.trunc(numeric);
  return truncated >= 1 ? truncated : undefined;
}

function coerceNumber(value: CellValue | undefined, fallback: number): number | undefined {
  if (value === undefined) {
    return fallback;
  }
  const numeric = toNumber(value);
  return numeric !== undefined && Number.isFinite(numeric) ? numeric : undefined;
}

function integerValue(value: CellValue | undefined, fallback?: number): number | undefined {
  if (value === undefined) {
    return fallback;
  }
  const numeric = toNumber(value);
  if (numeric === undefined || !Number.isFinite(numeric)) {
    return undefined;
  }
  return Math.trunc(numeric);
}

function toInteger(value: CellValue, fallback?: number): number | undefined {
  if (value === undefined) {
    return fallback;
  }
  const numeric = toNumber(value);
  if (numeric === undefined || !Number.isFinite(numeric)) {
    return undefined;
  }
  const truncated = Math.trunc(numeric);
  return Number.isSafeInteger(truncated) ? truncated : undefined;
}

function nonNegativeIntegerValue(
  value: CellValue | undefined,
  fallback?: number,
): number | undefined {
  const integer = integerValue(value, fallback);
  return integer !== undefined && integer >= 0 ? integer : undefined;
}

function positiveIntegerValue(value: CellValue | undefined, fallback?: number): number | undefined {
  const integer = integerValue(value, fallback);
  return integer !== undefined && integer >= 1 ? integer : undefined;
}

function sequenceResult(
  rowsArg: CellValue | undefined,
  colsArg: CellValue | undefined,
  startArg: CellValue | undefined,
  stepArg: CellValue | undefined,
): ArrayValue | CellValue {
  const rows = coercePositiveInteger(rowsArg, 1);
  const cols = coercePositiveInteger(colsArg, 1);
  const start = coerceNumber(startArg, 1);
  const step = coerceNumber(stepArg, 1);
  if (rows === undefined || cols === undefined || start === undefined || step === undefined) {
    return valueError();
  }

  const values: CellValue[] = [];
  for (let index = 0; index < rows * cols; index += 1) {
    values.push(numberResult(start + index * step));
  }
  return {
    kind: "array",
    rows,
    cols,
    values,
  };
}

function coerceDateSerial(value: CellValue | undefined): number | undefined {
  if (value === undefined) {
    return undefined;
  }
  const serial = toNumber(value);
  return serial !== undefined && Number.isFinite(serial) ? Math.trunc(serial) : undefined;
}

function coerceBoolean(value: CellValue | undefined, fallback: boolean): boolean | undefined {
  if (value === undefined) {
    return fallback;
  }
  if (value.tag === ValueTag.Boolean) {
    return value.value;
  }
  const numeric = toNumber(value);
  return numeric === undefined ? undefined : numeric !== 0;
}

function isLeapYear(year: number): boolean {
  return year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0);
}

function isValidBasis(basis: number): boolean {
  return basis === 0 || basis === 1 || basis === 2 || basis === 3 || basis === 4;
}

function isValidFrequency(frequency: number): boolean {
  return frequency === 1 || frequency === 2 || frequency === 4;
}

function yearsDaysInYear(year: number): number {
  return isLeapYear(year) ? 366 : 365;
}

function yearFracByBasis(
  startSerial: number,
  endSerial: number,
  basis: number,
): number | undefined {
  if (!isValidBasis(basis)) {
    return undefined;
  }

  let start = startSerial;
  let end = endSerial;
  if (start > end) {
    [start, end] = [end, start];
  }

  const startParts = excelSerialToDateParts(start);
  const endParts = excelSerialToDateParts(end);
  if (startParts === undefined || endParts === undefined) {
    return undefined;
  }

  let startDay = startParts.day;
  let startMonth = startParts.month;
  let startYear = startParts.year;
  let endDay = endParts.day;
  let endMonth = endParts.month;
  let endYear = endParts.year;

  let totalDays: number;
  switch (basis) {
    case 0:
      if (startDay === 31) {
        startDay -= 1;
      }
      if (startDay === 30 && endDay === 31) {
        endDay -= 1;
      } else if (startMonth === 2 && startDay === (isLeapYear(startYear) ? 29 : 28)) {
        startDay = 30;
        if (endMonth === 2 && endDay === (isLeapYear(endYear) ? 29 : 28)) {
          endDay = 30;
        }
      }
      totalDays = (endYear - startYear) * 360 + (endMonth - startMonth) * 30 + (endDay - startDay);
      break;
    case 1:
    case 2:
    case 3:
      totalDays = end - start;
      break;
    case 4:
      if (startDay === 31) {
        startDay -= 1;
      }
      if (endDay === 31) {
        endDay -= 1;
      }
      totalDays = (endYear - startYear) * 360 + (endMonth - startMonth) * 30 + (endDay - startDay);
      break;
    default:
      return undefined;
  }

  let daysInYear: number;
  switch (basis) {
    case 1: {
      const isYearDifferent = startYear !== endYear;
      if (
        isYearDifferent &&
        (endYear !== startYear + 1 ||
          endMonth < startMonth ||
          (endMonth === startMonth && endDay > startDay))
      ) {
        let dayCount = 0;
        for (let year = startYear; year <= endYear; year += 1) {
          dayCount += yearsDaysInYear(year);
        }
        daysInYear = dayCount / (endYear - startYear + 1);
      } else if (isYearDifferent) {
        const crossesLeap =
          (isLeapYear(startYear) && (startMonth < 2 || (startMonth === 2 && startDay <= 29))) ||
          (isLeapYear(endYear) && (endMonth > 2 || (endMonth === 2 && endDay === 29)));
        daysInYear = crossesLeap ? 366 : 365;
      } else {
        daysInYear = yearsDaysInYear(startYear);
      }
      break;
    }
    case 3:
      daysInYear = 365;
      break;
    case 0:
    case 2:
    case 4:
      daysInYear = 360;
      break;
    default:
      return undefined;
  }

  return totalDays / daysInYear;
}

function securityAnnualizedYearFraction(
  settlement: number,
  maturity: number,
  basis: number,
): number | undefined {
  if (settlement >= maturity || !isValidBasis(basis)) {
    return undefined;
  }
  const years = yearFracByBasis(settlement, maturity, basis);
  return years !== undefined && years > 0 ? years : undefined;
}

function treasuryBillDays(settlement: number, maturity: number): number | undefined {
  if (settlement >= maturity) {
    return undefined;
  }
  const days = maturity - settlement;
  return days > 0 && days <= 365 ? days : undefined;
}

function maturityAtIssueFractions(
  settlement: number,
  maturity: number,
  issue: number,
  basis: number,
):
  | {
      issueToMaturity: number;
      settlementToMaturity: number;
      issueToSettlement: number;
    }
  | undefined {
  if (issue >= settlement || issue >= maturity || settlement >= maturity || !isValidBasis(basis)) {
    return undefined;
  }
  const issueToMaturity = yearFracByBasis(issue, maturity, basis);
  const settlementToMaturity = yearFracByBasis(settlement, maturity, basis);
  const issueToSettlement = yearFracByBasis(issue, settlement, basis);
  if (
    issueToMaturity === undefined ||
    settlementToMaturity === undefined ||
    issueToSettlement === undefined ||
    issueToMaturity <= 0 ||
    settlementToMaturity <= 0
  ) {
    return undefined;
  }
  return {
    issueToMaturity,
    settlementToMaturity,
    issueToSettlement,
  };
}

function oddLastCouponFractions(
  settlement: number,
  maturity: number,
  lastInterest: number,
  frequency: number,
  basis: number,
):
  | {
      accruedFraction: number;
      remainingFraction: number;
      totalFraction: number;
    }
  | undefined {
  if (
    lastInterest >= settlement ||
    settlement >= maturity ||
    !isValidFrequency(frequency) ||
    !isValidBasis(basis)
  ) {
    return undefined;
  }

  const stepMonths = 12 / frequency;
  let periodStart = lastInterest;
  let accruedFraction = 0;
  let remainingFraction = 0;
  let totalFraction = 0;
  let iterations = 0;

  while (periodStart < maturity && iterations < 32) {
    const normalEnd = addMonthsToExcelDate(periodStart, stepMonths);
    if (normalEnd === undefined || normalEnd <= periodStart) {
      return undefined;
    }
    const actualEnd = Math.min(normalEnd, maturity);
    const normalDays = couponDaysByBasis(periodStart, normalEnd, basis);
    const countedDays = couponDaysByBasis(periodStart, actualEnd, basis);
    if (
      normalDays === undefined ||
      countedDays === undefined ||
      normalDays <= 0 ||
      countedDays < 0
    ) {
      return undefined;
    }

    totalFraction += countedDays / normalDays;

    if (settlement > periodStart) {
      const accruedEnd = Math.min(settlement, actualEnd);
      const accruedDays = couponDaysByBasis(periodStart, accruedEnd, basis);
      if (accruedDays === undefined || accruedDays < 0) {
        return undefined;
      }
      accruedFraction += accruedDays / normalDays;
    }

    if (settlement < actualEnd) {
      const remainingStart = Math.max(settlement, periodStart);
      const remainingDays = couponDaysByBasis(remainingStart, actualEnd, basis);
      if (remainingDays === undefined || remainingDays < 0) {
        return undefined;
      }
      remainingFraction += remainingDays / normalDays;
    }

    periodStart = actualEnd;
    iterations += 1;
  }

  if (
    periodStart !== maturity ||
    iterations >= 32 ||
    remainingFraction <= 0 ||
    totalFraction <= 0
  ) {
    return undefined;
  }

  return {
    accruedFraction,
    remainingFraction,
    totalFraction,
  };
}

function oddFirstCouponMetrics(
  settlement: number,
  maturity: number,
  issue: number,
  firstCoupon: number,
  frequency: number,
  basis: number,
):
  | {
      accruedFraction: number;
      remainingFraction: number;
      totalFraction: number;
      regularPeriodsAfterFirst: number;
    }
  | undefined {
  if (
    issue >= settlement ||
    settlement >= firstCoupon ||
    firstCoupon >= maturity ||
    !isValidFrequency(frequency) ||
    !isValidBasis(basis)
  ) {
    return undefined;
  }

  const stepMonths = 12 / frequency;
  const segments: Array<{
    actualStart: number;
    periodEnd: number;
    normalDays: number;
    countedDays: number;
  }> = [];
  let periodEnd = firstCoupon;
  let iterations = 0;
  while (periodEnd > issue && iterations < 64) {
    const normalStart = addMonthsToExcelDate(periodEnd, -stepMonths);
    if (normalStart === undefined || normalStart >= periodEnd) {
      return undefined;
    }
    const actualStart = Math.max(normalStart, issue);
    const normalDays = couponDaysByBasis(normalStart, periodEnd, basis);
    const countedDays = couponDaysByBasis(actualStart, periodEnd, basis);
    if (
      normalDays === undefined ||
      countedDays === undefined ||
      normalDays <= 0 ||
      countedDays < 0
    ) {
      return undefined;
    }
    segments.unshift({ actualStart, periodEnd, normalDays, countedDays });
    periodEnd = actualStart;
    iterations += 1;
  }
  if (periodEnd !== issue || iterations >= 64 || segments.length === 0) {
    return undefined;
  }

  let accruedFraction = 0;
  let remainingFraction = 0;
  let totalFraction = 0;
  for (const segment of segments) {
    totalFraction += segment.countedDays / segment.normalDays;

    if (settlement > segment.actualStart) {
      const accruedEnd = Math.min(settlement, segment.periodEnd);
      const accruedDays = couponDaysByBasis(segment.actualStart, accruedEnd, basis);
      if (accruedDays === undefined || accruedDays < 0) {
        return undefined;
      }
      accruedFraction += accruedDays / segment.normalDays;
    }

    if (settlement < segment.periodEnd) {
      const remainingStart = Math.max(settlement, segment.actualStart);
      const remainingDays = couponDaysByBasis(remainingStart, segment.periodEnd, basis);
      if (remainingDays === undefined || remainingDays < 0) {
        return undefined;
      }
      remainingFraction += remainingDays / segment.normalDays;
    }
  }

  let regularPeriodsAfterFirst = 0;
  let couponDate = firstCoupon;
  while (couponDate < maturity && regularPeriodsAfterFirst < 256) {
    const nextCouponDate = addMonthsToExcelDate(couponDate, stepMonths);
    if (nextCouponDate === undefined || nextCouponDate <= couponDate) {
      return undefined;
    }
    couponDate = nextCouponDate;
    regularPeriodsAfterFirst += 1;
  }
  if (
    couponDate !== maturity ||
    regularPeriodsAfterFirst <= 0 ||
    regularPeriodsAfterFirst >= 256 ||
    remainingFraction <= 0 ||
    totalFraction <= 0
  ) {
    return undefined;
  }

  return {
    accruedFraction,
    remainingFraction,
    totalFraction,
    regularPeriodsAfterFirst,
  };
}

function oddFirstPriceValue(
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
  if (
    !Number.isFinite(rate) ||
    !Number.isFinite(yieldRate) ||
    !Number.isFinite(redemption) ||
    rate < 0 ||
    redemption <= 0
  ) {
    return undefined;
  }
  const metrics = oddFirstCouponMetrics(settlement, maturity, issue, firstCoupon, frequency, basis);
  if (metrics === undefined) {
    return undefined;
  }

  const discountBase = 1 + yieldRate / frequency;
  if (discountBase <= 0) {
    return undefined;
  }
  const coupon = (100 * rate) / frequency;
  let price =
    (coupon * metrics.totalFraction) / Math.pow(discountBase, metrics.remainingFraction) -
    coupon * metrics.accruedFraction;

  for (let period = 1; period <= metrics.regularPeriodsAfterFirst; period += 1) {
    const exponent = metrics.remainingFraction + period;
    const cashflow = period === metrics.regularPeriodsAfterFirst ? redemption + coupon : coupon;
    price += cashflow / Math.pow(discountBase, exponent);
  }
  return price;
}

function solveOddFirstYield(
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
  if (
    !Number.isFinite(rate) ||
    !Number.isFinite(price) ||
    !Number.isFinite(redemption) ||
    rate < 0 ||
    price <= 0 ||
    redemption <= 0
  ) {
    return undefined;
  }

  const metrics = oddFirstCouponMetrics(settlement, maturity, issue, firstCoupon, frequency, basis);
  if (metrics === undefined) {
    return undefined;
  }

  const priceAtYield = (yieldRate: number): number | undefined =>
    oddFirstPriceValue(
      settlement,
      maturity,
      issue,
      firstCoupon,
      rate,
      yieldRate,
      redemption,
      frequency,
      basis,
    );

  let lower = -frequency + 1e-10;
  let upper = Math.max(1, rate * 2 + 0.1);
  let lowerPrice = priceAtYield(lower);
  let upperPrice = priceAtYield(upper);
  for (
    let iteration = 0;
    iteration < 200 &&
    (lowerPrice === undefined ||
      upperPrice === undefined ||
      lowerPrice < price ||
      upperPrice > price);
    iteration += 1
  ) {
    if (upperPrice === undefined || upperPrice > price) {
      upper = upper * 2 + 1;
      upperPrice = priceAtYield(upper);
      continue;
    }
    lower = (lower - frequency) / 2;
    lowerPrice = priceAtYield(lower);
  }
  if (
    lowerPrice === undefined ||
    upperPrice === undefined ||
    lowerPrice < price ||
    upperPrice > price
  ) {
    return undefined;
  }

  let guess = Math.min(Math.max(rate, lower + 1e-8), upper - 1e-8);
  for (let iteration = 0; iteration < 200; iteration += 1) {
    const estimatedPrice = priceAtYield(guess);
    if (estimatedPrice === undefined) {
      return undefined;
    }
    const error = estimatedPrice - price;
    if (Math.abs(error) < 1e-14) {
      return guess;
    }

    const epsilon = Math.max(1e-7, Math.abs(guess) * 1e-6);
    const shiftedPrice = priceAtYield(guess + epsilon);
    const derivative =
      shiftedPrice === undefined ? undefined : (shiftedPrice - estimatedPrice) / epsilon;
    let nextGuess =
      derivative === undefined || !Number.isFinite(derivative) || derivative === 0
        ? (lower + upper) / 2
        : guess - error / derivative;
    if (!Number.isFinite(nextGuess) || nextGuess <= lower || nextGuess >= upper) {
      nextGuess = (lower + upper) / 2;
    }

    const boundedPrice = priceAtYield(nextGuess);
    if (boundedPrice === undefined) {
      return undefined;
    }
    if (boundedPrice > price) {
      lower = nextGuess;
    } else {
      upper = nextGuess;
    }
    guess = nextGuess;
    if (Math.abs(upper - lower) < 1e-14) {
      return (lower + upper) / 2;
    }
  }

  return (lower + upper) / 2;
}

function days360Us(startSerial: number, endSerial: number): number | undefined {
  const startParts = excelSerialToDateParts(startSerial);
  const endParts = excelSerialToDateParts(endSerial);
  if (startParts === undefined || endParts === undefined) {
    return undefined;
  }

  let startDay = startParts.day;
  let endDay = endParts.day;
  if (startDay === 31) {
    startDay = 30;
  }
  if (startDay === 30 && endDay === 31) {
    endDay = 30;
  } else if (startParts.month === 2 && startDay === (isLeapYear(startParts.year) ? 29 : 28)) {
    startDay = 30;
    if (endParts.month === 2 && endDay === (isLeapYear(endParts.year) ? 29 : 28)) {
      endDay = 30;
    }
  }

  return (
    (endParts.year - startParts.year) * 360 +
    (endParts.month - startParts.month) * 30 +
    (endDay - startDay)
  );
}

function days360European(startSerial: number, endSerial: number): number | undefined {
  const startParts = excelSerialToDateParts(startSerial);
  const endParts = excelSerialToDateParts(endSerial);
  if (startParts === undefined || endParts === undefined) {
    return undefined;
  }

  const startDay = startParts.day === 31 ? 30 : startParts.day;
  const endDay = endParts.day === 31 ? 30 : endParts.day;
  return (
    (endParts.year - startParts.year) * 360 +
    (endParts.month - startParts.month) * 30 +
    (endDay - startDay)
  );
}

function couponDaysByBasis(
  startSerial: number,
  endSerial: number,
  basis: number,
): number | undefined {
  if (!isValidBasis(basis) || startSerial > endSerial) {
    return undefined;
  }
  switch (basis) {
    case 0:
      return days360Us(startSerial, endSerial);
    case 4:
      return days360European(startSerial, endSerial);
    default:
      return endSerial - startSerial;
  }
}

interface CouponSchedule {
  previousCoupon: number;
  nextCoupon: number;
  periodsRemaining: number;
}

function couponSchedule(
  settlement: number,
  maturity: number,
  frequency: number,
): CouponSchedule | undefined {
  if (settlement >= maturity || !isValidFrequency(frequency)) {
    return undefined;
  }

  const stepMonths = 12 / frequency;
  let periodsRemaining = 1;
  let previousCoupon = addMonthsToExcelDate(maturity, -stepMonths);
  while (previousCoupon !== undefined && previousCoupon > settlement) {
    periodsRemaining += 1;
    previousCoupon = addMonthsToExcelDate(maturity, -periodsRemaining * stepMonths);
  }

  if (previousCoupon === undefined) {
    return undefined;
  }
  const nextCoupon = addMonthsToExcelDate(maturity, -(periodsRemaining - 1) * stepMonths);
  if (nextCoupon === undefined || nextCoupon <= settlement) {
    return undefined;
  }

  return {
    previousCoupon,
    nextCoupon,
    periodsRemaining,
  };
}

interface CouponMetrics extends CouponSchedule {
  accruedDays: number;
  daysToNextCoupon: number;
  daysInPeriod: number;
}

function couponMetrics(
  settlement: number,
  maturity: number,
  frequency: number,
  basis: number,
): CouponMetrics | undefined {
  const schedule = couponSchedule(settlement, maturity, frequency);
  if (schedule === undefined || !isValidBasis(basis)) {
    return undefined;
  }

  const accruedDays = couponDaysByBasis(schedule.previousCoupon, settlement, basis);
  const daysToNextCoupon = couponDaysByBasis(settlement, schedule.nextCoupon, basis);
  const daysInPeriod =
    basis === 1
      ? couponDaysByBasis(schedule.previousCoupon, schedule.nextCoupon, basis)
      : basis === 3
        ? 365 / frequency
        : 360 / frequency;
  if (
    accruedDays === undefined ||
    daysToNextCoupon === undefined ||
    daysInPeriod === undefined ||
    daysInPeriod <= 0
  ) {
    return undefined;
  }

  return {
    ...schedule,
    accruedDays,
    daysToNextCoupon,
    daysInPeriod,
  };
}

function pricePeriodicSecurity(
  metrics: CouponMetrics,
  rate: number,
  yieldRate: number,
  redemption: number,
  frequency: number,
): number | undefined {
  const coupon = (100 * rate) / frequency;
  const periodsToNextCoupon = metrics.daysToNextCoupon / metrics.daysInPeriod;
  if (metrics.periodsRemaining === 1) {
    const denominator = 1 + (yieldRate / frequency) * periodsToNextCoupon;
    return denominator <= 0
      ? undefined
      : (redemption + coupon) / denominator - coupon * (metrics.accruedDays / metrics.daysInPeriod);
  }

  const discountBase = 1 + yieldRate / frequency;
  if (discountBase <= 0) {
    return undefined;
  }

  let price = 0;
  for (let period = 1; period <= metrics.periodsRemaining; period += 1) {
    const periodsToCashflow = period - 1 + periodsToNextCoupon;
    price += coupon / Math.pow(discountBase, periodsToCashflow);
  }
  price += redemption / Math.pow(discountBase, metrics.periodsRemaining - 1 + periodsToNextCoupon);
  return price - coupon * (metrics.accruedDays / metrics.daysInPeriod);
}

function solvePeriodicSecurityYield(
  metrics: CouponMetrics,
  rate: number,
  price: number,
  redemption: number,
  frequency: number,
): number | undefined {
  const coupon = (100 * rate) / frequency;
  if (metrics.periodsRemaining === 1) {
    const dirtyPrice = price + coupon * (metrics.accruedDays / metrics.daysInPeriod);
    if (dirtyPrice <= 0 || metrics.daysToNextCoupon <= 0) {
      return undefined;
    }
    return (
      ((redemption + coupon) / dirtyPrice - 1) *
      frequency *
      (metrics.daysInPeriod / metrics.daysToNextCoupon)
    );
  }

  const targetPrice = price;
  const priceAtYield = (yieldRate: number): number | undefined =>
    pricePeriodicSecurity(metrics, rate, yieldRate, redemption, frequency);
  let lower = -frequency + 1e-10;
  let upper = Math.max(1, rate * 2 + 0.1);
  let lowerPrice = priceAtYield(lower);
  let upperPrice = priceAtYield(upper);
  for (
    let iteration = 0;
    iteration < 100 &&
    (lowerPrice === undefined ||
      upperPrice === undefined ||
      lowerPrice < targetPrice ||
      upperPrice > targetPrice);
    iteration += 1
  ) {
    if (upperPrice === undefined || upperPrice > targetPrice) {
      upper = upper * 2 + 1;
      upperPrice = priceAtYield(upper);
      continue;
    }
    lower = (lower - frequency) / 2;
    lowerPrice = priceAtYield(lower);
  }
  if (
    lowerPrice === undefined ||
    upperPrice === undefined ||
    lowerPrice < targetPrice ||
    upperPrice > targetPrice
  ) {
    return undefined;
  }

  let guess = Math.min(Math.max(rate, lower + 1e-8), upper - 1e-8);
  for (let iteration = 0; iteration < 100; iteration += 1) {
    const estimatedPrice = priceAtYield(guess);
    if (estimatedPrice === undefined) {
      return undefined;
    }
    const error = estimatedPrice - targetPrice;
    if (Math.abs(error) < 1e-12) {
      return guess;
    }

    const epsilon = Math.max(1e-7, Math.abs(guess) * 1e-6);
    const shiftedPrice = priceAtYield(guess + epsilon);
    const derivative =
      shiftedPrice === undefined ? undefined : (shiftedPrice - estimatedPrice) / epsilon;
    let nextGuess =
      derivative === undefined || !Number.isFinite(derivative) || derivative === 0
        ? (lower + upper) / 2
        : guess - error / derivative;
    if (!Number.isFinite(nextGuess) || nextGuess <= lower || nextGuess >= upper) {
      nextGuess = (lower + upper) / 2;
    }

    const boundedPrice = priceAtYield(nextGuess);
    if (boundedPrice === undefined) {
      return undefined;
    }
    if (boundedPrice > targetPrice) {
      lower = nextGuess;
    } else {
      upper = nextGuess;
    }
    guess = nextGuess;
    if (Math.abs(upper - lower) < 1e-12) {
      return guess;
    }
  }
  return guess;
}

function macaulayDuration(
  metrics: CouponMetrics,
  couponRate: number,
  yieldRate: number,
  frequency: number,
): number | undefined {
  const price = pricePeriodicSecurity(metrics, couponRate, yieldRate, 100, frequency);
  if (price === undefined || price <= 0) {
    return undefined;
  }

  const coupon = (100 * couponRate) / frequency;
  const periodsToNextCoupon = metrics.daysToNextCoupon / metrics.daysInPeriod;
  const discountBase = 1 + yieldRate / frequency;
  if (discountBase <= 0) {
    return undefined;
  }

  let weightedPresentValue = 0;
  for (let period = 1; period <= metrics.periodsRemaining; period += 1) {
    const periodsToCashflow = period - 1 + periodsToNextCoupon;
    const timeInYears = periodsToCashflow / frequency;
    const cashflow = period === metrics.periodsRemaining ? 100 + coupon : coupon;
    weightedPresentValue += (timeInYears * cashflow) / Math.pow(discountBase, periodsToCashflow);
  }
  return weightedPresentValue / price;
}

function getAmordegrc(
  cost: number,
  datePurchased: number,
  firstPeriod: number,
  salvage: number,
  period: number,
  rate: number,
  basis: number,
): number | undefined {
  if (
    datePurchased > firstPeriod ||
    rate <= 0 ||
    salvage > cost ||
    cost <= 0 ||
    salvage < 0 ||
    period < 0
  ) {
    return undefined;
  }

  const serialPer = Math.trunc(period);
  if (!Number.isFinite(serialPer) || serialPer < 0) {
    return undefined;
  }

  const fractionalRate = yearFracByBasis(datePurchased, firstPeriod, basis);
  if (fractionalRate === undefined) {
    return undefined;
  }

  const useRate = 1 / rate;
  let amortizationCoefficient = 1.0;
  if (useRate < 3.0) {
    amortizationCoefficient = 1.0;
  } else if (useRate < 5.0) {
    amortizationCoefficient = 1.5;
  } else if (useRate <= 6.0) {
    amortizationCoefficient = 2.0;
  } else {
    amortizationCoefficient = 2.5;
  }

  const adjustedRate = rate * amortizationCoefficient;
  let currentRate = Math.round(fractionalRate * adjustedRate * cost);
  let currentCost = cost - currentRate;
  let remaining = currentCost - salvage;

  for (let step = 0; step < serialPer; step += 1) {
    currentRate = Math.round(adjustedRate * currentCost);
    remaining -= currentRate;
    if (remaining < 0.0) {
      if (serialPer - step === 0 || serialPer - step === 1) {
        return Math.round(currentCost * 0.5);
      }
      return 0.0;
    }
    currentCost -= currentRate;
  }

  return currentRate;
}

function getAmorlinc(
  cost: number,
  datePurchased: number,
  firstPeriod: number,
  salvage: number,
  period: number,
  rate: number,
  basis: number,
): number | undefined {
  if (
    datePurchased > firstPeriod ||
    rate <= 0 ||
    salvage > cost ||
    cost <= 0 ||
    salvage < 0 ||
    period < 0
  ) {
    return undefined;
  }

  const serialPer = Math.trunc(period);
  if (!Number.isFinite(serialPer) || serialPer < 0) {
    return undefined;
  }

  const fractionalRate = yearFracByBasis(datePurchased, firstPeriod, basis);
  if (fractionalRate === undefined) {
    return undefined;
  }

  const fullRate = cost * rate;
  const remainingCost = cost - salvage;
  const firstRate = fractionalRate * rate * cost;
  const fullPeriods = Math.trunc((cost - salvage - firstRate) / fullRate);

  let result = 0.0;
  if (serialPer === 0) {
    result = firstRate;
  } else if (serialPer <= fullPeriods) {
    result = fullRate;
  } else if (serialPer === fullPeriods + 1) {
    result = remainingCost - fullRate * fullPeriods - firstRate;
  }

  return result > 0.0 ? result : 0.0;
}

function roundToDigits(value: number, digits: number): number {
  if (digits >= 0) {
    const factor = 10 ** digits;
    return Math.round(value * factor) / factor;
  }
  const factor = 10 ** -digits;
  return Math.round(value / factor) * factor;
}

function roundUpToDigits(value: number, digits: number): number {
  if (digits >= 0) {
    const factor = 10 ** digits;
    return (value >= 0 ? Math.ceil(value * factor) : Math.floor(value * factor)) / factor;
  }
  const factor = 10 ** -digits;
  return (value >= 0 ? Math.ceil(value / factor) : Math.floor(value / factor)) * factor;
}

function roundDownToDigits(value: number, digits: number): number {
  if (digits >= 0) {
    const factor = 10 ** digits;
    return (value >= 0 ? Math.floor(value * factor) : Math.ceil(value * factor)) / factor;
  }
  const factor = 10 ** -digits;
  return (value >= 0 ? Math.floor(value / factor) : Math.ceil(value / factor)) * factor;
}

function roundTowardZero(value: number, digits: number): number {
  if (digits >= 0) {
    const factor = 10 ** digits;
    return Math.trunc(value * factor) / factor;
  }
  const factor = 10 ** -digits;
  return Math.trunc(value / factor) * factor;
}

function roundWith(value: CellValue, digits: CellValue | undefined): CellValue {
  const numberValue = toNumber(value);
  const digitValue = digits === undefined ? 0 : toNumber(digits);
  if (numberValue === undefined || digitValue === undefined) {
    return valueError();
  }
  return numberResult(roundToDigits(numberValue, Math.trunc(digitValue)));
}

function floorWith(value: CellValue, significance?: CellValue): CellValue {
  const numberValue = toNumber(value);
  const significanceValue = significance === undefined ? 1 : toNumber(significance);
  if (numberValue === undefined || significanceValue === undefined) {
    return valueError();
  }
  if (significanceValue === 0) {
    return { tag: ValueTag.Error, code: ErrorCode.Div0 };
  }
  return numberResult(Math.floor(numberValue / significanceValue) * significanceValue);
}

function ceilingWith(value: CellValue, significance?: CellValue): CellValue {
  const numberValue = toNumber(value);
  const significanceValue = significance === undefined ? 1 : toNumber(significance);
  if (numberValue === undefined || significanceValue === undefined) {
    return valueError();
  }
  if (significanceValue === 0) {
    return { tag: ValueTag.Error, code: ErrorCode.Div0 };
  }
  return numberResult(Math.ceil(numberValue / significanceValue) * significanceValue);
}

function unaryMath(value: CellValue, evaluate: (numeric: number) => number): CellValue {
  return numericResultOrError(evaluate(toNumber(value) ?? 0));
}

function binaryMath(
  left: CellValue,
  right: CellValue,
  evaluate: (left: number, right: number) => number,
): CellValue {
  return numericResultOrError(evaluate(toNumber(left) ?? 0, toNumber(right) ?? 0));
}

function collectNumericArgs(args: CellValue[]): number[] {
  return args.map(toNumber).filter((value): value is number => value !== undefined);
}

function collectStatNumericArgs(args: CellValue[]): number[] {
  const values: number[] = [];
  for (const arg of args) {
    if (arg.tag === ValueTag.Number) {
      values.push(arg.value);
      continue;
    }
    if (arg.tag === ValueTag.Boolean) {
      values.push(arg.value ? 1 : 0);
    }
  }
  return values;
}

function factorialValue(value: number): number | undefined {
  if (!Number.isFinite(value) || value < 0) {
    return undefined;
  }
  const truncated = Math.trunc(value);
  let result = 1;
  for (let index = 2; index <= truncated; index += 1) {
    result *= index;
  }
  return result;
}

function doubleFactorialValue(value: number): number | undefined {
  if (!Number.isFinite(value) || value < 0) {
    return undefined;
  }
  const truncated = Math.trunc(value);
  let result = 1;
  for (let index = truncated; index >= 2; index -= 2) {
    result *= index;
  }
  return result;
}

function gcdPair(left: number, right: number): number {
  let a = Math.abs(Math.trunc(left));
  let b = Math.abs(Math.trunc(right));
  while (b !== 0) {
    const next = a % b;
    a = b;
    b = next;
  }
  return a;
}

function lcmPair(left: number, right: number): number {
  if (left === 0 || right === 0) {
    return 0;
  }
  return Math.abs(Math.trunc(left * right)) / gcdPair(left, right);
}

function evenValue(numberValue: number): number {
  const sign = numberValue < 0 ? -1 : 1;
  const rounded = Math.ceil(Math.abs(numberValue) / 2) * 2;
  return sign * rounded;
}

function oddValue(numberValue: number): number {
  const sign = numberValue < 0 ? -1 : 1;
  const rounded = Math.ceil(Math.abs(numberValue));
  const odd = rounded % 2 === 0 ? rounded + 1 : rounded;
  return sign * odd;
}

function buildIdentityMatrix(size: number): ArrayValue {
  const values: CellValue[] = [];
  for (let row = 0; row < size; row += 1) {
    for (let col = 0; col < size; col += 1) {
      values.push(numberResult(row === col ? 1 : 0));
    }
  }
  return { kind: "array", rows: size, cols: size, values };
}

const complexBuiltins = createComplexBuiltins({ toNumber, numberResult, valueError });
const radixBuiltins = createRadixBuiltins({
  toNumber,
  integerValue,
  nonNegativeIntegerValue,
  valueError,
  numberResult,
});

const LANCZOS_G = 7;
const LANCZOS_COEFFICIENTS = [
  676.5203681218851, -1259.1392167224028, 771.3234287776531, -176.6150291621406, 12.507343278686905,
  -0.13857109526572012, 9.984369578019572e-6, 1.5056327351493116e-7,
] as const;

function logGamma(value: number): number {
  if (!Number.isFinite(value) || value <= 0) {
    return Number.NaN;
  }
  let sum = 0.9999999999998099;
  const shifted = value - 1;
  LANCZOS_COEFFICIENTS.forEach((coefficient, index) => {
    sum += coefficient / (shifted + index + 1);
  });
  const t = shifted + LANCZOS_G + 0.5;
  return 0.5 * Math.log(2 * Math.PI) + (shifted + 0.5) * Math.log(t) - t + Math.log(sum);
}

const BESSEL_EPSILON = 1e-8;

function gammaReal(value: number): number {
  if (!Number.isFinite(value) || (value <= 0 && value === Math.trunc(value))) {
    return Number.NaN;
  }
  if (value < 0.5) {
    const sine = Math.sin(Math.PI * value);
    return sine === 0 ? Number.NaN : Math.PI / (sine * gammaReal(1 - value));
  }
  return Math.exp(logGamma(value));
}

function besselSeries(order: number, x: number, alternating: boolean): number {
  const half = x / 2;
  let term = half ** order / gammaReal(order + 1);
  if (!Number.isFinite(term)) {
    return Number.NaN;
  }
  let sum = term;
  for (let index = 0; index < 400; index += 1) {
    const denominator = (index + 1) * (index + order + 1);
    if (denominator === 0) {
      return Number.NaN;
    }
    term *= ((alternating ? -1 : 1) * half * half) / denominator;
    sum += term;
    if (Math.abs(term) <= Math.abs(sum) * 1e-15) {
      break;
    }
  }
  return sum;
}

function besselJValue(x: number, order: number): number {
  if (x === 0) {
    return order === 0 ? 1 : 0;
  }
  const absolute = Math.abs(x);
  const result = besselSeries(order, absolute, true);
  return x < 0 && order % 2 === 1 ? -result : result;
}

function besselIValue(x: number, order: number): number {
  if (x === 0) {
    return order === 0 ? 1 : 0;
  }
  const absolute = Math.abs(x);
  const result = besselSeries(order, absolute, false);
  return x < 0 && order % 2 === 1 ? -result : result;
}

function besselYValue(x: number, order: number): number {
  const shiftedOrder = order + BESSEL_EPSILON;
  return (
    (besselSeries(shiftedOrder, x, true) * Math.cos(Math.PI * shiftedOrder) -
      besselSeries(-shiftedOrder, x, true)) /
    Math.sin(Math.PI * shiftedOrder)
  );
}

function besselKValue(x: number, order: number): number {
  const shiftedOrder = order + BESSEL_EPSILON;
  return (
    ((Math.PI / 2) *
      (besselSeries(-shiftedOrder, x, false) - besselSeries(shiftedOrder, x, false))) /
    Math.sin(Math.PI * shiftedOrder)
  );
}

function gammaFunction(value: number): number {
  if (!Number.isFinite(value) || (Number.isInteger(value) && value <= 0)) {
    return Number.NaN;
  }
  if (value < 0.5) {
    const sine = Math.sin(Math.PI * value);
    if (sine === 0) {
      return Number.NaN;
    }
    return Math.PI / (sine * gammaFunction(1 - value));
  }
  return Math.exp(logGamma(value));
}

function regularizedLowerGamma(shape: number, x: number): number {
  if (!Number.isFinite(shape) || !Number.isFinite(x) || shape <= 0 || x < 0) {
    return Number.NaN;
  }
  if (x === 0) {
    return 0;
  }
  const logGammaShape = logGamma(shape);
  if (!Number.isFinite(logGammaShape)) {
    return Number.NaN;
  }
  if (x < shape + 1) {
    let term = 1 / shape;
    let sum = term;
    for (let iteration = 1; iteration < 1000; iteration += 1) {
      term *= x / (shape + iteration);
      sum += term;
      if (Math.abs(term) <= Math.abs(sum) * 1e-14) {
        break;
      }
    }
    return sum * Math.exp(-x + shape * Math.log(x) - logGammaShape);
  }

  let b = x + 1 - shape;
  let c = 1 / 1e-300;
  let d = 1 / b;
  let h = d;
  for (let iteration = 1; iteration < 1000; iteration += 1) {
    const factor = -iteration * (iteration - shape);
    b += 2;
    d = factor * d + b;
    if (Math.abs(d) < 1e-300) {
      d = 1e-300;
    }
    c = b + factor / c;
    if (Math.abs(c) < 1e-300) {
      c = 1e-300;
    }
    d = 1 / d;
    const delta = d * c;
    h *= delta;
    if (Math.abs(delta - 1) <= 1e-14) {
      break;
    }
  }
  return 1 - Math.exp(-x + shape * Math.log(x) - logGammaShape) * h;
}

function regularizedUpperGamma(shape: number, x: number): number {
  const lower = regularizedLowerGamma(shape, x);
  return Number.isFinite(lower) ? 1 - lower : Number.NaN;
}

function logBeta(alpha: number, beta: number): number {
  return logGamma(alpha) + logGamma(beta) - logGamma(alpha + beta);
}

function betaContinuedFraction(x: number, alpha: number, beta: number): number {
  const maxIterations = 200;
  const epsilon = 1e-14;
  const tiny = 1e-300;
  const qab = alpha + beta;
  const qap = alpha + 1;
  const qam = alpha - 1;
  let c = 1;
  let d = 1 - (qab * x) / qap;
  if (Math.abs(d) < tiny) {
    d = tiny;
  }
  d = 1 / d;
  let h = d;
  for (let iteration = 1; iteration <= maxIterations; iteration += 1) {
    const step = iteration * 2;
    let factor = (iteration * (beta - iteration) * x) / ((qam + step) * (alpha + step));
    d = 1 + factor * d;
    if (Math.abs(d) < tiny) {
      d = tiny;
    }
    c = 1 + factor / c;
    if (Math.abs(c) < tiny) {
      c = tiny;
    }
    d = 1 / d;
    h *= d * c;

    factor = (-(alpha + iteration) * (qab + iteration) * x) / ((alpha + step) * (qap + step));
    d = 1 + factor * d;
    if (Math.abs(d) < tiny) {
      d = tiny;
    }
    c = 1 + factor / c;
    if (Math.abs(c) < tiny) {
      c = tiny;
    }
    d = 1 / d;
    const delta = d * c;
    h *= delta;
    if (Math.abs(delta - 1) <= epsilon) {
      break;
    }
  }
  return h;
}

function regularizedBeta(x: number, alpha: number, beta: number): number {
  if (
    !Number.isFinite(x) ||
    !Number.isFinite(alpha) ||
    !Number.isFinite(beta) ||
    alpha <= 0 ||
    beta <= 0 ||
    x < 0 ||
    x > 1
  ) {
    return Number.NaN;
  }
  if (x === 0) {
    return 0;
  }
  if (x === 1) {
    return 1;
  }
  const logTerm = alpha * Math.log(x) + beta * Math.log(1 - x) - logBeta(alpha, beta);
  if (!Number.isFinite(logTerm)) {
    return Number.NaN;
  }
  const front = Math.exp(logTerm);
  if (x < (alpha + 1) / (alpha + beta + 2)) {
    return (front * betaContinuedFraction(x, alpha, beta)) / alpha;
  }
  return 1 - (front * betaContinuedFraction(1 - x, beta, alpha)) / beta;
}

function betaDistributionDensity(
  x: number,
  alpha: number,
  beta: number,
  lowerBound = 0,
  upperBound = 1,
): number {
  if (
    !Number.isFinite(x) ||
    !Number.isFinite(alpha) ||
    !Number.isFinite(beta) ||
    !Number.isFinite(lowerBound) ||
    !Number.isFinite(upperBound) ||
    alpha <= 0 ||
    beta <= 0 ||
    upperBound <= lowerBound ||
    x < lowerBound ||
    x > upperBound
  ) {
    return Number.NaN;
  }
  const scale = upperBound - lowerBound;
  const normalized = (x - lowerBound) / scale;
  if (normalized === 0) {
    if (alpha === 1) {
      return beta / scale;
    }
    return alpha < 1 ? Number.POSITIVE_INFINITY : 0;
  }
  if (normalized === 1) {
    if (beta === 1) {
      return alpha / scale;
    }
    return beta < 1 ? Number.POSITIVE_INFINITY : 0;
  }
  return Math.exp(
    (alpha - 1) * Math.log(normalized) +
      (beta - 1) * Math.log(1 - normalized) -
      logBeta(alpha, beta) -
      Math.log(scale),
  );
}

function betaDistributionCdf(
  x: number,
  alpha: number,
  beta: number,
  lowerBound = 0,
  upperBound = 1,
): number {
  if (
    !Number.isFinite(x) ||
    !Number.isFinite(alpha) ||
    !Number.isFinite(beta) ||
    !Number.isFinite(lowerBound) ||
    !Number.isFinite(upperBound) ||
    alpha <= 0 ||
    beta <= 0 ||
    upperBound <= lowerBound ||
    x < lowerBound ||
    x > upperBound
  ) {
    return Number.NaN;
  }
  return regularizedBeta((x - lowerBound) / (upperBound - lowerBound), alpha, beta);
}

function inverseRegularizedBeta(
  probability: number,
  alpha: number,
  beta: number,
): number | undefined {
  if (
    !Number.isFinite(probability) ||
    !Number.isFinite(alpha) ||
    !Number.isFinite(beta) ||
    !(probability > 0 && probability < 1) ||
    alpha <= 0 ||
    beta <= 0
  ) {
    return undefined;
  }
  let lower = 0;
  let upper = 1;
  for (let iteration = 0; iteration < 100; iteration += 1) {
    const midpoint = (lower + upper) / 2;
    const cdf = regularizedBeta(midpoint, alpha, beta);
    if (!Number.isFinite(cdf)) {
      return undefined;
    }
    if (cdf < probability) {
      lower = midpoint;
    } else {
      upper = midpoint;
    }
  }
  return (lower + upper) / 2;
}

function betaDistributionInverse(
  probability: number,
  alpha: number,
  beta: number,
  lowerBound = 0,
  upperBound = 1,
): number | undefined {
  if (
    !Number.isFinite(probability) ||
    !Number.isFinite(alpha) ||
    !Number.isFinite(beta) ||
    !Number.isFinite(lowerBound) ||
    !Number.isFinite(upperBound) ||
    !(probability > 0 && probability < 1) ||
    alpha <= 0 ||
    beta <= 0 ||
    upperBound <= lowerBound
  ) {
    return undefined;
  }
  const normalized = inverseRegularizedBeta(probability, alpha, beta);
  return normalized === undefined ? undefined : lowerBound + (upperBound - lowerBound) * normalized;
}

function fDistributionDensity(x: number, degreesFreedom1: number, degreesFreedom2: number): number {
  if (
    !Number.isFinite(x) ||
    !Number.isFinite(degreesFreedom1) ||
    !Number.isFinite(degreesFreedom2) ||
    x < 0 ||
    degreesFreedom1 < 1 ||
    degreesFreedom2 < 1
  ) {
    return Number.NaN;
  }
  if (x === 0) {
    return degreesFreedom1 === 2 ? 1 : degreesFreedom1 < 2 ? Number.POSITIVE_INFINITY : 0;
  }
  const a = degreesFreedom1 / 2;
  const b = degreesFreedom2 / 2;
  return Math.exp(
    a * Math.log(degreesFreedom1) +
      b * Math.log(degreesFreedom2) +
      (a - 1) * Math.log(x) -
      (a + b) * Math.log(degreesFreedom1 * x + degreesFreedom2) -
      logBeta(a, b),
  );
}

function fDistributionCdf(x: number, degreesFreedom1: number, degreesFreedom2: number): number {
  if (
    !Number.isFinite(x) ||
    !Number.isFinite(degreesFreedom1) ||
    !Number.isFinite(degreesFreedom2) ||
    x < 0 ||
    degreesFreedom1 < 1 ||
    degreesFreedom2 < 1
  ) {
    return Number.NaN;
  }
  const a = degreesFreedom1 / 2;
  const b = degreesFreedom2 / 2;
  const transformed = (degreesFreedom1 * x) / (degreesFreedom1 * x + degreesFreedom2);
  return regularizedBeta(transformed, a, b);
}

function inverseFDistribution(
  probability: number,
  degreesFreedom1: number,
  degreesFreedom2: number,
): number | undefined {
  if (
    !Number.isFinite(probability) ||
    !Number.isFinite(degreesFreedom1) ||
    !Number.isFinite(degreesFreedom2) ||
    !(probability > 0 && probability < 1) ||
    degreesFreedom1 < 1 ||
    degreesFreedom2 < 1
  ) {
    return undefined;
  }
  const transformed = inverseRegularizedBeta(probability, degreesFreedom1 / 2, degreesFreedom2 / 2);
  if (transformed === undefined || transformed >= 1) {
    return undefined;
  }
  return (degreesFreedom2 * transformed) / (degreesFreedom1 * (1 - transformed));
}

function studentTDensity(x: number, degreesFreedom: number): number {
  if (!Number.isFinite(x) || !Number.isFinite(degreesFreedom) || degreesFreedom < 1) {
    return Number.NaN;
  }
  const halfDegrees = degreesFreedom / 2;
  return Math.exp(
    logGamma((degreesFreedom + 1) / 2) -
      logGamma(halfDegrees) -
      0.5 * (Math.log(degreesFreedom) + Math.log(Math.PI)) -
      ((degreesFreedom + 1) / 2) * Math.log(1 + (x * x) / degreesFreedom),
  );
}

function studentTCdf(x: number, degreesFreedom: number): number {
  if (!Number.isFinite(x) || !Number.isFinite(degreesFreedom) || degreesFreedom < 1) {
    return Number.NaN;
  }
  if (x === 0) {
    return 0.5;
  }
  const transformed = degreesFreedom / (degreesFreedom + x * x);
  const tail = regularizedBeta(transformed, degreesFreedom / 2, 0.5);
  if (!Number.isFinite(tail)) {
    return Number.NaN;
  }
  return x > 0 ? 1 - tail / 2 : tail / 2;
}

function inverseStudentT(probability: number, degreesFreedom: number): number | undefined {
  if (
    !Number.isFinite(probability) ||
    !Number.isFinite(degreesFreedom) ||
    !(probability > 0 && probability < 1) ||
    degreesFreedom < 1
  ) {
    return undefined;
  }
  if (probability === 0.5) {
    return 0;
  }
  if (probability < 0.5) {
    const mirrored = inverseStudentT(1 - probability, degreesFreedom);
    return mirrored === undefined ? undefined : -mirrored;
  }
  let lower = 0;
  let upper = 1;
  let upperCdf = studentTCdf(upper, degreesFreedom);
  while (Number.isFinite(upperCdf) && upperCdf < probability && upper < 1e10) {
    lower = upper;
    upper *= 2;
    upperCdf = studentTCdf(upper, degreesFreedom);
  }
  if (!Number.isFinite(upperCdf)) {
    return undefined;
  }
  for (let iteration = 0; iteration < 100; iteration += 1) {
    const midpoint = (lower + upper) / 2;
    const cdf = studentTCdf(midpoint, degreesFreedom);
    if (!Number.isFinite(cdf)) {
      return undefined;
    }
    if (cdf < probability) {
      lower = midpoint;
    } else {
      upper = midpoint;
    }
  }
  return (lower + upper) / 2;
}

function logCombination(total: number, chosen: number): number {
  if (!Number.isInteger(total) || !Number.isInteger(chosen) || chosen < 0 || chosen > total) {
    return Number.NaN;
  }
  return logGamma(total + 1) - logGamma(chosen + 1) - logGamma(total - chosen + 1);
}

function binomialProbability(successes: number, trials: number, probability: number): number {
  if (
    !Number.isInteger(successes) ||
    !Number.isInteger(trials) ||
    successes < 0 ||
    trials < 0 ||
    successes > trials ||
    probability < 0 ||
    probability > 1
  ) {
    return Number.NaN;
  }
  if (probability === 0) {
    return successes === 0 ? 1 : 0;
  }
  if (probability === 1) {
    return successes === trials ? 1 : 0;
  }
  return Math.exp(
    logCombination(trials, successes) +
      successes * Math.log(probability) +
      (trials - successes) * Math.log(1 - probability),
  );
}

function negativeBinomialProbability(
  failures: number,
  successes: number,
  probability: number,
): number {
  if (
    !Number.isInteger(failures) ||
    !Number.isInteger(successes) ||
    failures < 0 ||
    successes <= 0 ||
    probability < 0 ||
    probability > 1
  ) {
    return Number.NaN;
  }
  if (probability === 0) {
    return 0;
  }
  if (probability === 1) {
    return failures === 0 ? 1 : 0;
  }
  return Math.exp(
    logCombination(failures + successes - 1, failures) +
      failures * Math.log(1 - probability) +
      successes * Math.log(probability),
  );
}

function hypergeometricProbability(
  sampleSuccesses: number,
  sampleSize: number,
  populationSuccesses: number,
  populationSize: number,
): number {
  if (
    !Number.isInteger(sampleSuccesses) ||
    !Number.isInteger(sampleSize) ||
    !Number.isInteger(populationSuccesses) ||
    !Number.isInteger(populationSize) ||
    sampleSuccesses < 0 ||
    sampleSize < 0 ||
    populationSuccesses < 0 ||
    populationSize <= 0 ||
    sampleSize > populationSize ||
    populationSuccesses > populationSize
  ) {
    return Number.NaN;
  }
  const minimum = Math.max(0, sampleSize - (populationSize - populationSuccesses));
  const maximum = Math.min(sampleSize, populationSuccesses);
  if (sampleSuccesses < minimum || sampleSuccesses > maximum) {
    return 0;
  }
  return Math.exp(
    logCombination(populationSuccesses, sampleSuccesses) +
      logCombination(populationSize - populationSuccesses, sampleSize - sampleSuccesses) -
      logCombination(populationSize, sampleSize),
  );
}

function poissonProbability(events: number, mean: number): number {
  if (!Number.isInteger(events) || events < 0 || !Number.isFinite(mean) || mean < 0) {
    return Number.NaN;
  }
  if (mean === 0) {
    return events === 0 ? 1 : 0;
  }
  return Math.exp(events * Math.log(mean) - mean - logGamma(events + 1));
}

function gammaDistributionDensity(x: number, alpha: number, beta: number): number {
  if (!Number.isFinite(x) || !Number.isFinite(alpha) || !Number.isFinite(beta)) {
    return Number.NaN;
  }
  if (x < 0 || alpha <= 0 || beta <= 0) {
    return Number.NaN;
  }
  if (x === 0) {
    if (alpha === 1) {
      return 1 / beta;
    }
    return alpha < 1 ? Number.POSITIVE_INFINITY : 0;
  }
  return Math.exp((alpha - 1) * Math.log(x) - x / beta - logGamma(alpha) - alpha * Math.log(beta));
}

function gammaDistributionCdf(x: number, alpha: number, beta: number): number {
  return regularizedLowerGamma(alpha, x / beta);
}

function inverseGammaDistribution(
  probability: number,
  alpha: number,
  beta: number,
): number | undefined {
  if (
    !Number.isFinite(probability) ||
    !Number.isFinite(alpha) ||
    !Number.isFinite(beta) ||
    !(probability > 0 && probability < 1) ||
    !(alpha > 0) ||
    !(beta > 0)
  ) {
    return undefined;
  }

  let estimate = alpha * beta;
  if (!(estimate > 0) || !Number.isFinite(estimate)) {
    estimate = 1;
  }

  let lower = 0;
  let upper = Math.max(estimate, 1);
  let upperCdf = gammaDistributionCdf(upper, alpha, beta);
  for (let iteration = 0; iteration < 64 && upperCdf < probability; iteration += 1) {
    upper *= 2;
    upperCdf = gammaDistributionCdf(upper, alpha, beta);
  }
  if (!(upperCdf >= probability)) {
    return undefined;
  }

  let current = Math.min(Math.max(estimate, lower), upper);
  for (let iteration = 0; iteration < 20; iteration += 1) {
    const cdf = gammaDistributionCdf(current, alpha, beta);
    if (!Number.isFinite(cdf)) {
      break;
    }
    if (cdf < probability) {
      lower = current;
    } else {
      upper = current;
    }
    const density = gammaDistributionDensity(current, alpha, beta);
    if (!(density > 0) || !Number.isFinite(density)) {
      current = (lower + upper) / 2;
      continue;
    }
    const next = current - (cdf - probability) / density;
    current = Number.isFinite(next) && next > lower && next < upper ? next : (lower + upper) / 2;
  }

  for (let iteration = 0; iteration < 60; iteration += 1) {
    const midpoint = (lower + upper) / 2;
    const cdf = gammaDistributionCdf(midpoint, alpha, beta);
    if (!Number.isFinite(cdf)) {
      return undefined;
    }
    if (cdf < probability) {
      lower = midpoint;
    } else {
      upper = midpoint;
    }
  }

  return (lower + upper) / 2;
}

function chiSquareDensity(x: number, degreesFreedom: number): number {
  return gammaDistributionDensity(x, degreesFreedom / 2, 2);
}

function chiSquareCdf(x: number, degreesFreedom: number): number {
  return gammaDistributionCdf(x, degreesFreedom / 2, 2);
}

function inverseChiSquare(probability: number, degreesFreedom: number): number | undefined {
  if (
    !Number.isFinite(probability) ||
    !Number.isFinite(degreesFreedom) ||
    !(probability > 0 && probability < 1) ||
    !(degreesFreedom >= 1)
  ) {
    return undefined;
  }

  const z = inverseStandardNormal(probability);
  const approximationFactor =
    z === undefined
      ? Number.NaN
      : 1 - 2 / (9 * degreesFreedom) + z * Math.sqrt(2 / (9 * degreesFreedom));
  let estimate =
    Number.isFinite(approximationFactor) && approximationFactor > 0
      ? degreesFreedom * approximationFactor ** 3
      : degreesFreedom;
  if (!(estimate > 0) || !Number.isFinite(estimate)) {
    estimate = Math.max(degreesFreedom, 1);
  }

  let lower = 0;
  let upper = Math.max(estimate, 1);
  let upperCdf = chiSquareCdf(upper, degreesFreedom);
  for (let iteration = 0; iteration < 64 && upperCdf < probability; iteration += 1) {
    upper *= 2;
    upperCdf = chiSquareCdf(upper, degreesFreedom);
  }
  if (!(upperCdf >= probability)) {
    return undefined;
  }

  let current = Math.min(Math.max(estimate, lower), upper);
  for (let iteration = 0; iteration < 20; iteration += 1) {
    const cdf = chiSquareCdf(current, degreesFreedom);
    if (!Number.isFinite(cdf)) {
      break;
    }
    if (cdf < probability) {
      lower = current;
    } else {
      upper = current;
    }
    const density = chiSquareDensity(current, degreesFreedom);
    if (!(density > 0) || !Number.isFinite(density)) {
      current = (lower + upper) / 2;
      continue;
    }
    const next = current - (cdf - probability) / density;
    current = Number.isFinite(next) && next > lower && next < upper ? next : (lower + upper) / 2;
  }

  for (let iteration = 0; iteration < 60; iteration += 1) {
    const midpoint = (lower + upper) / 2;
    const cdf = chiSquareCdf(midpoint, degreesFreedom);
    if (!Number.isFinite(cdf)) {
      return undefined;
    }
    if (cdf < probability) {
      lower = midpoint;
    } else {
      upper = midpoint;
    }
  }

  return (lower + upper) / 2;
}

function standardNormalPdf(value: number): number {
  return Math.exp(-(value * value) / 2) / Math.sqrt(2 * Math.PI);
}

function standardNormalCdf(value: number): number {
  return 0.5 * (1 + erfApprox(value / Math.SQRT2));
}

function inverseStandardNormal(probability: number): number | undefined {
  if (!(probability > 0 && probability < 1)) {
    return undefined;
  }
  const a = [
    -39.69683028665376, 220.9460984245205, -275.9285104469687, 138.357751867269, -30.66479806614716,
    2.506628277459239,
  ];
  const b = [
    -54.47609879822406, 161.5858368580409, -155.6989798598866, 66.80131188771972,
    -13.28068155288572,
  ];
  const c = [
    -0.007784894002430293, -0.3223964580411365, -2.400758277161838, -2.549732539343734,
    4.374664141464968, 2.938163982698783,
  ];
  const d = [0.007784695709041462, 0.3224671290700398, 2.445134137142996, 3.754408661907416];
  const lower = 0.02425;
  const upper = 1 - lower;

  if (probability < lower) {
    const q = Math.sqrt(-2 * Math.log(probability));
    return (
      (((((c[0]! * q + c[1]!) * q + c[2]!) * q + c[3]!) * q + c[4]!) * q + c[5]!) /
      ((((d[0]! * q + d[1]!) * q + d[2]!) * q + d[3]!) * q + 1)
    );
  }
  if (probability > upper) {
    const q = Math.sqrt(-2 * Math.log(1 - probability));
    return -(
      (((((c[0]! * q + c[1]!) * q + c[2]!) * q + c[3]!) * q + c[4]!) * q + c[5]!) /
      ((((d[0]! * q + d[1]!) * q + d[2]!) * q + d[3]!) * q + 1)
    );
  }
  const q = probability - 0.5;
  const r = q * q;
  return (
    ((((((a[0]! * r + a[1]!) * r + a[2]!) * r + a[3]!) * r + a[4]!) * r + a[5]!) * q) /
    (((((b[0]! * r + b[1]!) * r + b[2]!) * r + b[3]!) * r + b[4]!) * r + 1)
  );
}

function skewSample(numbers: number[]): number | undefined {
  if (numbers.length < 3) {
    return undefined;
  }
  const mean = numbers.reduce((sum, value) => sum + value, 0) / numbers.length;
  const stddev = sampleStandardDeviation(numbers);
  if (!(stddev > 0)) {
    return undefined;
  }
  const moment3 = numbers.reduce((sum, value) => sum + (value - mean) ** 3, 0);
  const n = numbers.length;
  return (n * moment3) / ((n - 1) * (n - 2) * stddev ** 3);
}

function skewPopulation(numbers: number[]): number | undefined {
  if (numbers.length === 0) {
    return undefined;
  }
  const mean = numbers.reduce((sum, value) => sum + value, 0) / numbers.length;
  const stddev = populationStandardDeviation(numbers);
  if (!(stddev > 0)) {
    return undefined;
  }
  const moment3 = numbers.reduce((sum, value) => sum + (value - mean) ** 3, 0) / numbers.length;
  return moment3 / stddev ** 3;
}

function kurtosis(numbers: number[]): number | undefined {
  if (numbers.length < 4) {
    return undefined;
  }
  const mean = numbers.reduce((sum, value) => sum + value, 0) / numbers.length;
  const stddev = sampleStandardDeviation(numbers);
  if (!(stddev > 0)) {
    return undefined;
  }
  const n = numbers.length;
  const sum4 = numbers.reduce((sum, value) => sum + ((value - mean) / stddev) ** 4, 0);
  return (
    (n * (n + 1) * sum4) / ((n - 1) * (n - 2) * (n - 3)) - (3 * (n - 1) ** 2) / ((n - 2) * (n - 3))
  );
}

function percentileNormal(mean: number, standardDeviation: number, value: number): number {
  return standardNormalCdf((value - mean) / standardDeviation);
}

function inverseNormal(
  probability: number,
  mean: number,
  standardDeviation: number,
): number | undefined {
  const z = inverseStandardNormal(probability);
  return z === undefined ? undefined : mean + standardDeviation * z;
}

function coercePaymentType(value: CellValue | undefined, fallback: number): number | undefined {
  const type = integerValue(value, fallback);
  return type === 0 || type === 1 ? type : undefined;
}

function futureValue(
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

function presentValue(
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

function periodicPayment(
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

function totalPeriods(
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

function solveRate(
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

function dbDepreciation(
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

function ddbDepreciation(
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

function vdbDepreciation(
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

function interestPayment(
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

function principalPayment(
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

function cumulativePeriodicPayment(
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

function toZeroNumericValue(value: CellValue): number | undefined {
  if (value.tag === ValueTag.String) {
    return 0;
  }
  return toNumber(value);
}

function toColumnLabel(column: number): string | undefined {
  if (!Number.isInteger(column) || column < 1) {
    return undefined;
  }
  let current = column;
  let label = "";
  while (current > 0) {
    const offset = (current - 1) % 26;
    label = String.fromCharCode(65 + offset) + label;
    current = Math.floor((current - 1) / 26);
  }
  return label;
}

function formatThousands(value: string): string {
  return value.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}

function formatFixed(value: number, decimals: number, includeThousands: boolean): string {
  if (!Number.isFinite(value) || !Number.isInteger(decimals)) {
    return "";
  }
  const rounded = roundToDigits(value, decimals);
  const sign = rounded < 0 ? "-" : "";
  const unsigned = Math.abs(rounded);
  const fixedDecimals = decimals >= 0 ? decimals : 0;
  const fixed = unsigned.toFixed(fixedDecimals);
  const [integerPart = "0", fractionPart] = fixed.split(".");
  const normalizedInteger = includeThousands ? formatThousands(integerPart) : integerPart;
  return `${sign}${normalizedInteger}${fractionPart === undefined ? "" : `.${fractionPart}`}`;
}

function countLeadingZeros(value: number): number {
  if (value <= 0) {
    return 1;
  }
  return Math.max(1, Math.ceil(Math.log10(value)));
}

function isValidDollarFraction(fraction: number): boolean {
  if (!Number.isInteger(fraction) || fraction <= 0) {
    return false;
  }
  if (fraction === 1) {
    return true;
  }
  return Number.isInteger(Math.log2(fraction));
}

function parseDollarDecimal(value: number): { integerPart: number; fractionalNumerator: number } {
  const absolute = Math.abs(value);
  const parts = absolute.toString().split(".");
  const integerPart = Number(parts[0] ?? 0);
  const fractionalText = parts[1] ?? "0";
  return { integerPart, fractionalNumerator: Number(fractionalText) };
}

function aggregateByCode(functionNum: number, values: CellValue[]): CellValue {
  const normalized = functionNum > 100 ? functionNum - 100 : functionNum;
  const numericValues = collectNumericArgs(values);
  switch (normalized) {
    case 1:
      return numericValues.length === 0
        ? numberResult(0)
        : numberResult(numericValues.reduce((sum, value) => sum + value, 0) / numericValues.length);
    case 2:
      return numberResult(
        values.filter((value) => value.tag === ValueTag.Number || value.tag === ValueTag.Boolean)
          .length,
      );
    case 3:
      return numberResult(values.filter((value) => value.tag !== ValueTag.Empty).length);
    case 4:
      return numberResult(numericValues.length === 0 ? 0 : Math.max(...numericValues));
    case 5:
      return numberResult(numericValues.length === 0 ? 0 : Math.min(...numericValues));
    case 6:
      return numberResult(
        numericValues.length === 0
          ? 0
          : numericValues.reduce((product, value) => product * value, 1),
      );
    case 7:
      return numberResult(Math.sqrt(sampleVariance(numericValues)));
    case 8:
      return numberResult(Math.sqrt(populationVariance(numericValues)));
    case 9:
      return numberResult(numericValues.reduce((sum, value) => sum + value, 0));
    case 10:
      return numberResult(sampleVariance(numericValues));
    case 11:
      return numberResult(populationVariance(numericValues));
    default:
      return valueError();
  }
}

const scalarPlaceholderBuiltins = createBlockedBuiltinMap(scalarPlaceholderBuiltinNames);

const externalScalarBuiltinNames = [
  "CALL",
  "CUBEKPIMEMBER",
  "CUBEMEMBER",
  "CUBEMEMBERPROPERTY",
  "CUBERANKEDMEMBER",
  "CUBESET",
  "CUBESETCOUNT",
  "CUBEVALUE",
  "DDE",
  "DETECTLANGUAGE",
  "HYPERLINK",
  "IMAGE",
  "INFO",
  "REGISTER.ID",
  "RTD",
  "TRANSLATE",
  "WEBSERVICE",
] as const;

function createExternalScalarBuiltin(name: string): Builtin {
  return (...args) => {
    const existingError = firstError(args);
    if (existingError) {
      return existingError;
    }
    const external = getExternalScalarFunction(name);
    return external ? external(...args) : blockedError();
  };
}

const externalScalarBuiltins = Object.fromEntries(
  externalScalarBuiltinNames.map((name) => [name, createExternalScalarBuiltin(name)]),
) as Record<string, Builtin>;

const scalarBuiltins: Record<string, Builtin> = {
  SUM: (...args) => {
    const error = firstError(args);
    if (error) return error;
    return numberResult(args.reduce((sum, arg) => sum + (toNumber(arg) ?? 0), 0));
  },
  AVERAGEA: (...args) => {
    const error = firstError(args);
    if (error) return error;
    const numbers = args
      .map((arg) => toZeroNumericValue(arg))
      .filter((value): value is number => value !== undefined);
    return numberResult(
      numbers.length === 0 ? 0 : numbers.reduce((sum, value) => sum + value, 0) / numbers.length,
    );
  },
  AVERAGE: (...args) => {
    const error = firstError(args);
    if (error) return error;
    const numbers = collectNumericArgs(args);
    if (numbers.length === 0) return numberResult(0);
    return numberResult(numbers.reduce((sum, value) => sum + value, 0) / numbers.length);
  },
  AVG: (...args) => {
    const error = firstError(args);
    if (error) return error;
    const numbers = collectNumericArgs(args);
    if (numbers.length === 0) return numberResult(0);
    return numberResult(numbers.reduce((sum, value) => sum + value, 0) / numbers.length);
  },
  CHOOSE: (indexValue, ...values) => {
    const index = integerValue(indexValue);
    if (index === undefined || index < 1 || index > values.length) {
      return valueError();
    }
    const value = values[index - 1];
    return value === undefined ? valueError() : value;
  },
  MIN: (...args) =>
    numberResult(Math.min(...args.map((arg) => toNumber(arg) ?? Number.POSITIVE_INFINITY))),
  MAX: (...args) =>
    numberResult(Math.max(...args.map((arg) => toNumber(arg) ?? Number.NEGATIVE_INFINITY))),
  MAXA: (...args) => {
    const error = firstError(args);
    if (error) return error;
    const values = args
      .map((arg) => toZeroNumericValue(arg))
      .filter((value): value is number => value !== undefined);
    return values.length === 0 ? numberResult(0) : numberResult(Math.max(...values));
  },
  MINA: (...args) => {
    const error = firstError(args);
    if (error) return error;
    const values = args
      .map((arg) => toZeroNumericValue(arg))
      .filter((value): value is number => value !== undefined);
    return values.length === 0 ? numberResult(0) : numberResult(Math.min(...values));
  },
  COUNT: (...args) =>
    numberResult(
      args.filter((arg) => arg.tag === ValueTag.Number || arg.tag === ValueTag.Boolean).length,
    ),
  COUNTA: (...args) => numberResult(args.filter((arg) => arg.tag !== ValueTag.Empty).length),
  COUNTBLANK: (...args) => {
    let blanks = 0;
    for (const arg of args) {
      if (arg.tag === ValueTag.Empty) {
        blanks += 1;
      }
    }
    return numberResult(blanks);
  },
  ABS: (value) => numberResult(Math.abs(toNumber(value) ?? 0)),
  ADDRESS: (
    rowArg,
    columnArg,
    absNumArg = { tag: ValueTag.Number, value: 1 },
    refStyleArg = { tag: ValueTag.Number, value: 1 },
    sheetTextArg,
  ) => {
    const row = positiveIntegerValue(rowArg);
    const column = positiveIntegerValue(columnArg);
    const absNum = integerValue(absNumArg, 1);
    const refStyle = integerValue(refStyleArg, 1);
    if (
      row === undefined ||
      column === undefined ||
      absNum === undefined ||
      refStyle === undefined
    ) {
      return valueError();
    }
    if (![1, 2, 3, 4].includes(absNum) || ![1, 2].includes(refStyle)) {
      return valueError();
    }
    if (
      sheetTextArg !== undefined &&
      sheetTextArg.tag !== ValueTag.String &&
      sheetTextArg.tag !== ValueTag.Empty
    ) {
      return valueError();
    }
    if (sheetTextArg?.tag === ValueTag.Empty) {
      return valueError();
    }
    const columnLabel = toColumnLabel(column);
    if (columnLabel === undefined) {
      return valueError();
    }
    const sheetPrefix =
      sheetTextArg?.tag === ValueTag.String ? `'${sheetTextArg.value.replace(/'/g, "''")}'!` : "";
    if (refStyle === 2) {
      const rowLabel = absNum === 1 || absNum === 2 ? String(row) : `[${row}]`;
      const colLabel = absNum === 1 || absNum === 3 ? String(column) : `[${column}]`;
      return {
        tag: ValueTag.String,
        value: `${sheetPrefix}R${rowLabel}C${colLabel}`,
        stringId: 0,
      };
    }
    const rowLabel = absNum === 1 || absNum === 2 ? `$${row}` : `${row}`;
    const colLabel = absNum === 1 || absNum === 3 ? `$${columnLabel}` : columnLabel;
    return {
      tag: ValueTag.String,
      value: `${sheetPrefix}${colLabel}${rowLabel}`,
      stringId: 0,
    };
  },
  DOLLAR: (valueArg, decimalsArg = { tag: ValueTag.Number, value: 2 }, noCommasArg) => {
    const value = toNumber(valueArg);
    const decimals = toInteger(decimalsArg, 2);
    const noCommasValue = noCommasArg === undefined ? 0 : (toNumber(noCommasArg) ?? 0);
    if (value === undefined || decimals === undefined) {
      return valueError();
    }
    const text = formatFixed(value, decimals, noCommasValue === 0);
    if (text === "") {
      return valueError();
    }
    const normalizedText = text.startsWith("-") ? text.slice(1) : text;
    return {
      tag: ValueTag.String,
      value: value < 0 ? `-$${normalizedText}` : `$${text}`,
      stringId: 0,
    };
  },
  FIXED: (valueArg, decimalsArg = { tag: ValueTag.Number, value: 2 }, noCommasArg) => {
    const value = toNumber(valueArg);
    const decimals = toInteger(decimalsArg, 2);
    const noCommasValue = noCommasArg === undefined ? 0 : (toNumber(noCommasArg) ?? 0);
    if (value === undefined || decimals === undefined) {
      return valueError();
    }
    const text = formatFixed(value, decimals, noCommasValue === 0);
    return text === "" ? valueError() : { tag: ValueTag.String, value: text, stringId: 0 };
  },
  DOLLARDE: (valueArg, fractionArg) => {
    const value = toNumber(valueArg);
    const fraction = toInteger(fractionArg);
    if (value === undefined || fraction === undefined || !isValidDollarFraction(fraction)) {
      return valueError();
    }
    const { integerPart, fractionalNumerator } = parseDollarDecimal(value);
    if (fractionalNumerator >= fraction || !Number.isInteger(fractionalNumerator)) {
      return valueError();
    }
    const sign = value < 0 ? -1 : 1;
    return numberResult(sign * (integerPart + fractionalNumerator / fraction));
  },
  DOLLARFR: (valueArg, fractionArg) => {
    const value = toNumber(valueArg);
    const fraction = toInteger(fractionArg);
    if (value === undefined || fraction === undefined || !isValidDollarFraction(fraction)) {
      return valueError();
    }
    const sign = value < 0 ? -1 : 1;
    const absolute = Math.abs(value);
    const integerPart = Math.floor(absolute);
    const fractional = absolute - integerPart;
    const width = countLeadingZeros(fraction);
    const scaledNumerator = Math.round(fractional * fraction);
    const carry = Math.floor(scaledNumerator / fraction);
    const numerator = scaledNumerator - carry * fraction;
    const outputValue = `${integerPart + carry}.${String(numerator).padStart(width, "0")}`;
    return numberResult(sign * Number(outputValue));
  },
  GEOMEAN: (...args) => {
    const error = firstError(args);
    if (error) return error;
    const numbers = collectNumericArgs(args);
    if (numbers.length === 0) {
      return valueError();
    }
    if (numbers.some((value) => value < 0)) {
      return valueError();
    }
    if (numbers.some((value) => value === 0)) {
      return numberResult(0);
    }
    const logSum = numbers.reduce((sum, value) => sum + Math.log(value), 0);
    return numberResult(Math.exp(logSum / numbers.length));
  },
  HARMEAN: (...args) => {
    const error = firstError(args);
    if (error) return error;
    const numbers = collectNumericArgs(args);
    if (numbers.length === 0 || numbers.some((value) => value <= 0)) {
      return valueError();
    }
    return numberResult(numbers.length / numbers.reduce((sum, value) => sum + 1 / value, 0));
  },
  SIN: (value) => unaryMath(value, Math.sin),
  COS: (value) => unaryMath(value, Math.cos),
  TAN: (value) => unaryMath(value, Math.tan),
  ASIN: (value) => unaryMath(value, Math.asin),
  ACOS: (value) => unaryMath(value, Math.acos),
  ATAN: (value) => unaryMath(value, Math.atan),
  ATAN2: (left, right) => binaryMath(left, right, Math.atan2),
  DEGREES: (value) => unaryMath(value, (numeric) => (numeric * 180) / Math.PI),
  RADIANS: (value) => unaryMath(value, (numeric) => (numeric * Math.PI) / 180),
  EXP: (value) => unaryMath(value, Math.exp),
  LN: (value) => unaryMath(value, Math.log),
  LOG10: (value) => unaryMath(value, Math.log10),
  LOG: (value, base) => {
    const numeric = toNumber(value) ?? 0;
    const baseValue = base === undefined ? 10 : (toNumber(base) ?? 10);
    const result =
      base === undefined ? Math.log10(numeric) : Math.log(numeric) / Math.log(baseValue);
    return numericResultOrError(result);
  },
  POWER: (base, exponent) => binaryMath(base, exponent, Math.pow),
  SQRT: (value) => unaryMath(value, Math.sqrt),
  PI: () => numberResult(Math.PI),
  SINH: (value) => unaryMath(value, Math.sinh),
  COSH: (value) => unaryMath(value, Math.cosh),
  TANH: (value) => unaryMath(value, Math.tanh),
  ASINH: (value) => unaryMath(value, Math.asinh),
  ACOSH: (value) => unaryMath(value, Math.acosh),
  ATANH: (value) => unaryMath(value, Math.atanh),
  ACOT: (value) => {
    const numeric = toNumber(value) ?? 0;
    return numberResult(numeric === 0 ? Math.PI / 2 : Math.atan(1 / numeric));
  },
  ACOTH: (value) => {
    const numeric = toNumber(value) ?? 0;
    return numericResultOrError(0.5 * Math.log((numeric + 1) / (numeric - 1)));
  },
  COT: (value) => {
    const tangent = Math.tan(toNumber(value) ?? 0);
    return tangent === 0
      ? { tag: ValueTag.Error, code: ErrorCode.Div0 }
      : numberResult(1 / tangent);
  },
  COTH: (value) => {
    const hyperbolic = Math.tanh(toNumber(value) ?? 0);
    return hyperbolic === 0
      ? { tag: ValueTag.Error, code: ErrorCode.Div0 }
      : numberResult(1 / hyperbolic);
  },
  CSC: (value) => {
    const sine = Math.sin(toNumber(value) ?? 0);
    return sine === 0 ? { tag: ValueTag.Error, code: ErrorCode.Div0 } : numberResult(1 / sine);
  },
  CSCH: (value) => {
    const hyperbolic = Math.sinh(toNumber(value) ?? 0);
    return hyperbolic === 0
      ? { tag: ValueTag.Error, code: ErrorCode.Div0 }
      : numberResult(1 / hyperbolic);
  },
  SEC: (value) => {
    const cosine = Math.cos(toNumber(value) ?? 0);
    return cosine === 0 ? { tag: ValueTag.Error, code: ErrorCode.Div0 } : numberResult(1 / cosine);
  },
  SECH: (value) => unaryMath(value, (numeric) => 1 / Math.cosh(numeric)),
  SIGN: (value) => {
    const numeric = toNumber(value);
    if (numeric === undefined) {
      return valueError();
    }
    return numberResult(numeric === 0 ? 0 : numeric > 0 ? 1 : -1);
  },
  ROUND: (value, digits) => roundWith(value, digits),
  FLOOR: (value, significance) => floorWith(value, significance),
  CEILING: (value, significance) => ceilingWith(value, significance),
  "FLOOR.MATH": (value, significance, mode) => {
    const numberValue = toNumber(value);
    const significanceValue = Math.abs(
      toNumber(significance ?? { tag: ValueTag.Number, value: 1 }) ?? 1,
    );
    const modeValue = toNumber(mode ?? { tag: ValueTag.Number, value: 0 }) ?? 0;
    if (numberValue === undefined || significanceValue === 0) {
      return valueError();
    }
    if (numberValue >= 0) {
      return numberResult(Math.floor(numberValue / significanceValue) * significanceValue);
    }
    const magnitude =
      modeValue === 0
        ? Math.ceil(Math.abs(numberValue) / significanceValue)
        : Math.floor(Math.abs(numberValue) / significanceValue);
    return numberResult(-magnitude * significanceValue);
  },
  "FLOOR.PRECISE": (value, significance) => {
    const numberValue = toNumber(value);
    const significanceValue = Math.abs(
      toNumber(significance ?? { tag: ValueTag.Number, value: 1 }) ?? 1,
    );
    if (numberValue === undefined || significanceValue === 0) {
      return valueError();
    }
    return numberResult(Math.floor(numberValue / significanceValue) * significanceValue);
  },
  "CEILING.MATH": (value, significance, mode) => {
    const numberValue = toNumber(value);
    const significanceValue = Math.abs(
      toNumber(significance ?? { tag: ValueTag.Number, value: 1 }) ?? 1,
    );
    const modeValue = toNumber(mode ?? { tag: ValueTag.Number, value: 0 }) ?? 0;
    if (numberValue === undefined || significanceValue === 0) {
      return valueError();
    }
    if (numberValue >= 0) {
      return numberResult(Math.ceil(numberValue / significanceValue) * significanceValue);
    }
    const magnitude =
      modeValue === 0
        ? Math.floor(Math.abs(numberValue) / significanceValue)
        : Math.ceil(Math.abs(numberValue) / significanceValue);
    return numberResult(-magnitude * significanceValue);
  },
  "CEILING.PRECISE": (value, significance) => {
    const numberValue = toNumber(value);
    const significanceValue = Math.abs(
      toNumber(significance ?? { tag: ValueTag.Number, value: 1 }) ?? 1,
    );
    if (numberValue === undefined || significanceValue === 0) {
      return valueError();
    }
    return numberResult(Math.ceil(numberValue / significanceValue) * significanceValue);
  },
  "ISO.CEILING": (value, significance) => {
    const numberValue = toNumber(value);
    const significanceValue = Math.abs(
      toNumber(significance ?? { tag: ValueTag.Number, value: 1 }) ?? 1,
    );
    if (numberValue === undefined || significanceValue === 0) {
      return valueError();
    }
    return numberResult(Math.ceil(numberValue / significanceValue) * significanceValue);
  },
  MOD: (left, right) => {
    const divisor = toNumber(right) ?? 0;
    if (divisor === 0) return { tag: ValueTag.Error, code: ErrorCode.Div0 };
    return numberResult((toNumber(left) ?? 0) % divisor);
  },
  BITAND: (...args) => {
    if (args.length < 2) {
      return valueError();
    }
    let value = toBitwiseUnsigned(args[0]);
    if (value === undefined) {
      return valueError();
    }
    for (let index = 1; index < args.length; index += 1) {
      const current = toBitwiseUnsigned(args[index]);
      if (current === undefined) {
        return valueError();
      }
      value &= current;
    }
    return numberResult(value >>> 0);
  },
  BITOR: (...args) => {
    if (args.length < 2) {
      return valueError();
    }
    let value = toBitwiseUnsigned(args[0]);
    if (value === undefined) {
      return valueError();
    }
    for (let index = 1; index < args.length; index += 1) {
      const current = toBitwiseUnsigned(args[index]);
      if (current === undefined) {
        return valueError();
      }
      value |= current;
    }
    return numberResult(value >>> 0);
  },
  BITXOR: (...args) => {
    if (args.length < 2) {
      return valueError();
    }
    let value = toBitwiseUnsigned(args[0]);
    if (value === undefined) {
      return valueError();
    }
    for (let index = 1; index < args.length; index += 1) {
      const current = toBitwiseUnsigned(args[index]);
      if (current === undefined) {
        return valueError();
      }
      value ^= current;
    }
    return numberResult(value >>> 0);
  },
  BITLSHIFT: (valueArg, shiftArg) => {
    const value = toBitwiseUnsigned(valueArg);
    const shift = coerceShiftAmount(shiftArg);
    if (value === undefined || shift === undefined) {
      return valueError();
    }
    return numberResult((value << (shift & 31)) >>> 0);
  },
  BITRSHIFT: (valueArg, shiftArg) => {
    const value = toBitwiseUnsigned(valueArg);
    const shift = coerceShiftAmount(shiftArg);
    if (value === undefined || shift === undefined) {
      return valueError();
    }
    return numberResult((value >>> (shift & 31)) >>> 0);
  },
  BESSELI: (xArg, orderArg) => {
    const x = toNumber(xArg);
    const order = integerValue(orderArg);
    if (x === undefined || order === undefined) {
      return valueError();
    }
    if (order < 0) {
      return numError();
    }
    const result = besselIValue(x, order);
    return Number.isFinite(result) ? numberResult(result) : numError();
  },
  BESSELJ: (xArg, orderArg) => {
    const x = toNumber(xArg);
    const order = integerValue(orderArg);
    if (x === undefined || order === undefined) {
      return valueError();
    }
    if (order < 0) {
      return numError();
    }
    const result = besselJValue(x, order);
    return Number.isFinite(result) ? numberResult(result) : numError();
  },
  BESSELK: (xArg, orderArg) => {
    const x = toNumber(xArg);
    const order = integerValue(orderArg);
    if (x === undefined || order === undefined) {
      return valueError();
    }
    if (x <= 0 || order < 0) {
      return numError();
    }
    const result = besselKValue(x, order);
    return Number.isFinite(result) ? numberResult(result) : numError();
  },
  BESSELY: (xArg, orderArg) => {
    const x = toNumber(xArg);
    const order = integerValue(orderArg);
    if (x === undefined || order === undefined) {
      return valueError();
    }
    if (x <= 0 || order < 0) {
      return numError();
    }
    const result = besselYValue(x, order);
    return Number.isFinite(result) ? numberResult(result) : numError();
  },
  INT: (value) => {
    const numberValue = toNumber(value);
    if (numberValue === undefined) {
      return valueError();
    }
    return numberResult(Math.floor(numberValue));
  },
  ROUNDUP: (value, digits) => {
    const numberValue = toNumber(value);
    const digitValue = digits === undefined ? 0 : toNumber(digits);
    if (numberValue === undefined || digitValue === undefined) {
      return valueError();
    }
    return numberResult(roundUpToDigits(numberValue, Math.trunc(digitValue)));
  },
  ROUNDDOWN: (value, digits) => {
    const numberValue = toNumber(value);
    const digitValue = digits === undefined ? 0 : toNumber(digits);
    if (numberValue === undefined || digitValue === undefined) {
      return valueError();
    }
    return numberResult(roundDownToDigits(numberValue, Math.trunc(digitValue)));
  },
  TRUNC: (value, digits) => {
    const numberValue = toNumber(value);
    const digitValue = digits === undefined ? 0 : toNumber(digits);
    if (numberValue === undefined || digitValue === undefined) {
      return valueError();
    }
    return numberResult(roundTowardZero(numberValue, Math.trunc(digitValue)));
  },
  EVEN: (value) => {
    const numberValue = toNumber(value);
    return numberValue === undefined ? valueError() : numberResult(evenValue(numberValue));
  },
  ODD: (value) => {
    const numberValue = toNumber(value);
    return numberValue === undefined ? valueError() : numberResult(oddValue(numberValue));
  },
  FACT: (value) => {
    const factorial = factorialValue(toNumber(value) ?? Number.NaN);
    return factorial === undefined ? valueError() : numberResult(factorial);
  },
  FACTDOUBLE: (value) => {
    const factorial = doubleFactorialValue(toNumber(value) ?? Number.NaN);
    return factorial === undefined ? valueError() : numberResult(factorial);
  },
  COMBIN: (numberArg, chosenArg) => {
    const numberValue = nonNegativeIntegerValue(numberArg);
    const chosenValue = nonNegativeIntegerValue(chosenArg);
    if (numberValue === undefined || chosenValue === undefined || chosenValue > numberValue) {
      return valueError();
    }
    const numerator = factorialValue(numberValue);
    const denominator = factorialValue(chosenValue);
    const remainder = factorialValue(numberValue - chosenValue);
    return numerator === undefined || denominator === undefined || remainder === undefined
      ? valueError()
      : numberResult(numerator / (denominator * remainder));
  },
  COMBINA: (numberArg, chosenArg) => {
    const numberValue = nonNegativeIntegerValue(numberArg);
    const chosenValue = nonNegativeIntegerValue(chosenArg);
    if (numberValue === undefined || chosenValue === undefined) {
      return valueError();
    }
    if (chosenValue === 0) {
      return numberResult(1);
    }
    if (numberValue === 0) {
      return numberResult(0);
    }
    const combined = numberValue + chosenValue - 1;
    const numerator = factorialValue(combined);
    const denominator = factorialValue(chosenValue);
    const remainder = factorialValue(numberValue - 1);
    return numerator === undefined || denominator === undefined || remainder === undefined
      ? valueError()
      : numberResult(numerator / (denominator * remainder));
  },
  GCD: (...args) => {
    const numbers = collectNumericArgs(args).map((value) => Math.abs(Math.trunc(value)));
    if (numbers.length === 0) {
      return valueError();
    }
    return numberResult(numbers.reduce((acc, value) => gcdPair(acc, value)));
  },
  LCM: (...args) => {
    const numbers = collectNumericArgs(args).map((value) => Math.abs(Math.trunc(value)));
    if (numbers.length === 0) {
      return valueError();
    }
    return numberResult(numbers.reduce((acc, value) => lcmPair(acc, value)));
  },
  MROUND: (value, multiple) => {
    const numberValue = toNumber(value);
    const multipleValue = toNumber(multiple);
    if (numberValue === undefined || multipleValue === undefined || multipleValue === 0) {
      return valueError();
    }
    if (numberValue !== 0 && Math.sign(numberValue) !== Math.sign(multipleValue)) {
      return valueError();
    }
    return numberResult(Math.round(numberValue / multipleValue) * multipleValue);
  },
  MULTINOMIAL: (...args) => {
    const numbers = collectNumericArgs(args).map((value) => Math.trunc(value));
    if (numbers.some((value) => value < 0)) {
      return valueError();
    }
    const numerator = factorialValue(numbers.reduce((sum, value) => sum + value, 0));
    const denominator = numbers.reduce(
      (product, value) => product * (factorialValue(value) ?? Number.NaN),
      1,
    );
    return numerator === undefined || Number.isNaN(denominator)
      ? valueError()
      : numberResult(numerator / denominator);
  },
  PRODUCT: (...args) => {
    const error = firstError(args);
    if (error) return error;
    const numbers = collectNumericArgs(args);
    return numberResult(
      numbers.length === 0 ? 0 : numbers.reduce((product, value) => product * value, 1),
    );
  },
  QUOTIENT: (numeratorArg, denominatorArg) => {
    const numerator = toNumber(numeratorArg);
    const denominator = toNumber(denominatorArg);
    if (numerator === undefined || denominator === undefined) {
      return valueError();
    }
    if (denominator === 0) {
      return { tag: ValueTag.Error, code: ErrorCode.Div0 };
    }
    return numberResult(Math.trunc(numerator / denominator));
  },
  ACCRINT: (
    issueArg,
    firstInterestArg,
    settlementArg,
    rateArg,
    parArg,
    frequencyArg,
    basisArg,
    calcMethodArg,
  ) => {
    const issue = coerceDateSerial(issueArg);
    const firstInterest = coerceDateSerial(firstInterestArg);
    const settlement = coerceDateSerial(settlementArg);
    const rate = toNumber(rateArg);
    const par = coerceNumber(parArg, 1000);
    const frequency = integerValue(frequencyArg);
    const basis = integerValue(basisArg, 0);
    const calcMethod = coerceBoolean(calcMethodArg, true);
    if (
      issue === undefined ||
      firstInterest === undefined ||
      settlement === undefined ||
      rate === undefined ||
      frequency === undefined ||
      basis === undefined ||
      calcMethod === undefined ||
      par === undefined ||
      rate <= 0 ||
      par <= 0 ||
      issue >= settlement ||
      firstInterest <= issue ||
      firstInterest >= settlement ||
      !isValidFrequency(frequency) ||
      !isValidBasis(basis)
    ) {
      return valueError();
    }
    const accrualStart = settlement > firstInterest && !calcMethod ? firstInterest : issue;
    const years = yearFracByBasis(accrualStart, settlement, basis);
    return years === undefined ? valueError() : numberResult(par * rate * years);
  },
  ACCRINTM: (issueArg, settlementArg, rateArg, parArg, basisArg) => {
    const issue = coerceDateSerial(issueArg);
    const settlement = coerceDateSerial(settlementArg);
    const rate = toNumber(rateArg);
    const par = coerceNumber(parArg, 1000);
    const basis = integerValue(basisArg, 0);
    if (
      issue === undefined ||
      settlement === undefined ||
      rate === undefined ||
      basis === undefined ||
      par === undefined ||
      rate <= 0 ||
      par <= 0 ||
      issue >= settlement ||
      !isValidBasis(basis)
    ) {
      return valueError();
    }
    const years = yearFracByBasis(issue, settlement, basis);
    return years === undefined ? valueError() : numberResult(par * rate * years);
  },
  COUPDAYBS: (settlementArg, maturityArg, frequencyArg, basisArg) => {
    const settlement = coerceDateSerial(settlementArg);
    const maturity = coerceDateSerial(maturityArg);
    const frequency = integerValue(frequencyArg);
    const basis = integerValue(basisArg, 0);
    if (
      settlement === undefined ||
      maturity === undefined ||
      frequency === undefined ||
      basis === undefined ||
      !isValidFrequency(frequency) ||
      !isValidBasis(basis)
    ) {
      return valueError();
    }
    const metrics = couponMetrics(settlement, maturity, frequency, basis);
    return metrics === undefined ? valueError() : numberResult(metrics.accruedDays);
  },
  COUPDAYS: (settlementArg, maturityArg, frequencyArg, basisArg) => {
    const settlement = coerceDateSerial(settlementArg);
    const maturity = coerceDateSerial(maturityArg);
    const frequency = integerValue(frequencyArg);
    const basis = integerValue(basisArg, 0);
    if (
      settlement === undefined ||
      maturity === undefined ||
      frequency === undefined ||
      basis === undefined ||
      !isValidFrequency(frequency) ||
      !isValidBasis(basis)
    ) {
      return valueError();
    }
    const metrics = couponMetrics(settlement, maturity, frequency, basis);
    return metrics === undefined ? valueError() : numberResult(metrics.daysInPeriod);
  },
  COUPDAYSNC: (settlementArg, maturityArg, frequencyArg, basisArg) => {
    const settlement = coerceDateSerial(settlementArg);
    const maturity = coerceDateSerial(maturityArg);
    const frequency = integerValue(frequencyArg);
    const basis = integerValue(basisArg, 0);
    if (
      settlement === undefined ||
      maturity === undefined ||
      frequency === undefined ||
      basis === undefined ||
      !isValidFrequency(frequency) ||
      !isValidBasis(basis)
    ) {
      return valueError();
    }
    const metrics = couponMetrics(settlement, maturity, frequency, basis);
    return metrics === undefined ? valueError() : numberResult(metrics.daysToNextCoupon);
  },
  COUPNCD: (settlementArg, maturityArg, frequencyArg, basisArg) => {
    const settlement = coerceDateSerial(settlementArg);
    const maturity = coerceDateSerial(maturityArg);
    const frequency = integerValue(frequencyArg);
    const basis = integerValue(basisArg, 0);
    if (
      settlement === undefined ||
      maturity === undefined ||
      frequency === undefined ||
      basis === undefined ||
      !isValidFrequency(frequency) ||
      !isValidBasis(basis)
    ) {
      return valueError();
    }
    const schedule = couponSchedule(settlement, maturity, frequency);
    return schedule === undefined ? valueError() : numberResult(schedule.nextCoupon);
  },
  COUPNUM: (settlementArg, maturityArg, frequencyArg, basisArg) => {
    const settlement = coerceDateSerial(settlementArg);
    const maturity = coerceDateSerial(maturityArg);
    const frequency = integerValue(frequencyArg);
    const basis = integerValue(basisArg, 0);
    if (
      settlement === undefined ||
      maturity === undefined ||
      frequency === undefined ||
      basis === undefined ||
      !isValidFrequency(frequency) ||
      !isValidBasis(basis)
    ) {
      return valueError();
    }
    const schedule = couponSchedule(settlement, maturity, frequency);
    return schedule === undefined ? valueError() : numberResult(schedule.periodsRemaining);
  },
  COUPPCD: (settlementArg, maturityArg, frequencyArg, basisArg) => {
    const settlement = coerceDateSerial(settlementArg);
    const maturity = coerceDateSerial(maturityArg);
    const frequency = integerValue(frequencyArg);
    const basis = integerValue(basisArg, 0);
    if (
      settlement === undefined ||
      maturity === undefined ||
      frequency === undefined ||
      basis === undefined ||
      !isValidFrequency(frequency) ||
      !isValidBasis(basis)
    ) {
      return valueError();
    }
    const schedule = couponSchedule(settlement, maturity, frequency);
    return schedule === undefined ? valueError() : numberResult(schedule.previousCoupon);
  },
  DISC: (settlementArg, maturityArg, priceArg, redemptionArg, basisArg) => {
    const settlement = coerceDateSerial(settlementArg);
    const maturity = coerceDateSerial(maturityArg);
    const price = toNumber(priceArg);
    const redemption = toNumber(redemptionArg);
    const basis = integerValue(basisArg, 0);
    if (
      settlement === undefined ||
      maturity === undefined ||
      price === undefined ||
      redemption === undefined ||
      basis === undefined ||
      price <= 0 ||
      redemption <= 0
    ) {
      return valueError();
    }
    const years = securityAnnualizedYearFraction(settlement, maturity, basis);
    return years === undefined
      ? valueError()
      : numberResult((redemption - price) / redemption / years);
  },
  INTRATE: (settlementArg, maturityArg, investmentArg, redemptionArg, basisArg) => {
    const settlement = coerceDateSerial(settlementArg);
    const maturity = coerceDateSerial(maturityArg);
    const investment = toNumber(investmentArg);
    const redemption = toNumber(redemptionArg);
    const basis = integerValue(basisArg, 0);
    if (
      settlement === undefined ||
      maturity === undefined ||
      investment === undefined ||
      redemption === undefined ||
      basis === undefined ||
      investment <= 0 ||
      redemption <= 0
    ) {
      return valueError();
    }
    const years = securityAnnualizedYearFraction(settlement, maturity, basis);
    return years === undefined
      ? valueError()
      : numberResult((redemption - investment) / investment / years);
  },
  RECEIVED: (settlementArg, maturityArg, investmentArg, discountArg, basisArg) => {
    const settlement = coerceDateSerial(settlementArg);
    const maturity = coerceDateSerial(maturityArg);
    const investment = toNumber(investmentArg);
    const discount = toNumber(discountArg);
    const basis = integerValue(basisArg, 0);
    if (
      settlement === undefined ||
      maturity === undefined ||
      investment === undefined ||
      discount === undefined ||
      basis === undefined ||
      investment <= 0 ||
      discount <= 0
    ) {
      return valueError();
    }
    const years = securityAnnualizedYearFraction(settlement, maturity, basis);
    if (years === undefined) {
      return valueError();
    }
    const denominator = 1 - discount * years;
    return denominator <= 0 ? valueError() : numberResult(investment / denominator);
  },
  PRICEDISC: (settlementArg, maturityArg, discountArg, redemptionArg, basisArg) => {
    const settlement = coerceDateSerial(settlementArg);
    const maturity = coerceDateSerial(maturityArg);
    const discount = toNumber(discountArg);
    const redemption = toNumber(redemptionArg);
    const basis = integerValue(basisArg, 0);
    if (
      settlement === undefined ||
      maturity === undefined ||
      discount === undefined ||
      redemption === undefined ||
      basis === undefined ||
      discount <= 0 ||
      redemption <= 0
    ) {
      return valueError();
    }
    const years = securityAnnualizedYearFraction(settlement, maturity, basis);
    return years === undefined ? valueError() : numberResult(redemption * (1 - discount * years));
  },
  YIELDDISC: (settlementArg, maturityArg, priceArg, redemptionArg, basisArg) => {
    const settlement = coerceDateSerial(settlementArg);
    const maturity = coerceDateSerial(maturityArg);
    const price = toNumber(priceArg);
    const redemption = toNumber(redemptionArg);
    const basis = integerValue(basisArg, 0);
    if (
      settlement === undefined ||
      maturity === undefined ||
      price === undefined ||
      redemption === undefined ||
      basis === undefined ||
      price <= 0 ||
      redemption <= 0
    ) {
      return valueError();
    }
    const years = securityAnnualizedYearFraction(settlement, maturity, basis);
    return years === undefined ? valueError() : numberResult((redemption - price) / price / years);
  },
  TBILLPRICE: (settlementArg, maturityArg, discountArg) => {
    const settlement = coerceDateSerial(settlementArg);
    const maturity = coerceDateSerial(maturityArg);
    const discount = toNumber(discountArg);
    if (
      settlement === undefined ||
      maturity === undefined ||
      discount === undefined ||
      discount <= 0
    ) {
      return valueError();
    }
    const days = treasuryBillDays(settlement, maturity);
    return days === undefined ? valueError() : numberResult(100 * (1 - (discount * days) / 360));
  },
  TBILLYIELD: (settlementArg, maturityArg, priceArg) => {
    const settlement = coerceDateSerial(settlementArg);
    const maturity = coerceDateSerial(maturityArg);
    const price = toNumber(priceArg);
    if (settlement === undefined || maturity === undefined || price === undefined || price <= 0) {
      return valueError();
    }
    const days = treasuryBillDays(settlement, maturity);
    return days === undefined ? valueError() : numberResult(((100 - price) * 360) / (price * days));
  },
  TBILLEQ: (settlementArg, maturityArg, discountArg) => {
    const settlement = coerceDateSerial(settlementArg);
    const maturity = coerceDateSerial(maturityArg);
    const discount = toNumber(discountArg);
    if (
      settlement === undefined ||
      maturity === undefined ||
      discount === undefined ||
      discount <= 0
    ) {
      return valueError();
    }
    const days = treasuryBillDays(settlement, maturity);
    if (days === undefined) {
      return valueError();
    }
    const denominator = 360 - discount * days;
    return denominator === 0 ? valueError() : numberResult((365 * discount) / denominator);
  },
  PRICEMAT: (settlementArg, maturityArg, issueArg, rateArg, yieldArg, basisArg) => {
    const settlement = coerceDateSerial(settlementArg);
    const maturity = coerceDateSerial(maturityArg);
    const issue = coerceDateSerial(issueArg);
    const rate = toNumber(rateArg);
    const yieldRate = toNumber(yieldArg);
    const basis = integerValue(basisArg, 0);
    if (
      settlement === undefined ||
      maturity === undefined ||
      issue === undefined ||
      rate === undefined ||
      yieldRate === undefined ||
      basis === undefined ||
      rate < 0 ||
      yieldRate < 0
    ) {
      return valueError();
    }
    const fractions = maturityAtIssueFractions(settlement, maturity, issue, basis);
    if (fractions === undefined) {
      return valueError();
    }
    const maturityValue = 100 * (1 + rate * fractions.issueToMaturity);
    const accruedInterest = 100 * rate * fractions.issueToSettlement;
    const denominator = 1 + yieldRate * fractions.settlementToMaturity;
    return numberResult(maturityValue / denominator - accruedInterest);
  },
  YIELDMAT: (settlementArg, maturityArg, issueArg, rateArg, priceArg, basisArg) => {
    const settlement = coerceDateSerial(settlementArg);
    const maturity = coerceDateSerial(maturityArg);
    const issue = coerceDateSerial(issueArg);
    const rate = toNumber(rateArg);
    const price = toNumber(priceArg);
    const basis = integerValue(basisArg, 0);
    if (
      settlement === undefined ||
      maturity === undefined ||
      issue === undefined ||
      rate === undefined ||
      price === undefined ||
      basis === undefined ||
      rate < 0 ||
      price <= 0
    ) {
      return valueError();
    }
    const fractions = maturityAtIssueFractions(settlement, maturity, issue, basis);
    if (fractions === undefined) {
      return valueError();
    }
    const settlementValue = price + 100 * rate * fractions.issueToSettlement;
    const maturityValue = 100 * (1 + rate * fractions.issueToMaturity);
    return numberResult((maturityValue / settlementValue - 1) / fractions.settlementToMaturity);
  },
  ODDFPRICE: (
    settlementArg,
    maturityArg,
    issueArg,
    firstCouponArg,
    rateArg,
    yieldArg,
    redemptionArg,
    frequencyArg,
    basisArg,
  ) => {
    const settlement = coerceDateSerial(settlementArg);
    const maturity = coerceDateSerial(maturityArg);
    const issue = coerceDateSerial(issueArg);
    const firstCoupon = coerceDateSerial(firstCouponArg);
    const rate = toNumber(rateArg);
    const yieldRate = toNumber(yieldArg);
    const redemption = toNumber(redemptionArg);
    const frequency = integerValue(frequencyArg);
    const basis = integerValue(basisArg, 0);
    if (
      settlement === undefined ||
      maturity === undefined ||
      issue === undefined ||
      firstCoupon === undefined ||
      rate === undefined ||
      yieldRate === undefined ||
      redemption === undefined ||
      frequency === undefined ||
      basis === undefined ||
      rate < 0 ||
      yieldRate < 0 ||
      redemption <= 0 ||
      !isValidFrequency(frequency) ||
      !isValidBasis(basis)
    ) {
      return valueError();
    }
    const price = oddFirstPriceValue(
      settlement,
      maturity,
      issue,
      firstCoupon,
      rate,
      yieldRate,
      redemption,
      frequency,
      basis,
    );
    return price === undefined ? valueError() : numberResult(price);
  },
  ODDFYIELD: (
    settlementArg,
    maturityArg,
    issueArg,
    firstCouponArg,
    rateArg,
    priceArg,
    redemptionArg,
    frequencyArg,
    basisArg,
  ) => {
    const settlement = coerceDateSerial(settlementArg);
    const maturity = coerceDateSerial(maturityArg);
    const issue = coerceDateSerial(issueArg);
    const firstCoupon = coerceDateSerial(firstCouponArg);
    const rate = toNumber(rateArg);
    const price = toNumber(priceArg);
    const redemption = toNumber(redemptionArg);
    const frequency = integerValue(frequencyArg);
    const basis = integerValue(basisArg, 0);
    if (
      settlement === undefined ||
      maturity === undefined ||
      issue === undefined ||
      firstCoupon === undefined ||
      rate === undefined ||
      price === undefined ||
      redemption === undefined ||
      frequency === undefined ||
      basis === undefined ||
      rate < 0 ||
      price <= 0 ||
      redemption <= 0 ||
      !isValidFrequency(frequency) ||
      !isValidBasis(basis)
    ) {
      return valueError();
    }
    const yieldRate = solveOddFirstYield(
      settlement,
      maturity,
      issue,
      firstCoupon,
      rate,
      price,
      redemption,
      frequency,
      basis,
    );
    return yieldRate === undefined ? valueError() : numberResult(yieldRate);
  },
  ODDLPRICE: (
    settlementArg,
    maturityArg,
    lastInterestArg,
    rateArg,
    yieldArg,
    redemptionArg,
    frequencyArg,
    basisArg,
  ) => {
    const settlement = coerceDateSerial(settlementArg);
    const maturity = coerceDateSerial(maturityArg);
    const lastInterest = coerceDateSerial(lastInterestArg);
    const rate = toNumber(rateArg);
    const yieldRate = toNumber(yieldArg);
    const redemption = toNumber(redemptionArg);
    const frequency = integerValue(frequencyArg);
    const basis = integerValue(basisArg, 0);
    if (
      settlement === undefined ||
      maturity === undefined ||
      lastInterest === undefined ||
      rate === undefined ||
      yieldRate === undefined ||
      redemption === undefined ||
      frequency === undefined ||
      basis === undefined ||
      rate < 0 ||
      yieldRate < 0 ||
      redemption <= 0 ||
      !isValidFrequency(frequency) ||
      !isValidBasis(basis)
    ) {
      return valueError();
    }

    const fractions = oddLastCouponFractions(settlement, maturity, lastInterest, frequency, basis);
    if (fractions === undefined) {
      return valueError();
    }

    const coupon = (100 * rate) / frequency;
    const maturityValue = redemption + coupon * fractions.totalFraction;
    const denominator = 1 + (yieldRate * fractions.remainingFraction) / frequency;
    if (denominator <= 0) {
      return valueError();
    }
    return numberResult(maturityValue / denominator - coupon * fractions.accruedFraction);
  },
  ODDLYIELD: (
    settlementArg,
    maturityArg,
    lastInterestArg,
    rateArg,
    priceArg,
    redemptionArg,
    frequencyArg,
    basisArg,
  ) => {
    const settlement = coerceDateSerial(settlementArg);
    const maturity = coerceDateSerial(maturityArg);
    const lastInterest = coerceDateSerial(lastInterestArg);
    const rate = toNumber(rateArg);
    const price = toNumber(priceArg);
    const redemption = toNumber(redemptionArg);
    const frequency = integerValue(frequencyArg);
    const basis = integerValue(basisArg, 0);
    if (
      settlement === undefined ||
      maturity === undefined ||
      lastInterest === undefined ||
      rate === undefined ||
      price === undefined ||
      redemption === undefined ||
      frequency === undefined ||
      basis === undefined ||
      rate < 0 ||
      price <= 0 ||
      redemption <= 0 ||
      !isValidFrequency(frequency) ||
      !isValidBasis(basis)
    ) {
      return valueError();
    }

    const fractions = oddLastCouponFractions(settlement, maturity, lastInterest, frequency, basis);
    if (fractions === undefined) {
      return valueError();
    }

    const coupon = (100 * rate) / frequency;
    const dirtyPrice = price + coupon * fractions.accruedFraction;
    if (dirtyPrice <= 0 || fractions.remainingFraction <= 0) {
      return valueError();
    }
    const maturityValue = redemption + coupon * fractions.totalFraction;
    return numberResult(
      ((maturityValue - dirtyPrice) / dirtyPrice) * (frequency / fractions.remainingFraction),
    );
  },
  PRICE: (settlementArg, maturityArg, rateArg, yieldArg, redemptionArg, frequencyArg, basisArg) => {
    const settlement = coerceDateSerial(settlementArg);
    const maturity = coerceDateSerial(maturityArg);
    const rate = toNumber(rateArg);
    const yieldRate = toNumber(yieldArg);
    const redemption = toNumber(redemptionArg);
    const frequency = integerValue(frequencyArg);
    const basis = integerValue(basisArg, 0);
    if (
      settlement === undefined ||
      maturity === undefined ||
      rate === undefined ||
      yieldRate === undefined ||
      redemption === undefined ||
      frequency === undefined ||
      basis === undefined ||
      rate < 0 ||
      yieldRate < 0 ||
      redemption <= 0 ||
      !isValidFrequency(frequency) ||
      !isValidBasis(basis)
    ) {
      return valueError();
    }
    const metrics = couponMetrics(settlement, maturity, frequency, basis);
    const price = metrics
      ? pricePeriodicSecurity(metrics, rate, yieldRate, redemption, frequency)
      : undefined;
    return price === undefined ? valueError() : numberResult(price);
  },
  YIELD: (settlementArg, maturityArg, rateArg, priceArg, redemptionArg, frequencyArg, basisArg) => {
    const settlement = coerceDateSerial(settlementArg);
    const maturity = coerceDateSerial(maturityArg);
    const rate = toNumber(rateArg);
    const price = toNumber(priceArg);
    const redemption = toNumber(redemptionArg);
    const frequency = integerValue(frequencyArg);
    const basis = integerValue(basisArg, 0);
    if (
      settlement === undefined ||
      maturity === undefined ||
      rate === undefined ||
      price === undefined ||
      redemption === undefined ||
      frequency === undefined ||
      basis === undefined ||
      rate < 0 ||
      price <= 0 ||
      redemption <= 0 ||
      !isValidFrequency(frequency) ||
      !isValidBasis(basis)
    ) {
      return valueError();
    }
    const metrics = couponMetrics(settlement, maturity, frequency, basis);
    const yieldRate = metrics
      ? solvePeriodicSecurityYield(metrics, rate, price, redemption, frequency)
      : undefined;
    return yieldRate === undefined ? valueError() : numberResult(yieldRate);
  },
  DURATION: (settlementArg, maturityArg, couponArg, yieldArg, frequencyArg, basisArg) => {
    const settlement = coerceDateSerial(settlementArg);
    const maturity = coerceDateSerial(maturityArg);
    const coupon = toNumber(couponArg);
    const yieldRate = toNumber(yieldArg);
    const frequency = integerValue(frequencyArg);
    const basis = integerValue(basisArg, 0);
    if (
      settlement === undefined ||
      maturity === undefined ||
      coupon === undefined ||
      yieldRate === undefined ||
      frequency === undefined ||
      basis === undefined ||
      coupon < 0 ||
      yieldRate < 0 ||
      !isValidFrequency(frequency) ||
      !isValidBasis(basis)
    ) {
      return valueError();
    }
    const metrics = couponMetrics(settlement, maturity, frequency, basis);
    const duration = metrics ? macaulayDuration(metrics, coupon, yieldRate, frequency) : undefined;
    return duration === undefined ? valueError() : numberResult(duration);
  },
  MDURATION: (settlementArg, maturityArg, couponArg, yieldArg, frequencyArg, basisArg) => {
    const settlement = coerceDateSerial(settlementArg);
    const maturity = coerceDateSerial(maturityArg);
    const coupon = toNumber(couponArg);
    const yieldRate = toNumber(yieldArg);
    const frequency = integerValue(frequencyArg);
    const basis = integerValue(basisArg, 0);
    if (
      settlement === undefined ||
      maturity === undefined ||
      coupon === undefined ||
      yieldRate === undefined ||
      frequency === undefined ||
      basis === undefined ||
      coupon < 0 ||
      yieldRate < 0 ||
      !isValidFrequency(frequency) ||
      !isValidBasis(basis)
    ) {
      return valueError();
    }
    const metrics = couponMetrics(settlement, maturity, frequency, basis);
    const duration = metrics ? macaulayDuration(metrics, coupon, yieldRate, frequency) : undefined;
    return duration === undefined
      ? valueError()
      : numberResult(duration / (1 + yieldRate / frequency));
  },
  AMORDEGRC: (
    costArg,
    datePurchasedArg,
    firstPeriodArg,
    salvageArg,
    periodArg,
    rateArg,
    basisArg,
  ) => {
    const cost = toNumber(costArg);
    const datePurchased = coerceDateSerial(datePurchasedArg);
    const firstPeriod = coerceDateSerial(firstPeriodArg);
    const salvage = toNumber(salvageArg);
    const period = toNumber(periodArg);
    const rate = toNumber(rateArg);
    const basis = integerValue(basisArg, 0);
    if (
      cost === undefined ||
      datePurchased === undefined ||
      firstPeriod === undefined ||
      salvage === undefined ||
      period === undefined ||
      rate === undefined ||
      basis === undefined ||
      !isValidBasis(basis)
    ) {
      return valueError();
    }
    const depreciation = getAmordegrc(
      cost,
      datePurchased,
      firstPeriod,
      salvage,
      period,
      rate,
      basis,
    );
    return depreciation === undefined ? valueError() : numberResult(depreciation);
  },
  AMORLINC: (
    costArg,
    datePurchasedArg,
    firstPeriodArg,
    salvageArg,
    periodArg,
    rateArg,
    basisArg,
  ) => {
    const cost = toNumber(costArg);
    const datePurchased = coerceDateSerial(datePurchasedArg);
    const firstPeriod = coerceDateSerial(firstPeriodArg);
    const salvage = toNumber(salvageArg);
    const period = toNumber(periodArg);
    const rate = toNumber(rateArg);
    const basis = integerValue(basisArg, 0);
    if (
      cost === undefined ||
      datePurchased === undefined ||
      firstPeriod === undefined ||
      salvage === undefined ||
      period === undefined ||
      rate === undefined ||
      basis === undefined ||
      !isValidBasis(basis)
    ) {
      return valueError();
    }
    const depreciation = getAmorlinc(
      cost,
      datePurchased,
      firstPeriod,
      salvage,
      period,
      rate,
      basis,
    );
    return depreciation === undefined ? valueError() : numberResult(depreciation);
  },
  RANDBETWEEN: (bottomArg, topArg) => {
    const bottom = integerValue(bottomArg);
    const top = integerValue(topArg);
    if (bottom === undefined || top === undefined || top < bottom) {
      return valueError();
    }
    return numberResult(Math.floor(Math.random() * (top - bottom + 1)) + bottom);
  },
  RANDARRAY: (rowsArg, colsArg, minArg, maxArg, wholeArg) => {
    const rows = coercePositiveInteger(rowsArg, 1);
    const cols = coercePositiveInteger(colsArg, 1);
    const min = coerceNumber(minArg, 0);
    const max = coerceNumber(maxArg, 1);
    const whole = wholeArg === undefined ? false : (toNumber(wholeArg) ?? 0) !== 0;
    if (
      rows === undefined ||
      cols === undefined ||
      min === undefined ||
      max === undefined ||
      max < min
    ) {
      return valueError();
    }
    const values: CellValue[] = [];
    for (let index = 0; index < rows * cols; index += 1) {
      const value = whole
        ? Math.floor(Math.random() * (Math.trunc(max) - Math.trunc(min) + 1)) + Math.trunc(min)
        : Math.random() * (max - min) + min;
      values.push(numberResult(value));
    }
    return { kind: "array", rows, cols, values };
  },
  MUNIT: (sizeArg) => {
    const size = positiveIntegerValue(sizeArg);
    return size === undefined ? valueError() : buildIdentityMatrix(size);
  },
  SERIESSUM: (xArg, nArg, mArg, ...coefficientArgs) => {
    const x = toNumber(xArg);
    const n = integerValue(nArg);
    const m = integerValue(mArg);
    if (x === undefined || n === undefined || m === undefined) {
      return valueError();
    }
    let sum = 0;
    coefficientArgs.forEach((coefficientArg, index) => {
      const coefficient = toNumber(coefficientArg) ?? 0;
      sum += coefficient * x ** (n + index * m);
    });
    return numberResult(sum);
  },
  SQRTPI: (value) => {
    const numeric = toNumber(value);
    return numeric === undefined
      ? valueError()
      : numericResultOrError(Math.sqrt(numeric * Math.PI));
  },
  SUMSQ: (...args) => {
    const error = firstError(args);
    if (error) return error;
    return numberResult(collectNumericArgs(args).reduce((sum, value) => sum + value ** 2, 0));
  },
  CONVERT: (numberArg, fromUnitArg, toUnitArg) => convertBuiltin(numberArg, fromUnitArg, toUnitArg),
  EUROCONVERT: (numberArg, sourceArg, targetArg, fullPrecisionArg, triangulationPrecisionArg) =>
    euroconvertBuiltin(
      numberArg,
      sourceArg,
      targetArg,
      fullPrecisionArg,
      triangulationPrecisionArg,
    ),
  ...radixBuiltins,
  ...complexBuiltins,
  T: (value = { tag: ValueTag.Empty }) => {
    if (value.tag === ValueTag.Error) {
      return value;
    }
    return value.tag === ValueTag.String ? value : { tag: ValueTag.Empty };
  },
  ISOMITTED: (...args) => {
    if (args.length !== 1) {
      return valueError();
    }
    return { tag: ValueTag.Boolean, value: false };
  },
  N: (value = { tag: ValueTag.Empty }) => {
    if (value.tag === ValueTag.Error) {
      return value;
    }
    return numberResult(toNumber(value) ?? 0);
  },
  TYPE: (value = { tag: ValueTag.Empty }) => {
    if (value.tag === ValueTag.Error) {
      return numberResult(16);
    }
    if ((value as EvaluationResult & { kind?: string }).kind === "array") {
      return numberResult(64);
    }
    switch (value.tag) {
      case ValueTag.Number:
      case ValueTag.Empty:
        return numberResult(1);
      case ValueTag.String:
        return numberResult(2);
      case ValueTag.Boolean:
        return numberResult(4);
    }
  },
  DELTA: (leftArg, rightArg = { tag: ValueTag.Number, value: 0 }) => {
    const left = toNumber(leftArg);
    const right = toNumber(rightArg);
    if (left === undefined || right === undefined) {
      return valueError();
    }
    return numberResult(left === right ? 1 : 0);
  },
  GESTEP: (numberArg, stepArg = { tag: ValueTag.Number, value: 0 }) => {
    const numberValue = toNumber(numberArg);
    const stepValue = toNumber(stepArg);
    if (numberValue === undefined || stepValue === undefined) {
      return valueError();
    }
    return numberResult(numberValue >= stepValue ? 1 : 0);
  },
  GAUSS: (value) => {
    const numeric = toNumber(value);
    return numeric === undefined ? valueError() : numberResult(standardNormalCdf(numeric) - 0.5);
  },
  PHI: (value) => {
    const numeric = toNumber(value);
    return numeric === undefined ? valueError() : numberResult(standardNormalPdf(numeric));
  },
  STANDARDIZE: (xArg, meanArg, standardDeviationArg) => {
    const x = toNumber(xArg);
    const mean = toNumber(meanArg);
    const standardDeviation = toNumber(standardDeviationArg);
    if (
      x === undefined ||
      mean === undefined ||
      standardDeviation === undefined ||
      standardDeviation <= 0
    ) {
      return valueError();
    }
    return numberResult((x - mean) / standardDeviation);
  },
  MODE: (...args) => {
    const error = firstError(args);
    if (error) return error;
    const mode = modeSingle(collectNumericArgs(args));
    return mode === undefined ? { tag: ValueTag.Error, code: ErrorCode.NA } : numberResult(mode);
  },
  "MODE.SNGL": (...args) => {
    const error = firstError(args);
    if (error) return error;
    const mode = modeSingle(collectNumericArgs(args));
    return mode === undefined ? { tag: ValueTag.Error, code: ErrorCode.NA } : numberResult(mode);
  },
  STDEV: (...args) => {
    const error = firstError(args);
    if (error) return error;
    const numbers = collectStatNumericArgs(args);
    return numbers.length < 2
      ? valueError()
      : numericResultOrError(sampleStandardDeviation(numbers));
  },
  "STDEV.S": (...args) => {
    const error = firstError(args);
    if (error) return error;
    const numbers = collectStatNumericArgs(args);
    return numbers.length < 2
      ? valueError()
      : numericResultOrError(sampleStandardDeviation(numbers));
  },
  STDEVP: (...args) => {
    const error = firstError(args);
    if (error) return error;
    const numbers = collectStatNumericArgs(args);
    return numbers.length === 0
      ? valueError()
      : numericResultOrError(populationStandardDeviation(numbers));
  },
  "STDEV.P": (...args) => {
    const error = firstError(args);
    if (error) return error;
    const numbers = collectStatNumericArgs(args);
    return numbers.length === 0
      ? valueError()
      : numericResultOrError(populationStandardDeviation(numbers));
  },
  STDEVA: (...args) => {
    const error = firstError(args);
    if (error) return error;
    const numbers = collectAStyleNumericArgs(args);
    return numbers.length < 2
      ? valueError()
      : numericResultOrError(sampleStandardDeviation(numbers));
  },
  STDEVPA: (...args) => {
    const error = firstError(args);
    if (error) return error;
    const numbers = collectAStyleNumericArgs(args);
    return numbers.length === 0
      ? valueError()
      : numericResultOrError(populationStandardDeviation(numbers));
  },
  VAR: (...args) => {
    const error = firstError(args);
    if (error) return error;
    const numbers = collectStatNumericArgs(args);
    return numbers.length < 2 ? valueError() : numberResult(sampleVariance(numbers));
  },
  "VAR.S": (...args) => {
    const error = firstError(args);
    if (error) return error;
    const numbers = collectStatNumericArgs(args);
    return numbers.length < 2 ? valueError() : numberResult(sampleVariance(numbers));
  },
  VARP: (...args) => {
    const error = firstError(args);
    if (error) return error;
    const numbers = collectStatNumericArgs(args);
    return numbers.length === 0 ? valueError() : numberResult(populationVariance(numbers));
  },
  "VAR.P": (...args) => {
    const error = firstError(args);
    if (error) return error;
    const numbers = collectStatNumericArgs(args);
    return numbers.length === 0 ? valueError() : numberResult(populationVariance(numbers));
  },
  VARA: (...args) => {
    const error = firstError(args);
    if (error) return error;
    const numbers = collectAStyleNumericArgs(args);
    return numbers.length < 2 ? valueError() : numberResult(sampleVariance(numbers));
  },
  VARPA: (...args) => {
    const error = firstError(args);
    if (error) return error;
    const numbers = collectAStyleNumericArgs(args);
    return numbers.length === 0 ? valueError() : numberResult(populationVariance(numbers));
  },
  SKEW: (...args) => {
    const error = firstError(args);
    if (error) return error;
    const skew = skewSample(collectStatNumericArgs(args));
    return skew === undefined ? valueError() : numberResult(skew);
  },
  "SKEW.P": (...args) => {
    const error = firstError(args);
    if (error) return error;
    const skew = skewPopulation(collectStatNumericArgs(args));
    return skew === undefined ? valueError() : numberResult(skew);
  },
  SKEWP: (...args) => {
    return scalarBuiltins["SKEW.P"]!(...args);
  },
  KURT: (...args) => {
    const error = firstError(args);
    if (error) return error;
    const value = kurtosis(collectStatNumericArgs(args));
    return value === undefined ? valueError() : numberResult(value);
  },
  NORMDIST: (xArg, meanArg, standardDeviationArg, cumulativeArg) => {
    const x = toNumber(xArg);
    const mean = toNumber(meanArg);
    const standardDeviation = toNumber(standardDeviationArg);
    const cumulative = coerceBoolean(cumulativeArg, false);
    if (
      x === undefined ||
      mean === undefined ||
      standardDeviation === undefined ||
      cumulative === undefined ||
      standardDeviation <= 0
    ) {
      return valueError();
    }
    return numberResult(
      cumulative
        ? percentileNormal(mean, standardDeviation, x)
        : standardNormalPdf((x - mean) / standardDeviation) / standardDeviation,
    );
  },
  "NORM.DIST": (xArg, meanArg, standardDeviationArg, cumulativeArg) => {
    return scalarBuiltins["NORMDIST"]!(xArg, meanArg, standardDeviationArg, cumulativeArg);
  },
  NORMINV: (probabilityArg, meanArg, standardDeviationArg) => {
    const probability = toNumber(probabilityArg);
    const mean = toNumber(meanArg);
    const standardDeviation = toNumber(standardDeviationArg);
    if (
      probability === undefined ||
      mean === undefined ||
      standardDeviation === undefined ||
      standardDeviation <= 0
    ) {
      return valueError();
    }
    const result = inverseNormal(probability, mean, standardDeviation);
    return result === undefined ? valueError() : numberResult(result);
  },
  "NORM.INV": (probabilityArg, meanArg, standardDeviationArg) => {
    return scalarBuiltins["NORMINV"]!(probabilityArg, meanArg, standardDeviationArg);
  },
  NORMSDIST: (value) => {
    const numeric = toNumber(value);
    return numeric === undefined ? valueError() : numberResult(standardNormalCdf(numeric));
  },
  "LEGACY.NORMSDIST": (value) => {
    return scalarBuiltins["NORMSDIST"]!(value);
  },
  "NORM.S.DIST": (value, cumulativeArg = { tag: ValueTag.Boolean, value: true }) => {
    const numeric = toNumber(value);
    const cumulative = coerceBoolean(cumulativeArg, true);
    if (numeric === undefined || cumulative === undefined) {
      return valueError();
    }
    return numberResult(cumulative ? standardNormalCdf(numeric) : standardNormalPdf(numeric));
  },
  NORMSINV: (value) => {
    const numeric = toNumber(value);
    if (numeric === undefined) {
      return valueError();
    }
    const result = inverseStandardNormal(numeric);
    return result === undefined ? valueError() : numberResult(result);
  },
  "LEGACY.NORMSINV": (value) => {
    return scalarBuiltins["NORMSINV"]!(value);
  },
  "NORM.S.INV": (value) => {
    return scalarBuiltins["NORMSINV"]!(value);
  },
  LOGINV: (probabilityArg, meanArg, standardDeviationArg) => {
    const probability = toNumber(probabilityArg);
    const mean = toNumber(meanArg);
    const standardDeviation = toNumber(standardDeviationArg);
    if (
      probability === undefined ||
      mean === undefined ||
      standardDeviation === undefined ||
      standardDeviation <= 0
    ) {
      return valueError();
    }
    const normal = inverseNormal(probability, mean, standardDeviation);
    return normal === undefined ? valueError() : numberResult(Math.exp(normal));
  },
  "LOGNORM.INV": (probabilityArg, meanArg, standardDeviationArg) => {
    return scalarBuiltins["LOGINV"]!(probabilityArg, meanArg, standardDeviationArg);
  },
  LOGNORMDIST: (xArg, meanArg, standardDeviationArg) => {
    const x = toNumber(xArg);
    const mean = toNumber(meanArg);
    const standardDeviation = toNumber(standardDeviationArg);
    if (
      x === undefined ||
      mean === undefined ||
      standardDeviation === undefined ||
      standardDeviation <= 0 ||
      x <= 0
    ) {
      return valueError();
    }
    return numberResult(percentileNormal(mean, standardDeviation, Math.log(x)));
  },
  "LOGNORM.DIST": (
    xArg,
    meanArg,
    standardDeviationArg,
    cumulativeArg = { tag: ValueTag.Boolean, value: true },
  ) => {
    const x = toNumber(xArg);
    const mean = toNumber(meanArg);
    const standardDeviation = toNumber(standardDeviationArg);
    const cumulative = coerceBoolean(cumulativeArg, true);
    if (
      x === undefined ||
      mean === undefined ||
      standardDeviation === undefined ||
      cumulative === undefined ||
      standardDeviation <= 0 ||
      x <= 0
    ) {
      return valueError();
    }
    const z = (Math.log(x) - mean) / standardDeviation;
    return numberResult(
      cumulative ? standardNormalCdf(z) : standardNormalPdf(z) / (x * standardDeviation),
    );
  },
  EFFECT: (nominalRateArg, periodsArg) => {
    const nominalRate = toNumber(nominalRateArg);
    const periods = positiveIntegerValue(periodsArg);
    if (nominalRate === undefined || periods === undefined) {
      return valueError();
    }
    return numberResult((1 + nominalRate / periods) ** periods - 1);
  },
  NOMINAL: (effectiveRateArg, periodsArg) => {
    const effectiveRate = toNumber(effectiveRateArg);
    const periods = positiveIntegerValue(periodsArg);
    if (effectiveRate === undefined || periods === undefined || effectiveRate <= -1) {
      return valueError();
    }
    return numberResult(periods * ((1 + effectiveRate) ** (1 / periods) - 1));
  },
  PDURATION: (rateArg, presentArg, futureArg) => {
    const rate = toNumber(rateArg);
    const present = toNumber(presentArg);
    const future = toNumber(futureArg);
    if (
      rate === undefined ||
      present === undefined ||
      future === undefined ||
      rate <= 0 ||
      present <= 0 ||
      future <= 0
    ) {
      return valueError();
    }
    return numberResult(Math.log(future / present) / Math.log(1 + rate));
  },
  RRI: (periodsArg, presentArg, futureArg) => {
    const periods = toNumber(periodsArg);
    const present = toNumber(presentArg);
    const future = toNumber(futureArg);
    if (
      periods === undefined ||
      present === undefined ||
      future === undefined ||
      periods <= 0 ||
      present === 0
    ) {
      return valueError();
    }
    return numericResultOrError((future / present) ** (1 / periods) - 1);
  },
  FV: (rateArg, periodsArg, paymentArg, presentArg, typeArg) => {
    const rate = toNumber(rateArg);
    const periods = toNumber(periodsArg);
    const payment = toNumber(paymentArg);
    const present = coerceNumber(presentArg, 0);
    const type = coercePaymentType(typeArg, 0);
    if (
      rate === undefined ||
      periods === undefined ||
      payment === undefined ||
      present === undefined ||
      type === undefined
    ) {
      return valueError();
    }
    return numberResult(futureValue(rate, periods, payment, present, type));
  },
  FVSCHEDULE: (principalArg, ...scheduleArgs) => {
    const principal = toNumber(principalArg);
    if (principal === undefined) {
      return valueError();
    }
    let result = principal;
    for (const scheduleArg of scheduleArgs) {
      const rate = toNumber(scheduleArg);
      if (rate === undefined) {
        return valueError();
      }
      result *= 1 + rate;
    }
    return numberResult(result);
  },
  DB: (costArg, salvageArg, lifeArg, periodArg, monthArg) => {
    const cost = toNumber(costArg);
    const salvage = toNumber(salvageArg);
    const life = toNumber(lifeArg);
    const period = toNumber(periodArg);
    const month = coerceNumber(monthArg, 12);
    if (
      cost === undefined ||
      salvage === undefined ||
      life === undefined ||
      period === undefined ||
      month === undefined
    ) {
      return valueError();
    }
    const depreciation = dbDepreciation(cost, salvage, life, period, month);
    return depreciation === undefined ? valueError() : numberResult(depreciation);
  },
  DDB: (costArg, salvageArg, lifeArg, periodArg, factorArg) => {
    const cost = toNumber(costArg);
    const salvage = toNumber(salvageArg);
    const life = toNumber(lifeArg);
    const period = toNumber(periodArg);
    const factor = coerceNumber(factorArg, 2);
    if (
      cost === undefined ||
      salvage === undefined ||
      life === undefined ||
      period === undefined ||
      factor === undefined
    ) {
      return valueError();
    }
    const depreciation = ddbDepreciation(cost, salvage, life, period, factor);
    return depreciation === undefined ? valueError() : numberResult(depreciation);
  },
  VDB: (costArg, salvageArg, lifeArg, startArg, endArg, factorArg, noSwitchArg) => {
    const cost = toNumber(costArg);
    const salvage = toNumber(salvageArg);
    const life = toNumber(lifeArg);
    const start = toNumber(startArg);
    const end = toNumber(endArg);
    const factor = coerceNumber(factorArg, 2);
    const noSwitch = coerceBoolean(noSwitchArg, false);
    if (
      cost === undefined ||
      salvage === undefined ||
      life === undefined ||
      start === undefined ||
      end === undefined ||
      factor === undefined ||
      noSwitch === undefined
    ) {
      return valueError();
    }
    const depreciation = vdbDepreciation(cost, salvage, life, start, end, factor, noSwitch);
    return depreciation === undefined ? valueError() : numberResult(depreciation);
  },
  PV: (rateArg, periodsArg, paymentArg, futureArg, typeArg) => {
    const rate = toNumber(rateArg);
    const periods = toNumber(periodsArg);
    const payment = toNumber(paymentArg);
    const future = coerceNumber(futureArg, 0);
    const type = coercePaymentType(typeArg, 0);
    if (
      rate === undefined ||
      periods === undefined ||
      payment === undefined ||
      future === undefined ||
      type === undefined
    ) {
      return valueError();
    }
    return numberResult(presentValue(rate, periods, payment, future, type));
  },
  PMT: (rateArg, periodsArg, presentArg, futureArg, typeArg) => {
    const rate = toNumber(rateArg);
    const periods = toNumber(periodsArg);
    const present = toNumber(presentArg);
    const future = coerceNumber(futureArg, 0);
    const type = coercePaymentType(typeArg, 0);
    if (
      rate === undefined ||
      periods === undefined ||
      present === undefined ||
      future === undefined ||
      type === undefined
    ) {
      return valueError();
    }
    const payment = periodicPayment(rate, periods, present, future, type);
    return payment === undefined ? valueError() : numberResult(payment);
  },
  RATE: (periodsArg, paymentArg, presentArg, futureArg, typeArg, guessArg) => {
    const periods = toNumber(periodsArg);
    const payment = toNumber(paymentArg);
    const present = toNumber(presentArg);
    const future = coerceNumber(futureArg, 0);
    const type = coercePaymentType(typeArg, 0);
    const guess = coerceNumber(guessArg, 0.1);
    if (
      periods === undefined ||
      payment === undefined ||
      present === undefined ||
      future === undefined ||
      type === undefined ||
      guess === undefined
    ) {
      return valueError();
    }
    const rate = solveRate(periods, payment, present, future, type, guess);
    return rate === undefined ? valueError() : numberResult(rate);
  },
  SLN: (costArg, salvageArg, lifeArg) => {
    const cost = toNumber(costArg);
    const salvage = toNumber(salvageArg);
    const life = toNumber(lifeArg);
    if (cost === undefined || salvage === undefined || life === undefined || life <= 0) {
      return valueError();
    }
    return numberResult((cost - salvage) / life);
  },
  SYD: (costArg, salvageArg, lifeArg, periodArg) => {
    const cost = toNumber(costArg);
    const salvage = toNumber(salvageArg);
    const life = toNumber(lifeArg);
    const period = toNumber(periodArg);
    if (
      cost === undefined ||
      salvage === undefined ||
      life === undefined ||
      period === undefined ||
      life <= 0 ||
      period <= 0 ||
      period > life
    ) {
      return valueError();
    }
    return numberResult(((cost - salvage) * (life - period + 1) * 2) / (life * (life + 1)));
  },
  NPER: (rateArg, paymentArg, presentArg, futureArg, typeArg) => {
    const rate = toNumber(rateArg);
    const payment = toNumber(paymentArg);
    const present = toNumber(presentArg);
    const future = coerceNumber(futureArg, 0);
    const type = coercePaymentType(typeArg, 0);
    if (
      rate === undefined ||
      payment === undefined ||
      present === undefined ||
      future === undefined ||
      type === undefined
    ) {
      return valueError();
    }
    const periods = totalPeriods(rate, payment, present, future, type);
    return periods === undefined ? valueError() : numberResult(periods);
  },
  NPV: (rateArg, ...valueArgs) => {
    const rate = toNumber(rateArg);
    if (rate === undefined || valueArgs.length === 0) {
      return valueError();
    }
    let result = 0;
    for (let index = 0; index < valueArgs.length; index += 1) {
      const value = toNumber(valueArgs[index]!);
      if (value === undefined) {
        return valueError();
      }
      result += value / (1 + rate) ** (index + 1);
    }
    return numberResult(result);
  },
  IPMT: (rateArg, periodArg, periodsArg, presentArg, futureArg, typeArg) => {
    const rate = toNumber(rateArg);
    const period = toNumber(periodArg);
    const periods = toNumber(periodsArg);
    const present = toNumber(presentArg);
    const future = coerceNumber(futureArg, 0);
    const type = coercePaymentType(typeArg, 0);
    if (
      rate === undefined ||
      period === undefined ||
      periods === undefined ||
      present === undefined ||
      future === undefined ||
      type === undefined
    ) {
      return valueError();
    }
    const interest = interestPayment(rate, period, periods, present, future, type);
    return interest === undefined ? valueError() : numberResult(interest);
  },
  PPMT: (rateArg, periodArg, periodsArg, presentArg, futureArg, typeArg) => {
    const rate = toNumber(rateArg);
    const period = toNumber(periodArg);
    const periods = toNumber(periodsArg);
    const present = toNumber(presentArg);
    const future = coerceNumber(futureArg, 0);
    const type = coercePaymentType(typeArg, 0);
    if (
      rate === undefined ||
      period === undefined ||
      periods === undefined ||
      present === undefined ||
      future === undefined ||
      type === undefined
    ) {
      return valueError();
    }
    const principal = principalPayment(rate, period, periods, present, future, type);
    return principal === undefined ? valueError() : numberResult(principal);
  },
  ISPMT: (rateArg, periodArg, periodsArg, presentArg) => {
    const rate = toNumber(rateArg);
    const period = toNumber(periodArg);
    const periods = toNumber(periodsArg);
    const present = toNumber(presentArg);
    if (
      rate === undefined ||
      period === undefined ||
      periods === undefined ||
      present === undefined ||
      periods <= 0 ||
      period < 1 ||
      period > periods
    ) {
      return valueError();
    }
    return numberResult(present * rate * (period / periods - 1));
  },
  CUMIPMT: (rateArg, periodsArg, presentArg, startPeriodArg, endPeriodArg, typeArg) => {
    const rate = toNumber(rateArg);
    const periods = toNumber(periodsArg);
    const present = toNumber(presentArg);
    const startPeriod = integerValue(startPeriodArg);
    const endPeriod = integerValue(endPeriodArg);
    const type = coercePaymentType(typeArg, 0);
    if (
      rate === undefined ||
      periods === undefined ||
      present === undefined ||
      startPeriod === undefined ||
      endPeriod === undefined ||
      type === undefined
    ) {
      return valueError();
    }
    const total = cumulativePeriodicPayment(
      rate,
      periods,
      present,
      startPeriod,
      endPeriod,
      type,
      false,
    );
    return total === undefined ? valueError() : numberResult(total);
  },
  CUMPRINC: (rateArg, periodsArg, presentArg, startPeriodArg, endPeriodArg, typeArg) => {
    const rate = toNumber(rateArg);
    const periods = toNumber(periodsArg);
    const present = toNumber(presentArg);
    const startPeriod = integerValue(startPeriodArg);
    const endPeriod = integerValue(endPeriodArg);
    const type = coercePaymentType(typeArg, 0);
    if (
      rate === undefined ||
      periods === undefined ||
      present === undefined ||
      startPeriod === undefined ||
      endPeriod === undefined ||
      type === undefined
    ) {
      return valueError();
    }
    const total = cumulativePeriodicPayment(
      rate,
      periods,
      present,
      startPeriod,
      endPeriod,
      type,
      true,
    );
    return total === undefined ? valueError() : numberResult(total);
  },
  PERMUT: (numberArg, chosenArg) => {
    const numberValue = nonNegativeIntegerValue(numberArg);
    const chosenValue = nonNegativeIntegerValue(chosenArg);
    if (numberValue === undefined || chosenValue === undefined || chosenValue > numberValue) {
      return valueError();
    }
    let result = 1;
    for (let index = 0; index < chosenValue; index += 1) {
      result *= numberValue - index;
    }
    return numberResult(result);
  },
  PERMUTATIONA: (numberArg, chosenArg) => {
    const numberValue = nonNegativeIntegerValue(numberArg);
    const chosenValue = nonNegativeIntegerValue(chosenArg);
    if (numberValue === undefined || chosenValue === undefined) {
      return valueError();
    }
    return numberResult(numberValue ** chosenValue);
  },
  "CONFIDENCE.NORM": (alphaArg, standardDeviationArg, sizeArg) => {
    const alpha = toNumber(alphaArg);
    const standardDeviation = toNumber(standardDeviationArg);
    const size = toNumber(sizeArg);
    if (
      alpha === undefined ||
      standardDeviation === undefined ||
      size === undefined ||
      !(alpha > 0 && alpha < 1) ||
      standardDeviation <= 0 ||
      size < 1
    ) {
      return valueError();
    }
    const criticalValue = inverseStandardNormal(1 - alpha / 2);
    return criticalValue === undefined
      ? valueError()
      : numberResult((criticalValue * standardDeviation) / Math.sqrt(size));
  },
  ERF: (lowerArg, upperArg) => {
    const lower = toNumber(lowerArg);
    if (lower === undefined) {
      return valueError();
    }
    if (upperArg === undefined) {
      return numberResult(erfApprox(lower));
    }
    const upper = toNumber(upperArg);
    return upper === undefined ? valueError() : numberResult(erfApprox(upper) - erfApprox(lower));
  },
  "ERF.PRECISE": (valueArg) => {
    const value = toNumber(valueArg);
    return value === undefined ? valueError() : numberResult(erfApprox(value));
  },
  ERFC: (valueArg) => {
    const value = toNumber(valueArg);
    return value === undefined ? valueError() : numberResult(1 - erfApprox(value));
  },
  "ERFC.PRECISE": (valueArg) => {
    const value = toNumber(valueArg);
    return value === undefined ? valueError() : numberResult(1 - erfApprox(value));
  },
  FISHER: (valueArg) => {
    const value = toNumber(valueArg);
    if (value === undefined || value <= -1 || value >= 1) {
      return valueError();
    }
    return numberResult(0.5 * Math.log((1 + value) / (1 - value)));
  },
  FISHERINV: (valueArg) => {
    const value = toNumber(valueArg);
    if (value === undefined) {
      return valueError();
    }
    const exponent = Math.exp(2 * value);
    return numberResult((exponent - 1) / (exponent + 1));
  },
  GAMMALN: (valueArg) => {
    const value = toNumber(valueArg);
    return value === undefined ? valueError() : numericResultOrError(logGamma(value));
  },
  "GAMMALN.PRECISE": (valueArg) => {
    return scalarBuiltins["GAMMALN"]!(valueArg);
  },
  GAMMA: (valueArg) => {
    const value = toNumber(valueArg);
    return value === undefined ? valueError() : numericResultOrError(gammaFunction(value));
  },
  CONFIDENCE: (alphaArg, standardDeviationArg, sizeArg) => {
    return scalarBuiltins["CONFIDENCE.NORM"]!(alphaArg, standardDeviationArg, sizeArg);
  },
  "CONFIDENCE.T": (alphaArg, standardDeviationArg, sizeArg) => {
    const alpha = toNumber(alphaArg);
    const standardDeviation = toNumber(standardDeviationArg);
    const size = integerValue(sizeArg);
    if (
      alpha === undefined ||
      standardDeviation === undefined ||
      size === undefined ||
      !(alpha > 0 && alpha < 1) ||
      !(standardDeviation > 0) ||
      size < 2
    ) {
      return valueError();
    }
    const critical = inverseStudentT(1 - alpha / 2, size - 1);
    return critical === undefined
      ? valueError()
      : numericResultOrError((critical * standardDeviation) / Math.sqrt(size));
  },
  "BETA.DIST": (xArg, alphaArg, betaArg, cumulativeArg, lowerBoundArg, upperBoundArg) => {
    const x = toNumber(xArg);
    const alpha = toNumber(alphaArg);
    const beta = toNumber(betaArg);
    const cumulative = coerceBoolean(cumulativeArg, false);
    const lowerBound = coerceNumber(lowerBoundArg, 0);
    const upperBound = coerceNumber(upperBoundArg, 1);
    if (
      x === undefined ||
      alpha === undefined ||
      beta === undefined ||
      cumulative === undefined ||
      lowerBound === undefined ||
      upperBound === undefined
    ) {
      return valueError();
    }
    return numericResultOrError(
      cumulative
        ? betaDistributionCdf(x, alpha, beta, lowerBound, upperBound)
        : betaDistributionDensity(x, alpha, beta, lowerBound, upperBound),
    );
  },
  BETADIST: (xArg, alphaArg, betaArg, lowerBoundArg, upperBoundArg) => {
    return scalarBuiltins["BETA.DIST"]!(
      xArg,
      alphaArg,
      betaArg,
      { tag: ValueTag.Boolean, value: true },
      lowerBoundArg,
      upperBoundArg,
    );
  },
  "BETA.INV": (probabilityArg, alphaArg, betaArg, lowerBoundArg, upperBoundArg) => {
    const probability = toNumber(probabilityArg);
    const alpha = toNumber(alphaArg);
    const beta = toNumber(betaArg);
    const lowerBound = coerceNumber(lowerBoundArg, 0);
    const upperBound = coerceNumber(upperBoundArg, 1);
    if (
      probability === undefined ||
      alpha === undefined ||
      beta === undefined ||
      lowerBound === undefined ||
      upperBound === undefined
    ) {
      return valueError();
    }
    const result = betaDistributionInverse(probability, alpha, beta, lowerBound, upperBound);
    return result === undefined ? valueError() : numericResultOrError(result);
  },
  BETAINV: (probabilityArg, alphaArg, betaArg, lowerBoundArg, upperBoundArg) => {
    return scalarBuiltins["BETA.INV"]!(
      probabilityArg,
      alphaArg,
      betaArg,
      lowerBoundArg,
      upperBoundArg,
    );
  },
  EXPONDIST: (xArg, lambdaArg, cumulativeArg) => {
    const x = toNumber(xArg);
    const lambda = toNumber(lambdaArg);
    const cumulative = coerceBoolean(cumulativeArg, false);
    if (
      x === undefined ||
      lambda === undefined ||
      cumulative === undefined ||
      x < 0 ||
      lambda <= 0
    ) {
      return valueError();
    }
    return numberResult(cumulative ? 1 - Math.exp(-lambda * x) : lambda * Math.exp(-lambda * x));
  },
  "EXPON.DIST": (xArg, lambdaArg, cumulativeArg) => {
    return scalarBuiltins["EXPONDIST"]!(xArg, lambdaArg, cumulativeArg);
  },
  POISSON: (eventsArg, meanArg, cumulativeArg) => {
    const events = nonNegativeIntegerValue(eventsArg);
    const mean = toNumber(meanArg);
    const cumulative = coerceBoolean(cumulativeArg, false);
    if (events === undefined || mean === undefined || cumulative === undefined || mean < 0) {
      return valueError();
    }
    if (!cumulative) {
      return numericResultOrError(poissonProbability(events, mean));
    }
    let total = 0;
    for (let index = 0; index <= events; index += 1) {
      total += poissonProbability(index, mean);
    }
    return numericResultOrError(total);
  },
  "POISSON.DIST": (eventsArg, meanArg, cumulativeArg) => {
    return scalarBuiltins["POISSON"]!(eventsArg, meanArg, cumulativeArg);
  },
  WEIBULL: (xArg, alphaArg, betaArg, cumulativeArg) => {
    const x = toNumber(xArg);
    const alpha = toNumber(alphaArg);
    const beta = toNumber(betaArg);
    const cumulative = coerceBoolean(cumulativeArg, false);
    if (
      x === undefined ||
      alpha === undefined ||
      beta === undefined ||
      cumulative === undefined ||
      x < 0 ||
      alpha <= 0 ||
      beta <= 0
    ) {
      return valueError();
    }
    if (cumulative) {
      return numberResult(1 - Math.exp(-((x / beta) ** alpha)));
    }
    if (x === 0) {
      return numberResult(alpha === 1 ? 1 / beta : alpha < 1 ? Number.POSITIVE_INFINITY : 0);
    }
    return numberResult(
      (alpha / beta ** alpha) * x ** (alpha - 1) * Math.exp(-((x / beta) ** alpha)),
    );
  },
  "WEIBULL.DIST": (xArg, alphaArg, betaArg, cumulativeArg) => {
    return scalarBuiltins["WEIBULL"]!(xArg, alphaArg, betaArg, cumulativeArg);
  },
  GAMMADIST: (xArg, alphaArg, betaArg, cumulativeArg) => {
    const x = toNumber(xArg);
    const alpha = toNumber(alphaArg);
    const beta = toNumber(betaArg);
    const cumulative = coerceBoolean(cumulativeArg, false);
    if (
      x === undefined ||
      alpha === undefined ||
      beta === undefined ||
      cumulative === undefined ||
      x < 0 ||
      alpha <= 0 ||
      beta <= 0
    ) {
      return valueError();
    }
    return numberResult(
      cumulative ? gammaDistributionCdf(x, alpha, beta) : gammaDistributionDensity(x, alpha, beta),
    );
  },
  "GAMMA.DIST": (xArg, alphaArg, betaArg, cumulativeArg) => {
    return scalarBuiltins["GAMMADIST"]!(xArg, alphaArg, betaArg, cumulativeArg);
  },
  "GAMMA.INV": (probabilityArg, alphaArg, betaArg) => {
    const probability = toNumber(probabilityArg);
    const alpha = toNumber(alphaArg);
    const beta = toNumber(betaArg);
    if (
      probability === undefined ||
      alpha === undefined ||
      beta === undefined ||
      !(probability > 0 && probability < 1) ||
      !(alpha > 0) ||
      !(beta > 0)
    ) {
      return valueError();
    }
    const result = inverseGammaDistribution(probability, alpha, beta);
    return result === undefined ? valueError() : numericResultOrError(result);
  },
  GAMMAINV: (probabilityArg, alphaArg, betaArg) => {
    return scalarBuiltins["GAMMA.INV"]!(probabilityArg, alphaArg, betaArg);
  },
  CHIDIST: (xArg, degreesArg) => {
    const x = toNumber(xArg);
    const degrees = toNumber(degreesArg);
    if (x === undefined || degrees === undefined || x < 0 || degrees < 1) {
      return valueError();
    }
    return numericResultOrError(regularizedUpperGamma(degrees / 2, x / 2));
  },
  "LEGACY.CHIDIST": (xArg, degreesArg) => {
    return scalarBuiltins["CHIDIST"]!(xArg, degreesArg);
  },
  CHIINV: (probabilityArg, degreesArg) => {
    return scalarBuiltins["CHISQ.INV.RT"]!(probabilityArg, degreesArg);
  },
  "CHISQ.DIST.RT": (xArg, degreesArg) => {
    return scalarBuiltins["CHIDIST"]!(xArg, degreesArg);
  },
  CHISQDIST: (xArg, degreesArg) => {
    return scalarBuiltins["CHISQ.DIST.RT"]!(xArg, degreesArg);
  },
  "CHISQ.DIST": (xArg, degreesArg, cumulativeArg) => {
    const x = toNumber(xArg);
    const degrees = toNumber(degreesArg);
    const cumulative = coerceBoolean(cumulativeArg, false);
    if (
      x === undefined ||
      degrees === undefined ||
      cumulative === undefined ||
      x < 0 ||
      degrees < 1
    ) {
      return valueError();
    }
    return numberResult(cumulative ? chiSquareCdf(x, degrees) : chiSquareDensity(x, degrees));
  },
  "CHISQ.INV.RT": (probabilityArg, degreesArg) => {
    const probability = toNumber(probabilityArg);
    const degrees = toNumber(degreesArg);
    if (
      probability === undefined ||
      degrees === undefined ||
      !(probability > 0 && probability < 1) ||
      !(degrees >= 1)
    ) {
      return valueError();
    }
    const result = inverseChiSquare(1 - probability, degrees);
    return result === undefined
      ? { tag: ValueTag.Error, code: ErrorCode.NA }
      : numberResult(result);
  },
  CHISQINV: (probabilityArg, degreesArg) => {
    return scalarBuiltins["CHISQ.INV.RT"]!(probabilityArg, degreesArg);
  },
  "LEGACY.CHIINV": (probabilityArg, degreesArg) => {
    return scalarBuiltins["CHISQ.INV.RT"]!(probabilityArg, degreesArg);
  },
  "CHISQ.INV": (probabilityArg, degreesArg) => {
    const probability = toNumber(probabilityArg);
    const degrees = toNumber(degreesArg);
    if (
      probability === undefined ||
      degrees === undefined ||
      !(probability > 0 && probability < 1) ||
      !(degrees >= 1)
    ) {
      return valueError();
    }
    const result = inverseChiSquare(probability, degrees);
    return result === undefined
      ? { tag: ValueTag.Error, code: ErrorCode.NA }
      : numberResult(result);
  },
  "F.DIST": (xArg, degrees1Arg, degrees2Arg, cumulativeArg) => {
    const x = toNumber(xArg);
    const degrees1 = integerValue(degrees1Arg);
    const degrees2 = integerValue(degrees2Arg);
    const cumulative = coerceBoolean(cumulativeArg, false);
    if (
      x === undefined ||
      degrees1 === undefined ||
      degrees2 === undefined ||
      cumulative === undefined
    ) {
      return valueError();
    }
    return numericResultOrError(
      cumulative
        ? fDistributionCdf(x, degrees1, degrees2)
        : fDistributionDensity(x, degrees1, degrees2),
    );
  },
  "F.DIST.RT": (xArg, degrees1Arg, degrees2Arg) => {
    const x = toNumber(xArg);
    const degrees1 = integerValue(degrees1Arg);
    const degrees2 = integerValue(degrees2Arg);
    if (x === undefined || degrees1 === undefined || degrees2 === undefined) {
      return valueError();
    }
    return numericResultOrError(1 - fDistributionCdf(x, degrees1, degrees2));
  },
  FDIST: (xArg, degrees1Arg, degrees2Arg) => {
    return scalarBuiltins["F.DIST.RT"]!(xArg, degrees1Arg, degrees2Arg);
  },
  "LEGACY.FDIST": (xArg, degrees1Arg, degrees2Arg) => {
    return scalarBuiltins["F.DIST.RT"]!(xArg, degrees1Arg, degrees2Arg);
  },
  "F.INV": (probabilityArg, degrees1Arg, degrees2Arg) => {
    const probability = toNumber(probabilityArg);
    const degrees1 = integerValue(degrees1Arg);
    const degrees2 = integerValue(degrees2Arg);
    if (probability === undefined || degrees1 === undefined || degrees2 === undefined) {
      return valueError();
    }
    const result = inverseFDistribution(probability, degrees1, degrees2);
    return result === undefined ? valueError() : numericResultOrError(result);
  },
  "F.INV.RT": (probabilityArg, degrees1Arg, degrees2Arg) => {
    const probability = toNumber(probabilityArg);
    const degrees1 = integerValue(degrees1Arg);
    const degrees2 = integerValue(degrees2Arg);
    if (probability === undefined || degrees1 === undefined || degrees2 === undefined) {
      return valueError();
    }
    const result = inverseFDistribution(1 - probability, degrees1, degrees2);
    return result === undefined ? valueError() : numericResultOrError(result);
  },
  FINV: (probabilityArg, degrees1Arg, degrees2Arg) => {
    return scalarBuiltins["F.INV.RT"]!(probabilityArg, degrees1Arg, degrees2Arg);
  },
  "LEGACY.FINV": (probabilityArg, degrees1Arg, degrees2Arg) => {
    return scalarBuiltins["F.INV.RT"]!(probabilityArg, degrees1Arg, degrees2Arg);
  },
  "T.DIST": (xArg, degreesArg, cumulativeArg) => {
    const x = toNumber(xArg);
    const degrees = integerValue(degreesArg);
    const cumulative = coerceBoolean(cumulativeArg, false);
    if (x === undefined || degrees === undefined || cumulative === undefined) {
      return valueError();
    }
    return numericResultOrError(cumulative ? studentTCdf(x, degrees) : studentTDensity(x, degrees));
  },
  "T.DIST.RT": (xArg, degreesArg) => {
    const x = toNumber(xArg);
    const degrees = integerValue(degreesArg);
    if (x === undefined || degrees === undefined) {
      return valueError();
    }
    return numericResultOrError(1 - studentTCdf(x, degrees));
  },
  "T.DIST.2T": (xArg, degreesArg) => {
    const x = toNumber(xArg);
    const degrees = integerValue(degreesArg);
    if (x === undefined || degrees === undefined || x < 0) {
      return valueError();
    }
    return numericResultOrError(Math.min(1, 2 * (1 - studentTCdf(x, degrees))));
  },
  TDIST: (xArg, degreesArg, tailsArg) => {
    const x = toNumber(xArg);
    const degrees = integerValue(degreesArg);
    const tails = integerValue(tailsArg);
    if (x === undefined || degrees === undefined || tails === undefined || x < 0) {
      return valueError();
    }
    if (tails !== 1 && tails !== 2) {
      return valueError();
    }
    const upperTail = 1 - studentTCdf(x, degrees);
    return numericResultOrError(tails === 1 ? upperTail : Math.min(1, upperTail * 2));
  },
  "T.INV": (probabilityArg, degreesArg) => {
    const probability = toNumber(probabilityArg);
    const degrees = integerValue(degreesArg);
    if (probability === undefined || degrees === undefined) {
      return valueError();
    }
    const result = inverseStudentT(probability, degrees);
    return result === undefined ? valueError() : numericResultOrError(result);
  },
  "T.INV.2T": (probabilityArg, degreesArg) => {
    const probability = toNumber(probabilityArg);
    const degrees = integerValue(degreesArg);
    if (
      probability === undefined ||
      degrees === undefined ||
      !(probability > 0 && probability < 1)
    ) {
      return valueError();
    }
    const result = inverseStudentT(1 - probability / 2, degrees);
    return result === undefined ? valueError() : numericResultOrError(result);
  },
  TINV: (probabilityArg, degreesArg) => {
    return scalarBuiltins["T.INV.2T"]!(probabilityArg, degreesArg);
  },
  BINOMDIST: (successesArg, trialsArg, probabilityArg, cumulativeArg) => {
    const successes = nonNegativeIntegerValue(successesArg);
    const trials = nonNegativeIntegerValue(trialsArg);
    const probability = toNumber(probabilityArg);
    const cumulative = coerceBoolean(cumulativeArg, false);
    if (
      successes === undefined ||
      trials === undefined ||
      probability === undefined ||
      cumulative === undefined ||
      successes > trials ||
      probability < 0 ||
      probability > 1
    ) {
      return valueError();
    }
    if (!cumulative) {
      return numericResultOrError(binomialProbability(successes, trials, probability));
    }
    let total = 0;
    for (let index = 0; index <= successes; index += 1) {
      total += binomialProbability(index, trials, probability);
    }
    return numericResultOrError(total);
  },
  "BINOM.DIST": (successesArg, trialsArg, probabilityArg, cumulativeArg) => {
    return scalarBuiltins["BINOMDIST"]!(successesArg, trialsArg, probabilityArg, cumulativeArg);
  },
  "BINOM.DIST.RANGE": (trialsArg, probabilityArg, successesArg, upperSuccessesArg) => {
    const trials = nonNegativeIntegerValue(trialsArg);
    const probability = toNumber(probabilityArg);
    const lower = nonNegativeIntegerValue(successesArg);
    const upper = nonNegativeIntegerValue(upperSuccessesArg, lower);
    if (
      trials === undefined ||
      probability === undefined ||
      lower === undefined ||
      upper === undefined ||
      lower > upper ||
      upper > trials ||
      probability < 0 ||
      probability > 1
    ) {
      return valueError();
    }
    let total = 0;
    for (let index = lower; index <= upper; index += 1) {
      total += binomialProbability(index, trials, probability);
    }
    return numericResultOrError(total);
  },
  CRITBINOM: (trialsArg, probabilityArg, alphaArg) => {
    const trials = nonNegativeIntegerValue(trialsArg);
    const probability = toNumber(probabilityArg);
    const alpha = toNumber(alphaArg);
    if (
      trials === undefined ||
      probability === undefined ||
      alpha === undefined ||
      probability < 0 ||
      probability > 1 ||
      alpha <= 0 ||
      alpha >= 1
    ) {
      return valueError();
    }
    let cumulative = 0;
    for (let index = 0; index <= trials; index += 1) {
      cumulative += binomialProbability(index, trials, probability);
      if (cumulative >= alpha) {
        return numberResult(index);
      }
    }
    return numberResult(trials);
  },
  "BINOM.INV": (trialsArg, probabilityArg, alphaArg) => {
    return scalarBuiltins["CRITBINOM"]!(trialsArg, probabilityArg, alphaArg);
  },
  HYPGEOMDIST: (sampleSuccessesArg, sampleSizeArg, populationSuccessesArg, populationSizeArg) => {
    const sampleSuccesses = nonNegativeIntegerValue(sampleSuccessesArg);
    const sampleSize = nonNegativeIntegerValue(sampleSizeArg);
    const populationSuccesses = nonNegativeIntegerValue(populationSuccessesArg);
    const populationSize = positiveIntegerValue(populationSizeArg);
    if (
      sampleSuccesses === undefined ||
      sampleSize === undefined ||
      populationSuccesses === undefined ||
      populationSize === undefined
    ) {
      return valueError();
    }
    return numericResultOrError(
      hypergeometricProbability(sampleSuccesses, sampleSize, populationSuccesses, populationSize),
    );
  },
  "HYPGEOM.DIST": (
    sampleSuccessesArg,
    sampleSizeArg,
    populationSuccessesArg,
    populationSizeArg,
    cumulativeArg,
  ) => {
    const sampleSuccesses = nonNegativeIntegerValue(sampleSuccessesArg);
    const sampleSize = nonNegativeIntegerValue(sampleSizeArg);
    const populationSuccesses = nonNegativeIntegerValue(populationSuccessesArg);
    const populationSize = positiveIntegerValue(populationSizeArg);
    const cumulative = coerceBoolean(cumulativeArg, false);
    if (
      sampleSuccesses === undefined ||
      sampleSize === undefined ||
      populationSuccesses === undefined ||
      populationSize === undefined ||
      cumulative === undefined
    ) {
      return valueError();
    }
    if (!cumulative) {
      return numericResultOrError(
        hypergeometricProbability(sampleSuccesses, sampleSize, populationSuccesses, populationSize),
      );
    }
    const minimum = Math.max(0, sampleSize - (populationSize - populationSuccesses));
    let total = 0;
    for (let index = minimum; index <= sampleSuccesses; index += 1) {
      total += hypergeometricProbability(index, sampleSize, populationSuccesses, populationSize);
    }
    return numericResultOrError(total);
  },
  NEGBINOMDIST: (failuresArg, successesArg, probabilityArg) => {
    const failures = nonNegativeIntegerValue(failuresArg);
    const successes = positiveIntegerValue(successesArg);
    const probability = toNumber(probabilityArg);
    if (
      failures === undefined ||
      successes === undefined ||
      probability === undefined ||
      probability < 0 ||
      probability > 1
    ) {
      return valueError();
    }
    return numericResultOrError(negativeBinomialProbability(failures, successes, probability));
  },
  "NEGBINOM.DIST": (failuresArg, successesArg, probabilityArg, cumulativeArg) => {
    const failures = nonNegativeIntegerValue(failuresArg);
    const successes = positiveIntegerValue(successesArg);
    const probability = toNumber(probabilityArg);
    const cumulative = coerceBoolean(cumulativeArg, false);
    if (
      failures === undefined ||
      successes === undefined ||
      probability === undefined ||
      cumulative === undefined ||
      probability < 0 ||
      probability > 1
    ) {
      return valueError();
    }
    if (!cumulative) {
      return numericResultOrError(negativeBinomialProbability(failures, successes, probability));
    }
    let total = 0;
    for (let index = 0; index <= failures; index += 1) {
      total += negativeBinomialProbability(index, successes, probability);
    }
    return numericResultOrError(total);
  },
  SUBTOTAL: (functionNumArg, ...args) => {
    const functionNum = integerValue(functionNumArg);
    return functionNum === undefined ? valueError() : aggregateByCode(functionNum, args);
  },
  AGGREGATE: (functionNumArg, _optionsArg, ...args) => {
    const functionNum = integerValue(functionNumArg);
    return functionNum === undefined ? valueError() : aggregateByCode(functionNum, args);
  },
  SEQUENCE: (...args) => sequenceResult(args[0], args[1], args[2], args[3]),
  ...externalScalarBuiltins,
  ...scalarPlaceholderBuiltins,
};

const builtins: Record<string, Builtin> = {
  ...scalarBuiltins,
  ...logicalBuiltins,
  ...textBuiltins,
  ...datetimeBuiltins,
};

const builtinIdByName = new Map(
  BUILTINS.map((builtin) => [builtin.name.toUpperCase(), builtin.id]),
);
builtinIdByName.set("USE.THE.COUNTIF", BuiltinId.Countif);
builtinIdByName.set("FORECAST.LINEAR", BuiltinId.Forecast);

export function getBuiltin(name: string): Builtin | undefined {
  return builtins[name.toUpperCase()] ?? getExternalScalarFunction(name);
}

export function hasBuiltin(name: string): boolean {
  const upper = name.toUpperCase();
  return (
    builtins[upper] !== undefined ||
    lookupBuiltins[upper] !== undefined ||
    builtinIdByName.has(upper) ||
    builtinJsSpecialNames.has(upper) ||
    hasExternalFunction(upper)
  );
}

export function getBuiltinId(name: string): BuiltinId | undefined {
  return builtinIdByName.get(name.toUpperCase());
}
