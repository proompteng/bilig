import { BUILTINS, BuiltinId, ErrorCode, ValueTag } from "@bilig/protocol";
import type { CellValue } from "@bilig/protocol";
import { datetimeBuiltins, excelSerialToDateParts } from "./builtins/datetime.js";
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

function romanValue(numberValue: number): string | undefined {
  const number = Math.trunc(numberValue);
  if (!Number.isFinite(numberValue) || number < 1 || number > 3999) {
    return undefined;
  }
  const numerals: Array<[number, string]> = [
    [1000, "M"],
    [900, "CM"],
    [500, "D"],
    [400, "CD"],
    [100, "C"],
    [90, "XC"],
    [50, "L"],
    [40, "XL"],
    [10, "X"],
    [9, "IX"],
    [5, "V"],
    [4, "IV"],
    [1, "I"],
  ];
  let remaining = number;
  let result = "";
  for (const [value, numeral] of numerals) {
    while (remaining >= value) {
      result += numeral;
      remaining -= value;
    }
  }
  return result;
}

function arabicValue(text: string): number | undefined {
  const numerals = new Map<string, number>([
    ["I", 1],
    ["V", 5],
    ["X", 10],
    ["L", 50],
    ["C", 100],
    ["D", 500],
    ["M", 1000],
  ]);
  const upper = text.trim().toUpperCase();
  if (!/^[IVXLCDM]+$/.test(upper)) {
    return undefined;
  }
  let total = 0;
  let index = 0;
  while (index < upper.length) {
    const current = numerals.get(upper[index] ?? "");
    const next = numerals.get(upper[index + 1] ?? "");
    if (current === undefined) {
      return undefined;
    }
    if (next !== undefined && current < next) {
      total += next - current;
      index += 2;
      continue;
    }
    total += current;
    index += 1;
  }
  return romanValue(total) === upper ? total : undefined;
}

function isValidBaseDigits(raw: string, radix: number): boolean {
  const upper = raw.toUpperCase();
  const digits = new Set("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ".slice(0, radix));
  for (const char of upper) {
    if (!digits.has(char)) {
      return false;
    }
  }
  return true;
}

function baseString(
  numberArg: CellValue,
  radixArg: CellValue,
  minLengthArg?: CellValue,
): CellValue {
  const numberValue = integerValue(numberArg);
  const radixValue = integerValue(radixArg);
  const minLengthValue = nonNegativeIntegerValue(minLengthArg, 0);
  if (
    numberValue === undefined ||
    numberValue < 0 ||
    radixValue === undefined ||
    radixValue < 2 ||
    radixValue > 36 ||
    minLengthValue === undefined
  ) {
    return valueError();
  }
  return {
    tag: ValueTag.String,
    value: numberValue.toString(radixValue).toUpperCase().padStart(minLengthValue, "0"),
    stringId: 0,
  };
}

function decimalValue(textArg: CellValue, radixArg: CellValue): CellValue {
  if (textArg.tag === ValueTag.Error) {
    return textArg;
  }
  const radixValue = integerValue(radixArg);
  if (radixValue === undefined || radixValue < 2 || radixValue > 36) {
    return valueError();
  }
  const raw =
    textArg.tag === ValueTag.String
      ? textArg.value.trim()
      : String(Math.trunc(toNumber(textArg) ?? NaN));
  if (raw === "" || raw === "NaN" || !isValidBaseDigits(raw, radixValue)) {
    return valueError();
  }
  return numberResult(Number.parseInt(raw, radixValue));
}

function sampleVariance(numbers: number[]): number {
  if (numbers.length <= 1) {
    return 0;
  }
  const mean = numbers.reduce((sum, value) => sum + value, 0) / numbers.length;
  const squared = numbers.reduce((sum, value) => sum + (value - mean) ** 2, 0);
  return squared / (numbers.length - 1);
}

function populationVariance(numbers: number[]): number {
  if (numbers.length === 0) {
    return 0;
  }
  const mean = numbers.reduce((sum, value) => sum + value, 0) / numbers.length;
  const squared = numbers.reduce((sum, value) => sum + (value - mean) ** 2, 0);
  return squared / numbers.length;
}

function sampleStandardDeviation(numbers: number[]): number {
  const variance = sampleVariance(numbers);
  return variance < 0 ? Number.NaN : Math.sqrt(variance);
}

function populationStandardDeviation(numbers: number[]): number {
  const variance = populationVariance(numbers);
  return variance < 0 ? Number.NaN : Math.sqrt(variance);
}

function collectAStyleNumericArgs(args: CellValue[]): number[] {
  const values: number[] = [];
  for (const arg of args) {
    switch (arg.tag) {
      case ValueTag.Number:
        values.push(arg.value);
        break;
      case ValueTag.Boolean:
        values.push(arg.value ? 1 : 0);
        break;
      case ValueTag.String:
        values.push(0);
        break;
      case ValueTag.Empty:
      case ValueTag.Error:
        break;
    }
  }
  return values;
}

function modeSingle(numbers: number[]): number | undefined {
  const counts = new Map<number, number>();
  let bestValue: number | undefined;
  let bestCount = 1;
  for (const value of numbers) {
    const count = (counts.get(value) ?? 0) + 1;
    counts.set(value, count);
    if (
      count > bestCount ||
      (count === bestCount && bestValue !== undefined && value < bestValue)
    ) {
      bestCount = count;
      bestValue = value;
    }
    if (count > bestCount && bestValue === undefined) {
      bestValue = value;
    }
  }
  return bestCount >= 2 ? bestValue : undefined;
}

function erfApprox(value: number): number {
  const sign = value < 0 ? -1 : 1;
  const absolute = Math.abs(value);
  const t = 1 / (1 + 0.3275911 * absolute);
  const y =
    1 -
    ((((1.061405429 * t - 1.453152027) * t + 1.421413741) * t - 0.284496736) * t + 0.254829592) *
      t *
      Math.exp(-absolute * absolute);
  return sign * y;
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
  BASE: (numberArg, radixArg, minLengthArg) => baseString(numberArg, radixArg, minLengthArg),
  DECIMAL: (textArg, radixArg) => decimalValue(textArg, radixArg),
  ROMAN: (value) => {
    const roman = romanValue(toNumber(value) ?? Number.NaN);
    return roman === undefined ? valueError() : { tag: ValueTag.String, value: roman, stringId: 0 };
  },
  ARABIC: (value) => {
    if (value.tag !== ValueTag.String) {
      return valueError();
    }
    const numeric = arabicValue(value.value);
    return numeric === undefined ? valueError() : numberResult(numeric);
  },
  T: (value = { tag: ValueTag.Empty }) => {
    if (value.tag === ValueTag.Error) {
      return value;
    }
    return value.tag === ValueTag.String ? value : { tag: ValueTag.Empty };
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
  SUBTOTAL: (functionNumArg, ...args) => {
    const functionNum = integerValue(functionNumArg);
    return functionNum === undefined ? valueError() : aggregateByCode(functionNum, args);
  },
  AGGREGATE: (functionNumArg, _optionsArg, ...args) => {
    const functionNum = integerValue(functionNumArg);
    return functionNum === undefined ? valueError() : aggregateByCode(functionNum, args);
  },
  SEQUENCE: (...args) => sequenceResult(args[0], args[1], args[2], args[3]),
  ...scalarPlaceholderBuiltins,
};

const builtins: Record<string, Builtin> = {
  ...scalarBuiltins,
  ...logicalBuiltins,
  ...textBuiltins,
  ...datetimeBuiltins,
};

const jsSpecialBuiltins = new Set([
  "LET",
  "LAMBDA",
  "MAKEARRAY",
  "MAP",
  "REDUCE",
  "SCAN",
  "BYROW",
  "BYCOL",
]);

const builtinIdByName = new Map(
  BUILTINS.map((builtin) => [builtin.name.toUpperCase(), builtin.id]),
);

export function getBuiltin(name: string): Builtin | undefined {
  return builtins[name.toUpperCase()] ?? getExternalScalarFunction(name);
}

export function hasBuiltin(name: string): boolean {
  const upper = name.toUpperCase();
  return (
    builtins[upper] !== undefined ||
    lookupBuiltins[upper] !== undefined ||
    jsSpecialBuiltins.has(upper) ||
    hasExternalFunction(upper)
  );
}

export function getBuiltinId(name: string): BuiltinId | undefined {
  return builtinIdByName.get(name.toUpperCase());
}
