import { BuiltinId, ErrorCode, ValueTag } from "@bilig/protocol";
import type { CellValue } from "@bilig/protocol";
import { datetimeBuiltins } from "./builtins/datetime.js";
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

function nonNegativeIntegerValue(value: CellValue | undefined, fallback?: number): number | undefined {
  const integer = integerValue(value, fallback);
  return integer !== undefined && integer >= 0 ? integer : undefined;
}

function positiveIntegerValue(value: CellValue | undefined, fallback?: number): number | undefined {
  const integer = integerValue(value, fallback);
  return integer !== undefined && integer >= 1 ? integer : undefined;
}

function sequenceResult(rowsArg: CellValue | undefined, colsArg: CellValue | undefined, startArg: CellValue | undefined, stepArg: CellValue | undefined): ArrayValue | CellValue {
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
    values
  };
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

function binaryMath(left: CellValue, right: CellValue, evaluate: (left: number, right: number) => number): CellValue {
  return numericResultOrError(evaluate(toNumber(left) ?? 0, toNumber(right) ?? 0));
}

function collectNumericArgs(args: CellValue[]): number[] {
  return args.map(toNumber).filter((value): value is number => value !== undefined);
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
    [1000, "M"], [900, "CM"], [500, "D"], [400, "CD"],
    [100, "C"], [90, "XC"], [50, "L"], [40, "XL"],
    [10, "X"], [9, "IX"], [5, "V"], [4, "IV"], [1, "I"]
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
    ["I", 1], ["V", 5], ["X", 10], ["L", 50], ["C", 100], ["D", 500], ["M", 1000]
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
  const digits = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ".slice(0, radix);
  for (const char of Array.from(upper)) {
    if (!digits.includes(char)) {
      return false;
    }
  }
  return true;
}

function baseString(numberArg: CellValue, radixArg: CellValue, minLengthArg?: CellValue): CellValue {
  const numberValue = integerValue(numberArg);
  const radixValue = integerValue(radixArg);
  const minLengthValue = nonNegativeIntegerValue(minLengthArg, 0);
  if (
    numberValue === undefined
    || numberValue < 0
    || radixValue === undefined
    || radixValue < 2
    || radixValue > 36
    || minLengthValue === undefined
  ) {
    return valueError();
  }
  return {
    tag: ValueTag.String,
    value: numberValue.toString(radixValue).toUpperCase().padStart(minLengthValue, "0"),
    stringId: 0
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
  const raw = textArg.tag === ValueTag.String ? textArg.value.trim() : String(Math.trunc(toNumber(textArg) ?? NaN));
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

function aggregateByCode(functionNum: number, values: CellValue[]): CellValue {
  const normalized = functionNum > 100 ? functionNum - 100 : functionNum;
  const numericValues = collectNumericArgs(values);
  switch (normalized) {
    case 1:
      return numericValues.length === 0
        ? numberResult(0)
        : numberResult(numericValues.reduce((sum, value) => sum + value, 0) / numericValues.length);
    case 2:
      return numberResult(values.filter((value) => value.tag === ValueTag.Number || value.tag === ValueTag.Boolean).length);
    case 3:
      return numberResult(values.filter((value) => value.tag !== ValueTag.Empty).length);
    case 4:
      return numberResult(numericValues.length === 0 ? 0 : Math.max(...numericValues));
    case 5:
      return numberResult(numericValues.length === 0 ? 0 : Math.min(...numericValues));
    case 6:
      return numberResult(numericValues.length === 0 ? 0 : numericValues.reduce((product, value) => product * value, 1));
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
  MIN: (...args) => numberResult(Math.min(...args.map((arg) => toNumber(arg) ?? Number.POSITIVE_INFINITY))),
  MAX: (...args) => numberResult(Math.max(...args.map((arg) => toNumber(arg) ?? Number.NEGATIVE_INFINITY))),
  COUNT: (...args) => numberResult(args.filter((arg) => arg.tag === ValueTag.Number || arg.tag === ValueTag.Boolean).length),
  COUNTA: (...args) => numberResult(args.filter((arg) => arg.tag !== ValueTag.Empty).length),
  ABS: (value) => numberResult(Math.abs(toNumber(value) ?? 0)),
  SIN: (value) => unaryMath(value, Math.sin),
  COS: (value) => unaryMath(value, Math.cos),
  TAN: (value) => unaryMath(value, Math.tan),
  ASIN: (value) => unaryMath(value, Math.asin),
  ACOS: (value) => unaryMath(value, Math.acos),
  ATAN: (value) => unaryMath(value, Math.atan),
  ATAN2: (left, right) => binaryMath(left, right, Math.atan2),
  DEGREES: (value) => unaryMath(value, (numeric) => numeric * 180 / Math.PI),
  RADIANS: (value) => unaryMath(value, (numeric) => numeric * Math.PI / 180),
  EXP: (value) => unaryMath(value, Math.exp),
  LN: (value) => unaryMath(value, Math.log),
  LOG10: (value) => unaryMath(value, Math.log10),
  LOG: (value, base) => {
    const numeric = toNumber(value) ?? 0;
    const baseValue = base === undefined ? 10 : (toNumber(base) ?? 10);
    const result = base === undefined ? Math.log10(numeric) : Math.log(numeric) / Math.log(baseValue);
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
    return tangent === 0 ? { tag: ValueTag.Error, code: ErrorCode.Div0 } : numberResult(1 / tangent);
  },
  COTH: (value) => {
    const hyperbolic = Math.tanh(toNumber(value) ?? 0);
    return hyperbolic === 0 ? { tag: ValueTag.Error, code: ErrorCode.Div0 } : numberResult(1 / hyperbolic);
  },
  CSC: (value) => {
    const sine = Math.sin(toNumber(value) ?? 0);
    return sine === 0 ? { tag: ValueTag.Error, code: ErrorCode.Div0 } : numberResult(1 / sine);
  },
  CSCH: (value) => {
    const hyperbolic = Math.sinh(toNumber(value) ?? 0);
    return hyperbolic === 0 ? { tag: ValueTag.Error, code: ErrorCode.Div0 } : numberResult(1 / hyperbolic);
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
    const significanceValue = Math.abs(toNumber(significance ?? { tag: ValueTag.Number, value: 1 }) ?? 1);
    const modeValue = toNumber(mode ?? { tag: ValueTag.Number, value: 0 }) ?? 0;
    if (numberValue === undefined || significanceValue === 0) {
      return valueError();
    }
    if (numberValue >= 0) {
      return numberResult(Math.floor(numberValue / significanceValue) * significanceValue);
    }
    const magnitude = modeValue === 0 ? Math.ceil(Math.abs(numberValue) / significanceValue) : Math.floor(Math.abs(numberValue) / significanceValue);
    return numberResult(-magnitude * significanceValue);
  },
  "FLOOR.PRECISE": (value, significance) => {
    const numberValue = toNumber(value);
    const significanceValue = Math.abs(toNumber(significance ?? { tag: ValueTag.Number, value: 1 }) ?? 1);
    if (numberValue === undefined || significanceValue === 0) {
      return valueError();
    }
    return numberResult(Math.floor(numberValue / significanceValue) * significanceValue);
  },
  "CEILING.MATH": (value, significance, mode) => {
    const numberValue = toNumber(value);
    const significanceValue = Math.abs(toNumber(significance ?? { tag: ValueTag.Number, value: 1 }) ?? 1);
    const modeValue = toNumber(mode ?? { tag: ValueTag.Number, value: 0 }) ?? 0;
    if (numberValue === undefined || significanceValue === 0) {
      return valueError();
    }
    if (numberValue >= 0) {
      return numberResult(Math.ceil(numberValue / significanceValue) * significanceValue);
    }
    const magnitude = modeValue === 0 ? Math.floor(Math.abs(numberValue) / significanceValue) : Math.ceil(Math.abs(numberValue) / significanceValue);
    return numberResult(-magnitude * significanceValue);
  },
  "CEILING.PRECISE": (value, significance) => {
    const numberValue = toNumber(value);
    const significanceValue = Math.abs(toNumber(significance ?? { tag: ValueTag.Number, value: 1 }) ?? 1);
    if (numberValue === undefined || significanceValue === 0) {
      return valueError();
    }
    return numberResult(Math.ceil(numberValue / significanceValue) * significanceValue);
  },
  "ISO.CEILING": (value, significance) => {
    const numberValue = toNumber(value);
    const significanceValue = Math.abs(toNumber(significance ?? { tag: ValueTag.Number, value: 1 }) ?? 1);
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
    if (
      numberValue === undefined
      || chosenValue === undefined
      || chosenValue > numberValue
    ) {
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
    const denominator = numbers.reduce((product, value) => product * (factorialValue(value) ?? Number.NaN), 1);
    return numerator === undefined || Number.isNaN(denominator) ? valueError() : numberResult(numerator / denominator);
  },
  PRODUCT: (...args) => {
    const error = firstError(args);
    if (error) return error;
    const numbers = collectNumericArgs(args);
    return numberResult(numbers.length === 0 ? 0 : numbers.reduce((product, value) => product * value, 1));
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
    if (rows === undefined || cols === undefined || min === undefined || max === undefined || max < min) {
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
      sum += coefficient * (x ** (n + index * m));
    });
    return numberResult(sum);
  },
  SQRTPI: (value) => {
    const numeric = toNumber(value);
    return numeric === undefined ? valueError() : numericResultOrError(Math.sqrt(numeric * Math.PI));
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
  SUBTOTAL: (functionNumArg, ...args) => {
    const functionNum = integerValue(functionNumArg);
    return functionNum === undefined ? valueError() : aggregateByCode(functionNum, args);
  },
  AGGREGATE: (functionNumArg, _optionsArg, ...args) => {
    const functionNum = integerValue(functionNumArg);
    return functionNum === undefined ? valueError() : aggregateByCode(functionNum, args);
  },
  SEQUENCE: (...args) => sequenceResult(args[0], args[1], args[2], args[3]),
  ...scalarPlaceholderBuiltins
};

const builtins: Record<string, Builtin> = {
  ...scalarBuiltins,
  ...logicalBuiltins,
  ...textBuiltins,
  ...datetimeBuiltins
};

const jsSpecialBuiltins = new Set(["LET", "LAMBDA", "MAKEARRAY", "MAP", "REDUCE", "SCAN", "BYROW", "BYCOL"]);

function isBuiltinIdKey(value: string): value is keyof typeof BuiltinId {
  return value in BuiltinId;
}

export function getBuiltin(name: string): Builtin | undefined {
  return builtins[name.toUpperCase()] ?? getExternalScalarFunction(name);
}

export function hasBuiltin(name: string): boolean {
  const upper = name.toUpperCase();
  return builtins[upper] !== undefined || lookupBuiltins[upper] !== undefined || jsSpecialBuiltins.has(upper) || hasExternalFunction(upper);
}

export function getBuiltinId(name: string): BuiltinId | undefined {
  const first = name.charAt(0);
  if (!first) return undefined;
  const key = `${first.toUpperCase()}${name.slice(1).toLowerCase()}`;
  if (!isBuiltinIdKey(key)) {
    return undefined;
  }
  const builtinId = BuiltinId[key];
  return typeof builtinId === "number" ? builtinId : undefined;
}
