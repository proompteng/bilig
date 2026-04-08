import { BUILTINS, BuiltinId, ErrorCode, ValueTag } from "@bilig/protocol";
import type { CellValue } from "@bilig/protocol";
import { builtinJsSpecialNames } from "./builtin-capabilities.js";
import { createComplexBuiltins } from "./builtins/complex.js";
import {
  besselIValue,
  besselJValue,
  besselKValue,
  besselYValue,
} from "./builtins/distributions.js";
import { createDistributionBuiltins } from "./builtins/distribution-builtins.js";
import { createFinancialBuiltins } from "./builtins/financial-builtins.js";
import { createFixedIncomeBuiltins } from "./builtins/fixed-income-builtins.js";
import {
  countLeadingZeros,
  formatFixed,
  isValidDollarFraction,
  parseDollarDecimal,
  toColumnLabel,
} from "./builtins/formatting.js";
import {
  buildIdentityMatrix,
  collectNumericArgs,
  createNumericBuiltinHelpers,
  doubleFactorialValue,
  evenValue,
  factorialValue,
  gcdPair,
  lcmPair,
  oddValue,
  roundDownToDigits,
  roundTowardZero,
  roundUpToDigits,
} from "./builtins/numeric.js";
import { createRadixBuiltins } from "./builtins/radix.js";
import { populationVariance, sampleVariance } from "./builtins/statistics.js";
import { createStatisticalBuiltins } from "./builtins/statistical-builtins.js";
import { datetimeBuiltins } from "./builtins/datetime.js";
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

const complexBuiltins = createComplexBuiltins({ toNumber, numberResult, valueError });
const numericBuiltinHelpers = createNumericBuiltinHelpers({
  toNumber,
  numberResult,
  valueError,
  numericResultOrError,
});
const { binaryMath, ceilingWith, floorWith, roundWith, unaryMath } = numericBuiltinHelpers;
const radixBuiltins = createRadixBuiltins({
  toNumber,
  integerValue,
  nonNegativeIntegerValue,
  valueError,
  numberResult,
});
const fixedIncomeBuiltins = createFixedIncomeBuiltins({
  toNumber,
  coerceBoolean,
  coerceDateSerial,
  coerceNumber,
  integerValue,
  numberResult,
  valueError,
});
const statisticalBuiltins = createStatisticalBuiltins({
  toNumber,
  coerceBoolean,
  firstError,
  numberResult,
  numericResultOrError,
  valueError,
});
const distributionBuiltins = createDistributionBuiltins({
  toNumber,
  coerceBoolean,
  coerceNumber,
  integerValue,
  nonNegativeIntegerValue,
  positiveIntegerValue,
  numberResult,
  numericResultOrError,
  valueError,
});
const financialBuiltins = createFinancialBuiltins({
  toNumber,
  coerceBoolean,
  coerceNumber,
  coercePaymentType,
  integerValue,
  positiveIntegerValue,
  numberResult,
  numericResultOrError,
  valueError,
});

function coercePaymentType(value: CellValue | undefined, fallback: number): number | undefined {
  const type = integerValue(value, fallback);
  return type === 0 || type === 1 ? type : undefined;
}

function toZeroNumericValue(value: CellValue): number | undefined {
  if (value.tag === ValueTag.String) {
    return 0;
  }
  return toNumber(value);
}

function aggregateByCode(functionNum: number, values: CellValue[]): CellValue {
  const normalized = functionNum > 100 ? functionNum - 100 : functionNum;
  const numericValues = collectNumericArgs(values, toNumber);
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
    const numbers = collectNumericArgs(args, toNumber);
    if (numbers.length === 0) return numberResult(0);
    return numberResult(numbers.reduce((sum, value) => sum + value, 0) / numbers.length);
  },
  AVG: (...args) => {
    const error = firstError(args);
    if (error) return error;
    const numbers = collectNumericArgs(args, toNumber);
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
    const numbers = collectNumericArgs(args, toNumber);
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
    const numbers = collectNumericArgs(args, toNumber);
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
    const numbers = collectNumericArgs(args, toNumber).map((value) => Math.abs(Math.trunc(value)));
    if (numbers.length === 0) {
      return valueError();
    }
    return numberResult(numbers.reduce((acc, value) => gcdPair(acc, value)));
  },
  LCM: (...args) => {
    const numbers = collectNumericArgs(args, toNumber).map((value) => Math.abs(Math.trunc(value)));
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
    const numbers = collectNumericArgs(args, toNumber).map((value) => Math.trunc(value));
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
    const numbers = collectNumericArgs(args, toNumber);
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
  ...fixedIncomeBuiltins,
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
    return size === undefined ? valueError() : buildIdentityMatrix(size, numberResult);
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
    return numberResult(
      collectNumericArgs(args, toNumber).reduce((sum, value) => sum + value ** 2, 0),
    );
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
  ...statisticalBuiltins,
  ...financialBuiltins,
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
  ...distributionBuiltins,
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
