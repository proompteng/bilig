import { BUILTINS, BuiltinId, ErrorCode, ValueTag } from "@bilig/protocol";
import type { CellValue } from "@bilig/protocol";
import { builtinJsSpecialNames } from "./builtin-capabilities.js";
import { createComplexBuiltins } from "./builtins/complex.js";
import {
  betaDistributionCdf,
  betaDistributionDensity,
  betaDistributionInverse,
  binomialProbability,
  besselIValue,
  besselJValue,
  besselKValue,
  besselYValue,
  chiSquareCdf,
  chiSquareDensity,
  fDistributionCdf,
  fDistributionDensity,
  gammaDistributionCdf,
  gammaDistributionDensity,
  gammaFunction,
  hypergeometricProbability,
  inverseChiSquare,
  inverseFDistribution,
  inverseGammaDistribution,
  inverseNormal,
  inverseStandardNormal,
  inverseStudentT,
  kurtosis,
  logGamma,
  negativeBinomialProbability,
  percentileNormal,
  poissonProbability,
  regularizedUpperGamma,
  skewPopulation,
  skewSample,
  standardNormalCdf,
  standardNormalPdf,
  studentTCdf,
  studentTDensity,
} from "./builtins/distributions.js";
import {
  cumulativePeriodicPayment,
  dbDepreciation,
  ddbDepreciation,
  futureValue,
  interestPayment,
  periodicPayment,
  presentValue,
  principalPayment,
  solveRate,
  totalPeriods,
  vdbDepreciation,
} from "./builtins/financial.js";
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
  collectStatNumericArgs,
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
import {
  collectAStyleNumericArgs,
  erfApprox,
  modeSingle,
  populationStandardDeviation,
  populationVariance,
  sampleStandardDeviation,
  sampleVariance,
} from "./builtins/statistics.js";
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
    const mode = modeSingle(collectNumericArgs(args, toNumber));
    return mode === undefined ? { tag: ValueTag.Error, code: ErrorCode.NA } : numberResult(mode);
  },
  "MODE.SNGL": (...args) => {
    const error = firstError(args);
    if (error) return error;
    const mode = modeSingle(collectNumericArgs(args, toNumber));
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
