import { ErrorCode, ValueTag } from "@bilig/protocol";
import type { CellValue } from "@bilig/protocol";
import { besselIValue, besselJValue, besselKValue, besselYValue } from "./distributions.js";
import { collectNumericArgs } from "./numeric.js";
import type { EvaluationResult } from "../runtime-values.js";

type Builtin = (...args: CellValue[]) => EvaluationResult;

interface MathBuiltinDeps {
  toNumber: (value: CellValue) => number | undefined;
  toBitwiseUnsigned: (value: CellValue | undefined) => number | undefined;
  coerceShiftAmount: (value: CellValue | undefined) => number | undefined;
  integerValue: (value: CellValue | undefined, fallback?: number) => number | undefined;
  nonNegativeIntegerValue: (value: CellValue | undefined, fallback?: number) => number | undefined;
  firstError: (args: CellValue[]) => CellValue | undefined;
  numberResult: (value: number) => EvaluationResult;
  valueError: () => EvaluationResult;
  numError: () => EvaluationResult;
  numericResultOrError: (value: number) => EvaluationResult;
  unaryMath: (value: CellValue, operation: (numeric: number) => number) => EvaluationResult;
  binaryMath: (
    left: CellValue,
    right: CellValue,
    operation: (leftNumeric: number, rightNumeric: number) => number,
  ) => EvaluationResult;
  ceilingWith: (value: CellValue, significance: CellValue) => EvaluationResult;
  floorWith: (value: CellValue, significance: CellValue) => EvaluationResult;
  roundWith: (value: CellValue, digits: CellValue) => EvaluationResult;
  roundUpToDigits: (value: number, digits: number) => number;
  roundDownToDigits: (value: number, digits: number) => number;
  roundTowardZero: (value: number, digits: number) => number;
  evenValue: (value: number) => number;
  oddValue: (value: number) => number;
  factorialValue: (value: number) => number | undefined;
  doubleFactorialValue: (value: number) => number | undefined;
  gcdPair: (left: number, right: number) => number;
  lcmPair: (left: number, right: number) => number;
}

function div0Error(): EvaluationResult {
  return { tag: ValueTag.Error, code: ErrorCode.Div0 };
}

export function createMathBuiltins({
  toNumber,
  toBitwiseUnsigned,
  coerceShiftAmount,
  integerValue,
  nonNegativeIntegerValue,
  firstError,
  numberResult,
  valueError,
  numError,
  numericResultOrError,
  unaryMath,
  binaryMath,
  ceilingWith,
  floorWith,
  roundWith,
  roundUpToDigits,
  roundDownToDigits,
  roundTowardZero,
  evenValue,
  oddValue,
  factorialValue,
  doubleFactorialValue,
  gcdPair,
  lcmPair,
}: MathBuiltinDeps): Record<string, Builtin> {
  return {
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
      const numeric = toNumber(value);
      if (numeric === undefined) {
        return valueError();
      }
      const baseValue = base === undefined ? 10 : toNumber(base);
      if (baseValue === undefined) {
        return valueError();
      }
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
      const numeric = toNumber(value);
      if (numeric === undefined) {
        return valueError();
      }
      return numberResult(numeric === 0 ? Math.PI / 2 : Math.atan(1 / numeric));
    },
    ACOTH: (value) => {
      const numeric = toNumber(value);
      if (numeric === undefined) {
        return valueError();
      }
      return numericResultOrError(0.5 * Math.log((numeric + 1) / (numeric - 1)));
    },
    COT: (value) => {
      const numeric = toNumber(value);
      if (numeric === undefined) {
        return valueError();
      }
      const tangent = Math.tan(numeric);
      return tangent === 0 ? div0Error() : numberResult(1 / tangent);
    },
    COTH: (value) => {
      const numeric = toNumber(value);
      if (numeric === undefined) {
        return valueError();
      }
      const hyperbolic = Math.tanh(numeric);
      return hyperbolic === 0 ? div0Error() : numberResult(1 / hyperbolic);
    },
    CSC: (value) => {
      const numeric = toNumber(value);
      if (numeric === undefined) {
        return valueError();
      }
      const sine = Math.sin(numeric);
      return sine === 0 ? div0Error() : numberResult(1 / sine);
    },
    CSCH: (value) => {
      const numeric = toNumber(value);
      if (numeric === undefined) {
        return valueError();
      }
      const hyperbolic = Math.sinh(numeric);
      return hyperbolic === 0 ? div0Error() : numberResult(1 / hyperbolic);
    },
    SEC: (value) => {
      const numeric = toNumber(value);
      if (numeric === undefined) {
        return valueError();
      }
      const cosine = Math.cos(numeric);
      return cosine === 0 ? div0Error() : numberResult(1 / cosine);
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
      if (divisor === 0) {
        return div0Error();
      }
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
      const numbers = collectNumericArgs(args, toNumber).map((value) =>
        Math.abs(Math.trunc(value)),
      );
      if (numbers.length === 0) {
        return valueError();
      }
      return numberResult(numbers.reduce((acc, value) => gcdPair(acc, value)));
    },
    LCM: (...args) => {
      const numbers = collectNumericArgs(args, toNumber).map((value) =>
        Math.abs(Math.trunc(value)),
      );
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
      if (error) {
        return error;
      }
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
        return div0Error();
      }
      return numberResult(Math.trunc(numerator / denominator));
    },
  };
}
