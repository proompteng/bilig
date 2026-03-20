import { BuiltinId, ErrorCode, ValueTag } from "@bilig/protocol";
import type { CellValue } from "@bilig/protocol";
import { datetimeBuiltins } from "./builtins/datetime.js";
import { logicalBuiltins } from "./builtins/logical.js";
import { lookupBuiltins } from "./builtins/lookup.js";
import { textBuiltins } from "./builtins/text.js";

type Builtin = (...args: CellValue[]) => CellValue;

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

function firstError(args: CellValue[]): CellValue | undefined {
  return args.find((arg) => arg.tag === ValueTag.Error);
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

function roundWith(value: CellValue, digits: CellValue | undefined): CellValue {
  const numberValue = toNumber(value);
  const digitValue = digits === undefined ? 0 : toNumber(digits);
  if (numberValue === undefined || digitValue === undefined) {
    return { tag: ValueTag.Error, code: ErrorCode.Value };
  }
  return numberResult(roundToDigits(numberValue, Math.trunc(digitValue)));
}

function floorWith(value: CellValue, significance?: CellValue): CellValue {
  const numberValue = toNumber(value);
  const significanceValue = significance === undefined ? 1 : toNumber(significance);
  if (numberValue === undefined || significanceValue === undefined) {
    return { tag: ValueTag.Error, code: ErrorCode.Value };
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
    return { tag: ValueTag.Error, code: ErrorCode.Value };
  }
  if (significanceValue === 0) {
    return { tag: ValueTag.Error, code: ErrorCode.Div0 };
  }
  return numberResult(Math.ceil(numberValue / significanceValue) * significanceValue);
}

const scalarBuiltins: Record<string, Builtin> = {
  SUM: (...args) => {
    const error = firstError(args);
    if (error) return error;
    return numberResult(args.reduce((sum, arg) => sum + (toNumber(arg) ?? 0), 0));
  },
  AVERAGE: (...args) => {
    const error = firstError(args);
    if (error) return error;
    const numbers = args.map(toNumber).filter((value): value is number => value !== undefined);
    if (numbers.length === 0) return numberResult(0);
    return numberResult(numbers.reduce((sum, value) => sum + value, 0) / numbers.length);
  },
  AVG: (...args) => {
    const error = firstError(args);
    if (error) return error;
    const numbers = args.map(toNumber).filter((value): value is number => value !== undefined);
    if (numbers.length === 0) return numberResult(0);
    return numberResult(numbers.reduce((sum, value) => sum + value, 0) / numbers.length);
  },
  MIN: (...args) => numberResult(Math.min(...args.map((arg) => toNumber(arg) ?? Number.POSITIVE_INFINITY))),
  MAX: (...args) => numberResult(Math.max(...args.map((arg) => toNumber(arg) ?? Number.NEGATIVE_INFINITY))),
  COUNT: (...args) => numberResult(args.filter((arg) => arg.tag === ValueTag.Number || arg.tag === ValueTag.Boolean).length),
  COUNTA: (...args) => numberResult(args.filter((arg) => arg.tag !== ValueTag.Empty).length),
  ABS: (value) => numberResult(Math.abs(toNumber(value) ?? 0)),
  ROUND: (value, digits) => roundWith(value, digits),
  FLOOR: (value, significance) => floorWith(value, significance),
  CEILING: (value, significance) => ceilingWith(value, significance),
  MOD: (left, right) => {
    const divisor = toNumber(right) ?? 0;
    if (divisor === 0) return { tag: ValueTag.Error, code: ErrorCode.Div0 };
    return numberResult((toNumber(left) ?? 0) % divisor);
  },
  INT: (value) => {
    const numberValue = toNumber(value);
    if (numberValue === undefined) {
      return { tag: ValueTag.Error, code: ErrorCode.Value };
    }
    return numberResult(Math.floor(numberValue));
  },
  ROUNDUP: (value, digits) => {
    const numberValue = toNumber(value);
    const digitValue = digits === undefined ? 0 : toNumber(digits);
    if (numberValue === undefined || digitValue === undefined) {
      return { tag: ValueTag.Error, code: ErrorCode.Value };
    }
    return numberResult(roundUpToDigits(numberValue, Math.trunc(digitValue)));
  },
  ROUNDDOWN: (value, digits) => {
    const numberValue = toNumber(value);
    const digitValue = digits === undefined ? 0 : toNumber(digits);
    if (numberValue === undefined || digitValue === undefined) {
      return { tag: ValueTag.Error, code: ErrorCode.Value };
    }
    return numberResult(roundDownToDigits(numberValue, Math.trunc(digitValue)));
  }
};

const builtins: Record<string, Builtin> = {
  ...scalarBuiltins,
  ...logicalBuiltins,
  ...textBuiltins,
  ...datetimeBuiltins
};

function isBuiltinIdKey(value: string): value is keyof typeof BuiltinId {
  return value in BuiltinId;
}

export function getBuiltin(name: string): Builtin | undefined {
  return builtins[name.toUpperCase()];
}

export function hasBuiltin(name: string): boolean {
  const upper = name.toUpperCase();
  return builtins[upper] !== undefined || lookupBuiltins[upper] !== undefined;
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
