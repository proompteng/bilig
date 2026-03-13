import { BuiltinId, ErrorCode, ValueTag } from "@bilig/protocol";
import type { CellValue } from "@bilig/protocol";

type Builtin = (...args: CellValue[]) => CellValue;

function toNumber(value: CellValue): number | undefined {
  switch (value.tag) {
    case ValueTag.Number:
      return value.value;
    case ValueTag.Boolean:
      return value.value ? 1 : 0;
    case ValueTag.Empty:
      return 0;
    default:
      return undefined;
  }
}

function toStringValue(value: CellValue): string {
  switch (value.tag) {
    case ValueTag.Empty:
      return "";
    case ValueTag.Number:
      return String(value.value);
    case ValueTag.Boolean:
      return value.value ? "TRUE" : "FALSE";
    case ValueTag.String:
      return value.value;
    case ValueTag.Error:
      return `#${ErrorCode[value.code]}`;
  }
}

function numberResult(value: number): CellValue {
  return { tag: ValueTag.Number, value };
}

function booleanResult(value: boolean): CellValue {
  return { tag: ValueTag.Boolean, value };
}

function stringResult(value: string): CellValue {
  return { tag: ValueTag.String, value, stringId: 0 };
}

function firstError(args: CellValue[]): CellValue | undefined {
  return args.find((arg) => arg.tag === ValueTag.Error);
}

const builtins: Record<string, Builtin> = {
  SUM: (...args) => {
    const error = firstError(args);
    if (error) return error;
    return numberResult(args.reduce((sum, arg) => sum + (toNumber(arg) ?? 0), 0));
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
  COUNT: (...args) => numberResult(args.filter((arg) => toNumber(arg) !== undefined).length),
  COUNTA: (...args) => numberResult(args.filter((arg) => arg.tag !== ValueTag.Empty).length),
  ABS: (value) => numberResult(Math.abs(toNumber(value) ?? 0)),
  ROUND: (value) => numberResult(Math.round(toNumber(value) ?? 0)),
  FLOOR: (value) => numberResult(Math.floor(toNumber(value) ?? 0)),
  CEILING: (value) => numberResult(Math.ceil(toNumber(value) ?? 0)),
  MOD: (left, right) => {
    const divisor = toNumber(right) ?? 0;
    if (divisor === 0) return { tag: ValueTag.Error, code: ErrorCode.Div0 };
    return numberResult((toNumber(left) ?? 0) % divisor);
  },
  IF: (condition, truthy, falsy = { tag: ValueTag.Empty }) =>
    (toNumber(condition) ?? 0) !== 0 ? truthy : falsy,
  AND: (...args) => booleanResult(args.every((arg) => (toNumber(arg) ?? 0) !== 0)),
  OR: (...args) => booleanResult(args.some((arg) => (toNumber(arg) ?? 0) !== 0)),
  NOT: (value) => booleanResult((toNumber(value) ?? 0) === 0),
  LEN: (value) => numberResult(toStringValue(value).length),
  CONCAT: (...args) => stringResult(args.map(toStringValue).join(""))
};

export function getBuiltin(name: string): Builtin | undefined {
  return builtins[name.toUpperCase()];
}

export function getBuiltinId(name: string): BuiltinId | undefined {
  const first = name.charAt(0);
  if (!first) return undefined;
  return BuiltinId[`${first.toUpperCase()}${name.slice(1).toLowerCase()}` as keyof typeof BuiltinId];
}
