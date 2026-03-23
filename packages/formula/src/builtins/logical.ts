import { ErrorCode, ValueTag, type CellValue } from "@bilig/protocol";
import { createBlockedBuiltinMap, logicalPlaceholderBuiltinNames } from "./placeholder.js";

export type LogicalBuiltin = (...args: CellValue[]) => CellValue;

type LogicalCoercion =
  | { ok: true; value: boolean }
  | { ok: false; error: CellValue };

function emptyValue(): CellValue {
  return { tag: ValueTag.Empty };
}

function errorValue(code: ErrorCode): CellValue {
  return { tag: ValueTag.Error, code };
}

function booleanResult(value: boolean): CellValue {
  return { tag: ValueTag.Boolean, value };
}

function isError(value: CellValue): value is Extract<CellValue, { tag: ValueTag.Error }> {
  return value.tag === ValueTag.Error;
}

function coerceLogical(value: CellValue): LogicalCoercion {
  switch (value.tag) {
    case ValueTag.Boolean:
      return { ok: true, value: value.value };
    case ValueTag.Number:
      return { ok: true, value: value.value !== 0 };
    case ValueTag.Empty:
      return { ok: true, value: false };
    case ValueTag.Error:
      return { ok: false, error: value };
    case ValueTag.String:
    default:
      return { ok: false, error: errorValue(ErrorCode.Value) };
  }
}

const logicalPlaceholderBuiltins = createBlockedBuiltinMap(logicalPlaceholderBuiltinNames);

export const logicalBuiltins: Record<string, LogicalBuiltin> = {
  NA: (...args) => {
    if (args.length > 0) {
      return errorValue(ErrorCode.Value);
    }
    return errorValue(ErrorCode.NA);
  },
  IF: (condition, truthy, falsy = emptyValue()) => {
    if (condition === undefined || truthy === undefined) {
      return errorValue(ErrorCode.Value);
    }

    const coerced = coerceLogical(condition);
    if (!coerced.ok) {
      return coerced.error;
    }

    return coerced.value ? truthy : falsy;
  },
  IFERROR: (value, valueIfError = emptyValue()) => {
    if (value === undefined) {
      return errorValue(ErrorCode.Value);
    }
    return isError(value) ? valueIfError : value;
  },
  IFNA: (value, valueIfNa = emptyValue()) => {
    if (value === undefined) {
      return errorValue(ErrorCode.Value);
    }
    return isError(value) && value.code === ErrorCode.NA ? valueIfNa : value;
  },
  AND: (...args) => {
    if (args.length === 0) {
      return errorValue(ErrorCode.Value);
    }

    for (const arg of args) {
      const coerced = coerceLogical(arg);
      if (!coerced.ok) {
        return coerced.error;
      }
      if (!coerced.value) {
        return booleanResult(false);
      }
    }

    return booleanResult(true);
  },
  OR: (...args) => {
    if (args.length === 0) {
      return errorValue(ErrorCode.Value);
    }

    for (const arg of args) {
      const coerced = coerceLogical(arg);
      if (!coerced.ok) {
        return coerced.error;
      }
      if (coerced.value) {
        return booleanResult(true);
      }
    }

    return booleanResult(false);
  },
  NOT: (value) => {
    if (value === undefined) {
      return errorValue(ErrorCode.Value);
    }

    const coerced = coerceLogical(value);
    if (!coerced.ok) {
      return coerced.error;
    }

    return booleanResult(!coerced.value);
  },
  ISBLANK: (value = emptyValue()) => booleanResult(value.tag === ValueTag.Empty),
  ISNUMBER: (value = emptyValue()) => booleanResult(value.tag === ValueTag.Number),
  ISTEXT: (value = emptyValue()) => booleanResult(value.tag === ValueTag.String),
  ...logicalPlaceholderBuiltins
};

export function getLogicalBuiltin(name: string): LogicalBuiltin | undefined {
  return logicalBuiltins[name.toUpperCase()];
}
