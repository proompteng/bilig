import { ErrorCode, ValueTag, type CellValue } from "@bilig/protocol";
import { createBlockedBuiltinMap, logicalPlaceholderBuiltinNames } from "./placeholder.js";

export type LogicalBuiltin = (...args: CellValue[]) => CellValue;

type LogicalCoercion = { ok: true; value: boolean } | { ok: false; error: CellValue };

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

function compareText(left: string, right: string): number {
  const normalizedLeft = left.toUpperCase();
  const normalizedRight = right.toUpperCase();
  if (normalizedLeft === normalizedRight) {
    return 0;
  }
  return normalizedLeft < normalizedRight ? -1 : 1;
}

function compareScalars(left: CellValue, right: CellValue): number | undefined {
  const leftTextLike = left.tag === ValueTag.String || left.tag === ValueTag.Empty;
  const rightTextLike = right.tag === ValueTag.String || right.tag === ValueTag.Empty;
  if (leftTextLike && rightTextLike) {
    return compareText(
      left.tag === ValueTag.String ? left.value : "",
      right.tag === ValueTag.String ? right.value : "",
    );
  }
  const leftNumeric =
    left.tag === ValueTag.Boolean
      ? left.value
        ? 1
        : 0
      : left.tag === ValueTag.Empty
        ? 0
        : left.tag === ValueTag.Number
          ? left.value
          : undefined;
  const rightNumeric =
    right.tag === ValueTag.Boolean
      ? right.value
        ? 1
        : 0
      : right.tag === ValueTag.Empty
        ? 0
        : right.tag === ValueTag.Number
          ? right.value
          : undefined;
  if (leftNumeric === undefined || rightNumeric === undefined) {
    return undefined;
  }
  if (leftNumeric === rightNumeric) {
    return 0;
  }
  return leftNumeric < rightNumeric ? -1 : 1;
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

function coerceNumberLike(value: CellValue): number | undefined {
  switch (value.tag) {
    case ValueTag.Number:
      return value.value;
    case ValueTag.Boolean:
      return value.value ? 1 : 0;
    case ValueTag.Empty:
      return 0;
    case ValueTag.String: {
      const trimmed = value.value.trim();
      if (trimmed === "") {
        return 0;
      }
      const parsed = Number(trimmed);
      return Number.isFinite(parsed) ? parsed : undefined;
    }
    case ValueTag.Error:
    default:
      return undefined;
  }
}

function errorTypeCode(code: ErrorCode): number {
  return code;
}

const logicalPlaceholderBuiltins = createBlockedBuiltinMap(logicalPlaceholderBuiltinNames);

export const logicalBuiltins: Record<string, LogicalBuiltin> = {
  TRUE: (...args) => (args.length === 0 ? booleanResult(true) : errorValue(ErrorCode.Value)),
  FALSE: (...args) => (args.length === 0 ? booleanResult(false) : errorValue(ErrorCode.Value)),
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
  XOR: (...args) => {
    if (args.length === 0) {
      return errorValue(ErrorCode.Value);
    }

    let parity = false;
    for (const arg of args) {
      const coerced = coerceLogical(arg);
      if (!coerced.ok) {
        return coerced.error;
      }
      parity = parity !== coerced.value;
    }

    return booleanResult(parity);
  },
  IFS: (...args) => {
    if (args.length < 2 || args.length % 2 !== 0) {
      return errorValue(ErrorCode.Value);
    }

    for (let index = 0; index < args.length; index += 2) {
      const condition = args[index]!;
      const result = args[index + 1]!;
      const coerced = coerceLogical(condition);
      if (!coerced.ok) {
        return coerced.error;
      }
      if (coerced.value) {
        return result;
      }
    }

    return errorValue(ErrorCode.NA);
  },
  SWITCH: (...args) => {
    if (args.length < 3) {
      return errorValue(ErrorCode.Value);
    }

    const expression = args[0]!;
    if (isError(expression)) {
      return expression;
    }

    const hasDefault = (args.length - 1) % 2 === 1;
    const pairLimit = hasDefault ? args.length - 1 : args.length;
    for (let index = 1; index < pairLimit; index += 2) {
      const candidate = args[index]!;
      if (isError(candidate)) {
        return candidate;
      }
      if (compareScalars(expression, candidate) === 0) {
        return args[index + 1]!;
      }
    }

    return hasDefault ? args[args.length - 1]! : errorValue(ErrorCode.NA);
  },
  ISBLANK: (value = emptyValue()) => booleanResult(value.tag === ValueTag.Empty),
  ISNUMBER: (value = emptyValue()) => booleanResult(value.tag === ValueTag.Number),
  ISTEXT: (value = emptyValue()) => booleanResult(value.tag === ValueTag.String),
  ISERROR: (value = emptyValue()) => booleanResult(value.tag === ValueTag.Error),
  ISERR: (value = emptyValue()) => {
    if (value.tag !== ValueTag.Error) {
      return booleanResult(false);
    }
    return booleanResult(value.code !== ErrorCode.NA);
  },
  ISFORMULA: () => booleanResult(false),
  ISLOGICAL: (value = emptyValue()) => booleanResult(value.tag === ValueTag.Boolean),
  ISNONTEXT: (value = emptyValue()) => booleanResult(value.tag !== ValueTag.String),
  ISEVEN: (value = emptyValue()) => {
    const numberValue = coerceNumberLike(value);
    if (numberValue === undefined) {
      return errorValue(ErrorCode.Value);
    }
    return booleanResult(Math.trunc(numberValue) % 2 === 0);
  },
  ISODD: (value = emptyValue()) => {
    const numberValue = coerceNumberLike(value);
    if (numberValue === undefined) {
      return errorValue(ErrorCode.Value);
    }
    return booleanResult(Math.trunc(numberValue) % 2 !== 0);
  },
  ISNA: (value = emptyValue()) =>
    booleanResult(value.tag === ValueTag.Error && value.code === ErrorCode.NA),
  ISREF: (_value = emptyValue()) => booleanResult(false),
  "ERROR.TYPE": (value = emptyValue()) => {
    if (value.tag !== ValueTag.Error) {
      return errorValue(ErrorCode.NA);
    }
    return { tag: ValueTag.Number, value: errorTypeCode(value.code) };
  },
  ...logicalPlaceholderBuiltins,
};

export function getLogicalBuiltin(name: string): LogicalBuiltin | undefined {
  return logicalBuiltins[name.toUpperCase()];
}
