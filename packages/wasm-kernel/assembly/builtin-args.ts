import { ErrorCode, ValueTag } from "./protocol";
import {
  inputCellScalarValue,
  inputCellTag,
  inputColsFromSlot,
  inputRowsFromSlot,
  toNumberExact,
} from "./operands";
import { STACK_KIND_RANGE, STACK_KIND_SCALAR, UNRESOLVED_WASM_OPERAND } from "./result-io";

let statCollectionErrorCode = 0;

export function coercePositiveIntegerArg(tag: u8, value: f64, hasValue: bool, fallback: i32): i32 {
  if (!hasValue) {
    return fallback;
  }
  const numeric = toNumberExact(tag, value);
  if (!isFinite(numeric)) {
    return i32.MIN_VALUE;
  }
  const truncated = <i32>numeric;
  return truncated >= 1 ? truncated : i32.MIN_VALUE;
}

export function coerceNumberArg(tag: u8, value: f64, hasValue: bool, fallback: f64): f64 {
  if (!hasValue) {
    return fallback;
  }
  const numeric = toNumberExact(tag, value);
  return isFinite(numeric) ? numeric : NaN;
}

export function scalarErrorAt(
  base: i32,
  argc: i32,
  kindStack: Uint8Array,
  tagStack: Uint8Array,
  valueStack: Float64Array,
): f64 {
  for (let index = 0; index < argc; index++) {
    const slot = base + index;
    if (kindStack[slot] == STACK_KIND_SCALAR && tagStack[slot] == ValueTag.Error) {
      return valueStack[slot];
    }
  }
  return -1;
}

export function rangeErrorAt(
  base: i32,
  argc: i32,
  kindStack: Uint8Array,
  rangeIndexStack: Uint32Array,
  rangeOffsets: Uint32Array,
  rangeLengths: Uint32Array,
  rangeMembers: Uint32Array,
  cellTags: Uint8Array,
  cellErrors: Uint16Array,
): f64 {
  for (let index = 0; index < argc; index++) {
    const slot = base + index;
    if (kindStack[slot] != STACK_KIND_RANGE) {
      continue;
    }
    const rangeIndex = rangeIndexStack[slot];
    if (rangeIndex == UNRESOLVED_WASM_OPERAND) {
      return ErrorCode.Ref;
    }
    const start = rangeOffsets[rangeIndex];
    const length = <i32>rangeLengths[rangeIndex];
    for (let cursor = 0; cursor < length; cursor++) {
      const memberIndex = rangeMembers[start + cursor];
      if (cellTags[memberIndex] == ValueTag.Error) {
        return cellErrors[memberIndex];
      }
    }
  }
  return -1;
}

export function rangeSupportedScalarOnly(base: i32, argc: i32, kindStack: Uint8Array): bool {
  for (let index = 0; index < argc; index++) {
    if (kindStack[base + index] == STACK_KIND_RANGE) {
      return false;
    }
  }
  return true;
}

export function coerceLogical(tag: u8, value: f64): i32 {
  if (tag == ValueTag.Boolean || tag == ValueTag.Number) {
    return value != 0 ? 1 : 0;
  }
  if (tag == ValueTag.Empty) {
    return 0;
  }
  if (tag == ValueTag.Error) {
    return -(<i32>value) - 1;
  }
  return -(<i32>ErrorCode.Value) - 1;
}

export function scalarArgsOnly(base: i32, argc: i32, kindStack: Uint8Array): bool {
  for (let index = 0; index < argc; index++) {
    if (kindStack[base + index] != STACK_KIND_SCALAR) {
      return false;
    }
  }
  return true;
}

export function statScalarValue(tag: u8, value: f64, includeStringsAsZero: bool): f64 {
  if (tag == ValueTag.Number || tag == ValueTag.Boolean) {
    return value;
  }
  if (tag == ValueTag.String) {
    return includeStringsAsZero ? 0 : NaN;
  }
  return NaN;
}

export function collectScalarStatValues(
  base: i32,
  argc: i32,
  tagStack: Uint8Array,
  valueStack: Float64Array,
  includeStringsAsZero: bool,
): Array<f64> {
  const values = new Array<f64>();
  for (let index = 0; index < argc; index++) {
    const numeric = statScalarValue(
      tagStack[base + index],
      valueStack[base + index],
      includeStringsAsZero,
    );
    if (!isNaN(numeric)) {
      values.push(numeric);
    }
  }
  return values;
}

export function collectStatValuesFromArgs(
  base: i32,
  argc: i32,
  kindStack: Uint8Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  rangeIndexStack: Uint32Array,
  rangeOffsets: Uint32Array,
  rangeLengths: Uint32Array,
  rangeRowCounts: Uint32Array,
  rangeColCounts: Uint32Array,
  rangeMembers: Uint32Array,
  cellTags: Uint8Array,
  cellNumbers: Float64Array,
  cellStringIds: Uint32Array,
  cellErrors: Uint16Array,
  includeStringsAsZero: bool,
): Array<f64> | null {
  statCollectionErrorCode = 0;
  const values = new Array<f64>();
  for (let index = 0; index < argc; index += 1) {
    const slot = base + index;
    if (kindStack[slot] == STACK_KIND_SCALAR) {
      const tag = tagStack[slot];
      const raw = valueStack[slot];
      if (tag == ValueTag.Error) {
        statCollectionErrorCode = <i32>raw;
        return null;
      }
      if (tag == ValueTag.Number || tag == ValueTag.Boolean) {
        values.push(raw);
      } else if (tag == ValueTag.String && includeStringsAsZero) {
        values.push(0.0);
      }
      continue;
    }

    const rows = inputRowsFromSlot(slot, kindStack, rangeIndexStack, rangeRowCounts);
    const cols = inputColsFromSlot(slot, kindStack, rangeIndexStack, rangeColCounts);
    if (rows < 1 || cols < 1) {
      statCollectionErrorCode = ErrorCode.Value;
      return null;
    }

    for (let row = 0; row < rows; row += 1) {
      for (let col = 0; col < cols; col += 1) {
        const tag = inputCellTag(
          slot,
          row,
          col,
          kindStack,
          valueStack,
          tagStack,
          rangeIndexStack,
          rangeOffsets,
          rangeLengths,
          rangeRowCounts,
          rangeColCounts,
          rangeMembers,
          cellTags,
          cellNumbers,
        );
        const raw = inputCellScalarValue(
          slot,
          row,
          col,
          kindStack,
          valueStack,
          tagStack,
          rangeIndexStack,
          rangeOffsets,
          rangeLengths,
          rangeRowCounts,
          rangeColCounts,
          rangeMembers,
          cellTags,
          cellNumbers,
          cellStringIds,
          cellErrors,
        );
        if (tag == ValueTag.Error) {
          statCollectionErrorCode = <i32>raw;
          return null;
        }
        if (tag == ValueTag.Number) {
          values.push(raw);
        } else if (tag == ValueTag.Boolean && includeStringsAsZero) {
          values.push(raw);
        } else if (tag == ValueTag.String && includeStringsAsZero) {
          values.push(0.0);
        }
      }
    }
  }
  return values;
}

export function lastStatCollectionErrorCode(): i32 {
  return statCollectionErrorCode;
}

export function paymentType(tag: u8, value: f64, hasValue: bool): i32 {
  if (!hasValue) {
    return 0;
  }
  const numeric = toNumberExact(tag, value);
  if (!isFinite(numeric)) {
    return -1;
  }
  const truncated = <i32>numeric;
  return truncated == 0 || truncated == 1 ? truncated : -1;
}

export function isNumericResult(value: f64): bool {
  return !isNaN(value);
}
