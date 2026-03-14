import { BuiltinId, ErrorCode, ValueTag } from "./protocol";

export const STACK_KIND_SCALAR: u8 = 0;
export const STACK_KIND_RANGE: u8 = 1;

function toNumberOrNaN(tag: u8, value: f64): f64 {
  if (tag == ValueTag.Number || tag == ValueTag.Boolean) return value;
  if (tag == ValueTag.Empty) return 0;
  return NaN;
}

function toNumberOrZero(tag: u8, value: f64): f64 {
  const numeric = toNumberOrNaN(tag, value);
  return isNaN(numeric) ? 0 : numeric;
}

function writeResult(
  base: i32,
  kind: u8,
  tag: u8,
  value: f64,
  rangeIndexStack: Uint32Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  kindStack: Uint8Array
): i32 {
  rangeIndexStack[base] = 0;
  valueStack[base] = value;
  tagStack[base] = tag;
  kindStack[base] = kind;
  return base + 1;
}

function scalarErrorAt(base: i32, argc: i32, kindStack: Uint8Array, tagStack: Uint8Array, valueStack: Float64Array): f64 {
  for (let index = 0; index < argc; index++) {
    const slot = base + index;
    if (kindStack[slot] == STACK_KIND_SCALAR && tagStack[slot] == ValueTag.Error) {
      return valueStack[slot];
    }
  }
  return -1;
}

function rangeErrorAt(
  base: i32,
  argc: i32,
  kindStack: Uint8Array,
  rangeIndexStack: Uint32Array,
  rangeOffsets: Uint32Array,
  rangeLengths: Uint32Array,
  rangeMembers: Uint32Array,
  cellTags: Uint8Array,
  cellErrors: Uint16Array
): f64 {
  for (let index = 0; index < argc; index++) {
    const slot = base + index;
    if (kindStack[slot] != STACK_KIND_RANGE) {
      continue;
    }
    const rangeIndex = rangeIndexStack[slot];
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

function rangeSupportedScalarOnly(base: i32, argc: i32, kindStack: Uint8Array): bool {
  for (let index = 0; index < argc; index++) {
    if (kindStack[base + index] == STACK_KIND_RANGE) {
      return false;
    }
  }
  return true;
}

export function applyBuiltin(
  builtinId: i32,
  argc: i32,
  rangeIndexStack: Uint32Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  kindStack: Uint8Array,
  cellTags: Uint8Array,
  cellNumbers: Float64Array,
  cellErrors: Uint16Array,
  rangeOffsets: Uint32Array,
  rangeLengths: Uint32Array,
  rangeMembers: Uint32Array,
  sp: i32
): i32 {
  const base = sp - argc;

  if (builtinId == BuiltinId.Sum) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack);
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack);
    }
    const rangeError = rangeErrorAt(base, argc, kindStack, rangeIndexStack, rangeOffsets, rangeLengths, rangeMembers, cellTags, cellErrors);
    if (rangeError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, rangeError, rangeIndexStack, valueStack, tagStack, kindStack);
    }

    let sum = 0.0;
    for (let index = 0; index < argc; index++) {
      const slot = base + index;
      if (kindStack[slot] == STACK_KIND_RANGE) {
        const rangeIndex = rangeIndexStack[slot];
        const start = rangeOffsets[rangeIndex];
        const length = <i32>rangeLengths[rangeIndex];
        for (let cursor = 0; cursor < length; cursor++) {
          const memberIndex = rangeMembers[start + cursor];
          const numeric = toNumberOrNaN(cellTags[memberIndex], cellNumbers[memberIndex]);
          if (!isNaN(numeric)) {
            sum += numeric;
          }
        }
        continue;
      }
      const numeric = toNumberOrNaN(tagStack[slot], valueStack[slot]);
      if (!isNaN(numeric)) {
        sum += numeric;
      }
    }
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, sum, rangeIndexStack, valueStack, tagStack, kindStack);
  }

  if (builtinId == BuiltinId.Avg) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack);
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack);
    }
    const rangeError = rangeErrorAt(base, argc, kindStack, rangeIndexStack, rangeOffsets, rangeLengths, rangeMembers, cellTags, cellErrors);
    if (rangeError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, rangeError, rangeIndexStack, valueStack, tagStack, kindStack);
    }

    let sum = 0.0;
    let count = 0;
    for (let index = 0; index < argc; index++) {
      const slot = base + index;
      if (kindStack[slot] == STACK_KIND_RANGE) {
        const rangeIndex = rangeIndexStack[slot];
        const start = rangeOffsets[rangeIndex];
        const length = <i32>rangeLengths[rangeIndex];
        for (let cursor = 0; cursor < length; cursor++) {
          const memberIndex = rangeMembers[start + cursor];
          const numeric = toNumberOrNaN(cellTags[memberIndex], cellNumbers[memberIndex]);
          if (!isNaN(numeric)) {
            sum += numeric;
            count += 1;
          }
        }
        continue;
      }
      const numeric = toNumberOrNaN(tagStack[slot], valueStack[slot]);
      if (!isNaN(numeric)) {
        sum += numeric;
        count += 1;
      }
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      count == 0 ? 0 : sum / count,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack
    );
  }

  if (builtinId == BuiltinId.Min) {
    let min = Infinity;
    for (let index = 0; index < argc; index++) {
      const slot = base + index;
      if (kindStack[slot] == STACK_KIND_RANGE) {
        const rangeIndex = rangeIndexStack[slot];
        const start = rangeOffsets[rangeIndex];
        const length = <i32>rangeLengths[rangeIndex];
        for (let cursor = 0; cursor < length; cursor++) {
          const memberIndex = rangeMembers[start + cursor];
          const numeric = toNumberOrNaN(cellTags[memberIndex], cellNumbers[memberIndex]);
          if (!isNaN(numeric) && numeric < min) {
            min = numeric;
          }
        }
        continue;
      }
      const numeric = toNumberOrNaN(tagStack[slot], valueStack[slot]);
      if (!isNaN(numeric) && numeric < min) {
        min = numeric;
      }
    }
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, min, rangeIndexStack, valueStack, tagStack, kindStack);
  }

  if (builtinId == BuiltinId.Max) {
    let max = -Infinity;
    for (let index = 0; index < argc; index++) {
      const slot = base + index;
      if (kindStack[slot] == STACK_KIND_RANGE) {
        const rangeIndex = rangeIndexStack[slot];
        const start = rangeOffsets[rangeIndex];
        const length = <i32>rangeLengths[rangeIndex];
        for (let cursor = 0; cursor < length; cursor++) {
          const memberIndex = rangeMembers[start + cursor];
          const numeric = toNumberOrNaN(cellTags[memberIndex], cellNumbers[memberIndex]);
          if (!isNaN(numeric) && numeric > max) {
            max = numeric;
          }
        }
        continue;
      }
      const numeric = toNumberOrNaN(tagStack[slot], valueStack[slot]);
      if (!isNaN(numeric) && numeric > max) {
        max = numeric;
      }
    }
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, max, rangeIndexStack, valueStack, tagStack, kindStack);
  }

  if (builtinId == BuiltinId.Count) {
    let count = 0;
    for (let index = 0; index < argc; index++) {
      const slot = base + index;
      if (kindStack[slot] == STACK_KIND_RANGE) {
        const rangeIndex = rangeIndexStack[slot];
        const start = rangeOffsets[rangeIndex];
        const length = <i32>rangeLengths[rangeIndex];
        for (let cursor = 0; cursor < length; cursor++) {
          const memberIndex = rangeMembers[start + cursor];
          if (!isNaN(toNumberOrNaN(cellTags[memberIndex], cellNumbers[memberIndex]))) {
            count += 1;
          }
        }
        continue;
      }
      if (!isNaN(toNumberOrNaN(tagStack[slot], valueStack[slot]))) {
        count += 1;
      }
    }
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, count, rangeIndexStack, valueStack, tagStack, kindStack);
  }

  if (builtinId == BuiltinId.CountA) {
    let count = 0;
    for (let index = 0; index < argc; index++) {
      const slot = base + index;
      if (kindStack[slot] == STACK_KIND_RANGE) {
        const rangeIndex = rangeIndexStack[slot];
        const start = rangeOffsets[rangeIndex];
        const length = <i32>rangeLengths[rangeIndex];
        for (let cursor = 0; cursor < length; cursor++) {
          const memberIndex = rangeMembers[start + cursor];
          if (cellTags[memberIndex] != ValueTag.Empty) {
            count += 1;
          }
        }
        continue;
      }
      if (tagStack[slot] != ValueTag.Empty) {
        count += 1;
      }
    }
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, count, rangeIndexStack, valueStack, tagStack, kindStack);
  }

  if (!rangeSupportedScalarOnly(base, argc, kindStack)) {
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack);
  }

  if (builtinId == BuiltinId.Abs && argc == 1) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      Math.abs(toNumberOrZero(tagStack[base], valueStack[base])),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack
    );
  }

  if (builtinId == BuiltinId.Round && argc == 1) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      Math.round(toNumberOrZero(tagStack[base], valueStack[base])),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack
    );
  }

  if (builtinId == BuiltinId.Floor && argc == 1) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      Math.floor(toNumberOrZero(tagStack[base], valueStack[base])),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack
    );
  }

  if (builtinId == BuiltinId.Ceiling && argc == 1) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      Math.ceil(toNumberOrZero(tagStack[base], valueStack[base])),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack
    );
  }

  if (builtinId == BuiltinId.Mod && argc == 2) {
    const divisor = toNumberOrZero(tagStack[base + 1], valueStack[base + 1]);
    if (divisor == 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Div0, rangeIndexStack, valueStack, tagStack, kindStack);
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      toNumberOrZero(tagStack[base], valueStack[base]) % divisor,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack
    );
  }

  if (builtinId == BuiltinId.And) {
    let result = 1.0;
    for (let index = 0; index < argc; index++) {
      if (toNumberOrZero(tagStack[base + index], valueStack[base + index]) == 0) {
        result = 0;
        break;
      }
    }
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Boolean, result, rangeIndexStack, valueStack, tagStack, kindStack);
  }

  if (builtinId == BuiltinId.Or) {
    let result = 0.0;
    for (let index = 0; index < argc; index++) {
      if (toNumberOrZero(tagStack[base + index], valueStack[base + index]) != 0) {
        result = 1;
        break;
      }
    }
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Boolean, result, rangeIndexStack, valueStack, tagStack, kindStack);
  }

  if (builtinId == BuiltinId.Not && argc == 1) {
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Boolean,
      toNumberOrZero(tagStack[base], valueStack[base]) == 0 ? 1 : 0,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack
    );
  }

  return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack);
}
