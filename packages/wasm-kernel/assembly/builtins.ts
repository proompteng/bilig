import { BuiltinId, ErrorCode, ValueTag } from "./protocol";

function toNumberOrNaN(tag: u8, value: f64): f64 {
  if (tag == ValueTag.Number || tag == ValueTag.Boolean) return value;
  if (tag == ValueTag.Empty) return 0;
  return NaN;
}

function toNumberOrZero(tag: u8, value: f64): f64 {
  const numeric = toNumberOrNaN(tag, value);
  return isNaN(numeric) ? 0 : numeric;
}

function firstErrorIndex(argc: i32, valueStack: Float64Array, tagStack: Uint8Array, sp: i32): i32 {
  const base = sp - argc;
  for (let index = 0; index < argc; index++) {
    if (tagStack[base + index] == ValueTag.Error) {
      return base + index;
    }
  }
  return -1;
}

function writeResult(base: i32, tag: u8, value: f64, valueStack: Float64Array, tagStack: Uint8Array): i32 {
  valueStack[base] = value;
  tagStack[base] = tag;
  return base + 1;
}

export function applyBuiltin(
  builtinId: i32,
  argc: i32,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  sp: i32
): i32 {
  const base = sp - argc;

  if (builtinId == BuiltinId.Sum) {
    const errorIndex = firstErrorIndex(argc, valueStack, tagStack, sp);
    if (errorIndex >= 0) {
      return writeResult(base, <u8>ValueTag.Error, valueStack[errorIndex], valueStack, tagStack);
    }
    let sum = 0.0;
    for (let index = 0; index < argc; index++) {
      const numeric = toNumberOrNaN(tagStack[base + index], valueStack[base + index]);
      if (!isNaN(numeric)) {
        sum += numeric;
      }
    }
    return writeResult(base, <u8>ValueTag.Number, sum, valueStack, tagStack);
  }

  if (builtinId == BuiltinId.Avg) {
    const errorIndex = firstErrorIndex(argc, valueStack, tagStack, sp);
    if (errorIndex >= 0) {
      return writeResult(base, <u8>ValueTag.Error, valueStack[errorIndex], valueStack, tagStack);
    }
    let sum = 0.0;
    let count = 0;
    for (let index = 0; index < argc; index++) {
      const numeric = toNumberOrNaN(tagStack[base + index], valueStack[base + index]);
      if (!isNaN(numeric)) {
        sum += numeric;
        count += 1;
      }
    }
    return writeResult(base, <u8>ValueTag.Number, count == 0 ? 0 : sum / count, valueStack, tagStack);
  }

  if (builtinId == BuiltinId.Min) {
    let min = Infinity;
    for (let index = 0; index < argc; index++) {
      const numeric = toNumberOrNaN(tagStack[base + index], valueStack[base + index]);
      if (!isNaN(numeric) && numeric < min) {
        min = numeric;
      }
    }
    return writeResult(base, <u8>ValueTag.Number, min, valueStack, tagStack);
  }

  if (builtinId == BuiltinId.Max) {
    let max = -Infinity;
    for (let index = 0; index < argc; index++) {
      const numeric = toNumberOrNaN(tagStack[base + index], valueStack[base + index]);
      if (!isNaN(numeric) && numeric > max) {
        max = numeric;
      }
    }
    return writeResult(base, <u8>ValueTag.Number, max, valueStack, tagStack);
  }

  if (builtinId == BuiltinId.Count) {
    let count = 0;
    for (let index = 0; index < argc; index++) {
      if (!isNaN(toNumberOrNaN(tagStack[base + index], valueStack[base + index]))) {
        count += 1;
      }
    }
    return writeResult(base, <u8>ValueTag.Number, count, valueStack, tagStack);
  }

  if (builtinId == BuiltinId.CountA) {
    let count = 0;
    for (let index = 0; index < argc; index++) {
      if (tagStack[base + index] != ValueTag.Empty) {
        count += 1;
      }
    }
    return writeResult(base, <u8>ValueTag.Number, count, valueStack, tagStack);
  }

  if (builtinId == BuiltinId.Abs && argc == 1) {
    return writeResult(base, <u8>ValueTag.Number, Math.abs(toNumberOrZero(tagStack[base], valueStack[base])), valueStack, tagStack);
  }

  if (builtinId == BuiltinId.Round && argc == 1) {
    return writeResult(base, <u8>ValueTag.Number, Math.round(toNumberOrZero(tagStack[base], valueStack[base])), valueStack, tagStack);
  }

  if (builtinId == BuiltinId.Floor && argc == 1) {
    return writeResult(base, <u8>ValueTag.Number, Math.floor(toNumberOrZero(tagStack[base], valueStack[base])), valueStack, tagStack);
  }

  if (builtinId == BuiltinId.Ceiling && argc == 1) {
    return writeResult(base, <u8>ValueTag.Number, Math.ceil(toNumberOrZero(tagStack[base], valueStack[base])), valueStack, tagStack);
  }

  if (builtinId == BuiltinId.Mod && argc == 2) {
    const divisor = toNumberOrZero(tagStack[base + 1], valueStack[base + 1]);
    if (divisor == 0) {
      return writeResult(base, <u8>ValueTag.Error, ErrorCode.Div0, valueStack, tagStack);
    }
    return writeResult(
      base,
      <u8>ValueTag.Number,
      toNumberOrZero(tagStack[base], valueStack[base]) % divisor,
      valueStack,
      tagStack
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
    return writeResult(base, <u8>ValueTag.Boolean, result, valueStack, tagStack);
  }

  if (builtinId == BuiltinId.Or) {
    let result = 0.0;
    for (let index = 0; index < argc; index++) {
      if (toNumberOrZero(tagStack[base + index], valueStack[base + index]) != 0) {
        result = 1;
        break;
      }
    }
    return writeResult(base, <u8>ValueTag.Boolean, result, valueStack, tagStack);
  }

  if (builtinId == BuiltinId.Not && argc == 1) {
    return writeResult(
      base,
      <u8>ValueTag.Boolean,
      toNumberOrZero(tagStack[base], valueStack[base]) == 0 ? 1 : 0,
      valueStack,
      tagStack
    );
  }

  return writeResult(base, <u8>ValueTag.Error, ErrorCode.Value, valueStack, tagStack);
}
