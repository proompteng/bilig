import { BuiltinId, ErrorCode, ValueTag } from './protocol'
import { coerceLogical } from './builtin-args'
import { compareScalarValues } from './comparison'
import { coerceBitwiseUnsigned, coerceNonNegativeShift } from './numeric-core'
import { copySlotResult, STACK_KIND_SCALAR, writeResult } from './result-io'

function writeLogicInfoError(
  base: i32,
  error: i32,
  rangeIndexStack: Uint32Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  kindStack: Uint8Array,
): i32 {
  return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, error, rangeIndexStack, valueStack, tagStack, kindStack)
}

function writeLogicInfoBoolean(
  base: i32,
  value: bool,
  rangeIndexStack: Uint32Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  kindStack: Uint8Array,
): i32 {
  return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Boolean, value ? 1 : 0, rangeIndexStack, valueStack, tagStack, kindStack)
}

function writeLogicInfoNumber(
  base: i32,
  value: f64,
  rangeIndexStack: Uint32Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  kindStack: Uint8Array,
): i32 {
  return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, value, rangeIndexStack, valueStack, tagStack, kindStack)
}

export function tryApplyLogicInfoBuiltin(
  builtinId: i32,
  argc: i32,
  base: i32,
  rangeIndexStack: Uint32Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  kindStack: Uint8Array,
  stringOffsets: Uint32Array,
  stringLengths: Uint32Array,
  stringData: Uint16Array,
  outputStringOffsets: Uint32Array,
  outputStringLengths: Uint32Array,
  outputStringData: Uint16Array,
): i32 {
  if ((builtinId == BuiltinId.Bitand || builtinId == BuiltinId.Bitor || builtinId == BuiltinId.Bitxor) && argc >= 2) {
    let accumulatorValue = coerceBitwiseUnsigned(tagStack[base], valueStack[base])
    if (accumulatorValue == i64.MIN_VALUE) {
      return writeLogicInfoError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    let accumulator = <u32>accumulatorValue
    for (let index = 1; index < argc; index += 1) {
      const currentValue = coerceBitwiseUnsigned(tagStack[base + index], valueStack[base + index])
      if (currentValue == i64.MIN_VALUE) {
        return writeLogicInfoError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      const current = <u32>currentValue
      if (builtinId == BuiltinId.Bitand) {
        accumulator &= current
      } else if (builtinId == BuiltinId.Bitor) {
        accumulator |= current
      } else {
        accumulator ^= current
      }
    }
    return writeLogicInfoNumber(base, <f64>accumulator, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if ((builtinId == BuiltinId.Bitlshift || builtinId == BuiltinId.Bitrshift) && argc == 2) {
    const value = coerceBitwiseUnsigned(tagStack[base], valueStack[base])
    const shift = coerceNonNegativeShift(tagStack[base + 1], valueStack[base + 1])
    if (value == i64.MIN_VALUE || shift == i64.MIN_VALUE) {
      return writeLogicInfoError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const shiftAmount = <i32>(shift & 31)
    const numeric = <u32>value
    const result = builtinId == BuiltinId.Bitlshift ? <u32>(numeric << shiftAmount) : <u32>(numeric >>> shiftAmount)
    return writeLogicInfoNumber(base, <f64>result, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.And) {
    if (argc == 0) {
      return writeLogicInfoError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    for (let index = 0; index < argc; index += 1) {
      const coerced = coerceLogical(tagStack[base + index], valueStack[base + index])
      if (coerced < 0) {
        return writeLogicInfoError(base, -coerced - 1, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      if (coerced == 0) {
        return writeLogicInfoBoolean(base, false, rangeIndexStack, valueStack, tagStack, kindStack)
      }
    }
    return writeLogicInfoBoolean(base, true, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Or) {
    if (argc == 0) {
      return writeLogicInfoError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    for (let index = 0; index < argc; index += 1) {
      const coerced = coerceLogical(tagStack[base + index], valueStack[base + index])
      if (coerced < 0) {
        return writeLogicInfoError(base, -coerced - 1, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      if (coerced != 0) {
        return writeLogicInfoBoolean(base, true, rangeIndexStack, valueStack, tagStack, kindStack)
      }
    }
    return writeLogicInfoBoolean(base, false, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Xor) {
    if (argc == 0) {
      return writeLogicInfoError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    let parity = 0
    for (let index = 0; index < argc; index += 1) {
      const coerced = coerceLogical(tagStack[base + index], valueStack[base + index])
      if (coerced < 0) {
        return writeLogicInfoError(base, -coerced - 1, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      parity = parity ^ (coerced != 0 ? 1 : 0)
    }
    return writeLogicInfoBoolean(base, parity != 0, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Not && argc == 1) {
    const coerced = coerceLogical(tagStack[base], valueStack[base])
    if (coerced < 0) {
      return writeLogicInfoError(base, -coerced - 1, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeLogicInfoBoolean(base, coerced == 0, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Ifs) {
    if (argc < 2 || argc % 2 != 0) {
      return writeLogicInfoError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    for (let index = 0; index < argc; index += 2) {
      const coerced = coerceLogical(tagStack[base + index], valueStack[base + index])
      if (coerced < 0) {
        return writeLogicInfoError(base, -coerced - 1, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      if (coerced != 0) {
        return copySlotResult(base, base + index + 1, rangeIndexStack, valueStack, tagStack, kindStack)
      }
    }
    return writeLogicInfoError(base, ErrorCode.NA, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Switch) {
    if (argc < 3) {
      return writeLogicInfoError(base, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    if (tagStack[base] == ValueTag.Error) {
      return writeLogicInfoError(base, <i32>valueStack[base], rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const hasDefault = (argc - 1) % 2 == 1
    const pairLimit = hasDefault ? argc - 1 : argc
    for (let index = 1; index < pairLimit; index += 2) {
      if (tagStack[base + index] == ValueTag.Error) {
        return writeLogicInfoError(base, <i32>valueStack[base + index], rangeIndexStack, valueStack, tagStack, kindStack)
      }
      const comparison = compareScalarValues(
        tagStack[base],
        valueStack[base],
        tagStack[base + index],
        valueStack[base + index],
        null,
        stringOffsets,
        stringLengths,
        stringData,
        outputStringOffsets,
        outputStringLengths,
        outputStringData,
      )
      if (comparison == 0) {
        return copySlotResult(base, base + index + 1, rangeIndexStack, valueStack, tagStack, kindStack)
      }
    }
    return hasDefault
      ? copySlotResult(base, base + argc - 1, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeLogicInfoError(base, ErrorCode.NA, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.IsBlank && argc == 0) {
    return writeLogicInfoBoolean(base, true, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.IsBlank && argc == 1) {
    return writeLogicInfoBoolean(base, tagStack[base] == ValueTag.Empty, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.IsNumber && argc == 0) {
    return writeLogicInfoBoolean(base, false, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.IsNumber && argc == 1) {
    return writeLogicInfoBoolean(base, tagStack[base] == ValueTag.Number, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.IsText && argc == 0) {
    return writeLogicInfoBoolean(base, false, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.IsText && argc == 1) {
    return writeLogicInfoBoolean(base, tagStack[base] == ValueTag.String, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  return -1
}
