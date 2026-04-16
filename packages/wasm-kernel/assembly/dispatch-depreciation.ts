import { BuiltinId, ErrorCode, ValueTag } from './protocol'
import { dbDepreciation, ddbDepreciation, vdbDepreciation } from './date-finance'
import { toNumberExact } from './operands'
import { scalarErrorAt } from './builtin-args'
import { STACK_KIND_SCALAR, writeResult } from './result-io'

function coerceBoolean(tag: u8, value: f64): i32 {
  if (tag == ValueTag.Boolean || tag == ValueTag.Number) {
    return value != 0 ? 1 : 0
  }
  if (tag == ValueTag.Empty) {
    return 0
  }
  return -1
}

export function tryApplyDepreciationBuiltin(
  builtinId: i32,
  argc: i32,
  base: i32,
  rangeIndexStack: Uint32Array,
  valueStack: Float64Array,
  tagStack: Uint8Array,
  kindStack: Uint8Array,
): i32 {
  if (builtinId == BuiltinId.Db && (argc == 4 || argc == 5)) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const cost = toNumberExact(tagStack[base], valueStack[base])
    const salvage = toNumberExact(tagStack[base + 1], valueStack[base + 1])
    const life = toNumberExact(tagStack[base + 2], valueStack[base + 2])
    const period = toNumberExact(tagStack[base + 3], valueStack[base + 3])
    const month = argc == 5 ? toNumberExact(tagStack[base + 4], valueStack[base + 4]) : 12.0
    const depreciation = dbDepreciation(cost, salvage, life, period, month)
    return isNaN(depreciation)
      ? writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, depreciation, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Ddb && (argc == 4 || argc == 5)) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const cost = toNumberExact(tagStack[base], valueStack[base])
    const salvage = toNumberExact(tagStack[base + 1], valueStack[base + 1])
    const life = toNumberExact(tagStack[base + 2], valueStack[base + 2])
    const period = toNumberExact(tagStack[base + 3], valueStack[base + 3])
    const factor = argc == 5 ? toNumberExact(tagStack[base + 4], valueStack[base + 4]) : 2.0
    const depreciation = ddbDepreciation(cost, salvage, life, period, factor)
    return isNaN(depreciation)
      ? writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, depreciation, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Vdb && argc >= 5 && argc <= 7) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const cost = toNumberExact(tagStack[base], valueStack[base])
    const salvage = toNumberExact(tagStack[base + 1], valueStack[base + 1])
    const life = toNumberExact(tagStack[base + 2], valueStack[base + 2])
    const startPeriod = toNumberExact(tagStack[base + 3], valueStack[base + 3])
    const endPeriod = toNumberExact(tagStack[base + 4], valueStack[base + 4])
    const factor = argc >= 6 ? toNumberExact(tagStack[base + 5], valueStack[base + 5]) : 2.0
    const noSwitch = argc >= 7 ? coerceBoolean(tagStack[base + 6], valueStack[base + 6]) : 0
    const depreciation = noSwitch < 0 ? NaN : vdbDepreciation(cost, salvage, life, startPeriod, endPeriod, factor, noSwitch != 0)
    return isNaN(depreciation)
      ? writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      : writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, depreciation, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Sln && argc == 3) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const cost = toNumberExact(tagStack[base], valueStack[base])
    const salvage = toNumberExact(tagStack[base + 1], valueStack[base + 1])
    const life = toNumberExact(tagStack[base + 2], valueStack[base + 2])
    if (isNaN(cost) || isNaN(salvage) || isNaN(life) || life <= 0.0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      (cost - salvage) / life,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    )
  }

  if (builtinId == BuiltinId.Syd && argc == 4) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack)
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const cost = toNumberExact(tagStack[base], valueStack[base])
    const salvage = toNumberExact(tagStack[base + 1], valueStack[base + 1])
    const life = toNumberExact(tagStack[base + 2], valueStack[base + 2])
    const period = toNumberExact(tagStack[base + 3], valueStack[base + 3])
    if (isNaN(cost) || isNaN(salvage) || isNaN(life) || isNaN(period) || life <= 0.0 || period <= 0.0 || period > life) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const denominator = (life * (life + 1.0)) / 2.0
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      ((cost - salvage) * (life - period + 1.0)) / denominator,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    )
  }

  return -1
}
