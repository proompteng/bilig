import { BuiltinId, ErrorCode, ValueTag } from './protocol'
import {
  bytePositionToCharPositionUtf8,
  charPositionToBytePositionUtf8,
  leftBytesText,
  midBytesText,
  rightBytesText,
  scalarText,
  textLength,
  utf8ByteLength,
} from './text-codec'
import { STACK_KIND_SCALAR, writeResult, writeStringResult } from './result-io'
import { coerceLength, coercePositiveStart, excelTrim, findPosition } from './text-foundation'

export function tryApplyScalarTextBuiltin(
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
  if (builtinId == BuiltinId.Concat) {
    let scalarError = -1
    for (let index = 0; index < argc; index += 1) {
      if (tagStack[base + index] == ValueTag.Error) {
        scalarError = <i32>valueStack[base + index]
        break
      }
    }
    if (scalarError >= 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, scalarError, rangeIndexStack, valueStack, tagStack, kindStack)
    }

    for (let index = 0; index < argc; index += 1) {
      const len = textLength(tagStack[base + index], valueStack[base + index], stringLengths, outputStringLengths)
      if (len < 0) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
    }

    let text = ''
    for (let index = 0; index < argc; index += 1) {
      const part = scalarText(
        tagStack[base + index],
        valueStack[base + index],
        stringOffsets,
        stringLengths,
        stringData,
        outputStringOffsets,
        outputStringLengths,
        outputStringData,
      )
      if (part != null) {
        text += part
      }
    }
    return writeStringResult(base, text, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Len && argc == 1) {
    const length = textLength(tagStack[base], valueStack[base], stringLengths, outputStringLengths)
    if (length < 0) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, <f64>length, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Lenb && argc == 1) {
    const text = scalarText(
      tagStack[base],
      valueStack[base],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    if (text == null) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      <f64>utf8ByteLength(text),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    )
  }

  if (builtinId == BuiltinId.Exact && argc == 2) {
    const left = scalarText(
      tagStack[base],
      valueStack[base],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    const right = scalarText(
      tagStack[base + 1],
      valueStack[base + 1],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    if (left === null || right === null) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Boolean,
      left == right ? 1 : 0,
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    )
  }

  if ((builtinId == BuiltinId.Left || builtinId == BuiltinId.Right) && (argc == 1 || argc == 2)) {
    const text = scalarText(
      tagStack[base],
      valueStack[base],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    const count = argc == 2 ? coerceLength(tagStack[base + 1], valueStack[base + 1], 1) : 1
    if (text == null || count == i32.MIN_VALUE) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const result =
      builtinId == BuiltinId.Left ? text.slice(0, count) : count == 0 ? '' : count >= text.length ? text : text.slice(text.length - count)
    return writeStringResult(base, result, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if ((builtinId == BuiltinId.Leftb || builtinId == BuiltinId.Rightb) && (argc == 1 || argc == 2)) {
    const text = scalarText(
      tagStack[base],
      valueStack[base],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    const count = argc == 2 ? coerceLength(tagStack[base + 1], valueStack[base + 1], 1) : 1
    if (text == null || count == i32.MIN_VALUE) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const result = builtinId == BuiltinId.Leftb ? leftBytesText(text, count) : rightBytesText(text, count)
    return writeStringResult(base, result, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Mid && argc == 3) {
    const text = scalarText(
      tagStack[base],
      valueStack[base],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    const start = coercePositiveStart(tagStack[base + 1], valueStack[base + 1], 1)
    const count = coerceLength(tagStack[base + 2], valueStack[base + 2], 0)
    if (text == null || start == i32.MIN_VALUE || count == i32.MIN_VALUE) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeStringResult(base, text.slice(start - 1, start - 1 + count), rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Midb && argc == 3) {
    const text = scalarText(
      tagStack[base],
      valueStack[base],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    const start = coercePositiveStart(tagStack[base + 1], valueStack[base + 1], 1)
    const count = coerceLength(tagStack[base + 2], valueStack[base + 2], 0)
    if (text == null || start == i32.MIN_VALUE || count == i32.MIN_VALUE) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeStringResult(base, midBytesText(text, start, count), rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if (builtinId == BuiltinId.Trim && argc == 1) {
    const text = scalarText(
      tagStack[base],
      valueStack[base],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    if (text == null) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeStringResult(base, excelTrim(text), rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if ((builtinId == BuiltinId.Upper || builtinId == BuiltinId.Lower) && argc == 1) {
    const text = scalarText(
      tagStack[base],
      valueStack[base],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    if (text == null) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeStringResult(
      base,
      builtinId == BuiltinId.Upper ? text.toUpperCase() : text.toLowerCase(),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    )
  }

  if (builtinId == BuiltinId.Find && (argc == 2 || argc == 3)) {
    const needle = scalarText(
      tagStack[base],
      valueStack[base],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    const haystack = scalarText(
      tagStack[base + 1],
      valueStack[base + 1],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    const start = argc == 3 ? coercePositiveStart(tagStack[base + 2], valueStack[base + 2], 1) : 1
    if (needle == null || haystack == null || start == i32.MIN_VALUE) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const found = findPosition(needle, haystack, start, true, false)
    if (found == i32.MIN_VALUE) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, <f64>found, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  if ((builtinId == BuiltinId.Findb || builtinId == BuiltinId.Searchb) && (argc == 2 || argc == 3)) {
    const needle = scalarText(
      tagStack[base],
      valueStack[base],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    const haystack = scalarText(
      tagStack[base + 1],
      valueStack[base + 1],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    if (needle == null || haystack == null) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    let start = 1
    if (argc == 3) {
      const startByte = coercePositiveStart(tagStack[base + 2], valueStack[base + 2], 1)
      if (startByte == i32.MIN_VALUE || startByte > utf8ByteLength(haystack) + 1) {
        return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
      }
      start = bytePositionToCharPositionUtf8(haystack, startByte)
    }
    const found = findPosition(needle, haystack, start, builtinId == BuiltinId.Findb, builtinId == BuiltinId.Searchb)
    if (found == i32.MIN_VALUE) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeResult(
      base,
      STACK_KIND_SCALAR,
      <u8>ValueTag.Number,
      <f64>charPositionToBytePositionUtf8(haystack, found),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    )
  }

  if (builtinId == BuiltinId.Search && (argc == 2 || argc == 3)) {
    const needle = scalarText(
      tagStack[base],
      valueStack[base],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    const haystack = scalarText(
      tagStack[base + 1],
      valueStack[base + 1],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    )
    const start = argc == 3 ? coercePositiveStart(tagStack[base + 2], valueStack[base + 2], 1) : 1
    if (needle == null || haystack == null || start == i32.MIN_VALUE) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    const found = findPosition(needle, haystack, start, false, true)
    if (found == i32.MIN_VALUE) {
      return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Error, ErrorCode.Value, rangeIndexStack, valueStack, tagStack, kindStack)
    }
    return writeResult(base, STACK_KIND_SCALAR, <u8>ValueTag.Number, <f64>found, rangeIndexStack, valueStack, tagStack, kindStack)
  }

  return -1
}
