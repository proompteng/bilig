import { BuiltinId, ErrorCode, ValueTag } from "./protocol";
import { scalarErrorAt } from "./builtin-args";
import { replaceBytesText, scalarText } from "./text-codec";
import { repeatText, replaceText, substituteNthText, substituteText } from "./text-ops";
import { STACK_KIND_SCALAR, writeResult, writeStringResult } from "./result-io";
import { coerceLength, coerceNonNegativeLength, coercePositiveStart } from "./text-foundation";

export function tryApplyTextMutationBuiltin(
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
  if (builtinId == BuiltinId.Replaceb && argc == 4) {
    const text = scalarText(
      tagStack[base],
      valueStack[base],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    const start = coercePositiveStart(tagStack[base + 1], valueStack[base + 1], 1);
    const count = coerceLength(tagStack[base + 2], valueStack[base + 2], 0);
    const replacement = scalarText(
      tagStack[base + 3],
      valueStack[base + 3],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    if (text == null || start == i32.MIN_VALUE || count == i32.MIN_VALUE || replacement == null) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeStringResult(
      base,
      replaceBytesText(text, start, count, replacement),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Replace && argc == 4) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack);
    if (scalarError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        scalarError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const text = scalarText(
      tagStack[base],
      valueStack[base],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    const start = coercePositiveStart(tagStack[base + 1], valueStack[base + 1], 1);
    const count = coerceLength(tagStack[base + 2], valueStack[base + 2], 0);
    const replacement = scalarText(
      tagStack[base + 3],
      valueStack[base + 3],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    if (text == null || start == i32.MIN_VALUE || count == i32.MIN_VALUE || replacement == null) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeStringResult(
      base,
      replaceText(text, start, count, replacement),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Substitute && (argc == 3 || argc == 4)) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack);
    if (scalarError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        scalarError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const text = scalarText(
      tagStack[base],
      valueStack[base],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    const oldText = scalarText(
      tagStack[base + 1],
      valueStack[base + 1],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    const newText = scalarText(
      tagStack[base + 2],
      valueStack[base + 2],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    if (text == null || oldText == null || newText == null || oldText.length == 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    if (argc == 3) {
      return writeStringResult(
        base,
        substituteText(text, oldText, newText),
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const instance = coercePositiveStart(tagStack[base + 3], valueStack[base + 3], 1);
    if (instance == i32.MIN_VALUE) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeStringResult(
      base,
      substituteNthText(text, oldText, newText, instance),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  if (builtinId == BuiltinId.Rept && argc == 2) {
    const scalarError = scalarErrorAt(base, argc, kindStack, tagStack, valueStack);
    if (scalarError >= 0) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        scalarError,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    const text = scalarText(
      tagStack[base],
      valueStack[base],
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    const count = coerceNonNegativeLength(tagStack[base + 1], valueStack[base + 1]);
    if (text == null || count == i32.MIN_VALUE) {
      return writeResult(
        base,
        STACK_KIND_SCALAR,
        <u8>ValueTag.Error,
        ErrorCode.Value,
        rangeIndexStack,
        valueStack,
        tagStack,
        kindStack,
      );
    }
    return writeStringResult(
      base,
      repeatText(text, count),
      rangeIndexStack,
      valueStack,
      tagStack,
      kindStack,
    );
  }

  return -1;
}
