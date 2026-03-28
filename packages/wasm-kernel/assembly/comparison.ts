import { ValueTag } from "./protocol";
import { scalarText } from "./text-codec";
import { parseNumericText } from "./text-special";
import { toNumberOrNaN } from "./operands";

export function valueNumber(
  tag: u8,
  value: f64,
  stringOffsets: Uint32Array,
  stringLengths: Uint32Array,
  stringData: Uint16Array,
  outputStringOffsets: Uint32Array,
  outputStringLengths: Uint32Array,
  outputStringData: Uint16Array,
): f64 {
  if (tag == ValueTag.Number || tag == ValueTag.Boolean) {
    return value;
  }
  if (tag == ValueTag.Empty) {
    return 0;
  }
  if (tag != ValueTag.String) {
    return NaN;
  }
  const text = scalarText(
    tag,
    value,
    stringOffsets,
    stringLengths,
    stringData,
    outputStringOffsets,
    outputStringLengths,
    outputStringData,
  );
  return text == null ? NaN : parseNumericText(text);
}

export function compareScalarValues(
  leftTag: u8,
  leftValue: f64,
  rightTag: u8,
  rightValue: f64,
  rightText: string | null,
  stringOffsets: Uint32Array,
  stringLengths: Uint32Array,
  stringData: Uint16Array,
  outputStringOffsets: Uint32Array,
  outputStringLengths: Uint32Array,
  outputStringData: Uint16Array,
): i32 {
  const leftTextlike = leftTag == ValueTag.String || leftTag == ValueTag.Empty;
  const rightTextlike = rightTag == ValueTag.String || rightTag == ValueTag.Empty;
  if (leftTextlike && rightTextlike) {
    const leftText = scalarText(
      leftTag,
      leftValue,
      stringOffsets,
      stringLengths,
      stringData,
      outputStringOffsets,
      outputStringLengths,
      outputStringData,
    );
    const resolvedRightText =
      rightText != null
        ? rightText
        : scalarText(
            rightTag,
            rightValue,
            stringOffsets,
            stringLengths,
            stringData,
            outputStringOffsets,
            outputStringLengths,
            outputStringData,
          );
    if (leftText == null || resolvedRightText == null) {
      return i32.MIN_VALUE;
    }
    const normalizedLeft = leftText.toUpperCase();
    const normalizedRight = resolvedRightText.toUpperCase();
    if (normalizedLeft == normalizedRight) {
      return 0;
    }
    return normalizedLeft < normalizedRight ? -1 : 1;
  }

  const leftNumeric = toNumberOrNaN(leftTag, leftValue);
  const rightNumeric = toNumberOrNaN(rightTag, rightValue);
  if (isNaN(leftNumeric) || isNaN(rightNumeric)) {
    return i32.MIN_VALUE;
  }
  if (leftNumeric == rightNumeric) {
    return 0;
  }
  return leftNumeric < rightNumeric ? -1 : 1;
}
