import { ValueTag } from "./protocol";

const OUTPUT_STRING_BASE: f64 = 2147483648.0;

function outputStringIndex(value: f64): i32 {
  if (value < OUTPUT_STRING_BASE) {
    return -1;
  }
  return <i32>(value - OUTPUT_STRING_BASE);
}

export function textLength(
  tag: u8,
  value: f64,
  stringLengths: Uint32Array,
  outputStringLengths: Uint32Array,
): i32 {
  if (tag == ValueTag.Empty) {
    return 0;
  }
  if (tag == ValueTag.Boolean) {
    return value != 0 ? 4 : 5;
  }
  if (tag == ValueTag.Number) {
    return value.toString().length;
  }
  if (tag == ValueTag.String) {
    const outputIndex = outputStringIndex(value);
    if (outputIndex >= 0) {
      const index = outputIndex;
      if (index < 0 || index >= outputStringLengths.length) {
        return -1;
      }
      return <i32>outputStringLengths[index];
    }
    const stringId = <i32>value;
    if (stringId < 0 || stringId >= stringLengths.length) {
      return -1;
    }
    return <i32>stringLengths[stringId];
  }
  return -1;
}

export function utf8ByteLength(text: string): i32 {
  return String.UTF8.byteLength(text, false);
}

function utf8CodeUnitByteLength(text: string, index: i32): i32 {
  const code = text.charCodeAt(index);
  if (code < 0x80) return 1;
  if (code < 0x800) return 2;
  if ((code & 0xfc00) == 0xd800 && index + 1 < text.length) {
    const next = text.charCodeAt(index + 1);
    if ((next & 0xfc00) == 0xdc00) {
      return 4;
    }
  }
  return 3;
}

function utf8DecodeReplace(bytes: Uint8Array): string {
  let result = "";
  let index = 0;
  while (index < bytes.length) {
    const b0 = <u32>bytes[index];
    if (b0 < 0x80) {
      result += String.fromCharCode(<i32>b0);
      index += 1;
      continue;
    }
    if (b0 >= 0xc2 && b0 <= 0xdf) {
      if (index + 1 < bytes.length) {
        const b1 = <u32>bytes[index + 1];
        if ((b1 & 0xc0) == 0x80) {
          result += String.fromCharCode(<i32>(((b0 & 0x1f) << 6) | (b1 & 0x3f)));
          index += 2;
          continue;
        }
      }
      result += "\ufffd";
      index += 1;
      continue;
    }
    if (b0 >= 0xe0 && b0 <= 0xef) {
      if (index + 2 < bytes.length) {
        const b1 = <u32>bytes[index + 1];
        const b2 = <u32>bytes[index + 2];
        const validSecond =
          (b1 & 0xc0) == 0x80 &&
          (b2 & 0xc0) == 0x80 &&
          (b0 != 0xe0 || b1 >= 0xa0) &&
          (b0 != 0xed || b1 < 0xa0);
        if (validSecond) {
          const codePoint = ((b0 & 0x0f) << 12) | ((b1 & 0x3f) << 6) | (b2 & 0x3f);
          result += String.fromCharCode(<i32>codePoint);
          index += 3;
          continue;
        }
      }
      result += "\ufffd";
      index += 1;
      continue;
    }
    if (b0 >= 0xf0 && b0 <= 0xf4) {
      if (index + 3 < bytes.length) {
        const b1 = <u32>bytes[index + 1];
        const b2 = <u32>bytes[index + 2];
        const b3 = <u32>bytes[index + 3];
        const validFour =
          (b1 & 0xc0) == 0x80 &&
          (b2 & 0xc0) == 0x80 &&
          (b3 & 0xc0) == 0x80 &&
          (b0 != 0xf0 || b1 >= 0x90) &&
          (b0 != 0xf4 || b1 < 0x90);
        if (validFour) {
          let codePoint =
            ((b0 & 0x07) << 18) | ((b1 & 0x3f) << 12) | ((b2 & 0x3f) << 6) | (b3 & 0x3f);
          codePoint -= 0x10000;
          result += String.fromCharCode(
            <i32>(0xd800 | (codePoint >> 10)),
            <i32>(0xdc00 | (codePoint & 0x03ff)),
          );
          index += 4;
          continue;
        }
      }
      result += "\ufffd";
      index += 1;
      continue;
    }
    result += "\ufffd";
    index += 1;
  }
  return result;
}

export function leftBytesText(text: string, count: i32): string {
  const bytes = Uint8Array.wrap(String.UTF8.encode(text));
  const normalized = max<i32>(0, min<i32>(count, bytes.length));
  return utf8DecodeReplace(bytes.subarray(0, normalized));
}

export function rightBytesText(text: string, count: i32): string {
  const bytes = Uint8Array.wrap(String.UTF8.encode(text));
  const normalized = max<i32>(0, min<i32>(count, bytes.length));
  return utf8DecodeReplace(bytes.subarray(bytes.length - normalized));
}

export function midBytesText(text: string, start: i32, count: i32): string {
  if (count <= 0) {
    return "";
  }
  const bytes = Uint8Array.wrap(String.UTF8.encode(text));
  const zeroBasedStart = max<i32>(0, start - 1);
  if (zeroBasedStart >= bytes.length) {
    return "";
  }
  const zeroBasedEnd = min<i32>(bytes.length, zeroBasedStart + count);
  return utf8DecodeReplace(bytes.subarray(zeroBasedStart, zeroBasedEnd));
}

export function replaceBytesText(
  text: string,
  start: i32,
  count: i32,
  replacement: string,
): string {
  const textBytes = Uint8Array.wrap(String.UTF8.encode(text));
  const zeroBasedStart = max<i32>(0, start - 1);
  if (zeroBasedStart >= textBytes.length) {
    return text;
  }
  const zeroBasedEnd = min<i32>(textBytes.length, zeroBasedStart + max<i32>(0, count));
  const replacementBytes = Uint8Array.wrap(String.UTF8.encode(replacement));
  const resultBytes = new Uint8Array(
    zeroBasedStart + replacementBytes.length + (textBytes.length - zeroBasedEnd),
  );
  let cursor = 0;
  for (let index = 0; index < zeroBasedStart; index += 1) {
    resultBytes[cursor] = textBytes[index];
    cursor += 1;
  }
  for (let index = 0; index < replacementBytes.length; index += 1) {
    resultBytes[cursor] = replacementBytes[index];
    cursor += 1;
  }
  for (let index = zeroBasedEnd; index < textBytes.length; index += 1) {
    resultBytes[cursor] = textBytes[index];
    cursor += 1;
  }
  return utf8DecodeReplace(resultBytes);
}

export function bytePositionToCharPositionUtf8(text: string, startByte: i32): i32 {
  if (startByte <= 1) {
    return 1;
  }
  const targetOffset = startByte - 1;
  let byteOffset = 0;
  let index = 0;
  while (index < text.length && byteOffset < targetOffset) {
    const step = utf8CodeUnitByteLength(text, index);
    if (byteOffset + step > targetOffset) {
      return index + 2;
    }
    byteOffset += step;
    if (step == 4 && index + 1 < text.length) {
      const next = text.charCodeAt(index + 1);
      if ((next & 0xfc00) == 0xdc00) {
        index += 2;
        continue;
      }
    }
    index += 1;
  }
  return byteOffset == targetOffset ? index + 1 : text.length + 1;
}

export function charPositionToBytePositionUtf8(text: string, charPosition: i32): i32 {
  const end = max<i32>(0, min<i32>(text.length, charPosition - 1));
  let byteOffset = 0;
  let index = 0;
  while (index < end) {
    const step = utf8CodeUnitByteLength(text, index);
    byteOffset += step;
    if (step == 4 && index + 1 < end) {
      const next = text.charCodeAt(index + 1);
      if ((next & 0xfc00) == 0xdc00) {
        index += 2;
        continue;
      }
    }
    index += 1;
  }
  return byteOffset + 1;
}

export function poolString(
  stringId: i32,
  stringOffsets: Uint32Array,
  stringLengths: Uint32Array,
  stringData: Uint16Array,
): string | null {
  if (stringId < 0 || stringId >= stringLengths.length) {
    return null;
  }
  const offset = <i32>stringOffsets[stringId];
  const length = <i32>stringLengths[stringId];
  let text = "";
  for (let index = 0; index < length; index++) {
    text += String.fromCharCode(stringData[offset + index]);
  }
  return text;
}

export function scalarText(
  tag: u8,
  value: f64,
  stringOffsets: Uint32Array,
  stringLengths: Uint32Array,
  stringData: Uint16Array,
  outputStringOffsets: Uint32Array,
  outputStringLengths: Uint32Array,
  outputStringData: Uint16Array,
): string | null {
  if (tag == ValueTag.Empty) {
    return "";
  }
  if (tag == ValueTag.Boolean) {
    return value != 0 ? "TRUE" : "FALSE";
  }
  if (tag == ValueTag.Number) {
    return value.toString();
  }
  if (tag == ValueTag.String) {
    const outputIndex = outputStringIndex(value);
    if (outputIndex >= 0) {
      const index = outputIndex;
      if (index < 0 || index >= outputStringLengths.length) return null;
      const offset = <i32>outputStringOffsets[index];
      const length = <i32>outputStringLengths[index];
      let text = "";
      for (let i = 0; i < length; i++) {
        text += String.fromCharCode(outputStringData[offset + i]);
      }
      return text;
    }
    const stringId = <i32>value;
    return poolString(stringId, stringOffsets, stringLengths, stringData);
  }
  return null;
}

export function trimAsciiWhitespace(input: string): string {
  let start = 0;
  let end = input.length;
  while (start < end && input.charCodeAt(start) <= 32) {
    start += 1;
  }
  while (end > start && input.charCodeAt(end - 1) <= 32) {
    end -= 1;
  }
  return input.slice(start, end);
}
