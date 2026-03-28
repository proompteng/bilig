import { ValueTag } from "./protocol";
import { toNumberExact } from "./operands";

export function coerceLength(tag: u8, value: f64, defaultValue: i32): i32 {
  if (tag == ValueTag.Empty) {
    return defaultValue;
  }
  const numeric = toNumberExact(tag, value);
  if (isNaN(numeric)) {
    return i32.MIN_VALUE;
  }
  const truncated = <i32>numeric;
  return truncated >= 0 ? truncated : i32.MIN_VALUE;
}

export function coercePositiveStart(tag: u8, value: f64, defaultValue: i32): i32 {
  if (tag == ValueTag.Empty) {
    return defaultValue;
  }
  const numeric = toNumberExact(tag, value);
  if (isNaN(numeric)) {
    return i32.MIN_VALUE;
  }
  const truncated = <i32>numeric;
  return truncated >= 1 ? truncated : i32.MIN_VALUE;
}

export function coerceNonNegativeLength(tag: u8, value: f64): i32 {
  const numeric = toNumberExact(tag, value);
  if (isNaN(numeric)) {
    return i32.MIN_VALUE;
  }
  const truncated = <i32>numeric;
  return truncated >= 0 ? truncated : i32.MIN_VALUE;
}

export function excelTrim(input: string): string {
  let start = 0;
  let end = input.length;
  while (start < end && input.charCodeAt(start) == 32) {
    start += 1;
  }
  while (end > start && input.charCodeAt(end - 1) == 32) {
    end -= 1;
  }
  let result = "";
  let previousSpace = false;
  for (let index = start; index < end; index += 1) {
    const char = input.charCodeAt(index);
    if (char == 32) {
      if (!previousSpace) {
        result += " ";
      }
      previousSpace = true;
      continue;
    }
    previousSpace = false;
    result += String.fromCharCode(char);
  }
  return result;
}

function hasSearchSyntax(pattern: string): bool {
  for (let index = 0; index < pattern.length; index += 1) {
    const char = pattern.charCodeAt(index);
    if (char == 126 || char == 42 || char == 63) {
      return true;
    }
  }
  return false;
}

function wildcardMatchAt(
  pattern: string,
  haystack: string,
  patternIndex: i32,
  haystackIndex: i32,
): bool {
  let p = patternIndex;
  let h = haystackIndex;
  while (p < pattern.length) {
    const char = pattern.charCodeAt(p);
    if (char == 126) {
      const nextIndex = p + 1;
      const nextChar = nextIndex < pattern.length ? pattern.charCodeAt(nextIndex) : 126;
      if (h >= haystack.length || haystack.charCodeAt(h) != nextChar) {
        return false;
      }
      p = nextIndex < pattern.length ? nextIndex + 1 : nextIndex;
      h += 1;
      continue;
    }
    if (char == 42) {
      let nextPatternIndex = p + 1;
      while (nextPatternIndex < pattern.length && pattern.charCodeAt(nextPatternIndex) == 42) {
        nextPatternIndex += 1;
      }
      if (nextPatternIndex >= pattern.length) {
        return true;
      }
      for (let scan = h; scan <= haystack.length; scan += 1) {
        if (wildcardMatchAt(pattern, haystack, nextPatternIndex, scan)) {
          return true;
        }
      }
      return false;
    }
    if (h >= haystack.length) {
      return false;
    }
    if (char == 63) {
      p += 1;
      h += 1;
      continue;
    }
    if (haystack.charCodeAt(h) != char) {
      return false;
    }
    p += 1;
    h += 1;
  }
  return true;
}

export function findPosition(
  needle: string,
  haystack: string,
  start: i32,
  caseSensitive: bool,
  wildcardAware: bool,
): i32 {
  const startIndex = start - 1;
  if (needle.length == 0) {
    return start;
  }
  if (startIndex > haystack.length) {
    return i32.MIN_VALUE;
  }
  const normalizedHaystack = caseSensitive ? haystack : haystack.toLowerCase();
  const normalizedNeedle = caseSensitive ? needle : needle.toLowerCase();
  if (wildcardAware && hasSearchSyntax(normalizedNeedle)) {
    for (let index = startIndex; index <= normalizedHaystack.length; index += 1) {
      if (wildcardMatchAt(normalizedNeedle, normalizedHaystack, 0, index)) {
        return index + 1;
      }
    }
    return i32.MIN_VALUE;
  }
  const found = normalizedHaystack.indexOf(normalizedNeedle, startIndex);
  return found < 0 ? i32.MIN_VALUE : found + 1;
}
