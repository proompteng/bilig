import { strFromU8 } from 'fflate'

import type { LiteralInput } from '@bilig/protocol'
import type { LargeSimpleSharedStringEntry } from './xlsx-large-simple-shared-strings.js'
import { decodeXmlText, normalizeWorksheetText } from './xlsx-large-simple-worksheet-stream-text.js'

export interface LargeSimpleXmlTextRange {
  readonly start: number
  readonly end: number
}

export function readLargeSimpleCellValueFromTextRange(
  bytes: Uint8Array,
  rawValueRange: LargeSimpleXmlTextRange | null,
  type: string | null,
  sharedStrings: readonly LargeSimpleSharedStringEntry[],
): LiteralInput | undefined {
  switch (type) {
    case null:
      return rawValueRange ? readNumberValue(bytes, rawValueRange) : undefined
    case 's': {
      const index = readLargeSimpleSharedStringIndexFromTextRange(bytes, rawValueRange)
      return index === null ? undefined : sharedStrings[index]?.text
    }
    case 'inlineStr':
      return undefined
    case 'str':
      return rawValueRange ? normalizeWorksheetText(decodeXmlText(decodeBytes(bytes, rawValueRange.start, rawValueRange.end))) : undefined
    case 'b':
      return rawValueRange ? readBooleanValue(bytes, rawValueRange) : undefined
    case 'e':
      return rawValueRange ? decodeXmlText(decodeBytes(bytes, rawValueRange.start, rawValueRange.end)) : undefined
    case 'd':
      return undefined
    default:
      return undefined
  }
}

export function readLargeSimpleSharedStringIndexFromTextRange(
  bytes: Uint8Array,
  rawValueRange: LargeSimpleXmlTextRange | null,
): number | null {
  if (!rawValueRange) {
    return null
  }
  if (rangeContainsByte(bytes, rawValueRange, 38)) {
    const index = Number(decodeXmlText(decodeBytes(bytes, rawValueRange.start, rawValueRange.end)).trim())
    return Number.isSafeInteger(index) && index >= 0 ? index : null
  }
  const integerValue = readNonNegativeInteger(bytes, rawValueRange)
  if (integerValue !== null) {
    return integerValue
  }
  const index = Number(decodeBytes(bytes, rawValueRange.start, rawValueRange.end).trim())
  return Number.isSafeInteger(index) && index >= 0 ? index : null
}

function readNumberValue(bytes: Uint8Array, range: LargeSimpleXmlTextRange): number | undefined {
  if (rangeContainsByte(bytes, range, 38)) {
    const value = Number(decodeXmlText(decodeBytes(bytes, range.start, range.end)).trim())
    return Number.isFinite(value) ? value : undefined
  }
  const integerValue = readSafeIntegerNumber(bytes, range)
  if (integerValue !== null) {
    return integerValue
  }
  const value = Number(decodeBytes(bytes, range.start, range.end).trim())
  return Number.isFinite(value) ? value : undefined
}

function readBooleanValue(bytes: Uint8Array, range: LargeSimpleXmlTextRange): boolean | undefined {
  const trimmed = trimmedRange(bytes, range)
  if (trimmed.start + 1 === trimmed.end) {
    const byte = bytes[trimmed.start]
    return byte === 49 ? true : byte === 48 ? false : undefined
  }
  if (attributeValueMatches(bytes, trimmed.start, trimmed.end, 'true')) {
    return true
  }
  if (attributeValueMatches(bytes, trimmed.start, trimmed.end, 'false')) {
    return false
  }
  return undefined
}

function readNonNegativeInteger(bytes: Uint8Array, range: LargeSimpleXmlTextRange): number | null {
  const trimmed = trimmedRange(bytes, range)
  if (trimmed.start === trimmed.end) {
    return null
  }
  let value = 0
  for (let index = trimmed.start; index < trimmed.end; index += 1) {
    const byte = bytes[index] ?? 0
    if (byte < 48 || byte > 57) {
      return null
    }
    value = value * 10 + byte - 48
    if (!Number.isSafeInteger(value)) {
      return null
    }
  }
  return value
}

function readSafeIntegerNumber(bytes: Uint8Array, range: LargeSimpleXmlTextRange): number | null {
  const trimmed = trimmedRange(bytes, range)
  let index = trimmed.start
  const end = trimmed.end
  if (index === end) {
    return null
  }
  let sign = 1
  const first = bytes[index]
  if (first === 43 || first === 45) {
    sign = first === 45 ? -1 : 1
    index += 1
  }
  if (index === end) {
    return null
  }
  let value = 0
  while (index < end) {
    const byte = bytes[index] ?? 0
    if (byte < 48 || byte > 57) {
      return null
    }
    value = value * 10 + byte - 48
    if (!Number.isSafeInteger(value)) {
      return null
    }
    index += 1
  }
  return sign * value
}

function trimmedRange(bytes: Uint8Array, range: LargeSimpleXmlTextRange): LargeSimpleXmlTextRange {
  let start = range.start
  let end = range.end
  while (start < end && isAsciiWhitespace(bytes[start] ?? 0)) {
    start += 1
  }
  while (end > start && isAsciiWhitespace(bytes[end - 1] ?? 0)) {
    end -= 1
  }
  return { start, end }
}

function rangeContainsByte(bytes: Uint8Array, range: LargeSimpleXmlTextRange, target: number): boolean {
  for (let index = range.start; index < range.end; index += 1) {
    if (bytes[index] === target) {
      return true
    }
  }
  return false
}

function attributeValueMatches(bytes: Uint8Array, startIndex: number, endIndex: number, value: string): boolean {
  if (endIndex - startIndex !== value.length) {
    return false
  }
  for (let index = 0; index < value.length; index += 1) {
    if (bytes[startIndex + index] !== value.charCodeAt(index)) {
      return false
    }
  }
  return true
}

function decodeBytes(bytes: Uint8Array, startIndex: number, endIndex: number): string {
  return strFromU8(bytes.subarray(startIndex, endIndex))
}

function isAsciiWhitespace(byte: number): boolean {
  return byte === 9 || byte === 10 || byte === 12 || byte === 13 || byte === 32
}
