import type { WorkbookAxisEntrySnapshot, WorkbookAxisMetadataSnapshot } from '@bilig/protocol'
import { appendRowMetadata, type LargeSimpleWorksheetMergeRef } from './xlsx-large-simple-worksheet-metadata.js'
import { readKnownXmlLocalName } from './xlsx-large-simple-xml-name.js'

const lessThan = 60
const slash = 47
const colon = 58
const doubleQuote = 34
const singleQuote = 39
const greaterThan = 62

export function appendLargeSimpleRowMetadataTagFromBytes(
  entries: WorkbookAxisEntrySnapshot[],
  metadata: WorkbookAxisMetadataSnapshot[],
  bytes: Uint8Array,
  nameEnd: number,
  tagEnd: number,
): void {
  const row = readPositiveIntegerAttributeFromTag(bytes, nameEnd, tagEnd, 'r')
  if (row === null) {
    return
  }
  appendRowMetadata(entries, metadata, row - 1, {
    height: readNumberAttributeFromTag(bytes, nameEnd, tagEnd, 'ht'),
    styleIndex: readNonNegativeIntegerAttributeFromTag(bytes, nameEnd, tagEnd, 's'),
    hidden: readOptionalBooleanAttributeFromTag(bytes, nameEnd, tagEnd, 'hidden'),
    customFormat: readOptionalBooleanAttributeFromTag(bytes, nameEnd, tagEnd, 'customFormat'),
    customHeight: readOptionalBooleanAttributeFromTag(bytes, nameEnd, tagEnd, 'customHeight'),
    outlineLevel: readNonNegativeIntegerAttributeFromTag(bytes, nameEnd, tagEnd, 'outlineLevel'),
    collapsed: readOptionalBooleanAttributeFromTag(bytes, nameEnd, tagEnd, 'collapsed'),
    thickTop: readOptionalBooleanAttributeFromTag(bytes, nameEnd, tagEnd, 'thickTop'),
    thickBottom: readOptionalBooleanAttributeFromTag(bytes, nameEnd, tagEnd, 'thickBottom'),
  })
}

export function readLargeSimpleMergeRefsFromBytes(bytes: Uint8Array, startIndex: number, endIndex: number): LargeSimpleWorksheetMergeRef[] {
  const refs: LargeSimpleWorksheetMergeRef[] = []
  let index = startIndex
  while (index < endIndex) {
    if (bytes[index] !== lessThan) {
      index += 1
      continue
    }
    const tag = readXmlTagName(bytes, index + 1)
    if (!tag) {
      index += 1
      continue
    }
    const tagEnd = findTagEnd(bytes, tag.endIndex, endIndex)
    if (tagEnd === null) {
      return refs
    }
    if (tag.localName === 'mergeCell') {
      const ref = readAttributeFromTag(bytes, tag.endIndex, tagEnd, 'ref')
      const separator = ref === null ? -1 : ref.indexOf(':')
      const startAddress = separator > 0 ? ref?.slice(0, separator) : undefined
      const endAddress = separator > 0 ? ref?.slice(separator + 1) : undefined
      if (startAddress && endAddress && startAddress !== endAddress) {
        refs.push({ startAddress, endAddress })
      }
    }
    index = tagEnd + 1
  }
  return refs
}

function readNumberAttributeFromTag(bytes: Uint8Array, startIndex: number, tagEnd: number, attributeName: string): number | null {
  const range = readAttributeRangeFromTag(bytes, startIndex, tagEnd, attributeName)
  if (!range) {
    return null
  }
  const trimmed = trimAsciiWhitespace(bytes, range.start, range.end)
  if (trimmed.start === trimmed.end) {
    return null
  }
  const value = parseAsciiNumber(bytes, trimmed.start, trimmed.end)
  return Number.isFinite(value) ? value : null
}

function readPositiveIntegerAttributeFromTag(bytes: Uint8Array, startIndex: number, tagEnd: number, attributeName: string): number | null {
  const value = readNumberAttributeFromTag(bytes, startIndex, tagEnd, attributeName)
  return Number.isInteger(value) && value !== null && value > 0 ? value : null
}

function readNonNegativeIntegerAttributeFromTag(
  bytes: Uint8Array,
  startIndex: number,
  tagEnd: number,
  attributeName: string,
): number | null {
  const value = readNumberAttributeFromTag(bytes, startIndex, tagEnd, attributeName)
  return Number.isInteger(value) && value !== null && value >= 0 ? value : null
}

function readOptionalBooleanAttributeFromTag(bytes: Uint8Array, startIndex: number, tagEnd: number, attributeName: string): boolean | null {
  const range = readAttributeRangeFromTag(bytes, startIndex, tagEnd, attributeName)
  if (!range) {
    return null
  }
  if (range.end - range.start === 1) {
    if (bytes[range.start] === 49) {
      return true
    }
    if (bytes[range.start] === 48) {
      return false
    }
  }
  if (asciiEquals(bytes, range.start, range.end, 'true')) {
    return true
  }
  return asciiEquals(bytes, range.start, range.end, 'false') ? false : null
}

function readAttributeFromTag(bytes: Uint8Array, startIndex: number, tagEnd: number, attributeName: string): string | null {
  const range = readAttributeRangeFromTag(bytes, startIndex, tagEnd, attributeName)
  return range ? decodeAscii(bytes, range.start, range.end) : null
}

function readAttributeRangeFromTag(
  bytes: Uint8Array,
  startIndex: number,
  tagEnd: number,
  attributeName: string,
): { readonly start: number; readonly end: number } | null {
  let index = startIndex
  while (index < tagEnd) {
    while (index < tagEnd && isAsciiWhitespace(bytes[index] ?? 0)) {
      index += 1
    }
    const nameStart = index
    while (index < tagEnd && isXmlNameByte(bytes[index] ?? 0)) {
      index += 1
    }
    const nameEnd = index
    index = skipAsciiWhitespace(bytes, index, tagEnd)
    if (bytes[index] !== 61) {
      index += 1
      continue
    }
    index = skipAsciiWhitespace(bytes, index + 1, tagEnd)
    const quote = bytes[index]
    if (quote !== doubleQuote && quote !== singleQuote) {
      index += 1
      continue
    }
    const valueStart = index + 1
    index = valueStart
    while (index < tagEnd && bytes[index] !== quote) {
      index += 1
    }
    const valueEnd = index
    if (attributeNameMatches(bytes, nameStart, nameEnd, attributeName)) {
      return { start: valueStart, end: valueEnd }
    }
    index += 1
  }
  return null
}

function readXmlTagName(bytes: Uint8Array, startIndex: number): { readonly localName: string; readonly endIndex: number } | null {
  const first = bytes[startIndex]
  if (first === undefined || first === 33 || first === slash || first === 63) {
    return null
  }
  let index = startIndex
  let localNameStart = startIndex
  while (index < bytes.byteLength && isXmlNameByte(bytes[index] ?? 0)) {
    if (bytes[index] === colon) {
      localNameStart = index + 1
    }
    index += 1
  }
  return index === localNameStart
    ? null
    : { localName: readKnownXmlLocalName(bytes, localNameStart, index) ?? decodeAscii(bytes, localNameStart, index), endIndex: index }
}

function findTagEnd(bytes: Uint8Array, startIndex: number, endIndex: number): number | null {
  let quote: number | null = null
  for (let index = startIndex; index < endIndex; index += 1) {
    const byte = bytes[index] ?? 0
    if (quote !== null) {
      if (byte === quote) {
        quote = null
      }
      continue
    }
    if (byte === doubleQuote || byte === singleQuote) {
      quote = byte
      continue
    }
    if (byte === greaterThan) {
      return index
    }
  }
  return null
}

function attributeNameMatches(bytes: Uint8Array, startIndex: number, endIndex: number, attributeName: string): boolean {
  if (endIndex - startIndex !== attributeName.length) {
    return false
  }
  return asciiEquals(bytes, startIndex, endIndex, attributeName)
}

function asciiEquals(bytes: Uint8Array, startIndex: number, endIndex: number, value: string): boolean {
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

function parseAsciiNumber(bytes: Uint8Array, startIndex: number, endIndex: number): number {
  let index = startIndex
  let sign = 1
  if (bytes[index] === 43 || bytes[index] === 45) {
    sign = bytes[index] === 45 ? -1 : 1
    index += 1
  }
  let value = 0
  let digitSeen = false
  while (index < endIndex && isAsciiDigit(bytes[index] ?? 0)) {
    digitSeen = true
    value = value * 10 + (bytes[index] ?? 0) - 48
    index += 1
  }
  if (bytes[index] === 46) {
    index += 1
    let divisor = 10
    while (index < endIndex && isAsciiDigit(bytes[index] ?? 0)) {
      digitSeen = true
      value += ((bytes[index] ?? 0) - 48) / divisor
      divisor *= 10
      index += 1
    }
  }
  if (!digitSeen) {
    return Number.NaN
  }
  let exponent = 0
  let exponentSign = 1
  if (bytes[index] === 69 || bytes[index] === 101) {
    index += 1
    if (bytes[index] === 43 || bytes[index] === 45) {
      exponentSign = bytes[index] === 45 ? -1 : 1
      index += 1
    }
    const exponentStart = index
    while (index < endIndex && isAsciiDigit(bytes[index] ?? 0)) {
      exponent = exponent * 10 + (bytes[index] ?? 0) - 48
      index += 1
    }
    if (index === exponentStart) {
      return Number.NaN
    }
  }
  return index === endIndex ? sign * value * 10 ** (exponentSign * exponent) : Number.NaN
}

function trimAsciiWhitespace(bytes: Uint8Array, startIndex: number, endIndex: number): { readonly start: number; readonly end: number } {
  let start = startIndex
  let end = endIndex
  while (start < end && isAsciiWhitespace(bytes[start] ?? 0)) {
    start += 1
  }
  while (end > start && isAsciiWhitespace(bytes[end - 1] ?? 0)) {
    end -= 1
  }
  return { start, end }
}

function skipAsciiWhitespace(bytes: Uint8Array, startIndex: number, endIndex: number): number {
  let index = startIndex
  while (index < endIndex && isAsciiWhitespace(bytes[index] ?? 0)) {
    index += 1
  }
  return index
}

function decodeAscii(bytes: Uint8Array, startIndex: number, endIndex: number): string {
  let output = ''
  for (let index = startIndex; index < endIndex; index += 1) {
    output += String.fromCharCode(bytes[index] ?? 0)
  }
  return output
}

function isAsciiWhitespace(byte: number): boolean {
  return byte === 9 || byte === 10 || byte === 12 || byte === 13 || byte === 32
}

function isAsciiDigit(byte: number): boolean {
  return byte >= 48 && byte <= 57
}

function isXmlNameByte(byte: number): boolean {
  return (
    (byte >= 65 && byte <= 90) ||
    (byte >= 97 && byte <= 122) ||
    (byte >= 48 && byte <= 57) ||
    byte === 45 ||
    byte === 46 ||
    byte === colon ||
    byte === 95
  )
}
