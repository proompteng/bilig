import type {
  WorkbookAutoFilterColumnSnapshot,
  WorkbookAutoFilterCustomCriteriaSnapshot,
  WorkbookAutoFilterCustomCriterionSnapshot,
  WorkbookAutoFilterCustomOperator,
  WorkbookAutoFilterSnapshot,
  WorkbookAutoFilterValueCriteriaSnapshot,
} from '@bilig/protocol'

const lessThan = 60
const slash = 47
const colon = 58
const doubleQuote = 34
const singleQuote = 39
const greaterThan = 62

export function readLargeSimpleAutoFiltersFromBytes(
  sheetName: string,
  bytes: Uint8Array,
  startIndex: number,
  endIndex: number,
): WorkbookAutoFilterSnapshot[] {
  const filters: WorkbookAutoFilterSnapshot[] = []
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
      return filters
    }
    if (tag.localName === 'autoFilter') {
      const parsed = parseAutoFilterFromBytes(sheetName, bytes, tag.endIndex, tagEnd, endIndex)
      if (parsed.filter) {
        filters.push(parsed.filter)
      }
      index = parsed.endIndex
      continue
    }
    index = tagEnd + 1
  }
  return filters
}

function parseAutoFilterFromBytes(
  sheetName: string,
  bytes: Uint8Array,
  nameEnd: number,
  tagEnd: number,
  endIndex: number,
): { readonly filter: WorkbookAutoFilterSnapshot | null; readonly endIndex: number } {
  const ref = readAttributeFromTag(bytes, nameEnd, tagEnd, 'ref')
  const range = ref ? parseRangeRef(sheetName, ref) : null
  if (!range) {
    return { filter: null, endIndex: tagEnd + 1 }
  }
  if (isSelfClosingTag(bytes, tagEnd)) {
    return { filter: range, endIndex: tagEnd + 1 }
  }
  const closing = findClosingTag(bytes, tagEnd + 1, 'autoFilter', endIndex)
  if (!closing) {
    return { filter: range, endIndex: tagEnd + 1 }
  }
  const criteria = readFilterColumnsFromBytes(bytes, tagEnd + 1, closing.start)
  return {
    filter: criteria.length > 0 ? { ...range, criteria } : range,
    endIndex: closing.end,
  }
}

function readFilterColumnsFromBytes(bytes: Uint8Array, startIndex: number, endIndex: number): WorkbookAutoFilterColumnSnapshot[] {
  const columns: WorkbookAutoFilterColumnSnapshot[] = []
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
      return columns
    }
    if (tag.localName === 'filterColumn') {
      const parsed = parseFilterColumnFromBytes(bytes, tag.endIndex, tagEnd, endIndex)
      if (parsed.column) {
        columns.push(parsed.column)
      }
      index = parsed.endIndex
      continue
    }
    index = tagEnd + 1
  }
  return columns
}

function parseFilterColumnFromBytes(
  bytes: Uint8Array,
  nameEnd: number,
  tagEnd: number,
  endIndex: number,
): { readonly column: WorkbookAutoFilterColumnSnapshot | null; readonly endIndex: number } {
  const colId = readNonNegativeIntegerAttributeFromTag(bytes, nameEnd, tagEnd, 'colId')
  if (colId === null) {
    return { column: null, endIndex: tagEnd + 1 }
  }
  const hiddenButton = readOptionalBooleanAttributeFromTag(bytes, nameEnd, tagEnd, 'hiddenButton')
  const showButton = readOptionalBooleanAttributeFromTag(bytes, nameEnd, tagEnd, 'showButton')
  let column: WorkbookAutoFilterColumnSnapshot = {
    colId,
    ...(hiddenButton !== undefined ? { hiddenButton } : {}),
    ...(showButton !== undefined ? { showButton } : {}),
  }
  if (isSelfClosingTag(bytes, tagEnd)) {
    return { column, endIndex: tagEnd + 1 }
  }
  const closing = findClosingTag(bytes, tagEnd + 1, 'filterColumn', endIndex)
  if (!closing) {
    return { column, endIndex: tagEnd + 1 }
  }
  let index = tagEnd + 1
  while (index < closing.start) {
    if (bytes[index] !== lessThan) {
      index += 1
      continue
    }
    const tag = readXmlTagName(bytes, index + 1)
    if (!tag) {
      index += 1
      continue
    }
    const childTagEnd = findTagEnd(bytes, tag.endIndex, closing.start)
    if (childTagEnd === null) {
      break
    }
    if (tag.localName === 'filters') {
      const parsed = parseValueFilterCriteriaFromBytes(bytes, tag.endIndex, childTagEnd, closing.start)
      if (parsed.criteria) {
        column.filters = parsed.criteria
      }
      index = parsed.endIndex
      continue
    }
    if (tag.localName === 'customFilters') {
      const parsed = parseCustomFilterCriteriaFromBytes(bytes, tag.endIndex, childTagEnd, closing.start)
      if (parsed.criteria) {
        column.customFilters = parsed.criteria
      }
      index = parsed.endIndex
      continue
    }
    index = childTagEnd + 1
  }
  return { column, endIndex: closing.end }
}

function parseValueFilterCriteriaFromBytes(
  bytes: Uint8Array,
  nameEnd: number,
  tagEnd: number,
  endIndex: number,
): { readonly criteria: WorkbookAutoFilterValueCriteriaSnapshot | null; readonly endIndex: number } {
  const blank = readOptionalBooleanAttributeFromTag(bytes, nameEnd, tagEnd, 'blank')
  const values: string[] = []
  if (isSelfClosingTag(bytes, tagEnd)) {
    return {
      criteria: blank !== undefined ? { blank, values } : null,
      endIndex: tagEnd + 1,
    }
  }
  const closing = findClosingTag(bytes, tagEnd + 1, 'filters', endIndex)
  if (!closing) {
    return { criteria: blank !== undefined ? { blank, values } : null, endIndex: tagEnd + 1 }
  }
  let index = tagEnd + 1
  while (index < closing.start) {
    if (bytes[index] !== lessThan) {
      index += 1
      continue
    }
    const tag = readXmlTagName(bytes, index + 1)
    if (!tag) {
      index += 1
      continue
    }
    const childTagEnd = findTagEnd(bytes, tag.endIndex, closing.start)
    if (childTagEnd === null) {
      break
    }
    if (tag.localName === 'filter') {
      const value = readAttributeFromTag(bytes, tag.endIndex, childTagEnd, 'val')
      if (value !== null) {
        values.push(value)
      }
    }
    index = childTagEnd + 1
  }
  return {
    criteria: values.length > 0 || blank !== undefined ? { ...(blank !== undefined ? { blank } : {}), values } : null,
    endIndex: closing.end,
  }
}

function parseCustomFilterCriteriaFromBytes(
  bytes: Uint8Array,
  nameEnd: number,
  tagEnd: number,
  endIndex: number,
): { readonly criteria: WorkbookAutoFilterCustomCriteriaSnapshot | null; readonly endIndex: number } {
  const and = readOptionalBooleanAttributeFromTag(bytes, nameEnd, tagEnd, 'and')
  const filters: WorkbookAutoFilterCustomCriterionSnapshot[] = []
  if (isSelfClosingTag(bytes, tagEnd)) {
    return { criteria: null, endIndex: tagEnd + 1 }
  }
  const closing = findClosingTag(bytes, tagEnd + 1, 'customFilters', endIndex)
  if (!closing) {
    return { criteria: null, endIndex: tagEnd + 1 }
  }
  let index = tagEnd + 1
  while (index < closing.start) {
    if (bytes[index] !== lessThan) {
      index += 1
      continue
    }
    const tag = readXmlTagName(bytes, index + 1)
    if (!tag) {
      index += 1
      continue
    }
    const childTagEnd = findTagEnd(bytes, tag.endIndex, closing.start)
    if (childTagEnd === null) {
      break
    }
    if (tag.localName === 'customFilter') {
      const value = readAttributeFromTag(bytes, tag.endIndex, childTagEnd, 'val')
      if (value !== null) {
        const operator = parseAutoFilterCustomOperator(readAttributeFromTag(bytes, tag.endIndex, childTagEnd, 'operator'))
        filters.push({ ...(operator ? { operator } : {}), value })
      }
    }
    index = childTagEnd + 1
  }
  return {
    criteria: filters.length > 0 ? { ...(and !== undefined ? { and } : {}), filters } : null,
    endIndex: closing.end,
  }
}

function parseRangeRef(sheetName: string, ref: string): WorkbookAutoFilterSnapshot | null {
  const [startAddress, endAddress = startAddress] = ref.split(':').map((value) => encodeCellAddress(decodeCellAddress(value ?? '')))
  return startAddress && endAddress ? { sheetName, startAddress, endAddress } : null
}

function decodeCellAddress(address: string): { readonly row: number; readonly column: number } | null {
  const trimmed = address.replaceAll('$', '')
  let index = 0
  let column = 0
  let letterCount = 0
  while (index < trimmed.length) {
    const code = trimmed.charCodeAt(index)
    const upper = code >= 97 && code <= 122 ? code - 32 : code
    if (upper < 65 || upper > 90) {
      break
    }
    column = column * 26 + upper - 64
    letterCount += 1
    index += 1
  }
  if (letterCount === 0 || letterCount > 3 || index >= trimmed.length) {
    return null
  }
  let row = 0
  while (index < trimmed.length) {
    const code = trimmed.charCodeAt(index)
    if (code < 48 || code > 57) {
      return null
    }
    row = row * 10 + code - 48
    index += 1
  }
  return row > 0 && column > 0 ? { row: row - 1, column: column - 1 } : null
}

function encodeCellAddress(address: { readonly row: number; readonly column: number } | null): string | null {
  if (!address) {
    return null
  }
  let value = address.column + 1
  let columnName = ''
  while (value > 0) {
    value -= 1
    columnName = String.fromCharCode(65 + (value % 26)) + columnName
    value = Math.floor(value / 26)
  }
  return `${columnName}${String(address.row + 1)}`
}

function parseAutoFilterCustomOperator(value: string | null): WorkbookAutoFilterCustomOperator | undefined {
  if (value === null) {
    return undefined
  }
  switch (value) {
    case 'equal':
    case 'lessThan':
    case 'lessThanOrEqual':
    case 'notEqual':
    case 'greaterThanOrEqual':
    case 'greaterThan':
      return value
    default:
      return undefined
  }
}

function readAttributeFromTag(bytes: Uint8Array, startIndex: number, tagEnd: number, attributeName: string): string | null {
  const range = readAttributeRangeFromTag(bytes, startIndex, tagEnd, attributeName)
  return range ? decodeXmlText(decodeAscii(bytes, range.start, range.end)) : null
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

function readNonNegativeIntegerAttributeFromTag(
  bytes: Uint8Array,
  startIndex: number,
  tagEnd: number,
  attributeName: string,
): number | null {
  const raw = readAttributeFromTag(bytes, startIndex, tagEnd, attributeName)
  if (raw === null || raw.length === 0) {
    return null
  }
  let value = 0
  for (let index = 0; index < raw.length; index += 1) {
    const code = raw.charCodeAt(index)
    if (code < 48 || code > 57) {
      return null
    }
    value = value * 10 + code - 48
    if (!Number.isSafeInteger(value)) {
      return null
    }
  }
  return value
}

function readOptionalBooleanAttributeFromTag(
  bytes: Uint8Array,
  startIndex: number,
  tagEnd: number,
  attributeName: string,
): boolean | undefined {
  const raw = readAttributeFromTag(bytes, startIndex, tagEnd, attributeName)
  if (raw === '1' || raw === 'true') {
    return true
  }
  if (raw === '0' || raw === 'false') {
    return false
  }
  return undefined
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
  return index === localNameStart ? null : { localName: decodeAscii(bytes, localNameStart, index), endIndex: index }
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

function findClosingTag(
  bytes: Uint8Array,
  startIndex: number,
  localName: string,
  endIndex: number,
): { readonly start: number; readonly end: number } | null {
  let index = startIndex
  while (index < endIndex) {
    if (bytes[index] !== lessThan || bytes[index + 1] !== slash) {
      index += 1
      continue
    }
    const tag = readXmlTagName(bytes, index + 2)
    if (tag?.localName === localName) {
      const tagEnd = findTagEnd(bytes, tag.endIndex, endIndex)
      return tagEnd === null ? null : { start: index, end: tagEnd + 1 }
    }
    index += 1
  }
  return null
}

function isSelfClosingTag(bytes: Uint8Array, tagEnd: number): boolean {
  let index = tagEnd - 1
  while (index >= 0 && isAsciiWhitespace(bytes[index] ?? 0)) {
    index -= 1
  }
  return bytes[index] === slash
}

function attributeNameMatches(bytes: Uint8Array, startIndex: number, endIndex: number, attributeName: string): boolean {
  const unprefixedStart = findUnprefixedNameStart(bytes, startIndex, endIndex)
  if (endIndex - unprefixedStart !== attributeName.length) {
    return false
  }
  for (let index = 0; index < attributeName.length; index += 1) {
    if (bytes[unprefixedStart + index] !== attributeName.charCodeAt(index)) {
      return false
    }
  }
  return true
}

function findUnprefixedNameStart(bytes: Uint8Array, startIndex: number, endIndex: number): number {
  for (let index = startIndex; index < endIndex; index += 1) {
    if (bytes[index] === colon) {
      return index + 1
    }
  }
  return startIndex
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

function decodeXmlText(value: string): string {
  return value.replace(/&(#x[0-9a-fA-F]+|#[0-9]+|amp|lt|gt|quot|apos);/gu, (_match, entity: string) => {
    if (entity.startsWith('#x')) {
      const codePoint = Number.parseInt(entity.slice(2), 16)
      return isValidXmlCodePoint(codePoint) ? String.fromCodePoint(codePoint) : ''
    }
    if (entity.startsWith('#')) {
      const codePoint = Number.parseInt(entity.slice(1), 10)
      return isValidXmlCodePoint(codePoint) ? String.fromCodePoint(codePoint) : ''
    }
    switch (entity) {
      case 'amp':
        return '&'
      case 'lt':
        return '<'
      case 'gt':
        return '>'
      case 'quot':
        return '"'
      case 'apos':
        return "'"
      default:
        return ''
    }
  })
}

function isValidXmlCodePoint(value: number): boolean {
  return Number.isInteger(value) && value >= 0 && value <= 0x10ffff
}

function isAsciiWhitespace(byte: number): boolean {
  return byte === 9 || byte === 10 || byte === 12 || byte === 13 || byte === 32
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
