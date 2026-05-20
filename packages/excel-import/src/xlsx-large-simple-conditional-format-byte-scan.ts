import { strFromU8 } from 'fflate'

import type {
  CellRangeRef,
  LiteralInput,
  WorkbookConditionalFormatRuleSnapshot,
  WorkbookConditionalFormatSnapshot,
  WorkbookValidationComparisonOperator,
} from '@bilig/protocol'
import { readKnownXmlLocalName } from './xlsx-large-simple-xml-name.js'

const lessThan = 60
const slash = 47
const colon = 58
const doubleQuote = 34
const singleQuote = 39
const greaterThan = 62

export interface LargeSimpleConditionalFormattingByteScan {
  readonly conditionalFormats?: WorkbookConditionalFormatSnapshot[]
  readonly artifactXml?: string
  readonly ruleCount: number
}

interface ParsedConditionalFormatRule {
  readonly rule: WorkbookConditionalFormatRuleSnapshot
  readonly priority: number | undefined
  readonly stopIfTrue: boolean | undefined
}

export function readLargeSimpleConditionalFormattingFromBytes(
  sheetName: string,
  bytes: Uint8Array,
  startIndex: number,
  endIndex: number,
  firstId: number,
): LargeSimpleConditionalFormattingByteScan {
  const artifact = (): LargeSimpleConditionalFormattingByteScan => ({
    artifactXml: decodeBytes(bytes, startIndex, endIndex),
    ruleCount: readConditionalFormatRuleCount(bytes, startIndex, endIndex),
  })
  const root = readSingleElement(bytes, startIndex, endIndex, 'conditionalFormatting')
  if (!root || !tagHasOnlyAttributes(bytes, root.nameEnd, root.tagEnd, ['sqref'])) {
    return artifact()
  }
  const sqref = readXmlAttribute(bytes, root.nameEnd, root.tagEnd, 'sqref')
  const ranges = sqref ? parseSqrefRanges(sheetName, decodeXmlText(sqref)) : null
  if (!ranges || ranges.length === 0 || root.selfClosing) {
    return artifact()
  }

  const rules = readConditionalFormatRules(bytes, root.contentStart, root.contentEnd)
  if (!rules || rules.length === 0) {
    return artifact()
  }

  const conditionalFormats: WorkbookConditionalFormatSnapshot[] = []
  for (const parsed of rules) {
    for (const range of ranges) {
      conditionalFormats.push({
        id: `xlsx-cf:${sheetName}:${range.startAddress}:${range.endAddress}:${String(firstId + conditionalFormats.length)}`,
        range,
        rule: parsed.rule,
        style: {},
        ...(parsed.stopIfTrue !== undefined ? { stopIfTrue: parsed.stopIfTrue } : {}),
        ...(parsed.priority !== undefined ? { priority: parsed.priority } : {}),
      })
    }
  }
  return {
    conditionalFormats,
    ruleCount: conditionalFormats.length,
  }
}

function readConditionalFormatRules(bytes: Uint8Array, startIndex: number, endIndex: number): ParsedConditionalFormatRule[] | null {
  const rules: ParsedConditionalFormatRule[] = []
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
    if (tagEnd === null || tag.localName !== 'cfRule') {
      return null
    }
    const selfClosing = isSelfClosingTag(bytes, tagEnd)
    const contentStart = tagEnd + 1
    const closing = selfClosing ? { start: contentStart, end: contentStart } : findClosingTag(bytes, contentStart, 'cfRule', endIndex)
    if (!closing) {
      return null
    }
    const rule = readConditionalFormatRule(bytes, tag.endIndex, tagEnd, contentStart, closing.start)
    if (!rule) {
      return null
    }
    rules.push(rule)
    index = selfClosing ? tagEnd + 1 : closing.end
  }
  return rules
}

function readConditionalFormatRule(
  bytes: Uint8Array,
  nameEnd: number,
  tagEnd: number,
  contentStart: number,
  contentEnd: number,
): ParsedConditionalFormatRule | null {
  const type = readXmlAttribute(bytes, nameEnd, tagEnd, 'type')
  if (!type) {
    return null
  }
  const formulas = readFormulaTexts(bytes, contentStart, contentEnd)
  if (formulas === null) {
    return null
  }
  const priority = numberValue(readXmlAttribute(bytes, nameEnd, tagEnd, 'priority')) ?? undefined
  const stopIfTrue = booleanValue(readXmlAttribute(bytes, nameEnd, tagEnd, 'stopIfTrue'))
  switch (type) {
    case 'cellIs': {
      if (!tagHasOnlyAttributes(bytes, nameEnd, tagEnd, ['type', 'priority', 'stopIfTrue', 'operator']) || formulas.length === 0) {
        return null
      }
      const operator = parseComparisonOperator(readXmlAttribute(bytes, nameEnd, tagEnd, 'operator'))
      const values = formulas.flatMap((formula) => {
        const parsed = parseLiteralFormula(formula)
        return parsed === undefined ? [] : [parsed]
      })
      return operator && values.length > 0 ? { rule: { kind: 'cellIs', operator, values }, priority, stopIfTrue } : null
    }
    case 'expression': {
      if (!tagHasOnlyAttributes(bytes, nameEnd, tagEnd, ['type', 'priority', 'stopIfTrue']) || formulas.length === 0) {
        return null
      }
      const formula = formulas[0]?.trim()
      return formula
        ? { rule: { kind: 'formula', formula: formula.startsWith('=') ? formula : `=${formula}` }, priority, stopIfTrue }
        : null
    }
    case 'containsBlanks':
      return tagHasOnlyAttributes(bytes, nameEnd, tagEnd, ['type', 'priority', 'stopIfTrue']) && formulas.length === 0
        ? { rule: { kind: 'blanks' }, priority, stopIfTrue }
        : null
    case 'notContainsBlanks':
      return tagHasOnlyAttributes(bytes, nameEnd, tagEnd, ['type', 'priority', 'stopIfTrue']) && formulas.length === 0
        ? { rule: { kind: 'notBlanks' }, priority, stopIfTrue }
        : null
    default:
      return null
  }
}

function readFormulaTexts(bytes: Uint8Array, startIndex: number, endIndex: number): string[] | null {
  const formulas: string[] = []
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
    if (tagEnd === null || tag.localName !== 'formula' || !tagHasOnlyAttributes(bytes, tag.endIndex, tagEnd, [])) {
      return null
    }
    const selfClosing = isSelfClosingTag(bytes, tagEnd)
    const contentStart = tagEnd + 1
    const closing = selfClosing ? { start: contentStart, end: contentStart } : findClosingTag(bytes, contentStart, 'formula', endIndex)
    if (!closing || containsOpeningElement(bytes, contentStart, closing.start)) {
      return null
    }
    formulas.push(selfClosing ? '' : decodeXmlText(decodeBytes(bytes, contentStart, closing.start).trim()))
    index = selfClosing ? tagEnd + 1 : closing.end
  }
  return formulas
}

function readConditionalFormatRuleCount(bytes: Uint8Array, startIndex: number, endIndex: number): number {
  const root = readSingleElement(bytes, startIndex, endIndex, 'conditionalFormatting')
  const ranges = root
    ? (parseSqrefRanges('', decodeXmlText(readXmlAttribute(bytes, root.nameEnd, root.tagEnd, 'sqref') ?? ''))?.length ?? 1)
    : 1
  return Math.max(1, ranges * countOpeningTags(bytes, startIndex, endIndex, 'cfRule'))
}

function parseSqrefRanges(sheetName: string, value: string): CellRangeRef[] | null {
  const refs = value.trim().split(/\s+/u).filter(Boolean)
  if (refs.length === 0) {
    return null
  }
  const ranges: CellRangeRef[] = []
  for (const ref of refs) {
    const range = parseSqrefRange(sheetName, ref)
    if (!range) {
      return null
    }
    ranges.push(range)
  }
  return ranges
}

function parseSqrefRange(sheetName: string, ref: string): CellRangeRef | null {
  const separator = ref.indexOf(':')
  const start = decodeCellAddress((separator === -1 ? ref : ref.slice(0, separator)).replaceAll('$', ''))
  const end = decodeCellAddress((separator === -1 ? ref : ref.slice(separator + 1)).replaceAll('$', ''))
  return start && end
    ? {
        sheetName,
        startAddress: encodeCellAddress(start.row, start.column),
        endAddress: encodeCellAddress(end.row, end.column),
      }
    : null
}

function parseLiteralFormula(value: string): LiteralInput | undefined {
  const trimmed = value.trim()
  if (trimmed.length === 0) {
    return ''
  }
  if (trimmed === 'TRUE') {
    return true
  }
  if (trimmed === 'FALSE') {
    return false
  }
  if (trimmed.startsWith('"') && trimmed.endsWith('"')) {
    return trimmed.slice(1, -1).replaceAll('""', '"')
  }
  const number = Number(trimmed)
  return Number.isFinite(number) ? number : trimmed
}

function parseComparisonOperator(value: string | null): WorkbookValidationComparisonOperator | null {
  if (value === null) {
    return null
  }
  switch (value) {
    case 'between':
    case 'notBetween':
    case 'equal':
    case 'notEqual':
    case 'greaterThan':
    case 'greaterThanOrEqual':
    case 'lessThan':
    case 'lessThanOrEqual':
      return value
    default:
      return null
  }
}

function readSingleElement(
  bytes: Uint8Array,
  startIndex: number,
  endIndex: number,
  localName: string,
): {
  readonly nameEnd: number
  readonly tagEnd: number
  readonly contentStart: number
  readonly contentEnd: number
  readonly selfClosing: boolean
} | null {
  const tag = findNextOpeningTag(bytes, startIndex, localName, endIndex)
  if (!tag) {
    return null
  }
  const tagEnd = findTagEnd(bytes, tag.nameEnd, endIndex)
  if (tagEnd === null) {
    return null
  }
  const selfClosing = isSelfClosingTag(bytes, tagEnd)
  if (selfClosing) {
    return { nameEnd: tag.nameEnd, tagEnd, contentStart: tagEnd + 1, contentEnd: tagEnd + 1, selfClosing }
  }
  const closing = findClosingTag(bytes, tagEnd + 1, localName, endIndex)
  return closing ? { nameEnd: tag.nameEnd, tagEnd, contentStart: tagEnd + 1, contentEnd: closing.start, selfClosing } : null
}

function findNextOpeningTag(
  bytes: Uint8Array,
  startIndex: number,
  localName: string,
  endIndex: number,
): { readonly start: number; readonly nameEnd: number } | null {
  let index = startIndex
  while (index < endIndex) {
    if (bytes[index] !== lessThan) {
      index += 1
      continue
    }
    const tag = readXmlTagName(bytes, index + 1)
    if (tag?.localName === localName) {
      return { start: index, nameEnd: tag.endIndex }
    }
    index += 1
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

function isSelfClosingTag(bytes: Uint8Array, tagEnd: number): boolean {
  let index = tagEnd - 1
  while (index >= 0 && isAsciiWhitespace(bytes[index] ?? 0)) {
    index -= 1
  }
  return bytes[index] === slash
}

function countOpeningTags(bytes: Uint8Array, startIndex: number, endIndex: number, localName: string): number {
  let count = 0
  let index = startIndex
  while (index < endIndex) {
    if (bytes[index] !== lessThan) {
      index += 1
      continue
    }
    const tag = readXmlTagName(bytes, index + 1)
    if (tag?.localName === localName) {
      count += 1
      index = tag.endIndex
      continue
    }
    index += 1
  }
  return count
}

function containsOpeningElement(bytes: Uint8Array, startIndex: number, endIndex: number): boolean {
  for (let index = startIndex; index < endIndex; index += 1) {
    if (bytes[index] === lessThan && readXmlTagName(bytes, index + 1)) {
      return true
    }
  }
  return false
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

function tagHasOnlyAttributes(bytes: Uint8Array, startIndex: number, tagEnd: number, allowedNames: readonly string[]): boolean {
  const allowed = new Set(allowedNames)
  let index = startIndex
  while (index < tagEnd) {
    while (index < tagEnd && isAsciiWhitespace(bytes[index] ?? 0)) {
      index += 1
    }
    if (index >= tagEnd || bytes[index] === slash) {
      return true
    }
    const nameStart = index
    while (index < tagEnd && isXmlNameByte(bytes[index] ?? 0)) {
      index += 1
    }
    const nameEnd = index
    index = skipAsciiWhitespace(bytes, index, tagEnd)
    if (bytes[index] !== 61) {
      return false
    }
    index = skipAsciiWhitespace(bytes, index + 1, tagEnd)
    const quote = bytes[index]
    if (quote !== doubleQuote && quote !== singleQuote) {
      return false
    }
    index += 1
    while (index < tagEnd && bytes[index] !== quote) {
      index += 1
    }
    if (index >= tagEnd) {
      return false
    }
    const attributeName = decodeAscii(bytes, nameStart, nameEnd)
    if (!isNamespaceDeclaration(attributeName) && !allowed.has(attributeName)) {
      return false
    }
    index += 1
  }
  return true
}

function readXmlAttribute(bytes: Uint8Array, startIndex: number, tagEnd: number, attributeName: string): string | null {
  const range = readXmlAttributeRange(bytes, startIndex, tagEnd, attributeName)
  return range ? decodeXmlText(decodeBytes(bytes, range.start, range.end)) : null
}

function readXmlAttributeRange(
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

function attributeNameMatches(bytes: Uint8Array, startIndex: number, endIndex: number, attributeName: string): boolean {
  if (endIndex - startIndex !== attributeName.length) {
    return false
  }
  for (let index = 0; index < attributeName.length; index += 1) {
    if (bytes[startIndex + index] !== attributeName.charCodeAt(index)) {
      return false
    }
  }
  return true
}

function decodeCellAddress(value: string): { readonly row: number; readonly column: number } | null {
  const match = /^([A-Z]{1,3})([1-9][0-9]*)$/iu.exec(value)
  if (!match) {
    return null
  }
  let column = 0
  for (const letter of match[1]?.toUpperCase() ?? '') {
    column = column * 26 + letter.charCodeAt(0) - 64
  }
  const row = Number(match[2])
  return Number.isSafeInteger(row) && row > 0 && column > 0 ? { row: row - 1, column: column - 1 } : null
}

function encodeCellAddress(row: number, column: number): string {
  let value = column + 1
  let columnName = ''
  while (value > 0) {
    value -= 1
    columnName = String.fromCharCode(65 + (value % 26)) + columnName
    value = Math.floor(value / 26)
  }
  return `${columnName}${String(row + 1)}`
}

function numberValue(value: string | null): number | null {
  if (value === null || value.trim().length === 0) {
    return null
  }
  const number = Number(value)
  return Number.isFinite(number) ? number : null
}

function booleanValue(value: string | null): boolean | undefined {
  if (value === '1' || value === 'true') {
    return true
  }
  if (value === '0' || value === 'false') {
    return false
  }
  return undefined
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

function isNamespaceDeclaration(attributeName: string): boolean {
  return attributeName === 'xmlns' || attributeName.startsWith('xmlns:')
}

function skipAsciiWhitespace(bytes: Uint8Array, startIndex: number, endIndex: number): number {
  let index = startIndex
  while (index < endIndex && isAsciiWhitespace(bytes[index] ?? 0)) {
    index += 1
  }
  return index
}

function decodeBytes(bytes: Uint8Array, startIndex: number, endIndex: number): string {
  return strFromU8(bytes.subarray(startIndex, endIndex))
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
