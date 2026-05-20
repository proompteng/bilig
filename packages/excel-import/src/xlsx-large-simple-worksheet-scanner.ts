import { strFromU8 } from 'fflate'

import type { LiteralInput, WorkbookRichTextCellSnapshot } from '@bilig/protocol'
import { toLiteralInput } from './workbook-import-helpers.js'
import { decodeExcelEscapedText } from './xlsx-escaped-text.js'
import {
  LargeSimpleFormulaRecords,
  parseLargeSimpleSharedFormulaIndex,
  readLargeSimpleFormulaTypeCode,
} from './xlsx-large-simple-formula-records.js'
import {
  ImportedWorkbookArena,
  ImportedWorksheetStyleIndexArena,
  type ImportedWorkbookArenaDedupeMode,
  type ImportedWorksheetCellScan,
} from './xlsx-large-simple-arena.js'
import type { LargeSimpleSharedStringEntry } from './xlsx-large-simple-shared-strings.js'
import type { ImportedWorkbookStringPool } from './xlsx-large-simple-string-pool.js'

const lessThan = 60
const slash = 47
const greaterThan = 62
const doubleQuote = 34
const singleQuote = 39
const unsupportedWorksheetTagNames = new Set(['dataValidations', 'oleObjects', 'picture'])
const metadataWorksheetTagNames = new Set([
  'autoFilter',
  'colBreaks',
  'cols',
  'conditionalFormatting',
  'drawing',
  'headerFooter',
  'hyperlinks',
  'mergeCells',
  'pageMargins',
  'pageSetup',
  'printOptions',
  'rowBreaks',
  'sheetFormatPr',
  'tableParts',
])
const richTextRunPattern = /<(?:[A-Za-z_][\w.-]*:)?r\b/u

export function parseLargeSimpleWorksheetCells(
  bytes: Uint8Array,
  sharedStrings: readonly LargeSimpleSharedStringEntry[],
  sheetIndex: number,
  options: {
    readonly retainCells?: boolean
    readonly stringPool?: ImportedWorkbookStringPool
    readonly deduplicateStrings?: ImportedWorkbookArenaDedupeMode
    readonly deduplicateFormulas?: ImportedWorkbookArenaDedupeMode
    readonly allowUnsupportedFormulaText?: boolean
    readonly preserveBlankStyleCells?: boolean
  } = {},
): ImportedWorksheetCellScan | null {
  const arena = new ImportedWorkbookArena(options.stringPool, {
    ...(options.deduplicateStrings === undefined ? {} : { deduplicateStrings: options.deduplicateStrings }),
    ...(options.deduplicateFormulas === undefined ? {} : { deduplicateFormulas: options.deduplicateFormulas }),
  })
  const richTextCells: WorkbookRichTextCellSnapshot[] = []
  const styleIndexes = new ImportedWorksheetStyleIndexArena()
  const allowUnsupportedFormulaText = options.allowUnsupportedFormulaText === true
  const formulas = new LargeSimpleFormulaRecords(allowUnsupportedFormulaText)
  const retainCells = options.retainCells !== false
  const preserveBlankStyleCells = options.preserveBlankStyleCells !== false
  const dimension = readWorksheetDimensionFromBytes(bytes)
  let rowCount = dimension?.rowCount ?? 0
  let columnCount = dimension?.columnCount ?? 0
  let cellCount = 0
  let valueCellCount = 0
  let formulaCellCount = 0
  let blankStyleCellCount = 0
  let minRow = Number.POSITIVE_INFINITY
  let minColumn = Number.POSITIVE_INFINITY
  let maxRow = -1
  let maxColumn = -1
  let index = 0

  while (index < bytes.byteLength) {
    const tag = findNextOpeningTag(bytes, index, 'c')
    if (!tag) {
      break
    }
    const tagEnd = findTagEnd(bytes, tag.nameEnd)
    if (tagEnd === null) {
      return null
    }
    const decodedAddress = readCellAddressAttributeFromTag(bytes, tag.nameEnd, tagEnd)
    if (!decodedAddress) {
      return null
    }
    const selfClosing = isSelfClosingTag(bytes, tagEnd)
    const contentStart = tagEnd + 1
    const closing = selfClosing ? { start: contentStart, end: contentStart } : findClosingTag(bytes, contentStart, 'c')
    if (!closing) {
      return null
    }
    rowCount = Math.max(rowCount, decodedAddress.row + 1)
    columnCount = Math.max(columnCount, decodedAddress.column + 1)
    const shouldDecodeValue = retainCells
    const cellType = readXmlAttributeFromTag(bytes, tag.nameEnd, tagEnd, 't')
    if (!retainCells && cellType === 'inlineStr') {
      return null
    }
    const styleIndex = readCellStyleIndexFromTag(bytes, tag.nameEnd, tagEnd)
    const rawValue = shouldDecodeValue ? readElementText(bytes, contentStart, closing.start, 'v') : null
    const value = shouldDecodeValue
      ? readCellValue(bytes, contentStart, closing.start, cellType, rawValue, sharedStrings)
      : hasCellValue(bytes, contentStart, closing.start)
        ? null
        : undefined
    const formula = retainCells
      ? readFormulaSpec(bytes, contentStart, closing.start, allowUnsupportedFormulaText)
      : readCompactFormulaSpec(bytes, contentStart, closing.start)
    if (formula === null) {
      return null
    }
    const hasValue = value !== undefined
    if (hasValue || formula !== undefined) {
      cellCount += 1
      minRow = Math.min(minRow, decodedAddress.row)
      minColumn = Math.min(minColumn, decodedAddress.column)
      maxRow = Math.max(maxRow, decodedAddress.row)
      maxColumn = Math.max(maxColumn, decodedAddress.column)
      if (hasValue) {
        valueCellCount += 1
      }
      if (formula !== undefined) {
        formulaCellCount += 1
      }
      const shouldRetainCell = retainCells
      const cellIndex = shouldRetainCell
        ? arena.addCell({
            sheetIndex,
            row: decodedAddress.row,
            column: decodedAddress.column,
            value: shouldDecodeValue ? value : undefined,
          })
        : -1
      if (formula && retainCells) {
        formulas.add(cellIndex, decodedAddress.row, decodedAddress.column, formula.typeCode, formula.sharedIndex, formula.rawFormula)
      }
      if (retainCells && styleIndex !== null) {
        styleIndexes.add(decodedAddress.row, decodedAddress.column, styleIndex)
      }
      const richTextCell = readRichTextCellArtifact(bytes, contentStart, closing.start, decodedAddress, cellType, rawValue, sharedStrings)
      if (richTextCell) {
        if (!retainCells) {
          return null
        }
        richTextCells.push(richTextCell)
      }
    } else if (styleIndex !== null) {
      blankStyleCellCount += 1
      if (retainCells && preserveBlankStyleCells) {
        styleIndexes.add(decodedAddress.row, decodedAddress.column, styleIndex)
      }
    }
    index = selfClosing ? tagEnd + 1 : closing.end
  }

  if (formulas.count > 0 && !formulas.resolveIntoArena(arena)) {
    return null
  }
  return {
    arena,
    sheetIndex,
    richTextCells,
    styleIndexes,
    blankStyleCellCount,
    cellCount,
    valueCellCount,
    formulaCellCount,
    rowCount,
    columnCount,
    usedRange:
      cellCount > 0
        ? {
            startRow: minRow,
            startColumn: minColumn,
            endRow: maxRow,
            endColumn: maxColumn,
          }
        : null,
  }
}

export function hasUnsupportedLargeSimpleWorksheetTags(bytes: Uint8Array): boolean {
  return hasAnyLocalNameTag(bytes, unsupportedWorksheetTagNames) || hasCellMetadataAttributes(bytes)
}

export function needsLargeSimpleWorksheetMetadataXml(bytes: Uint8Array): boolean {
  return hasAnyLocalNameTag(bytes, metadataWorksheetTagNames) || hasRowMetadataAttributes(bytes)
}

export function readLargeSimpleWorksheetMetadataXml(bytes: Uint8Array): string | undefined {
  const snippets: string[] = []
  let index = 0
  while (index < bytes.byteLength) {
    if (bytes[index] !== lessThan) {
      index += 1
      continue
    }
    const tag = readXmlTagName(bytes, index + 1)
    if (!tag) {
      index += 1
      continue
    }
    const tagEnd = findTagEnd(bytes, tag.endIndex)
    if (tagEnd === null) {
      return undefined
    }
    if (tag.localName === 'row') {
      const openingTag = decodeBytes(bytes, index, tagEnd + 1)
      if (/\b(?:ht|hidden|customHeight|s|customFormat|outlineLevel|collapsed|thickTop|thickBottom)=/u.test(openingTag)) {
        snippets.push(openingTag.endsWith('/>') ? openingTag : openingTag.replace(/>$/u, '/>'))
      }
      index = tagEnd + 1
      continue
    }
    if (!metadataWorksheetTagNames.has(tag.localName)) {
      index += 1
      continue
    }
    if (isSelfClosingTag(bytes, tagEnd)) {
      snippets.push(decodeBytes(bytes, index, tagEnd + 1))
      index = tagEnd + 1
      continue
    }
    const closing = findClosingTag(bytes, tagEnd + 1, tag.localName)
    if (!closing) {
      return undefined
    }
    snippets.push(decodeBytes(bytes, index, closing.end))
    index = closing.end
  }
  return snippets.length > 0 ? `<worksheet>${snippets.join('')}</worksheet>` : undefined
}

function readCellValue(
  bytes: Uint8Array,
  contentStart: number,
  contentEnd: number,
  type: string | null,
  rawValue: string | null,
  sharedStrings: readonly LargeSimpleSharedStringEntry[],
): LiteralInput | undefined {
  switch (type) {
    case null:
      break
    case 's': {
      const index = rawValue === null ? -1 : Number(decodeXmlText(rawValue.trim()))
      return Number.isSafeInteger(index) && index >= 0 ? sharedStrings[index]?.text : undefined
    }
    case 'inlineStr': {
      const inlineStringXml = readElementXml(bytes, contentStart, contentEnd, 'is')
      return inlineStringXml ? stringItemText(inlineStringXml) : undefined
    }
    case 'str':
      return rawValue === null ? undefined : normalizeWorksheetText(decodeXmlText(rawValue))
    case 'b': {
      const normalized = rawValue?.trim()
      return normalized === '1' || normalized === 'true' ? true : normalized === '0' || normalized === 'false' ? false : undefined
    }
    case 'e':
      return rawValue === null ? undefined : decodeXmlText(rawValue)
    case 'd':
      return undefined
    default:
      return undefined
  }
  if (rawValue === null) {
    return undefined
  }
  const number = Number(decodeXmlText(rawValue).trim())
  return Number.isFinite(number) ? number : undefined
}

function readFormulaSpec(
  bytes: Uint8Array,
  contentStart: number,
  contentEnd: number,
  allowUnsupportedFormulaText: boolean,
): { readonly typeCode: number; readonly sharedIndex: number | null; readonly rawFormula: string } | null | undefined {
  const tag = findNextOpeningTag(bytes, contentStart, 'f', contentEnd)
  if (!tag) {
    return undefined
  }
  const tagEnd = findTagEnd(bytes, tag.nameEnd, contentEnd)
  if (tagEnd === null) {
    return null
  }
  const type = readXmlAttributeFromTag(bytes, tag.nameEnd, tagEnd, 't')
  if (!allowUnsupportedFormulaText && (type === 'array' || type === 'dataTable')) {
    return null
  }
  const selfClosing = isSelfClosingTag(bytes, tagEnd)
  const closing = selfClosing ? { start: tagEnd + 1, end: tagEnd + 1 } : findClosingTag(bytes, tagEnd + 1, 'f', contentEnd)
  if (!closing) {
    return null
  }
  return {
    typeCode: readLargeSimpleFormulaTypeCode(type),
    sharedIndex: parseLargeSimpleSharedFormulaIndex(readXmlAttributeFromTag(bytes, tag.nameEnd, tagEnd, 'si')),
    rawFormula: selfClosing ? '' : decodeBytes(bytes, tagEnd + 1, closing.start).trim(),
  }
}

function readCompactFormulaSpec(
  bytes: Uint8Array,
  contentStart: number,
  contentEnd: number,
): { readonly typeCode: number; readonly sharedIndex: number | null; readonly rawFormula: string } | null | undefined {
  return hasElement(bytes, contentStart, contentEnd, 'f')
    ? { typeCode: readLargeSimpleFormulaTypeCode(null), sharedIndex: null, rawFormula: '' }
    : undefined
}

function readRichTextCellArtifact(
  bytes: Uint8Array,
  contentStart: number,
  contentEnd: number,
  address: { readonly row: number; readonly column: number },
  type: string | null,
  rawValue: string | null,
  sharedStrings: readonly LargeSimpleSharedStringEntry[],
): WorkbookRichTextCellSnapshot | undefined {
  if (type === 's') {
    const index = rawValue === null ? -1 : Number(decodeXmlText(rawValue.trim()))
    const entry = Number.isSafeInteger(index) && index >= 0 ? sharedStrings[index] : undefined
    return entry?.rich
      ? {
          address: encodeCellAddress(address.row, address.column),
          text: entry.text,
          storage: 'sharedString',
          xml: entry.xml ?? '',
        }
      : undefined
  }
  if (type !== 'inlineStr') {
    return undefined
  }
  const inlineStringXml = readElementXml(bytes, contentStart, contentEnd, 'is')
  if (!inlineStringXml || !richTextRunPattern.test(inlineStringXml)) {
    return undefined
  }
  return {
    address: encodeCellAddress(address.row, address.column),
    text: stringItemText(inlineStringXml),
    storage: 'inlineString',
    xml: inlineStringXml,
  }
}

function readWorksheetDimensionFromBytes(bytes: Uint8Array): { readonly rowCount: number; readonly columnCount: number } | null {
  const tag = findNextOpeningTag(bytes, 0, 'dimension')
  if (!tag) {
    return null
  }
  const tagEnd = findTagEnd(bytes, tag.nameEnd)
  if (tagEnd === null) {
    return null
  }
  const ref = readXmlAttributeFromTag(bytes, tag.nameEnd, tagEnd, 'ref')
  if (!ref) {
    return null
  }
  const [startRef, endRef = startRef] = ref.split(':')
  const start = decodeCellAddress(startRef ?? '')
  const end = decodeCellAddress(endRef ?? '')
  if (!start || !end) {
    return null
  }
  return {
    rowCount: Math.max(start.row, end.row) + 1,
    columnCount: Math.max(start.column, end.column) + 1,
  }
}

function hasAnyLocalNameTag(bytes: Uint8Array, localNames: ReadonlySet<string>): boolean {
  let index = 0
  while (index < bytes.byteLength) {
    if (bytes[index] !== lessThan) {
      index += 1
      continue
    }
    const tag = readXmlTagName(bytes, index + 1)
    if (tag && localNames.has(tag.localName)) {
      return true
    }
    index += 1
  }
  return false
}

function hasRowMetadataAttributes(bytes: Uint8Array): boolean {
  let index = 0
  while (index < bytes.byteLength) {
    const tag = findNextOpeningTag(bytes, index, 'row')
    if (!tag) {
      return false
    }
    const tagEnd = findTagEnd(bytes, tag.nameEnd)
    if (tagEnd === null) {
      return false
    }
    const openingTag = decodeBytes(bytes, tag.start, tagEnd + 1)
    if (/\b(?:ht|hidden|customHeight|s|customFormat|outlineLevel|collapsed|thickTop|thickBottom)=/u.test(openingTag)) {
      return true
    }
    index = tagEnd + 1
  }
  return false
}

function hasCellMetadataAttributes(bytes: Uint8Array): boolean {
  let index = 0
  while (index < bytes.byteLength) {
    const tag = findNextOpeningTag(bytes, index, 'c')
    if (!tag) {
      return false
    }
    const tagEnd = findTagEnd(bytes, tag.nameEnd)
    if (tagEnd === null) {
      return true
    }
    if (readXmlAttributeRangeFromTag(bytes, tag.nameEnd, tagEnd, 'cm') || readXmlAttributeRangeFromTag(bytes, tag.nameEnd, tagEnd, 'vm')) {
      return true
    }
    index = tagEnd + 1
  }
  return false
}

function findNextOpeningTag(
  bytes: Uint8Array,
  startIndex: number,
  localName: string,
  endIndex: number = bytes.byteLength,
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
  endIndex: number = bytes.byteLength,
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

function findTagEnd(bytes: Uint8Array, startIndex: number, endIndex: number = bytes.byteLength): number | null {
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

function readElementText(bytes: Uint8Array, startIndex: number, endIndex: number, elementName: string): string | null {
  const tag = findNextOpeningTag(bytes, startIndex, elementName, endIndex)
  if (!tag) {
    return null
  }
  const tagEnd = findTagEnd(bytes, tag.nameEnd, endIndex)
  if (tagEnd === null || isSelfClosingTag(bytes, tagEnd)) {
    return null
  }
  const closing = findClosingTag(bytes, tagEnd + 1, elementName, endIndex)
  return closing ? decodeBytes(bytes, tagEnd + 1, closing.start) : null
}

function readElementXml(bytes: Uint8Array, startIndex: number, endIndex: number, elementName: string): string | null {
  const tag = findNextOpeningTag(bytes, startIndex, elementName, endIndex)
  if (!tag) {
    return null
  }
  const tagEnd = findTagEnd(bytes, tag.nameEnd, endIndex)
  if (tagEnd === null) {
    return null
  }
  if (isSelfClosingTag(bytes, tagEnd)) {
    return decodeBytes(bytes, tag.start, tagEnd + 1)
  }
  const closing = findClosingTag(bytes, tagEnd + 1, elementName, endIndex)
  return closing ? decodeBytes(bytes, tag.start, closing.end) : null
}

function hasCellValue(bytes: Uint8Array, startIndex: number, endIndex: number): boolean {
  return hasElement(bytes, startIndex, endIndex, 'v') || hasElement(bytes, startIndex, endIndex, 'is')
}

function hasElement(bytes: Uint8Array, startIndex: number, endIndex: number, elementName: string): boolean {
  return findNextOpeningTag(bytes, startIndex, elementName, endIndex) !== null
}

function readXmlTagName(bytes: Uint8Array, startIndex: number): { readonly localName: string; readonly endIndex: number } | null {
  const first = bytes[startIndex]
  if (first === undefined || first === 33 || first === slash || first === 63) {
    return null
  }
  let index = startIndex
  let localNameStart = startIndex
  while (index < bytes.byteLength && isXmlNameByte(bytes[index] ?? 0)) {
    if (bytes[index] === 58) {
      localNameStart = index + 1
    }
    index += 1
  }
  return index === localNameStart ? null : { localName: decodeAscii(bytes, localNameStart, index), endIndex: index }
}

function stringItemText(xml: string): string {
  return normalizeWorksheetText(
    [...xml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?t\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?t>/gu)]
      .map((match) => decodeXmlText(match[1] ?? ''))
      .join(''),
  )
}

function normalizeWorksheetText(value: string): string {
  const literal = toLiteralInput(decodeExcelEscapedText(value))
  return typeof literal === 'string' ? literal : value
}

function readXmlAttributeFromTag(bytes: Uint8Array, startIndex: number, tagEnd: number, attributeName: string): string | null {
  const range = readXmlAttributeRangeFromTag(bytes, startIndex, tagEnd, attributeName)
  return range ? decodeBytes(bytes, range.start, range.end) : null
}

function readCellAddressAttributeFromTag(
  bytes: Uint8Array,
  startIndex: number,
  tagEnd: number,
): { readonly row: number; readonly column: number } | null {
  const range = readXmlAttributeRangeFromTag(bytes, startIndex, tagEnd, 'r')
  return range ? decodeCellAddressBytes(bytes, range.start, range.end) : null
}

function readCellStyleIndexFromTag(bytes: Uint8Array, startIndex: number, tagEnd: number): number | null {
  const range = readXmlAttributeRangeFromTag(bytes, startIndex, tagEnd, 's')
  if (!range) {
    return null
  }
  let value = 0
  if (range.start === range.end) {
    return null
  }
  for (let index = range.start; index < range.end; index += 1) {
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

function readXmlAttributeRangeFromTag(
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

function skipAsciiWhitespace(bytes: Uint8Array, startIndex: number, endIndex: number): number {
  let index = startIndex
  while (index < endIndex && isAsciiWhitespace(bytes[index] ?? 0)) {
    index += 1
  }
  return index
}

function decodeXmlText(value: string): string {
  if (!value.includes('&')) {
    return value
  }
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

function decodeCellAddress(address: string): { readonly row: number; readonly column: number } | null {
  const match = /^([A-Z]{1,3})([1-9][0-9]*)$/iu.exec(address.replaceAll('$', ''))
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

function decodeCellAddressBytes(
  bytes: Uint8Array,
  startIndex: number,
  endIndex: number,
): { readonly row: number; readonly column: number } | null {
  let column = 0
  let row = 0
  let letterCount = 0
  let digitCount = 0
  for (let index = startIndex; index < endIndex; index += 1) {
    const byte = bytes[index] ?? 0
    if (byte === 36) {
      continue
    }
    const upper = byte >= 97 && byte <= 122 ? byte - 32 : byte
    if (upper >= 65 && upper <= 90 && digitCount === 0) {
      column = column * 26 + upper - 64
      letterCount += 1
      continue
    }
    if (byte >= 48 && byte <= 57 && letterCount > 0) {
      row = row * 10 + byte - 48
      digitCount += 1
      continue
    }
    return null
  }
  return letterCount > 0 && letterCount <= 3 && digitCount > 0 && row > 0 && column > 0 ? { row: row - 1, column: column - 1 } : null
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
    byte === 58 ||
    byte === 95
  )
}

function isValidXmlCodePoint(value: number): boolean {
  return Number.isInteger(value) && value >= 0 && value <= 0x10ffff
}
