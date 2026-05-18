import { strFromU8 } from 'fflate'

import type { LiteralInput, WorkbookRichTextCellSnapshot } from '@bilig/protocol'
import { toLiteralInput } from './workbook-import-helpers.js'
import { decodeExcelEscapedText } from './xlsx-escaped-text.js'
import {
  LargeSimpleFormulaRecords,
  parseLargeSimpleSharedFormulaIndex,
  readLargeSimpleFormulaTypeCode,
} from './xlsx-large-simple-formula-records.js'
import { ImportedWorkbookArena, ImportedWorksheetStyleIndexArena, type ImportedWorksheetCellScan } from './xlsx-large-simple-arena.js'
import type { LargeSimpleSharedStringEntry } from './xlsx-large-simple-shared-strings.js'

const lessThan = 60
const slash = 47
const greaterThan = 62
const doubleQuote = 34
const singleQuote = 39
const unsupportedWorksheetTagNames = new Set(['dataValidations', 'legacyDrawing', 'oleObjects', 'picture', 'sheetProtection'])
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
const rowMetadataAttributePattern = /\b(?:ht|hidden|customHeight|s|customFormat|outlineLevel|collapsed|thickTop|thickBottom)=/u
const richTextRunPattern = /<(?:[A-Za-z_][\w.-]*:)?r\b/u

export interface LargeSimpleWorksheetStreamScan {
  readonly cellScan: ImportedWorksheetCellScan
  readonly metadataXml: string | undefined
  readonly sharedStringIndexes: ReadonlySet<number>
}

export function parseLargeSimpleWorksheetCellsFromChunks(
  readChunks: (onChunk: (chunk: Uint8Array) => void) => boolean,
  sheetIndex: number,
  options: {
    readonly hasSharedStrings: boolean
    readonly retainCells?: boolean
    readonly sharedStrings?: readonly LargeSimpleSharedStringEntry[]
    readonly collectSharedStringIndexes?: boolean
    readonly allowInlineStringsWithoutRetention?: boolean
    readonly deferSharedStrings?: boolean
    readonly retainMetadataXml?: boolean
  },
): LargeSimpleWorksheetStreamScan | null {
  const scanner = new LargeSimpleWorksheetChunkScanner(sheetIndex, {
    hasSharedStrings: options.hasSharedStrings,
    retainCells: options.retainCells !== false,
    sharedStrings: options.sharedStrings ?? [],
    collectSharedStringIndexes: options.collectSharedStringIndexes === true,
    allowInlineStringsWithoutRetention: options.allowInlineStringsWithoutRetention === true,
    deferSharedStrings: options.deferSharedStrings === true,
    retainMetadataXml: options.retainMetadataXml !== false,
  })
  if (!readChunks((chunk) => scanner.push(chunk))) {
    return null
  }
  return scanner.finish()
}

class LargeSimpleWorksheetChunkScanner {
  private buffer = new Uint8Array()
  private index = 0
  private failed = false
  private readonly arena = new ImportedWorkbookArena()
  private readonly formulas = new LargeSimpleFormulaRecords()
  private readonly richTextCells: WorkbookRichTextCellSnapshot[] = []
  private readonly styleIndexes = new ImportedWorksheetStyleIndexArena()
  private rowCount = 0
  private columnCount = 0
  private cellCount = 0
  private valueCellCount = 0
  private formulaCellCount = 0
  private blankStyleCellCount = 0
  private mergeCount = 0
  private conditionalFormatCount = 0
  private tableCount = 0
  private minRow = Number.POSITIVE_INFINITY
  private minColumn = Number.POSITIVE_INFINITY
  private maxRow = -1
  private maxColumn = -1
  private readonly metadataSnippets: string[] = []
  private readonly hasSharedStrings: boolean
  private readonly retainCells: boolean
  private readonly sharedStrings: readonly LargeSimpleSharedStringEntry[]
  private readonly collectSharedStringIndexes: boolean
  private readonly allowInlineStringsWithoutRetention: boolean
  private readonly deferSharedStrings: boolean
  private readonly retainMetadataXml: boolean
  private readonly sharedStringIndexes = new Set<number>()

  constructor(
    private readonly sheetIndex: number,
    options: {
      readonly hasSharedStrings: boolean
      readonly retainCells: boolean
      readonly sharedStrings: readonly LargeSimpleSharedStringEntry[]
      readonly collectSharedStringIndexes: boolean
      readonly allowInlineStringsWithoutRetention: boolean
      readonly deferSharedStrings: boolean
      readonly retainMetadataXml: boolean
    },
  ) {
    this.hasSharedStrings = options.hasSharedStrings
    this.retainCells = options.retainCells
    this.sharedStrings = options.sharedStrings
    this.collectSharedStringIndexes = options.collectSharedStringIndexes
    this.allowInlineStringsWithoutRetention = options.allowInlineStringsWithoutRetention
    this.deferSharedStrings = options.deferSharedStrings
    this.retainMetadataXml = options.retainMetadataXml
  }

  push(chunk: Uint8Array): void {
    if (this.failed || chunk.byteLength === 0) {
      return
    }
    this.append(chunk)
    this.process(false)
    this.compact()
  }

  finish(): LargeSimpleWorksheetStreamScan | null {
    if (this.failed) {
      return null
    }
    this.process(true)
    this.compact()
    if (this.failed || (this.formulas.count > 0 && !this.formulas.resolveIntoArena(this.arena))) {
      return null
    }
    return {
      cellScan: {
        arena: this.arena,
        sheetIndex: this.sheetIndex,
        richTextCells: this.richTextCells,
        styleIndexes: this.styleIndexes,
        blankStyleCellCount: this.blankStyleCellCount,
        cellCount: this.cellCount,
        valueCellCount: this.valueCellCount,
        formulaCellCount: this.formulaCellCount,
        mergeCount: this.mergeCount,
        conditionalFormatCount: this.conditionalFormatCount,
        tableCount: this.tableCount,
        rowCount: this.rowCount,
        columnCount: this.columnCount,
        usedRange:
          this.cellCount > 0
            ? {
                startRow: this.minRow,
                startColumn: this.minColumn,
                endRow: this.maxRow,
                endColumn: this.maxColumn,
              }
            : null,
      },
      metadataXml: this.metadataSnippets.length > 0 ? `<worksheet>${this.metadataSnippets.join('')}</worksheet>` : undefined,
      sharedStringIndexes: this.sharedStringIndexes,
    }
  }

  private append(chunk: Uint8Array): void {
    if (this.index === this.buffer.byteLength) {
      this.buffer = new Uint8Array(chunk)
      this.index = 0
      return
    }
    const retained = this.buffer.subarray(this.index)
    const next = new Uint8Array(retained.byteLength + chunk.byteLength)
    next.set(retained)
    next.set(chunk, retained.byteLength)
    this.buffer = next
    this.index = 0
  }

  private compact(): void {
    if (this.index === 0) {
      return
    }
    if (this.index >= this.buffer.byteLength) {
      this.buffer = new Uint8Array()
      this.index = 0
      return
    }
    this.buffer = new Uint8Array(this.buffer.subarray(this.index))
    this.index = 0
  }

  private process(final: boolean): void {
    while (!this.failed && this.index < this.buffer.byteLength) {
      if (this.buffer[this.index] !== lessThan) {
        this.index += 1
        continue
      }
      const tag = readXmlTagName(this.buffer, this.index + 1)
      if (!tag) {
        this.index += 1
        continue
      }
      const tagEnd = findTagEnd(this.buffer, tag.endIndex)
      if (tagEnd === null) {
        if (final) {
          this.failed = true
        }
        return
      }
      if (unsupportedWorksheetTagNames.has(tag.localName) || this.hasUnsupportedCellMetadata(tag.localName, tag.endIndex, tagEnd)) {
        this.failed = true
        return
      }
      if (tag.localName === 'dimension') {
        this.readDimension(tag.endIndex, tagEnd)
        this.index = tagEnd + 1
        continue
      }
      if (tag.localName === 'row') {
        if (this.retainMetadataXml) {
          this.collectRowMetadata(tagEnd)
        }
        this.index = tagEnd + 1
        continue
      }
      if (tag.localName === 'c') {
        if (!this.readCell(tag.endIndex, tagEnd, final)) {
          return
        }
        continue
      }
      if (metadataWorksheetTagNames.has(tag.localName)) {
        if (!this.collectMetadataElement(tag.localName, tagEnd, final)) {
          return
        }
        continue
      }
      this.index = tagEnd + 1
    }
  }

  private hasUnsupportedCellMetadata(localName: string, nameEnd: number, tagEnd: number): boolean {
    return (
      localName === 'c' &&
      (readXmlAttributeRangeFromTag(this.buffer, nameEnd, tagEnd, 'cm') ||
        readXmlAttributeRangeFromTag(this.buffer, nameEnd, tagEnd, 'vm')) !== null
    )
  }

  private readDimension(nameEnd: number, tagEnd: number): void {
    const ref = readXmlAttributeFromTag(this.buffer, nameEnd, tagEnd, 'ref')
    if (!ref) {
      return
    }
    const [startRef, endRef = startRef] = ref.split(':')
    const start = decodeCellAddress(startRef ?? '')
    const end = decodeCellAddress(endRef ?? '')
    if (!start || !end) {
      return
    }
    this.rowCount = Math.max(this.rowCount, start.row + 1, end.row + 1)
    this.columnCount = Math.max(this.columnCount, start.column + 1, end.column + 1)
  }

  private collectRowMetadata(tagEnd: number): void {
    const openingTag = decodeBytes(this.buffer, this.index, tagEnd + 1)
    if (rowMetadataAttributePattern.test(openingTag)) {
      this.metadataSnippets.push(openingTag.endsWith('/>') ? openingTag : openingTag.replace(/>$/u, '/>'))
    }
  }

  private readCell(nameEnd: number, tagEnd: number, final: boolean): boolean {
    const selfClosing = isSelfClosingTag(this.buffer, tagEnd)
    const contentStart = tagEnd + 1
    const closing = selfClosing ? { start: contentStart, end: contentStart } : findClosingTag(this.buffer, contentStart, 'c')
    if (!closing) {
      if (final) {
        this.failed = true
      }
      return false
    }
    const decodedAddress = readCellAddressAttributeFromTag(this.buffer, nameEnd, tagEnd)
    if (!decodedAddress) {
      this.failed = true
      return false
    }
    const cellType = readXmlAttributeFromTag(this.buffer, nameEnd, tagEnd, 't')
    if (
      (!this.hasSharedStrings && cellType === 's') ||
      (!this.retainCells && cellType === 'inlineStr' && !this.allowInlineStringsWithoutRetention)
    ) {
      this.failed = true
      return false
    }
    const styleIndex = readCellStyleIndexFromTag(this.buffer, nameEnd, tagEnd)
    const shouldReadSharedStringIndex = cellType === 's' && (this.retainCells || this.collectSharedStringIndexes || this.deferSharedStrings)
    const rawValue = this.retainCells || shouldReadSharedStringIndex ? readElementText(this.buffer, contentStart, closing.start, 'v') : null
    const sharedStringIndex = shouldReadSharedStringIndex ? readSharedStringIndex(rawValue) : null
    if (shouldReadSharedStringIndex && rawValue !== null) {
      if (sharedStringIndex === null) {
        this.failed = true
        return false
      }
      if (this.collectSharedStringIndexes || this.deferSharedStrings) {
        this.sharedStringIndexes.add(sharedStringIndex)
      }
    }
    const deferSharedStringValue = this.retainCells && this.deferSharedStrings && cellType === 's' && sharedStringIndex !== null
    const value =
      this.retainCells && !deferSharedStringValue
        ? readCellValue(this.buffer, contentStart, closing.start, cellType, rawValue, this.sharedStrings)
        : hasElement(this.buffer, contentStart, closing.start, 'v') || hasElement(this.buffer, contentStart, closing.start, 'is')
          ? null
          : undefined
    const formula = this.retainCells
      ? readFormulaSpec(this.buffer, contentStart, closing.start)
      : hasElement(this.buffer, contentStart, closing.start, 'f')
        ? { typeCode: readLargeSimpleFormulaTypeCode(null), sharedIndex: null, rawFormula: '' }
        : undefined
    if (formula === null) {
      this.failed = true
      return false
    }
    const hasValue = deferSharedStringValue || value !== undefined
    const hasFormula = formula !== undefined
    if (hasValue || hasFormula) {
      this.cellCount += 1
      this.rowCount = Math.max(this.rowCount, decodedAddress.row + 1)
      this.columnCount = Math.max(this.columnCount, decodedAddress.column + 1)
      this.minRow = Math.min(this.minRow, decodedAddress.row)
      this.minColumn = Math.min(this.minColumn, decodedAddress.column)
      this.maxRow = Math.max(this.maxRow, decodedAddress.row)
      this.maxColumn = Math.max(this.maxColumn, decodedAddress.column)
      if (hasValue) {
        this.valueCellCount += 1
      }
      if (hasFormula) {
        this.formulaCellCount += 1
      }
      const cellIndex = this.retainCells
        ? deferSharedStringValue
          ? this.arena.addSharedStringCell({
              sheetIndex: this.sheetIndex,
              row: decodedAddress.row,
              column: decodedAddress.column,
              sharedStringIndex,
            })
          : this.arena.addCell({
              sheetIndex: this.sheetIndex,
              row: decodedAddress.row,
              column: decodedAddress.column,
              value,
            })
        : -1
      if (formula && this.retainCells) {
        this.formulas.add(cellIndex, decodedAddress.row, decodedAddress.column, formula.typeCode, formula.sharedIndex, formula.rawFormula)
      }
      if (this.retainCells && styleIndex !== null) {
        this.styleIndexes.add(decodedAddress.row, decodedAddress.column, styleIndex)
      }
      const richTextCell =
        this.retainCells && !deferSharedStringValue
          ? readRichTextCellArtifact(this.buffer, contentStart, closing.start, decodedAddress, cellType, rawValue, this.sharedStrings)
          : undefined
      if (richTextCell) {
        this.richTextCells.push(richTextCell)
      }
    } else if (styleIndex !== null) {
      this.blankStyleCellCount += 1
    }
    this.index = selfClosing ? tagEnd + 1 : closing.end
    return true
  }

  private collectMetadataElement(localName: string, tagEnd: number, final: boolean): boolean {
    if (isSelfClosingTag(this.buffer, tagEnd)) {
      this.countMetadataElement(localName, tagEnd + 1, tagEnd + 1)
      if (this.retainMetadataXml) {
        this.metadataSnippets.push(decodeBytes(this.buffer, this.index, tagEnd + 1))
      }
      this.index = tagEnd + 1
      return true
    }
    const closing = findClosingTag(this.buffer, tagEnd + 1, localName)
    if (!closing) {
      if (final) {
        this.failed = true
      }
      return false
    }
    this.countMetadataElement(localName, tagEnd + 1, closing.start)
    if (this.retainMetadataXml) {
      this.metadataSnippets.push(decodeBytes(this.buffer, this.index, closing.end))
    }
    this.index = closing.end
    return true
  }

  private countMetadataElement(localName: string, contentStart: number, contentEnd: number): void {
    if (localName === 'conditionalFormatting') {
      this.conditionalFormatCount += Math.max(1, countOpeningTags(this.buffer, contentStart, contentEnd, 'cfRule'))
      return
    }
    if (localName === 'mergeCells') {
      this.mergeCount += countOpeningTags(this.buffer, contentStart, contentEnd, 'mergeCell')
    } else if (localName === 'tableParts') {
      this.tableCount += countOpeningTags(this.buffer, contentStart, contentEnd, 'tablePart')
    }
  }
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

function readSharedStringIndex(rawValue: string | null): number | null {
  if (rawValue === null) {
    return null
  }
  const index = Number(decodeXmlText(rawValue.trim()))
  return Number.isSafeInteger(index) && index >= 0 ? index : null
}

function readFormulaSpec(
  bytes: Uint8Array,
  contentStart: number,
  contentEnd: number,
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
  if (type === 'array' || type === 'dataTable') {
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

function hasElement(bytes: Uint8Array, startIndex: number, endIndex: number, elementName: string): boolean {
  let index = startIndex
  while (index < endIndex) {
    if (bytes[index] !== lessThan) {
      index += 1
      continue
    }
    const tag = readXmlTagName(bytes, index + 1)
    if (tag?.localName === elementName) {
      return true
    }
    index += 1
  }
  return false
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
