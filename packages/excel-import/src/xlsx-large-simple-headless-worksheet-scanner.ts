import { readKnownXmlLocalName } from './xlsx-large-simple-xml-name.js'
import { countLargeSimpleDataValidationsFromBytes } from './xlsx-large-simple-data-validation-byte-scan.js'

const lessThan = 60
const slash = 47
const greaterThan = 62
const doubleQuote = 34
const singleQuote = 39
const packedAddressColumnFactor = 16_384
const cellContentHasValue = 1 << 0
const cellContentHasFormula = 1 << 1
const emptyBytes = new Uint8Array(0)
const unsupportedWorksheetTagNames = new Set(['picture', 'sheetProtection'])
const metadataWorksheetTagNames = new Set([
  'autoFilter',
  'colBreaks',
  'cols',
  'conditionalFormatting',
  'dataValidations',
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

export interface HeadlessLargeSimpleWorksheetScan {
  readonly sheetIndex: number
  readonly cellCount: number
  readonly valueCellCount: number
  readonly formulaCellCount: number
  readonly tableCount: number
  readonly mergeCount: number
  readonly conditionalFormatCount: number
  readonly dataValidationCount?: number
  readonly rowCount: number
  readonly columnCount: number
  readonly usedRange: {
    readonly startRow: number
    readonly startColumn: number
    readonly endRow: number
    readonly endColumn: number
  } | null
}

export interface HeadlessLargeSimpleWorksheetScanOptions {
  readonly hasSharedStrings: boolean
  readonly allowUnsupportedFeaturesForMetrics?: boolean
  readonly onRetainedBufferLength?: (length: number) => void
}

interface ActiveMetadataCount {
  readonly localName: string
  readonly childName: string | null
  readonly multiplier: number
  childCount: number
}

export function parseHeadlessLargeSimpleWorksheetFromChunks(
  readChunks: (onChunk: (chunk: Uint8Array) => void) => boolean,
  sheetIndex: number,
  options: HeadlessLargeSimpleWorksheetScanOptions,
): HeadlessLargeSimpleWorksheetScan | null {
  const scanner = new HeadlessLargeSimpleWorksheetChunkScanner(sheetIndex, options)
  if (!readChunks((chunk) => scanner.push(chunk))) {
    return null
  }
  return scanner.finish()
}

class HeadlessLargeSimpleWorksheetChunkScanner {
  private buffer: Uint8Array = new Uint8Array()
  private index = 0
  private failed = false
  private rowCount = 0
  private columnCount = 0
  private cellCount = 0
  private valueCellCount = 0
  private formulaCellCount = 0
  private tableCount = 0
  private mergeCount = 0
  private conditionalFormatCount = 0
  private dataValidationCount = 0
  private minRow = Number.POSITIVE_INFINITY
  private minColumn = Number.POSITIVE_INFINITY
  private maxRow = -1
  private maxColumn = -1
  private currentRow = -1
  private nextImplicitRow = 0
  private nextImplicitColumn = 0
  private cellContentFlags = 0
  private cellContentNextIndex = 0
  private cellPackedAddress = -1
  private cellHasSharedStringType = false
  private activeMetadata: ActiveMetadataCount | null = null

  constructor(
    private readonly sheetIndex: number,
    private readonly options: HeadlessLargeSimpleWorksheetScanOptions,
  ) {}

  push(chunk: Uint8Array): void {
    if (this.failed || chunk.byteLength === 0) {
      return
    }
    this.append(chunk)
    this.process(false)
    this.compact()
    this.reportRetainedBufferLength()
  }

  finish(): HeadlessLargeSimpleWorksheetScan | null {
    if (this.failed) {
      return null
    }
    this.process(true)
    if (this.activeMetadata !== null) {
      this.failed = true
    }
    this.compact()
    this.reportRetainedBufferLength()
    return this.failed
      ? null
      : {
          sheetIndex: this.sheetIndex,
          cellCount: this.cellCount,
          valueCellCount: this.valueCellCount,
          formulaCellCount: this.formulaCellCount,
          tableCount: this.tableCount,
          mergeCount: this.mergeCount,
          conditionalFormatCount: this.conditionalFormatCount,
          dataValidationCount: this.dataValidationCount,
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
        }
  }

  private append(chunk: Uint8Array): void {
    if (this.index === this.buffer.byteLength) {
      this.buffer = chunk
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
      this.buffer = emptyBytes
      this.index = 0
      return
    }
    this.buffer = new Uint8Array(this.buffer.subarray(this.index))
    this.index = 0
  }

  private process(final: boolean): void {
    while (!this.failed && this.index < this.buffer.byteLength) {
      if (this.activeMetadata !== null) {
        if (!this.processActiveMetadata(final)) {
          return
        }
        continue
      }
      if (this.buffer[this.index] !== lessThan) {
        this.index += 1
        continue
      }
      const tag = readXmlTagName(this.buffer, this.index + 1)
      if (!tag) {
        if (!final && this.index + 1 >= this.buffer.byteLength) {
          return
        }
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
      if (
        this.options.allowUnsupportedFeaturesForMetrics !== true &&
        (unsupportedWorksheetTagNames.has(tag.localName) || this.hasUnsupportedCellMetadata(tag.localName, tag.endIndex, tagEnd))
      ) {
        this.failed = true
        return
      }
      if (tag.localName === 'dimension') {
        this.readDimension(tag.endIndex, tagEnd)
        this.index = tagEnd + 1
        continue
      }
      if (tag.localName === 'row') {
        this.readRow(tag.endIndex, tagEnd)
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
        if (tag.localName === 'dataValidations') {
          if (!this.readDataValidations(tagEnd, final)) {
            return
          }
          continue
        }
        if (!this.countMetadataElement(tag.localName, tag.endIndex, tagEnd, final)) {
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
    const ref = readXmlAttributeRangeFromTag(this.buffer, nameEnd, tagEnd, 'ref')
    if (!ref) {
      return
    }
    const range = readDimensionAddressRange(this.buffer, ref.start, ref.end)
    if (!range) {
      return
    }
    this.rowCount = Math.max(this.rowCount, range.startRow + 1, range.endRow + 1)
    this.columnCount = Math.max(this.columnCount, range.startColumn + 1, range.endColumn + 1)
  }

  private readRow(nameEnd: number, tagEnd: number): void {
    const row = readPositiveIntegerAttributeFromTag(this.buffer, nameEnd, tagEnd, 'r')
    this.currentRow = row === null ? this.nextImplicitRow : row - 1
    this.nextImplicitRow = this.currentRow + 1
    this.nextImplicitColumn = 0
  }

  private readCell(nameEnd: number, tagEnd: number, final: boolean): boolean {
    const selfClosing = isSelfClosingTag(this.buffer, tagEnd)
    const contentStart = tagEnd + 1
    let contentFlags = 0
    let contentNextIndex = contentStart
    if (!selfClosing && !this.readCellContentSummary(contentStart)) {
      if (final) {
        this.failed = true
      }
      return false
    }
    if (!selfClosing) {
      contentFlags = this.cellContentFlags
      contentNextIndex = this.cellContentNextIndex
    }
    if (!this.readCellTagAttributes(nameEnd, tagEnd) || (!this.options.hasSharedStrings && this.cellHasSharedStringType)) {
      this.failed = true
      return false
    }
    const row = packedAddressRow(this.cellPackedAddress)
    const column = packedAddressColumn(this.cellPackedAddress)
    this.currentRow = row
    this.nextImplicitColumn = column + 1
    const hasValue = (contentFlags & cellContentHasValue) !== 0
    const hasFormula = (contentFlags & cellContentHasFormula) !== 0
    if (hasValue || hasFormula) {
      this.cellCount += 1
      this.rowCount = Math.max(this.rowCount, row + 1)
      this.columnCount = Math.max(this.columnCount, column + 1)
      this.minRow = Math.min(this.minRow, row)
      this.minColumn = Math.min(this.minColumn, column)
      this.maxRow = Math.max(this.maxRow, row)
      this.maxColumn = Math.max(this.maxColumn, column)
      if (hasValue) {
        this.valueCellCount += 1
      }
      if (hasFormula) {
        this.formulaCellCount += 1
      }
    }
    this.index = contentNextIndex
    return true
  }

  private readDataValidations(tagEnd: number, final: boolean): boolean {
    const selfClosing = isSelfClosingTag(this.buffer, tagEnd)
    const closing = selfClosing ? { end: tagEnd + 1 } : findClosingTag(this.buffer, tagEnd + 1, 'dataValidations')
    if (!closing) {
      if (final) {
        this.failed = true
      }
      return false
    }
    const count = countLargeSimpleDataValidationsFromBytes('Sheet1', this.buffer, this.index, closing.end)
    if (count === null) {
      if (this.options.allowUnsupportedFeaturesForMetrics === true) {
        this.index = closing.end
        return true
      }
      this.failed = true
      return true
    }
    this.dataValidationCount += count
    this.index = closing.end
    return true
  }

  private readCellContentSummary(startIndex: number): boolean {
    let flags = 0
    let index = startIndex
    while (index < this.buffer.byteLength) {
      if (this.buffer[index] !== lessThan) {
        index += 1
        continue
      }
      const next = this.buffer[index + 1]
      if (next === slash && this.buffer[index + 2] === 99 && !isXmlNameByte(this.buffer[index + 3] ?? 0)) {
        const tagEnd = findTagEnd(this.buffer, index + 3)
        if (tagEnd === null) {
          return false
        }
        this.cellContentFlags = flags
        this.cellContentNextIndex = tagEnd + 1
        return true
      }
      if (next === slash && this.buffer[index + 2] === 118 && !isXmlNameByte(this.buffer[index + 3] ?? 0)) {
        index += 3
        continue
      }
      if (next === slash && this.buffer[index + 2] === 102 && !isXmlNameByte(this.buffer[index + 3] ?? 0)) {
        index += 3
        continue
      }
      if (
        next === slash &&
        this.buffer[index + 2] === 105 &&
        this.buffer[index + 3] === 115 &&
        !isXmlNameByte(this.buffer[index + 4] ?? 0)
      ) {
        index += 4
        continue
      }
      if (next === 118 && !isXmlNameByte(this.buffer[index + 2] ?? 0)) {
        flags |= cellContentHasValue
        index += 2
        continue
      }
      if (next === 102 && !isXmlNameByte(this.buffer[index + 2] ?? 0)) {
        flags |= cellContentHasFormula
        index += 2
        continue
      }
      if (next === 105 && this.buffer[index + 2] === 115 && !isXmlNameByte(this.buffer[index + 3] ?? 0)) {
        flags |= cellContentHasValue
        index += 3
        continue
      }
      const closing = this.buffer[index + 1] === slash
      const tag = readXmlTagName(this.buffer, index + (closing ? 2 : 1))
      if (closing && tag?.localName === 'c') {
        const tagEnd = findTagEnd(this.buffer, tag.endIndex)
        if (tagEnd === null) {
          return false
        }
        this.cellContentFlags = flags
        this.cellContentNextIndex = tagEnd + 1
        return true
      }
      if (tag?.localName === 'v' || tag?.localName === 'is') {
        flags |= cellContentHasValue
      } else if (tag?.localName === 'f') {
        flags |= cellContentHasFormula
      }
      index = tag?.endIndex ?? index + 1
    }
    return false
  }

  private readCellTagAttributes(startIndex: number, tagEnd: number): boolean {
    let index = startIndex
    let packedAddress: number | null = null
    let hasSharedStringType = false
    while (index < tagEnd) {
      while (index < tagEnd && isAsciiWhitespace(this.buffer[index] ?? 0)) {
        index += 1
      }
      const nameStart = index
      while (index < tagEnd && isXmlNameByte(this.buffer[index] ?? 0)) {
        index += 1
      }
      const nameEnd = index
      index = skipAsciiWhitespace(this.buffer, index, tagEnd)
      if (this.buffer[index] !== 61) {
        index += 1
        continue
      }
      index = skipAsciiWhitespace(this.buffer, index + 1, tagEnd)
      const quote = this.buffer[index]
      if (quote !== doubleQuote && quote !== singleQuote) {
        index += 1
        continue
      }
      const valueStart = index + 1
      index = valueStart
      while (index < tagEnd && this.buffer[index] !== quote) {
        index += 1
      }
      const valueEnd = index
      if (attributeNameMatches(this.buffer, nameStart, nameEnd, 'r')) {
        packedAddress = decodePackedCellAddressBytes(this.buffer, valueStart, valueEnd)
        if (packedAddress === null) {
          return false
        }
      } else if (attributeNameMatches(this.buffer, nameStart, nameEnd, 't')) {
        hasSharedStringType = valueEnd - valueStart === 1 && this.buffer[valueStart] === 115
      }
      index += 1
    }
    this.cellPackedAddress = packedAddress ?? this.readImplicitCellPackedAddress()
    this.cellHasSharedStringType = hasSharedStringType
    return this.cellPackedAddress !== -1
  }

  private readImplicitCellPackedAddress(): number {
    return this.currentRow < 0 || this.nextImplicitColumn >= packedAddressColumnFactor
      ? -1
      : packCellAddress(this.currentRow, this.nextImplicitColumn)
  }

  private countMetadataElement(localName: string, nameEnd: number, tagEnd: number, final: boolean): boolean {
    if (isSelfClosingTag(this.buffer, tagEnd)) {
      this.finalizeMetadataCount({
        localName,
        childName: metadataChildName(localName),
        multiplier: metadataCountMultiplier(localName, this.buffer, nameEnd, tagEnd),
        childCount: 0,
      })
      this.index = tagEnd + 1
      return true
    }
    this.activeMetadata = {
      localName,
      childName: metadataChildName(localName),
      multiplier: metadataCountMultiplier(localName, this.buffer, nameEnd, tagEnd),
      childCount: 0,
    }
    this.index = tagEnd + 1
    if (!this.processActiveMetadata(final)) {
      if (final) {
        this.failed = true
      }
      return false
    }
    return true
  }

  private processActiveMetadata(final: boolean): boolean {
    const active = this.activeMetadata
    if (active === null) {
      return true
    }
    while (this.index < this.buffer.byteLength) {
      if (this.buffer[this.index] !== lessThan) {
        this.index += 1
        continue
      }
      const closing = this.buffer[this.index + 1] === slash
      const tagNameStart = this.index + (closing ? 2 : 1)
      const tag = readXmlTagName(this.buffer, tagNameStart)
      if (!tag) {
        if (!final && tagNameStart >= this.buffer.byteLength) {
          return false
        }
        this.index += 1
        continue
      }
      const tagEnd = findTagEnd(this.buffer, tag.endIndex)
      if (tagEnd === null) {
        if (final) {
          this.failed = true
        }
        return false
      }
      if (closing && tag.localName === active.localName) {
        this.finalizeMetadataCount(active)
        this.activeMetadata = null
        this.index = tagEnd + 1
        return true
      }
      if (!closing && active.childName !== null && tag.localName === active.childName) {
        active.childCount += 1
      }
      this.index = tagEnd + 1
    }
    if (final) {
      this.failed = true
    } else {
      this.index = this.buffer.byteLength
    }
    return false
  }

  private finalizeMetadataCount(active: ActiveMetadataCount): void {
    if (active.localName === 'conditionalFormatting') {
      this.conditionalFormatCount += active.multiplier * Math.max(1, active.childCount)
    } else if (active.localName === 'mergeCells') {
      this.mergeCount += active.childCount
    } else if (active.localName === 'tableParts') {
      this.tableCount += active.childCount
    }
  }

  private reportRetainedBufferLength(): void {
    this.options.onRetainedBufferLength?.(this.buffer.byteLength)
  }
}

function metadataChildName(localName: string): string | null {
  switch (localName) {
    case 'conditionalFormatting':
      return 'cfRule'
    case 'mergeCells':
      return 'mergeCell'
    case 'tableParts':
      return 'tablePart'
    default:
      return null
  }
}

function metadataCountMultiplier(localName: string, bytes: Uint8Array, nameEnd: number, tagEnd: number): number {
  return localName === 'conditionalFormatting' ? countSqrefRangesFromTag(bytes, nameEnd, tagEnd) : 1
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

function isSelfClosingTag(bytes: Uint8Array, tagEnd: number): boolean {
  let index = tagEnd - 1
  while (index >= 0 && isAsciiWhitespace(bytes[index] ?? 0)) {
    index -= 1
  }
  return bytes[index] === slash
}

function readDimensionAddressRange(
  bytes: Uint8Array,
  startIndex: number,
  endIndex: number,
): { readonly startRow: number; readonly startColumn: number; readonly endRow: number; readonly endColumn: number } | null {
  const trimmed = trimAsciiWhitespace(bytes, startIndex, endIndex)
  if (trimmed.start === trimmed.end) {
    return null
  }
  const separator = findByte(bytes, trimmed.start, trimmed.end, 58)
  const firstEnd = separator ?? trimmed.end
  const start = decodePackedCellAddressBytes(bytes, trimmed.start, firstEnd)
  const end = separator === null ? start : decodePackedCellAddressBytes(bytes, separator + 1, trimmed.end)
  return start === null || end === null
    ? null
    : {
        startRow: packedAddressRow(start),
        startColumn: packedAddressColumn(start),
        endRow: packedAddressRow(end),
        endColumn: packedAddressColumn(end),
      }
}

function countSqrefRangesFromTag(bytes: Uint8Array, nameEnd: number, tagEnd: number): number {
  const sqref = readXmlAttributeRangeFromTag(bytes, nameEnd, tagEnd, 'sqref')
  if (!sqref) {
    return 1
  }
  let count = 0
  let inToken = false
  for (let index = sqref.start; index < sqref.end; index += 1) {
    if (isAsciiWhitespace(bytes[index] ?? 0)) {
      inToken = false
      continue
    }
    if (!inToken) {
      count += 1
      inToken = true
    }
  }
  return Math.max(1, count)
}

function readPositiveIntegerAttributeFromTag(bytes: Uint8Array, nameEnd: number, tagEnd: number, attributeName: string): number | null {
  const range = readXmlAttributeRangeFromTag(bytes, nameEnd, tagEnd, attributeName)
  if (!range || range.start === range.end) {
    return null
  }
  let value = 0
  for (let index = range.start; index < range.end; index += 1) {
    const byte = bytes[index] ?? 0
    if (byte < 48 || byte > 57) {
      return null
    }
    value = value * 10 + byte - 48
  }
  return value > 0 && Number.isSafeInteger(value) ? value : null
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
  return index === localNameStart
    ? null
    : { localName: readKnownXmlLocalName(bytes, localNameStart, index) ?? decodeAscii(bytes, localNameStart, index), endIndex: index }
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

function findByte(bytes: Uint8Array, startIndex: number, endIndex: number, target: number): number | null {
  for (let index = startIndex; index < endIndex; index += 1) {
    if (bytes[index] === target) {
      return index
    }
  }
  return null
}

function decodePackedCellAddressBytes(bytes: Uint8Array, startIndex: number, endIndex: number): number | null {
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
  return letterCount > 0 && letterCount <= 3 && digitCount > 0 && row > 0 && column > 0 ? packCellAddress(row - 1, column - 1) : null
}

function packCellAddress(row: number, column: number): number {
  return row * packedAddressColumnFactor + column
}

function packedAddressRow(value: number): number {
  return Math.floor(value / packedAddressColumnFactor)
}

function packedAddressColumn(value: number): number {
  return value % packedAddressColumnFactor
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
