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

export interface HeadlessLargeSimpleWorksheetScan {
  readonly sheetIndex: number
  readonly cellCount: number
  readonly valueCellCount: number
  readonly formulaCellCount: number
  readonly tableCount: number
  readonly mergeCount: number
  readonly conditionalFormatCount: number
  readonly rowCount: number
  readonly columnCount: number
  readonly usedRange: {
    readonly startRow: number
    readonly startColumn: number
    readonly endRow: number
    readonly endColumn: number
  } | null
}

export function parseHeadlessLargeSimpleWorksheetFromChunks(
  readChunks: (onChunk: (chunk: Uint8Array) => void) => boolean,
  sheetIndex: number,
  options: { readonly hasSharedStrings: boolean },
): HeadlessLargeSimpleWorksheetScan | null {
  const scanner = new HeadlessLargeSimpleWorksheetChunkScanner(sheetIndex, options.hasSharedStrings)
  if (!readChunks((chunk) => scanner.push(chunk))) {
    return null
  }
  return scanner.finish()
}

class HeadlessLargeSimpleWorksheetChunkScanner {
  private buffer = new Uint8Array()
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
  private minRow = Number.POSITIVE_INFINITY
  private minColumn = Number.POSITIVE_INFINITY
  private maxRow = -1
  private maxColumn = -1

  constructor(
    private readonly sheetIndex: number,
    private readonly hasSharedStrings: boolean,
  ) {}

  push(chunk: Uint8Array): void {
    if (this.failed || chunk.byteLength === 0) {
      return
    }
    this.append(chunk)
    this.process(false)
    this.compact()
  }

  finish(): HeadlessLargeSimpleWorksheetScan | null {
    if (this.failed) {
      return null
    }
    this.process(true)
    this.compact()
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
      if (tag.localName === 'c') {
        if (!this.readCell(tag.endIndex, tagEnd, final)) {
          return
        }
        continue
      }
      if (metadataWorksheetTagNames.has(tag.localName)) {
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
    const address = readCellAddressAttributeFromTag(this.buffer, nameEnd, tagEnd)
    if (!address || (!this.hasSharedStrings && readXmlAttributeFromTag(this.buffer, nameEnd, tagEnd, 't') === 's')) {
      this.failed = true
      return false
    }
    const hasValue = hasElement(this.buffer, contentStart, closing.start, 'v') || hasElement(this.buffer, contentStart, closing.start, 'is')
    const hasFormula = hasElement(this.buffer, contentStart, closing.start, 'f')
    if (hasValue || hasFormula) {
      this.cellCount += 1
      this.rowCount = Math.max(this.rowCount, address.row + 1)
      this.columnCount = Math.max(this.columnCount, address.column + 1)
      this.minRow = Math.min(this.minRow, address.row)
      this.minColumn = Math.min(this.minColumn, address.column)
      this.maxRow = Math.max(this.maxRow, address.row)
      this.maxColumn = Math.max(this.maxColumn, address.column)
      if (hasValue) {
        this.valueCellCount += 1
      }
      if (hasFormula) {
        this.formulaCellCount += 1
      }
    }
    this.index = selfClosing ? tagEnd + 1 : closing.end
    return true
  }

  private countMetadataElement(localName: string, nameEnd: number, tagEnd: number, final: boolean): boolean {
    if (isSelfClosingTag(this.buffer, tagEnd)) {
      this.countMetadata(localName, nameEnd, tagEnd, tagEnd + 1, tagEnd + 1)
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
    this.countMetadata(localName, nameEnd, tagEnd, tagEnd + 1, closing.start)
    this.index = closing.end
    return true
  }

  private countMetadata(localName: string, nameEnd: number, tagEnd: number, contentStart: number, contentEnd: number): void {
    if (localName === 'conditionalFormatting') {
      const sqref = readXmlAttributeFromTag(this.buffer, nameEnd, tagEnd, 'sqref')
      const rangeCount = sqref ? Math.max(1, sqref.trim().split(/\s+/u).filter(Boolean).length) : 1
      this.conditionalFormatCount += rangeCount * Math.max(1, countOpeningTags(this.buffer, contentStart, contentEnd, 'cfRule'))
    } else if (localName === 'mergeCells') {
      this.mergeCount += countOpeningTags(this.buffer, contentStart, contentEnd, 'mergeCell')
    } else if (localName === 'tableParts') {
      this.tableCount += countOpeningTags(this.buffer, contentStart, contentEnd, 'tablePart')
    }
  }
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
  return range ? decodeAscii(bytes, range.start, range.end) : null
}

function readCellAddressAttributeFromTag(
  bytes: Uint8Array,
  startIndex: number,
  tagEnd: number,
): { readonly row: number; readonly column: number } | null {
  const range = readXmlAttributeRangeFromTag(bytes, startIndex, tagEnd, 'r')
  return range ? decodeCellAddressBytes(bytes, range.start, range.end) : null
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
