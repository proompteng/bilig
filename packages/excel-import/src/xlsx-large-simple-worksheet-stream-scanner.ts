import type {
  WorkbookAxisEntrySnapshot,
  WorkbookAxisMetadataSnapshot,
  WorkbookConditionalFormatSnapshot,
  WorkbookDataValidationSnapshot,
  WorkbookRichTextCellSnapshot,
} from '@bilig/protocol'
import { readLargeSimpleAutoFiltersFromBytes } from './xlsx-large-simple-autofilter-byte-scan.js'
import {
  readLargeSimpleCellValueFromTextRange,
  readLargeSimpleSharedStringIndexFromTextRange,
  type LargeSimpleXmlTextRange,
} from './xlsx-large-simple-cell-value-scan.js'
import { readLargeSimpleConditionalFormattingFromBytes } from './xlsx-large-simple-conditional-format-byte-scan.js'
import { readLargeSimpleDataValidationsFromBytes } from './xlsx-large-simple-data-validation-byte-scan.js'
import {
  LargeSimpleFormulaRecords,
  parseLargeSimpleSharedFormulaIndex,
  readLargeSimpleFormulaTypeCode,
} from './xlsx-large-simple-formula-records.js'
import { readLargeSimpleSheetHyperlinkRefsFromBytes } from './xlsx-large-simple-hyperlinks.js'
import { internLargeSimpleWorksheetMetadata } from './xlsx-large-simple-metadata-interning.js'
import { appendLargeSimplePrintPageSetupElement, isLargeSimplePrintPageSetupElementName } from './xlsx-large-simple-printer-settings.js'
import { rowTagHasMetadataAttribute } from './xlsx-large-simple-row-metadata-scan.js'
import { ImportedWorkbookArena, ImportedWorksheetStyleIndexArena } from './xlsx-large-simple-arena.js'
import {
  appendLargeSimpleColumnMetadataFromBytes,
  appendLargeSimpleRowMetadataTagFromBytes,
  readLargeSimpleDrawingRelationshipIdTagFromBytes,
  readLargeSimpleMergeRefsFromBytes,
  readLargeSimpleSheetFormatPrTagFromBytes,
  readLargeSimpleTableRelationshipIdsFromBytes,
} from './xlsx-large-simple-metadata-byte-scan.js'
import type { LargeSimpleSharedStringEntry } from './xlsx-large-simple-shared-strings.js'
import type { ImportedWorkbookStringPool } from './xlsx-large-simple-string-pool.js'
import { stringItemText } from './xlsx-large-simple-worksheet-stream-text.js'
import type { LargeSimpleWorksheetStreamScan, LargeSimpleWorksheetStreamScanOptions } from './xlsx-large-simple-worksheet-stream-types.js'
import { readKnownXmlLocalName } from './xlsx-large-simple-xml-name.js'
import {
  decodeAscii,
  decodeBytes,
  decodeCellAddress,
  decodePackedCellAddressBytes,
  encodeCellAddress,
  attributeNameMatches,
  isAsciiWhitespace,
  isXmlNameByte,
  packedAddressColumn,
  packedAddressRow,
  skipAsciiWhitespace,
} from './xlsx-large-simple-xml-byte-utils.js'
import {
  metadataWorksheetTagNames,
  richTextRunPattern,
  unsupportedWorksheetTagNames,
} from './xlsx-large-simple-worksheet-scan-constants.js'
import type { LargeSimpleWorksheetMergeRef, LargeSimpleWorksheetScannedMetadata } from './xlsx-large-simple-worksheet-metadata.js'

const lessThan = 60
const slash = 47
const greaterThan = 62
const doubleQuote = 34
const singleQuote = 39
const emptyBytes = new Uint8Array(0)
const maxDimensionArenaReserveCellCount = 1_000_000
type StreamedMetadataElement = 'mergeCells' | 'tableParts'

export function parseLargeSimpleWorksheetCellsFromChunks(
  readChunks: (onChunk: (chunk: Uint8Array) => void) => boolean,
  sheetIndex: number,
  options: LargeSimpleWorksheetStreamScanOptions,
): LargeSimpleWorksheetStreamScan | null {
  const scanner = new LargeSimpleWorksheetChunkScanner(sheetIndex, {
    hasSharedStrings: options.hasSharedStrings,
    retainCells: options.retainCells !== false,
    sharedStrings: options.sharedStrings ?? [],
    deferSharedStrings: options.deferSharedStrings === true,
    retainMetadataXml: options.retainMetadataXml !== false,
    sheetName: options.sheetName,
    stringPool: options.stringPool,
    onRetainedBufferLength: options.onRetainedBufferLength,
  })
  if (!readChunks((chunk) => scanner.push(chunk))) {
    return null
  }
  return scanner.finish()
}

class LargeSimpleWorksheetChunkScanner {
  private buffer: Uint8Array = new Uint8Array()
  private index = 0
  private failed = false
  private readonly arena: ImportedWorkbookArena
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
  private readonly sheetName: string | undefined
  private minRow = Number.POSITIVE_INFINITY
  private minColumn = Number.POSITIVE_INFINITY
  private maxRow = -1
  private maxColumn = -1
  private columnEntries: WorkbookAxisEntrySnapshot[] | undefined
  private columnMetadata: WorkbookAxisMetadataSnapshot[] | undefined
  private conditionalFormats: WorkbookConditionalFormatSnapshot[] | undefined
  private conditionalFormatIdCounter = 0
  private conditionalFormattingXml: string[] | undefined
  private dataValidations: WorkbookDataValidationSnapshot[] | undefined
  private drawingRelationshipId: string | undefined
  private filters: LargeSimpleWorksheetScannedMetadata['filters']
  private hyperlinks: LargeSimpleWorksheetScannedMetadata['hyperlinks']
  private rowEntries: WorkbookAxisEntrySnapshot[] | undefined
  private rowMetadata: WorkbookAxisMetadataSnapshot[] | undefined
  private mergeRefs: LargeSimpleWorksheetMergeRef[] | undefined
  private printPageSetup: LargeSimpleWorksheetScannedMetadata['printPageSetup']
  private sheetFormatPr: LargeSimpleWorksheetScannedMetadata['sheetFormatPr']
  private tableRelationshipIds: string[] | undefined
  private readonly metadataSnippets: string[] = []
  private readonly hasSharedStrings: boolean
  private readonly retainCells: boolean
  private readonly sharedStrings: readonly LargeSimpleSharedStringEntry[]
  private readonly deferSharedStrings: boolean
  private readonly retainMetadataXml: boolean
  private readonly stringPool: ImportedWorkbookStringPool | undefined
  private activeMetadataElement: StreamedMetadataElement | null = null

  constructor(
    private readonly sheetIndex: number,
    options: {
      readonly hasSharedStrings: boolean
      readonly retainCells: boolean
      readonly sharedStrings: readonly LargeSimpleSharedStringEntry[]
      readonly deferSharedStrings: boolean
      readonly retainMetadataXml: boolean
      readonly sheetName: string | undefined
      readonly stringPool: ImportedWorkbookStringPool | undefined
      readonly onRetainedBufferLength: ((length: number) => void) | undefined
    },
  ) {
    this.arena = new ImportedWorkbookArena(options.stringPool)
    this.hasSharedStrings = options.hasSharedStrings
    this.retainCells = options.retainCells
    this.sharedStrings = options.sharedStrings
    this.deferSharedStrings = options.deferSharedStrings
    this.retainMetadataXml = options.retainMetadataXml
    this.sheetName = options.sheetName
    this.stringPool = options.stringPool
    this.reportRetainedBufferLength = () => options.onRetainedBufferLength?.(this.buffer.byteLength)
  }

  push(chunk: Uint8Array): void {
    if (this.failed || chunk.byteLength === 0) {
      return
    }
    this.append(chunk)
    this.process(false)
    this.compact()
    this.reportRetainedBufferLength()
  }

  finish(): LargeSimpleWorksheetStreamScan | null {
    if (this.failed) {
      return null
    }
    this.process(true)
    if (this.activeMetadataElement !== null) {
      this.failed = true
    }
    this.compact()
    this.reportRetainedBufferLength()
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
      metadata: internLargeSimpleWorksheetMetadata(this.buildMetadataScan(), this.stringPool),
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

  private buildMetadataScan(): LargeSimpleWorksheetScannedMetadata | undefined {
    const columns =
      (this.columnEntries?.length ?? 0) > 0 || (this.columnMetadata?.length ?? 0) > 0
        ? { entries: this.columnEntries ?? [], metadata: this.columnMetadata ?? [] }
        : undefined
    const rows =
      (this.rowEntries?.length ?? 0) > 0 || (this.rowMetadata?.length ?? 0) > 0
        ? { entries: this.rowEntries ?? [], metadata: this.rowMetadata ?? [] }
        : undefined
    const metadata: LargeSimpleWorksheetScannedMetadata = {
      ...(columns ? { columns } : {}),
      ...(this.conditionalFormats && this.conditionalFormats.length > 0 ? { conditionalFormats: this.conditionalFormats } : {}),
      ...(this.conditionalFormattingXml && this.conditionalFormattingXml.length > 0
        ? { conditionalFormattingXml: this.conditionalFormattingXml }
        : {}),
      ...(this.dataValidations && this.dataValidations.length > 0 ? { dataValidations: this.dataValidations } : {}),
      ...(this.drawingRelationshipId ? { drawingRelationshipId: this.drawingRelationshipId } : {}),
      ...(this.filters && this.filters.length > 0 ? { filters: this.filters } : {}),
      ...(this.hyperlinks && this.hyperlinks.length > 0 ? { hyperlinks: this.hyperlinks } : {}),
      ...(rows ? { rows } : {}),
      ...(this.mergeRefs && this.mergeRefs.length > 0 ? { merges: this.mergeRefs } : {}),
      ...(this.printPageSetup ? { printPageSetup: this.printPageSetup } : {}),
      ...(this.sheetFormatPr ? { sheetFormatPr: this.sheetFormatPr } : {}),
      ...(this.tableRelationshipIds && this.tableRelationshipIds.length > 0 ? { tableRelationshipIds: this.tableRelationshipIds } : {}),
    }
    return Object.keys(metadata).length > 0 ? metadata : undefined
  }

  private process(final: boolean): void {
    while (!this.failed && this.index < this.buffer.byteLength) {
      if (this.activeMetadataElement !== null) {
        if (!this.processActiveMetadataElement(final)) {
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
          this.collectRowMetadata(tag.endIndex, tagEnd)
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
    const dimensionCellCount = (end.row - start.row + 1) * (end.column - start.column + 1)
    if (this.retainCells && dimensionCellCount > 0 && dimensionCellCount <= maxDimensionArenaReserveCellCount) {
      this.arena.reserveCellCapacity(dimensionCellCount)
    }
  }

  private collectRowMetadata(nameEnd: number, tagEnd: number): void {
    if (!rowTagHasMetadataAttribute(this.buffer, nameEnd, tagEnd)) {
      return
    }
    this.rowEntries ??= []
    this.rowMetadata ??= []
    appendLargeSimpleRowMetadataTagFromBytes(this.rowEntries, this.rowMetadata, this.buffer, nameEnd, tagEnd)
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
    const packedAddress = readPackedCellAddressAttributeFromTag(this.buffer, nameEnd, tagEnd)
    if (packedAddress === null) {
      this.failed = true
      return false
    }
    const row = packedAddressRow(packedAddress)
    const column = packedAddressColumn(packedAddress)
    const cellType = readXmlAttributeFromTag(this.buffer, nameEnd, tagEnd, 't')
    if ((!this.hasSharedStrings && cellType === 's') || (!this.retainCells && cellType === 'inlineStr')) {
      this.failed = true
      return false
    }
    const styleIndex = readCellStyleIndexFromTag(this.buffer, nameEnd, tagEnd)
    const shouldReadSharedStringIndex = cellType === 's' && (this.retainCells || this.deferSharedStrings)
    const rawValueRange =
      this.retainCells || shouldReadSharedStringIndex ? readElementTextRange(this.buffer, contentStart, closing.start, 'v') : null
    const sharedStringIndex = shouldReadSharedStringIndex ? readLargeSimpleSharedStringIndexFromTextRange(this.buffer, rawValueRange) : null
    if (shouldReadSharedStringIndex && rawValueRange !== null) {
      if (sharedStringIndex === null) {
        this.failed = true
        return false
      }
    }
    const deferSharedStringValue = this.retainCells && this.deferSharedStrings && cellType === 's' && sharedStringIndex !== null
    const value =
      this.retainCells && !deferSharedStringValue
        ? cellType === 'inlineStr'
          ? readInlineStringCellValue(this.buffer, contentStart, closing.start)
          : readLargeSimpleCellValueFromTextRange(this.buffer, rawValueRange, cellType, this.sharedStrings)
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
      const cellIndex = this.retainCells
        ? deferSharedStringValue
          ? this.arena.addSharedStringCell({
              sheetIndex: this.sheetIndex,
              row,
              column,
              sharedStringIndex,
            })
          : this.arena.addCell({
              sheetIndex: this.sheetIndex,
              row,
              column,
              value,
            })
        : -1
      if (formula && this.retainCells) {
        this.formulas.add(cellIndex, row, column, formula.typeCode, formula.sharedIndex, formula.rawFormula)
      }
      if (this.retainCells && styleIndex !== null) {
        this.styleIndexes.add(row, column, styleIndex)
      }
      const richTextCell =
        this.retainCells && !deferSharedStringValue
          ? readRichTextCellArtifact(this.buffer, contentStart, closing.start, row, column, cellType, sharedStringIndex, this.sharedStrings)
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
      const handled = this.retainMetadataXml && this.collectTypedMetadataElement(localName, this.index, tagEnd + 1)
      if (!handled) {
        this.countMetadataElement(localName, tagEnd + 1, tagEnd + 1)
      }
      if (this.retainMetadataXml && !handled) {
        this.metadataSnippets.push(decodeBytes(this.buffer, this.index, tagEnd + 1))
      }
      this.index = tagEnd + 1
      return true
    }
    if (localName === 'mergeCells' || localName === 'tableParts') {
      this.activeMetadataElement = localName
      this.index = tagEnd + 1
      if (!this.processActiveMetadataElement(final)) {
        if (final) {
          this.failed = true
        }
        return false
      }
      return true
    }
    const closing = findClosingTag(this.buffer, tagEnd + 1, localName)
    if (!closing) {
      if (final) {
        this.failed = true
      }
      return false
    }
    const handled = this.retainMetadataXml && this.collectTypedMetadataElement(localName, this.index, closing.end)
    if (!handled) {
      this.countMetadataElement(localName, tagEnd + 1, closing.start)
    }
    if (this.retainMetadataXml && !handled) {
      this.metadataSnippets.push(decodeBytes(this.buffer, this.index, closing.end))
    }
    this.index = closing.end
    return true
  }

  private collectTypedMetadataElement(localName: string, startIndex: number, endIndex: number): boolean {
    if (localName === 'mergeCells') {
      const refs = readLargeSimpleMergeRefsFromBytes(this.buffer, startIndex, endIndex)
      this.mergeCount += refs.length
      if (refs.length > 0) {
        this.mergeRefs ??= []
        this.mergeRefs.push(...refs)
      }
      return true
    }
    if (localName === 'cols') {
      this.columnEntries ??= []
      this.columnMetadata ??= []
      appendLargeSimpleColumnMetadataFromBytes(this.columnEntries, this.columnMetadata, this.buffer, startIndex, endIndex)
      return true
    }
    if (localName === 'sheetFormatPr') {
      this.sheetFormatPr = readLargeSimpleSheetFormatPrTagFromBytes(this.buffer, startIndex, endIndex) ?? this.sheetFormatPr
      return true
    }
    if (localName === 'drawing') {
      this.drawingRelationshipId = readLargeSimpleDrawingRelationshipIdTagFromBytes(this.buffer, startIndex, endIndex)
      return true
    }
    if (localName === 'autoFilter') {
      if (!this.sheetName) {
        return false
      }
      const filters = readLargeSimpleAutoFiltersFromBytes(this.sheetName, this.buffer, startIndex, endIndex)
      this.filters = [...(this.filters ?? []), ...filters]
      return true
    }
    if (localName === 'hyperlinks') {
      const refs = readLargeSimpleSheetHyperlinkRefsFromBytes(this.buffer, startIndex, endIndex)
      if (refs === null) {
        return false
      }
      this.hyperlinks = [...(this.hyperlinks ?? []), ...refs]
      return true
    }
    if (isLargeSimplePrintPageSetupElementName(localName)) {
      this.printPageSetup ??= {}
      appendLargeSimplePrintPageSetupElement(this.printPageSetup, localName, decodeBytes(this.buffer, startIndex, endIndex))
      return true
    }
    if (localName === 'tableParts') {
      const relationshipIds = readLargeSimpleTableRelationshipIdsFromBytes(this.buffer, startIndex, endIndex)
      this.tableCount += relationshipIds.length
      if (relationshipIds.length > 0) {
        this.tableRelationshipIds ??= []
        this.tableRelationshipIds.push(...relationshipIds)
      }
      return true
    }
    if (localName === 'conditionalFormatting') {
      if (!this.sheetName) {
        return false
      }
      const scan = readLargeSimpleConditionalFormattingFromBytes(
        this.sheetName,
        this.buffer,
        startIndex,
        endIndex,
        this.conditionalFormatIdCounter + 1,
      )
      this.conditionalFormatCount += scan.ruleCount
      if (scan.conditionalFormats && scan.conditionalFormats.length > 0) {
        this.conditionalFormats ??= []
        this.conditionalFormats.push(...scan.conditionalFormats)
        this.conditionalFormatIdCounter += scan.conditionalFormats.length
      }
      if (scan.artifactXml) {
        this.conditionalFormattingXml ??= []
        this.conditionalFormattingXml.push(scan.artifactXml)
      }
      return true
    }
    if (localName === 'dataValidations') {
      if (!this.sheetName) {
        return false
      }
      const validations = readLargeSimpleDataValidationsFromBytes(this.sheetName, this.buffer, startIndex, endIndex)
      if (validations === null) {
        this.failed = true
        return true
      }
      if (validations.length > 0) {
        this.dataValidations ??= []
        this.dataValidations.push(...validations)
      }
      return true
    }
    return false
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

  private processActiveMetadataElement(final: boolean): boolean {
    const activeElement = this.activeMetadataElement
    if (activeElement === null) {
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
      if (closing && tag.localName === activeElement) {
        this.activeMetadataElement = null
        this.index = tagEnd + 1
        return true
      }
      if (!closing) {
        this.collectActiveMetadataTag(activeElement, tag.localName, tag.endIndex, tagEnd)
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

  private collectActiveMetadataTag(activeElement: StreamedMetadataElement, localName: string, nameEnd: number, tagEnd: number): void {
    if (activeElement === 'mergeCells' && localName === 'mergeCell') {
      this.collectMergeCellTag(nameEnd, tagEnd)
      return
    }
    if (activeElement === 'tableParts' && localName === 'tablePart') {
      const relationshipId =
        readXmlAttributeFromTag(this.buffer, nameEnd, tagEnd, 'r:id') ?? readXmlAttributeFromTag(this.buffer, nameEnd, tagEnd, 'id')
      if (relationshipId) {
        this.tableCount += 1
        if (this.retainMetadataXml) {
          this.tableRelationshipIds ??= []
          this.tableRelationshipIds.push(relationshipId)
        }
      }
    }
  }

  private collectMergeCellTag(nameEnd: number, tagEnd: number): void {
    const ref = readXmlAttributeFromTag(this.buffer, nameEnd, tagEnd, 'ref')
    const [startAddress, endAddress] = ref?.split(':') ?? []
    if (!startAddress || !endAddress || startAddress === endAddress) {
      return
    }
    this.mergeCount += 1
    if (this.retainMetadataXml) {
      this.mergeRefs ??= []
      this.mergeRefs.push({ startAddress, endAddress })
    }
  }

  private reportRetainedBufferLength: () => void = () => {}
}

function readInlineStringCellValue(bytes: Uint8Array, contentStart: number, contentEnd: number): string | undefined {
  const inlineStringXml = readElementXml(bytes, contentStart, contentEnd, 'is')
  return inlineStringXml ? stringItemText(inlineStringXml) : undefined
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
  row: number,
  column: number,
  type: string | null,
  sharedStringIndex: number | null,
  sharedStrings: readonly LargeSimpleSharedStringEntry[],
): WorkbookRichTextCellSnapshot | undefined {
  if (type === 's') {
    const entry = sharedStringIndex === null ? undefined : sharedStrings[sharedStringIndex]
    return entry?.rich
      ? {
          address: encodeCellAddress(row, column),
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
    address: encodeCellAddress(row, column),
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

function readElementTextRange(
  bytes: Uint8Array,
  startIndex: number,
  endIndex: number,
  elementName: string,
): LargeSimpleXmlTextRange | null {
  const tag = findNextOpeningTag(bytes, startIndex, elementName, endIndex)
  if (!tag) {
    return null
  }
  const tagEnd = findTagEnd(bytes, tag.nameEnd, endIndex)
  if (tagEnd === null || isSelfClosingTag(bytes, tagEnd)) {
    return null
  }
  const closing = findClosingTag(bytes, tagEnd + 1, elementName, endIndex)
  return closing ? { start: tagEnd + 1, end: closing.start } : null
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
  return index === localNameStart
    ? null
    : { localName: readKnownXmlLocalName(bytes, localNameStart, index) ?? decodeAscii(bytes, localNameStart, index), endIndex: index }
}

function readXmlAttributeFromTag(bytes: Uint8Array, startIndex: number, tagEnd: number, attributeName: string): string | null {
  const range = readXmlAttributeRangeFromTag(bytes, startIndex, tagEnd, attributeName)
  return range ? decodeBytes(bytes, range.start, range.end) : null
}

function readPackedCellAddressAttributeFromTag(bytes: Uint8Array, startIndex: number, tagEnd: number): number | null {
  const range = readXmlAttributeRangeFromTag(bytes, startIndex, tagEnd, 'r')
  return range ? decodePackedCellAddressBytes(bytes, range.start, range.end) : null
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
