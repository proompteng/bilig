import type { WorkbookAutoFilterSnapshot, WorkbookRichTextCellSnapshot } from '@bilig/protocol'
import {
  readLargeSimpleAutoFilterRootFromBytes,
  readLargeSimpleAutoFiltersFromBytes,
  wrapLargeSimpleAutoFilterColumnBytes,
} from './xlsx-large-simple-autofilter-byte-scan.js'
import {
  readLargeSimpleCellValueFromTextRange,
  readLargeSimpleSharedStringIndexFromTextRange,
} from './xlsx-large-simple-cell-value-scan.js'
import { LargeSimpleFormulaRecords, readLargeSimpleFormulaTypeCode } from './xlsx-large-simple-formula-records.js'
import {
  ImportedWorkbookArena,
  ImportedWorksheetStyleIndexArena,
  type ImportedWorkbookArenaDedupeMode,
  type ImportedWorksheetCellScan,
} from './xlsx-large-simple-arena.js'
import type { LargeSimpleSharedStrings } from './xlsx-large-simple-shared-strings.js'
import type { ImportedWorkbookStringPool } from './xlsx-large-simple-string-pool.js'
import { decodeBytes, decodeCellAddress, packedAddressColumn, packedAddressRow } from './xlsx-large-simple-xml-byte-utils.js'
import {
  findClosingTag,
  findTagEnd,
  hasElement,
  isSelfClosingTag,
  readCellStyleIndexFromTag,
  readElementTextRange,
  readPackedCellAddressAttributeFromTag,
  readXmlAttributeFromTag,
  readXmlTagName,
} from './xlsx-large-simple-worksheet-stream-xml.js'
import { metadataWorksheetTagNames, unsupportedWorksheetTagNames } from './xlsx-large-simple-worksheet-scan-constants.js'
import {
  readFormulaSpec,
  readInlineStringCellValue,
  readPositiveIntegerAttributeFromTag,
  readRichTextCellArtifact,
} from './xlsx-large-simple-worksheet-stream-cell-readers.js'
import {
  type ActiveConditionalFormatting,
  LargeSimpleWorksheetStreamMetadataCollector,
  type StreamedMetadataElement,
} from './xlsx-large-simple-worksheet-stream-metadata.js'
import type { LargeSimpleWorksheetScannedMetadata } from './xlsx-large-simple-worksheet-metadata.js'

const lessThan = 60
const slash = 47
const emptyBytes = new Uint8Array(0)
const packedAddressColumnFactor = 16_384

interface ActiveAutoFilter {
  readonly rootTag: Uint8Array
  filter: WorkbookAutoFilterSnapshot | null
}

export interface LargeSimpleWorksheetStreamScan {
  readonly cellScan: ImportedWorksheetCellScan
  readonly metadataXml: string | undefined
  readonly metadata: LargeSimpleWorksheetScannedMetadata | undefined
}

export function parseLargeSimpleWorksheetCellsFromChunks(
  readChunks: (onChunk: (chunk: Uint8Array) => void) => boolean,
  sheetIndex: number,
  options: {
    readonly hasSharedStrings: boolean
    readonly retainCells?: boolean
    readonly sharedStrings?: LargeSimpleSharedStrings
    readonly deferSharedStrings?: boolean
    readonly retainMetadataXml?: boolean
    readonly sheetName?: string
    readonly stringPool?: ImportedWorkbookStringPool
    readonly deduplicateStrings?: ImportedWorkbookArenaDedupeMode
    readonly deduplicateFormulas?: ImportedWorkbookArenaDedupeMode
    readonly allowUnsupportedFormulaText?: boolean
    readonly allowUnsupportedCellMetadata?: boolean
    readonly preserveBlankStyleCells?: boolean
    readonly retainStyleIndexes?: boolean
    readonly retainStyleCoordinates?: boolean
    readonly maxDimensionCellPreallocation?: number
    readonly onRetainedBufferLength?: (length: number) => void
  },
): LargeSimpleWorksheetStreamScan | null {
  const retainCells = options.retainCells !== false
  const retainStyleIndexes = options.retainStyleIndexes ?? retainCells
  const scanner = new LargeSimpleWorksheetChunkScanner(sheetIndex, {
    hasSharedStrings: options.hasSharedStrings,
    retainCells,
    sharedStrings: options.sharedStrings ?? [],
    deferSharedStrings: options.deferSharedStrings === true,
    retainMetadataXml: options.retainMetadataXml !== false,
    sheetName: options.sheetName,
    stringPool: options.stringPool,
    deduplicateStrings: options.deduplicateStrings,
    deduplicateFormulas: options.deduplicateFormulas,
    allowUnsupportedFormulaText: options.allowUnsupportedFormulaText,
    allowUnsupportedCellMetadata: options.allowUnsupportedCellMetadata,
    preserveBlankStyleCells: options.preserveBlankStyleCells !== false,
    retainStyleIndexes,
    retainStyleCoordinates: options.retainStyleCoordinates ?? retainStyleIndexes,
    maxDimensionCellPreallocation: options.maxDimensionCellPreallocation,
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
  private readonly formulas: LargeSimpleFormulaRecords
  private readonly richTextCells: WorkbookRichTextCellSnapshot[] = []
  private readonly styleIndexes = new ImportedWorksheetStyleIndexArena()
  private readonly metadata: LargeSimpleWorksheetStreamMetadataCollector
  private rowCount = 0
  private columnCount = 0
  private cellCount = 0
  private valueCellCount = 0
  private formulaCellCount = 0
  private blankStyleCellCount = 0
  private readonly sheetName: string | undefined
  private minRow = Number.POSITIVE_INFINITY
  private minColumn = Number.POSITIVE_INFINITY
  private maxRow = -1
  private maxColumn = -1
  private currentRow = -1
  private nextImplicitRow = 0
  private nextImplicitColumn = 0
  private readonly hasSharedStrings: boolean
  private readonly retainCells: boolean
  private readonly sharedStrings: LargeSimpleSharedStrings
  private readonly deferSharedStrings: boolean
  private readonly retainMetadataXml: boolean
  private readonly allowUnsupportedFormulaText: boolean
  private readonly preserveBlankStyleCells: boolean
  private readonly retainStyleIndexes: boolean
  private readonly retainStyleCoordinates: boolean
  private readonly maxDimensionCellPreallocation: number
  private dimensionCellPreallocationApplied = false
  private activeMetadataElement: StreamedMetadataElement | null = null
  private activeAutoFilter: ActiveAutoFilter | null = null
  private activeConditionalFormatting: ActiveConditionalFormatting | null = null
  private activeDataValidations = false
  private activeHyperlinks = false

  constructor(
    private readonly sheetIndex: number,
    options: {
      readonly hasSharedStrings: boolean
      readonly retainCells: boolean
      readonly sharedStrings: LargeSimpleSharedStrings
      readonly deferSharedStrings: boolean
      readonly retainMetadataXml: boolean
      readonly sheetName: string | undefined
      readonly stringPool: ImportedWorkbookStringPool | undefined
      readonly deduplicateStrings: ImportedWorkbookArenaDedupeMode | undefined
      readonly deduplicateFormulas: ImportedWorkbookArenaDedupeMode | undefined
      readonly allowUnsupportedFormulaText: boolean | undefined
      readonly allowUnsupportedCellMetadata: boolean | undefined
      readonly preserveBlankStyleCells: boolean
      readonly retainStyleIndexes: boolean
      readonly retainStyleCoordinates: boolean
      readonly maxDimensionCellPreallocation: number | undefined
      readonly onRetainedBufferLength: ((length: number) => void) | undefined
    },
  ) {
    this.allowUnsupportedFormulaText = options.allowUnsupportedFormulaText === true
    this.preserveBlankStyleCells = options.preserveBlankStyleCells
    this.retainStyleIndexes = options.retainStyleIndexes
    this.retainStyleCoordinates = options.retainStyleCoordinates
    this.maxDimensionCellPreallocation = Math.max(0, Math.trunc(options.maxDimensionCellPreallocation ?? 0))
    this.formulas = new LargeSimpleFormulaRecords(this.allowUnsupportedFormulaText)
    this.arena = new ImportedWorkbookArena(options.stringPool, {
      ...(options.deduplicateStrings === undefined ? {} : { deduplicateStrings: options.deduplicateStrings }),
      ...(options.deduplicateFormulas === undefined ? {} : { deduplicateFormulas: options.deduplicateFormulas }),
    })
    this.hasSharedStrings = options.hasSharedStrings
    this.retainCells = options.retainCells
    this.sharedStrings = options.sharedStrings
    this.deferSharedStrings = options.deferSharedStrings
    this.retainMetadataXml = options.retainMetadataXml
    this.sheetName = options.sheetName
    this.metadata = new LargeSimpleWorksheetStreamMetadataCollector(this.sheetName, this.retainMetadataXml)
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
    if (
      this.activeMetadataElement !== null ||
      this.activeAutoFilter !== null ||
      this.activeConditionalFormatting !== null ||
      this.activeHyperlinks
    ) {
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
        mergeCount: this.metadata.mergeCount,
        conditionalFormatCount: this.metadata.conditionalFormatCount,
        dataValidationCount: this.metadata.dataValidationCount,
        tableCount: this.metadata.tableCount,
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
      metadataXml: undefined,
      metadata: this.metadata.buildMetadataScan(),
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
      if (this.activeMetadataElement !== null) {
        if (!this.processActiveMetadataElement(final)) {
          return
        }
        continue
      }
      if (this.activeAutoFilter !== null) {
        if (!this.processActiveAutoFilter(final)) {
          return
        }
        continue
      }
      if (this.activeConditionalFormatting !== null) {
        if (!this.processActiveConditionalFormatting(final)) {
          return
        }
        continue
      }
      if (this.activeDataValidations) {
        if (!this.processActiveDataValidations(final)) {
          return
        }
        continue
      }
      if (this.activeHyperlinks) {
        if (!this.processActiveHyperlinks(final)) {
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
      if (unsupportedWorksheetTagNames.has(tag.localName)) {
        this.failed = true
        return
      }
      if (tag.localName === 'worksheet') {
        if (this.retainMetadataXml) {
          this.metadata.setWorksheetRootOpenTag(decodeBytes(this.buffer, this.index, tagEnd + 1))
        }
        this.index = tagEnd + 1
        continue
      }
      if (tag.localName === 'dimension') {
        this.readDimension(tag.endIndex, tagEnd)
        this.index = tagEnd + 1
        continue
      }
      if (tag.localName === 'row') {
        this.readRow(tag.endIndex, tagEnd)
        if (this.retainMetadataXml) {
          this.metadata.collectRowMetadata(this.buffer, tag.endIndex, tagEnd, this.currentRow)
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

  private readRow(nameEnd: number, tagEnd: number): void {
    const row = readPositiveIntegerAttributeFromTag(this.buffer, nameEnd, tagEnd, 'r')
    this.currentRow = row === null ? this.nextImplicitRow : row - 1
    this.nextImplicitRow = this.currentRow + 1
    this.nextImplicitColumn = 0
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
    this.reserveDimensionCellCapacity(end.row + 1, end.column + 1)
  }

  private reserveDimensionCellCapacity(rowCount: number, columnCount: number): void {
    const cellCapacity = rowCount * columnCount
    if (
      this.dimensionCellPreallocationApplied ||
      !Number.isSafeInteger(cellCapacity) ||
      cellCapacity <= 0 ||
      cellCapacity > this.maxDimensionCellPreallocation
    ) {
      return
    }
    this.dimensionCellPreallocationApplied = true
    this.arena.reserveDenseRowMajorCellCapacity(this.sheetIndex, columnCount, rowCount)
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
    const packedAddress = this.readCellPackedAddress(nameEnd, tagEnd)
    if (packedAddress === null) {
      this.failed = true
      return false
    }
    const row = packedAddressRow(packedAddress)
    const column = packedAddressColumn(packedAddress)
    this.currentRow = row
    this.nextImplicitColumn = column + 1
    this.metadata.collectCellMetadataRef(this.buffer, row, column, nameEnd, tagEnd)
    const cellType = readXmlAttributeFromTag(this.buffer, nameEnd, tagEnd, 't')
    if (!this.hasSharedStrings && cellType === 's') {
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
      ? readFormulaSpec(this.buffer, contentStart, closing.start, this.allowUnsupportedFormulaText)
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
      if (styleIndex !== null) {
        this.recordStyleIndex(row, column, styleIndex)
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
      if (this.preserveBlankStyleCells) {
        this.recordStyleIndex(row, column, styleIndex)
      }
    }
    this.index = selfClosing ? tagEnd + 1 : closing.end
    return true
  }

  private readCellPackedAddress(nameEnd: number, tagEnd: number): number | null {
    const packedAddress = readPackedCellAddressAttributeFromTag(this.buffer, nameEnd, tagEnd)
    if (packedAddress !== null) {
      return packedAddress
    }
    return this.currentRow < 0 || this.nextImplicitColumn >= packedAddressColumnFactor
      ? null
      : this.currentRow * packedAddressColumnFactor + this.nextImplicitColumn
  }

  private recordStyleIndex(row: number, column: number, styleIndex: number): void {
    if (!this.retainStyleIndexes) {
      return
    }
    if (this.retainStyleCoordinates) {
      this.styleIndexes.add(row, column, styleIndex)
      return
    }
    this.styleIndexes.addRequiredStyleIndex(styleIndex)
  }

  private collectMetadataElement(localName: string, tagEnd: number, final: boolean): boolean {
    if (localName === 'dataValidations' && isSelfClosingTag(this.buffer, tagEnd)) {
      this.index = tagEnd + 1
      return true
    }
    if (localName === 'hyperlinks' && isSelfClosingTag(this.buffer, tagEnd)) {
      this.index = tagEnd + 1
      return true
    }
    if (isSelfClosingTag(this.buffer, tagEnd)) {
      const handled = this.retainMetadataXml && this.metadata.collectTypedMetadataElement(localName, this.buffer, this.index, tagEnd + 1)
      if (!handled) {
        this.metadata.countMetadataElement(localName, this.buffer, tagEnd + 1, tagEnd + 1)
      }
      if (this.retainMetadataXml && !handled) {
        this.failed = true
      }
      this.index = tagEnd + 1
      return true
    }
    if (localName === 'autoFilter') {
      if (this.retainMetadataXml && !this.sheetName) {
        this.failed = true
        return true
      }
      this.activeAutoFilter = {
        rootTag: this.buffer.slice(this.index, tagEnd + 1),
        filter:
          this.retainMetadataXml && this.sheetName
            ? readLargeSimpleAutoFilterRootFromBytes(this.sheetName, this.buffer, this.index, tagEnd + 1)
            : null,
      }
      this.index = tagEnd + 1
      if (!this.processActiveAutoFilter(final)) {
        if (final) {
          this.failed = true
        }
        return false
      }
      return true
    }
    if (localName === 'conditionalFormatting') {
      this.activeConditionalFormatting = {
        rootTag: this.buffer.slice(this.index, tagEnd + 1),
        ruleSeen: false,
      }
      this.index = tagEnd + 1
      if (!this.processActiveConditionalFormatting(final)) {
        if (final) {
          this.failed = true
        }
        return false
      }
      return true
    }
    if (localName === 'cols' || localName === 'mergeCells' || localName === 'tableParts') {
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
    if (localName === 'dataValidations') {
      this.activeDataValidations = true
      this.index = tagEnd + 1
      if (!this.processActiveDataValidations(final)) {
        if (final) {
          this.failed = true
        }
        return false
      }
      return true
    }
    if (localName === 'hyperlinks') {
      this.activeHyperlinks = true
      this.index = tagEnd + 1
      if (!this.processActiveHyperlinks(final)) {
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
    const handled = this.retainMetadataXml && this.metadata.collectTypedMetadataElement(localName, this.buffer, this.index, closing.end)
    if (!handled) {
      this.metadata.countMetadataElement(localName, this.buffer, tagEnd + 1, closing.start)
    }
    if (this.retainMetadataXml && !handled) {
      this.failed = true
    }
    this.index = closing.end
    return true
  }

  private processActiveAutoFilter(final: boolean): boolean {
    const active = this.activeAutoFilter
    if (active === null) {
      return true
    }
    while (!this.failed && this.index < this.buffer.byteLength) {
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
      if (closing && tag.localName === 'autoFilter') {
        this.metadata.addAutoFilter(active.filter)
        this.activeAutoFilter = null
        this.index = tagEnd + 1
        return true
      }
      if (!closing && tag.localName === 'filterColumn') {
        if (!this.processActiveAutoFilterColumn(active, tagEnd, final)) {
          return false
        }
        continue
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

  private processActiveAutoFilterColumn(active: ActiveAutoFilter, tagEnd: number, final: boolean): boolean {
    const startIndex = this.index
    const selfClosing = isSelfClosingTag(this.buffer, tagEnd)
    const contentStart = tagEnd + 1
    const closing = selfClosing ? { start: contentStart, end: contentStart } : findClosingTag(this.buffer, contentStart, 'filterColumn')
    if (!closing) {
      if (final) {
        this.failed = true
      }
      this.index = startIndex
      return false
    }
    const endIndex = selfClosing ? tagEnd + 1 : closing.end
    if (this.retainMetadataXml && this.sheetName) {
      const wrapped = wrapLargeSimpleAutoFilterColumnBytes(active.rootTag, this.buffer, startIndex, endIndex)
      const parsed = readLargeSimpleAutoFiltersFromBytes(this.sheetName, wrapped, 0, wrapped.byteLength)[0]
      if (parsed) {
        active.filter = mergeAutoFilterCriteria(active.filter ?? parsed, parsed.criteria)
      }
    }
    this.index = endIndex
    return true
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
        this.metadata.collectActiveMetadataTag(activeElement, tag.localName, this.buffer, this.index, tag.endIndex, tagEnd)
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

  private processActiveConditionalFormatting(final: boolean): boolean {
    const active = this.activeConditionalFormatting
    if (active === null) {
      return true
    }
    while (!this.failed && this.index < this.buffer.byteLength) {
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
      if (closing && tag.localName === 'conditionalFormatting') {
        if (!active.ruleSeen) {
          this.metadata.conditionalFormatCount += 1
        }
        this.activeConditionalFormatting = null
        this.index = tagEnd + 1
        return true
      }
      if (!closing && tag.localName === 'cfRule') {
        if (!this.processActiveConditionalFormatRule(active, tagEnd, final)) {
          return false
        }
        continue
      }
      if (!closing && this.retainMetadataXml) {
        this.failed = true
        return true
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

  private processActiveConditionalFormatRule(active: ActiveConditionalFormatting, tagEnd: number, final: boolean): boolean {
    const startIndex = this.index
    const selfClosing = isSelfClosingTag(this.buffer, tagEnd)
    const contentStart = tagEnd + 1
    const closing = selfClosing ? { start: contentStart, end: contentStart } : findClosingTag(this.buffer, contentStart, 'cfRule')
    if (!closing) {
      if (final) {
        this.failed = true
      }
      this.index = startIndex
      return false
    }
    active.ruleSeen = true
    const endIndex = selfClosing ? tagEnd + 1 : closing.end
    if (!this.retainMetadataXml) {
      this.metadata.countConditionalFormatRootRule(active.rootTag)
      this.index = endIndex
      return true
    }
    if (!this.metadata.collectConditionalFormattingRule(active.rootTag, this.buffer, startIndex, endIndex)) {
      this.failed = true
      return true
    }
    this.index = endIndex
    return true
  }

  private processActiveHyperlinks(final: boolean): boolean {
    while (!this.failed && this.index < this.buffer.byteLength) {
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
      if (closing && tag.localName === 'hyperlinks') {
        this.activeHyperlinks = false
        this.index = tagEnd + 1
        return true
      }
      if (!closing && tag.localName === 'hyperlink') {
        if (!this.metadata.collectActiveHyperlinkTag(this.buffer, this.index, tagEnd)) {
          this.failed = true
          return true
        }
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

  private processActiveDataValidations(final: boolean): boolean {
    while (!this.failed && this.index < this.buffer.byteLength) {
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
      if (closing && tag.localName === 'dataValidations') {
        this.activeDataValidations = false
        this.index = tagEnd + 1
        return true
      }
      if (!closing && tag.localName === 'dataValidation') {
        if (!this.processActiveDataValidationElement(tagEnd, final)) {
          return false
        }
        continue
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

  private processActiveDataValidationElement(tagEnd: number, final: boolean): boolean {
    const startIndex = this.index
    const selfClosing = isSelfClosingTag(this.buffer, tagEnd)
    const contentStart = tagEnd + 1
    const closing = selfClosing ? { start: contentStart, end: contentStart } : findClosingTag(this.buffer, contentStart, 'dataValidation')
    if (!closing) {
      if (final) {
        this.failed = true
      }
      this.index = startIndex
      return false
    }
    const endIndex = selfClosing ? tagEnd + 1 : closing.end
    if (!this.metadata.collectDataValidationElement(this.buffer, startIndex, endIndex)) {
      this.failed = true
      return true
    }
    this.index = endIndex
    return true
  }

  private reportRetainedBufferLength: () => void = () => {}
}

function mergeAutoFilterCriteria(
  filter: WorkbookAutoFilterSnapshot,
  criteria: WorkbookAutoFilterSnapshot['criteria'],
): WorkbookAutoFilterSnapshot {
  return criteria && criteria.length > 0 ? { ...filter, criteria: [...(filter.criteria ?? []), ...criteria] } : filter
}
