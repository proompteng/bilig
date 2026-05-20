import { Buffer } from 'node:buffer'
import { TextDecoder } from 'node:util'
import { createInflateRaw, inflateRawSync } from 'node:zlib'

import {
  forEachInflatedXlsxZipEntryChunk,
  getZipText,
  readXlsxZipEntriesLazyFromByteSource,
  readXlsxZipEntryMetadata,
  releaseLazyXlsxZipSource,
  type XlsxZipByteSource,
  type XlsxZipEntries,
  type XlsxZipEntryMetadata,
} from '../packages/excel-import/src/xlsx-zip.js'
import type { WorkbookExternalWorkbookReferenceSnapshot } from '../packages/protocol/src/types.js'
import { WorksheetDataValidationSupportScanner } from './public-workbook-corpus-xlsx-data-validation-footprint.ts'
import type { WorkbookFootprint } from './public-workbook-corpus-workbook.ts'

interface ZipEntryInfo {
  readonly path: string
  readonly compressionMethod: number
  readonly compressedSize: number
  readonly localHeaderOffset: number
}

interface ParsedRelationship {
  readonly id: string
  readonly target: string
  readonly type: string
  readonly targetMode?: string
}

interface ParsedSheet {
  readonly name: string
  readonly path: string | null
}

interface WorkbookPackageEntry {
  readonly path: string
  readonly compressionMethod: number
  readonly compressedSize: number
}

interface WorksheetFootprint {
  cellCount: number
  columnCount: number
  conditionalFormatCount: number
  dataValidationCount: number
  formulaCellCount: number
  largeSimpleUnsupportedElementCount: number
  largeSimpleUnsupportedDataValidationCount: number
  largeSimpleUnsupportedFormulaCellCount: number
  mergeCount: number
  rowCount: number
  sharedStringCellCount: number
  usedRange: WorkbookFootprint['workbookMetadata']['dimensions'][number]['usedRange']
  valueCellCount: number
  xmlCellCount: number
}

const decoder = new TextDecoder()
const eocdSignature = 0x06054b50
const centralDirectorySignature = 0x02014b50
const localFileHeaderSignature = 0x04034b50
const lessThanByte = 0x3c
const greaterThanByte = 0x3e
const slashByte = 0x2f
const equalsByte = 0x3d
const doubleQuoteByte = 0x22
const singleQuoteByte = 0x27
const whitespaceBytes = new Set([0x09, 0x0a, 0x0d, 0x20])

export function isZipWorkbook(bytes: Uint8Array): boolean {
  return bytes.length >= 4 && bytes[0] === 0x50 && bytes[1] === 0x4b && bytes[2] === 0x03 && bytes[3] === 0x04
}

export function inspectXlsxWorkbookFootprintLowMemory(bytes: Uint8Array, fileName: string): WorkbookFootprint {
  const context = readWorkbookPackageContext(bytes)
  const dimensions: WorkbookFootprint['workbookMetadata']['dimensions'] = []
  const totals = emptyWorksheetFootprint()

  for (const sheet of context.sheets) {
    const worksheetBytes = sheet.path ? readZipEntryBytes(bytes, context.entriesByPath.get(sheet.path)) : null
    const footprint = worksheetBytes ? inspectWorksheetXmlBytes(worksheetBytes) : emptyWorksheetFootprint()
    addWorksheetFootprint(totals, footprint)
    dimensions.push(worksheetDimension(sheet.name, footprint))
  }

  return buildWorkbookFootprint(bytes, fileName, context, totals, dimensions)
}

export async function inspectXlsxWorkbookFootprintLowMemoryAsync(bytes: Uint8Array, fileName: string): Promise<WorkbookFootprint> {
  const context = readWorkbookPackageContext(bytes)
  const dimensions: WorkbookFootprint['workbookMetadata']['dimensions'] = []
  const totals = emptyWorksheetFootprint()

  for (const sheet of context.sheets) {
    let footprint = emptyWorksheetFootprint()
    if (sheet.path) {
      // oxlint-disable-next-line eslint(no-await-in-loop) -- Sequential worksheet scanning keeps peak RSS bounded.
      footprint = await inspectWorksheetZipEntryLowMemory(bytes, context.entriesByPath.get(sheet.path))
    }
    addWorksheetFootprint(totals, footprint)
    dimensions.push(worksheetDimension(sheet.name, footprint))
  }

  return buildWorkbookFootprint(bytes, fileName, context, totals, dimensions)
}

export function inspectXlsxWorkbookFootprintLowMemoryFromByteSource(source: XlsxZipByteSource, fileName: string): WorkbookFootprint | null {
  const zip = readXlsxZipEntriesLazyFromByteSource(source)
  const entryMetadata = readXlsxZipEntryMetadata(source)
  if (!zip || !entryMetadata) {
    return null
  }
  try {
    const context = readWorkbookPackageContextFromZip(zip, entryMetadata)
    const dimensions: WorkbookFootprint['workbookMetadata']['dimensions'] = []
    const totals = emptyWorksheetFootprint()

    for (const sheet of context.sheets) {
      const footprint = sheet.path ? inspectWorksheetZipEntryFromLazyZip(zip, sheet.path) : emptyWorksheetFootprint()
      addWorksheetFootprint(totals, footprint)
      dimensions.push(worksheetDimension(sheet.name, footprint))
    }

    return buildWorkbookFootprintFromContext({
      fileName,
      context,
      totals,
      dimensions,
      readPartText: (path) => getZipText(zip, path),
    })
  } finally {
    releaseLazyXlsxZipSource(zip)
  }
}

function readWorkbookPackageContext(bytes: Uint8Array): {
  readonly entries: readonly WorkbookPackageEntry[]
  readonly entriesByPath: ReadonlyMap<string, WorkbookPackageEntry>
  readonly workbookRelationships: readonly ParsedRelationship[]
  readonly workbookXml: string
  readonly sheets: readonly ParsedSheet[]
} {
  const entries = readZipCentralDirectory(bytes)
  const entriesByPath = new Map(entries.map((entry) => [entry.path, entry]))
  const workbookXml = readZipEntryText(bytes, entriesByPath.get('xl/workbook.xml')) ?? ''
  const workbookRelationships = parseRelationships(readZipEntryText(bytes, entriesByPath.get('xl/_rels/workbook.xml.rels')))
  const sheets = readWorkbookSheets(workbookXml, workbookRelationships, entries)
  return { entries, entriesByPath, workbookRelationships, workbookXml, sheets }
}

function readWorkbookPackageContextFromZip(
  zip: XlsxZipEntries,
  entryMetadata: readonly XlsxZipEntryMetadata[],
): {
  readonly entries: readonly WorkbookPackageEntry[]
  readonly entriesByPath: ReadonlyMap<string, WorkbookPackageEntry>
  readonly workbookRelationships: readonly ParsedRelationship[]
  readonly workbookXml: string
  readonly sheets: readonly ParsedSheet[]
} {
  const entries = entryMetadata.map((entry) => ({
    path: entry.path,
    compressionMethod: entry.compressionMethod,
    compressedSize: entry.compressedSize,
  }))
  const entriesByPath = new Map(entries.map((entry) => [entry.path, entry]))
  const workbookXml = getZipText(zip, 'xl/workbook.xml') ?? ''
  const workbookRelationships = parseRelationships(getZipText(zip, 'xl/_rels/workbook.xml.rels'))
  const sheets = readWorkbookSheets(workbookXml, workbookRelationships, entries)
  return { entries, entriesByPath, workbookRelationships, workbookXml, sheets }
}

function buildWorkbookFootprint(
  bytes: Uint8Array,
  fileName: string,
  context: {
    readonly entries: readonly ZipEntryInfo[]
    readonly entriesByPath: ReadonlyMap<string, ZipEntryInfo>
    readonly workbookRelationships: readonly ParsedRelationship[]
    readonly workbookXml: string
    readonly sheets: readonly ParsedSheet[]
  },
  totals: WorksheetFootprint,
  dimensions: WorkbookFootprint['workbookMetadata']['dimensions'],
): WorkbookFootprint {
  return buildWorkbookFootprintFromContext({
    fileName,
    context,
    totals,
    dimensions,
    readPartText: (path) => readZipEntryText(bytes, context.entriesByPath.get(path)),
  })
}

function buildWorkbookFootprintFromContext(args: {
  readonly fileName: string
  readonly context: {
    readonly entries: readonly WorkbookPackageEntry[]
    readonly entriesByPath: ReadonlyMap<string, WorkbookPackageEntry>
    readonly workbookRelationships: readonly ParsedRelationship[]
    readonly workbookXml: string
    readonly sheets: readonly ParsedSheet[]
  }
  readonly totals: WorksheetFootprint
  readonly dimensions: WorkbookFootprint['workbookMetadata']['dimensions']
  readonly readPartText: (path: string) => string | null
}): WorkbookFootprint {
  return {
    featureCounts: {
      sheetCount: args.context.sheets.length,
      cellCount: args.totals.cellCount,
      formulaCellCount: args.totals.formulaCellCount,
      valueCellCount: args.totals.valueCellCount,
      definedNameCount: countElementStarts(args.context.workbookXml, 'definedName'),
      tableCount: countZipEntries(args.context.entries, /^xl\/tables\/table[1-9][0-9]*\.xml$/u),
      chartCount: countZipEntries(args.context.entries, /^xl\/charts\/chart[1-9][0-9]*\.xml$/u),
      pivotCount: countZipEntries(args.context.entries, /^xl\/pivotTables\/pivotTable[1-9][0-9]*\.xml$/u),
      mergeCount: args.totals.mergeCount,
      styleRangeCount: 0,
      conditionalFormatCount: args.totals.conditionalFormatCount,
      dataValidationCount: args.totals.dataValidationCount,
      macroPayloadCount: args.context.entriesByPath.has('xl/vbaProject.bin') ? 1 : 0,
      warningCount: 0,
    },
    workbookMetadata: {
      workbookName: args.fileName.replace(/\.(xlsx|xlsm|csv)$/iu, '') || args.fileName,
      sheetNames: args.context.sheets.map((sheet) => sheet.name),
      dimensions: args.dimensions,
    },
    externalWorkbookReferences: readExternalWorkbookReferences(
      args.readPartText,
      args.context.entries,
      args.context.workbookXml,
      args.context.workbookRelationships,
    ),
    largeSimpleXlsxImport: inspectLargeSimpleXlsxImportCompatibility(args.context, args.totals),
  }
}

function inspectLargeSimpleXlsxImportCompatibility(
  context: {
    readonly entries: readonly WorkbookPackageEntry[]
    readonly entriesByPath: ReadonlyMap<string, WorkbookPackageEntry>
    readonly workbookXml: string
  },
  totals: WorksheetFootprint,
): WorkbookFootprint['largeSimpleXlsxImport'] {
  const blockers: string[] = []
  const unsupportedPackagePaths = context.entries.filter((entry) =>
    /^xl\/(?:charts|chartSheets|comments|ctrlProps|externalLinks|metadata\.xml|model|pivotCache|pivotTables|threadedComments|vbaProject\.bin)/u.test(
      entry.path,
    ),
  )
  if (unsupportedPackagePaths.length > 0) {
    blockers.push(`unsupported-package-parts=${String(unsupportedPackagePaths.length)}`)
  }
  const sharedStringsEntry = context.entriesByPath.get('xl/sharedStrings.xml')
  const hasSharedStrings = sharedStringsEntry !== undefined && sharedStringsEntry.compressedSize > 0
  if (definedNamesReferenceExternalWorkbook(context.workbookXml)) {
    blockers.push('external-defined-name-reference')
  }
  if (totals.largeSimpleUnsupportedDataValidationCount > 0) {
    blockers.push(`unsupported-data-validations=${String(totals.largeSimpleUnsupportedDataValidationCount)}`)
  }
  if (totals.largeSimpleUnsupportedFormulaCellCount > 0) {
    blockers.push(`unsupported-formula-cells=${String(totals.largeSimpleUnsupportedFormulaCellCount)}`)
  }
  if (totals.largeSimpleUnsupportedElementCount > 0) {
    blockers.push(`unsupported-worksheet-elements=${String(totals.largeSimpleUnsupportedElementCount)}`)
  }
  if (!hasSharedStrings) {
    if (totals.sharedStringCellCount > 0) {
      blockers.push(`missing-shared-strings=${String(totals.sharedStringCellCount)}`)
    }
    if (totals.valueCellCount === 0) {
      blockers.push('no-value-cells')
    }
    const blankCellCount = Math.max(0, totals.xmlCellCount - totals.valueCellCount)
    if (blankCellCount > 50_000 && blankCellCount > totals.valueCellCount * 8) {
      blockers.push(`style-only-blank-cells=${String(blankCellCount)}`)
    }
  }
  return {
    eligible: blockers.length === 0,
    blockers,
  }
}

function definedNamesReferenceExternalWorkbook(workbookXml: string): boolean {
  return /<(?:[A-Za-z_][\w.-]*:)?definedName\b(?:[^>"']|"[^"]*"|'[^']*')*>[\s\S]*?(?:^|[=,+(*/\s])'?\[[^\]]+\][\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?definedName>/u.test(
    workbookXml,
  )
}

function addWorksheetFootprint(total: WorksheetFootprint, footprint: WorksheetFootprint): void {
  total.cellCount += footprint.cellCount
  total.formulaCellCount += footprint.formulaCellCount
  total.largeSimpleUnsupportedFormulaCellCount += footprint.largeSimpleUnsupportedFormulaCellCount
  total.valueCellCount += footprint.valueCellCount
  total.xmlCellCount += footprint.xmlCellCount
  total.sharedStringCellCount += footprint.sharedStringCellCount
  total.mergeCount += footprint.mergeCount
  total.conditionalFormatCount += footprint.conditionalFormatCount
  total.dataValidationCount += footprint.dataValidationCount
  total.largeSimpleUnsupportedElementCount += footprint.largeSimpleUnsupportedElementCount
  total.largeSimpleUnsupportedDataValidationCount += footprint.largeSimpleUnsupportedDataValidationCount
}

function worksheetDimension(sheetName: string, footprint: WorksheetFootprint): WorkbookFootprint['workbookMetadata']['dimensions'][number] {
  return {
    sheetName,
    rowCount: footprint.rowCount,
    columnCount: footprint.columnCount,
    nonEmptyCellCount: footprint.cellCount,
    usedRange: footprint.usedRange,
  }
}

function readZipCentralDirectory(bytes: Uint8Array): ZipEntryInfo[] {
  const eocdOffset = findEndOfCentralDirectory(bytes)
  const centralDirectorySize = readUint32(bytes, eocdOffset + 12)
  const centralDirectoryOffset = readUint32(bytes, eocdOffset + 16)
  if (centralDirectorySize === 0xffffffff || centralDirectoryOffset === 0xffffffff) {
    throw new Error('Zip64 central directories are not supported by the low-memory XLSX footprint scanner')
  }
  const entries: ZipEntryInfo[] = []
  let offset = centralDirectoryOffset
  const endOffset = centralDirectoryOffset + centralDirectorySize
  while (offset < endOffset) {
    if (readUint32(bytes, offset) !== centralDirectorySignature) {
      throw new Error('Invalid XLSX central directory entry')
    }
    const compressionMethod = readUint16(bytes, offset + 10)
    const compressedSize = readUint32(bytes, offset + 20)
    const fileNameLength = readUint16(bytes, offset + 28)
    const extraFieldLength = readUint16(bytes, offset + 30)
    const fileCommentLength = readUint16(bytes, offset + 32)
    const localHeaderOffset = readUint32(bytes, offset + 42)
    const path = normalizeZipPath(decodeBytes(bytes.subarray(offset + 46, offset + 46 + fileNameLength)))
    entries.push({ path, compressionMethod, compressedSize, localHeaderOffset })
    offset += 46 + fileNameLength + extraFieldLength + fileCommentLength
  }
  return entries
}

function findEndOfCentralDirectory(bytes: Uint8Array): number {
  const minOffset = Math.max(0, bytes.length - 65_557)
  for (let offset = bytes.length - 22; offset >= minOffset; offset -= 1) {
    if (readUint32(bytes, offset) === eocdSignature) {
      return offset
    }
  }
  throw new Error('Invalid XLSX zip: end of central directory not found')
}

function readZipEntryBytes(bytes: Uint8Array, entry: ZipEntryInfo | undefined): Uint8Array | null {
  const payload = readZipEntryPayload(bytes, entry)
  if (!payload) {
    return null
  }
  if (payload.compressionMethod === 0) {
    return payload.compressed
  }
  if (payload.compressionMethod === 8) {
    return inflateRawSync(payload.compressed)
  }
  throw new Error(`Unsupported XLSX zip compression method ${String(payload.compressionMethod)} for ${payload.path}`)
}

function readZipEntryPayload(
  bytes: Uint8Array,
  entry: ZipEntryInfo | undefined,
): { readonly compressed: Uint8Array; readonly compressionMethod: number; readonly path: string } | null {
  if (!entry) {
    return null
  }
  const offset = entry.localHeaderOffset
  if (readUint32(bytes, offset) !== localFileHeaderSignature) {
    throw new Error(`Invalid local file header for ${entry.path}`)
  }
  const fileNameLength = readUint16(bytes, offset + 26)
  const extraFieldLength = readUint16(bytes, offset + 28)
  const dataOffset = offset + 30 + fileNameLength + extraFieldLength
  const compressed = bytes.subarray(dataOffset, dataOffset + entry.compressedSize)
  return { compressed, compressionMethod: entry.compressionMethod, path: entry.path }
}

function readZipEntryText(bytes: Uint8Array, entry: ZipEntryInfo | undefined): string | null {
  const entryBytes = readZipEntryBytes(bytes, entry)
  return entryBytes ? decodeBytes(entryBytes) : null
}

function readWorkbookSheets(
  workbookXml: string,
  workbookRelationships: readonly ParsedRelationship[],
  entries: readonly WorkbookPackageEntry[],
): ParsedSheet[] {
  const worksheetRelationships = new Map(
    workbookRelationships
      .filter((relationship) => relationship.type.endsWith('/worksheet') || relationship.target.includes('worksheets/'))
      .map((relationship) => [relationship.id, resolveTargetPath('xl/workbook.xml', relationship.target)]),
  )
  const sheets = elementTags(workbookXml, 'sheet').map((tag, index) => {
    const name = readXmlAttribute(tag, 'name') ?? `Sheet${String(index + 1)}`
    const relationshipId = readXmlAttribute(tag, 'r:id') ?? readXmlAttribute(tag, 'id')
    return {
      name,
      path: relationshipId ? (worksheetRelationships.get(relationshipId) ?? null) : null,
    }
  })
  if (sheets.length > 0) {
    return sheets
  }
  return entries
    .map((entry) => /^xl\/worksheets\/sheet([1-9][0-9]*)\.xml$/u.exec(entry.path)?.[1])
    .flatMap((sheetNumber) => (sheetNumber ? [{ name: `Sheet${sheetNumber}`, path: `xl/worksheets/sheet${sheetNumber}.xml` }] : []))
    .toSorted((left, right) => left.name.localeCompare(right.name, 'en-US', { numeric: true }))
}

function inspectWorksheetXmlBytes(xml: Uint8Array): WorksheetFootprint {
  const scanner = new WorksheetXmlByteFootprintScanner()
  scanner.push(xml)
  return scanner.finish()
}

async function inspectWorksheetZipEntryLowMemory(bytes: Uint8Array, entry: ZipEntryInfo | undefined): Promise<WorksheetFootprint> {
  const payload = readZipEntryPayload(bytes, entry)
  if (!payload) {
    return emptyWorksheetFootprint()
  }
  if (payload.compressionMethod === 0) {
    return inspectWorksheetXmlBytes(payload.compressed)
  }
  if (payload.compressionMethod === 8) {
    return inspectDeflatedWorksheetXmlBytes(payload.compressed)
  }
  throw new Error(`Unsupported XLSX zip compression method ${String(payload.compressionMethod)} for ${payload.path}`)
}

function inspectWorksheetZipEntryFromLazyZip(zip: XlsxZipEntries, path: string): WorksheetFootprint {
  const scanner = new WorksheetXmlByteFootprintScanner()
  const found = forEachInflatedXlsxZipEntryChunk(zip, path, (chunk) => scanner.push(chunk))
  return found ? scanner.finish() : emptyWorksheetFootprint()
}

function inspectDeflatedWorksheetXmlBytes(compressed: Uint8Array): Promise<WorksheetFootprint> {
  const scanner = new WorksheetXmlByteFootprintScanner()
  const inflate = createInflateRaw()
  return new Promise<WorksheetFootprint>((resolvePromise, reject) => {
    inflate.on('data', (chunk: Uint8Array) => scanner.push(chunk))
    inflate.on('error', reject)
    inflate.on('end', () => {
      try {
        resolvePromise(scanner.finish())
      } catch (error) {
        reject(error)
      }
    })
    inflate.end(Buffer.from(compressed.buffer, compressed.byteOffset, compressed.byteLength))
  })
}

function emptyWorksheetFootprint(): WorksheetFootprint {
  return {
    cellCount: 0,
    formulaCellCount: 0,
    valueCellCount: 0,
    rowCount: 0,
    columnCount: 0,
    usedRange: null,
    mergeCount: 0,
    conditionalFormatCount: 0,
    dataValidationCount: 0,
    xmlCellCount: 0,
    sharedStringCellCount: 0,
    largeSimpleUnsupportedElementCount: 0,
    largeSimpleUnsupportedDataValidationCount: 0,
    largeSimpleUnsupportedFormulaCellCount: 0,
  }
}

class WorksheetXmlByteFootprintScanner {
  private buffer = new Uint8Array()
  private readonly sharedFormulaIndexes = new Set<string>()
  private readonly counter = new WorksheetElementStartCounter()
  private readonly dataValidationScanner = new WorksheetDataValidationSupportScanner()
  private readonly footprint = emptyWorksheetFootprint()

  push(chunk: Uint8Array): void {
    this.counter.push(chunk, false)
    this.dataValidationScanner.push(chunk, false)
    this.buffer = concatBytes(this.buffer, chunk)
    this.scanBufferedCells(false)
  }

  finish(): WorksheetFootprint {
    this.counter.finish()
    const dataValidationSupport = this.dataValidationScanner.finish()
    this.scanBufferedCells(true)
    const counted = this.counter.counts()
    return {
      ...this.footprint,
      mergeCount: counted.mergeCount,
      conditionalFormatCount: counted.conditionalFormatCount,
      dataValidationCount: dataValidationSupport.dataValidationCount,
      largeSimpleUnsupportedDataValidationCount: dataValidationSupport.unsupportedDataValidationCount,
      largeSimpleUnsupportedElementCount: counted.largeSimpleUnsupportedElementCount,
    }
  }

  private scanBufferedCells(final: boolean): void {
    let index = 0
    while (index < this.buffer.length) {
      const tagStart = indexOfElementCandidate(this.buffer, cElementNameBytes, index)
      if (tagStart < 0) {
        this.retainFrom(final ? this.buffer.length : Math.max(index, this.buffer.length - 1))
        return
      }
      if (!isElementStartBytes(this.buffer, tagStart, cElementNameBytes)) {
        index = tagStart + 2
        continue
      }
      const openingEnd = indexOfByte(this.buffer, greaterThanByte, tagStart + 2)
      if (openingEnd < 0) {
        if (final) {
          this.retainFrom(this.buffer.length)
          return
        }
        this.retainFrom(tagStart)
        return
      }
      const selfClosing = this.buffer[openingEnd - 1] === slashByte
      const closeStart = selfClosing ? openingEnd : indexOfBytes(this.buffer, closeCellElementBytes, openingEnd + 1, this.buffer.length)
      if (closeStart < 0 && !final) {
        this.retainFrom(tagStart)
        return
      }
      const effectiveCloseStart = closeStart < 0 ? openingEnd : closeStart
      const cellEnd = selfClosing ? openingEnd + 1 : closeStart < 0 ? openingEnd : closeStart + closeCellElementBytes.length
      this.scanCell(tagStart, openingEnd, effectiveCloseStart, selfClosing)
      index = Math.max(cellEnd, openingEnd + 1)
    }
    this.retainFrom(this.buffer.length)
  }

  private scanCell(tagStart: number, openingEnd: number, closeStart: number, selfClosing: boolean): void {
    let hasFormula = false
    let hasValue = false
    this.footprint.xmlCellCount += 1
    if (readXmlAttributeInBytesRange(this.buffer, tagStart, openingEnd + 1, tAttributeNameBytes) === 's') {
      this.footprint.sharedStringCellCount += 1
    }
    if (!selfClosing && closeStart > openingEnd) {
      hasFormula = containsElementStartBytes(this.buffer, openingEnd + 1, closeStart, fElementNameBytes)
      hasValue =
        containsElementStartBytes(this.buffer, openingEnd + 1, closeStart, vElementNameBytes) ||
        containsElementStartBytes(this.buffer, openingEnd + 1, closeStart, isElementNameBytes)
    }
    if (!hasFormula && !hasValue) {
      return
    }
    this.footprint.cellCount += 1
    if (hasFormula) {
      this.footprint.formulaCellCount += 1
      if (isUnsupportedLargeSimpleFormulaCell(this.buffer, openingEnd + 1, closeStart, this.sharedFormulaIndexes)) {
        this.footprint.largeSimpleUnsupportedFormulaCellCount += 1
      }
    }
    if (hasValue) {
      this.footprint.valueCellCount += 1
    }
    const address = readXmlAttributeInBytesRange(this.buffer, tagStart, openingEnd + 1, rAttributeNameBytes)
    const decoded = address ? decodeCellAddress(address) : null
    if (decoded) {
      this.footprint.rowCount = Math.max(this.footprint.rowCount, decoded.row + 1)
      this.footprint.columnCount = Math.max(this.footprint.columnCount, decoded.column + 1)
      this.footprint.usedRange = expandUsedRange(this.footprint.usedRange, decoded.row, decoded.column)
    }
  }

  private retainFrom(index: number): void {
    this.buffer = copyBytes(this.buffer.subarray(index))
  }
}

class WorksheetElementStartCounter {
  private tail = new Uint8Array()
  private mergeCellCount = 0
  private conditionalFormattingCount = 0
  private largeSimpleUnsupportedElementCount = 0

  push(chunk: Uint8Array, final: boolean): void {
    const buffer = concatBytes(this.tail, chunk)
    const scanEnd = final ? buffer.length : Math.max(0, buffer.length - countedElementTailLength)
    this.mergeCellCount += countElementStartsBytes(buffer, mergeCellElementNameBytes, 0, scanEnd)
    this.conditionalFormattingCount += countElementStartsBytes(buffer, conditionalFormattingElementNameBytes, 0, scanEnd)
    for (const elementName of unsupportedLargeSimpleElementNameBytes) {
      this.largeSimpleUnsupportedElementCount += countElementStartsBytes(buffer, elementName, 0, scanEnd)
    }
    this.tail = copyBytes(buffer.subarray(scanEnd))
  }

  finish(): void {
    this.push(new Uint8Array(), true)
  }

  counts(): Pick<WorksheetFootprint, 'conditionalFormatCount' | 'largeSimpleUnsupportedElementCount' | 'mergeCount'> {
    return {
      mergeCount: this.mergeCellCount,
      conditionalFormatCount: this.conditionalFormattingCount,
      largeSimpleUnsupportedElementCount: this.largeSimpleUnsupportedElementCount,
    }
  }
}

function readExternalWorkbookReferences(
  readPartText: (path: string) => string | null,
  entries: readonly WorkbookPackageEntry[],
  workbookXml: string,
  workbookRelationships: readonly ParsedRelationship[],
): readonly WorkbookExternalWorkbookReferenceSnapshot[] {
  const workbookTargets = readWorkbookExternalLinkTargets(workbookXml, workbookRelationships)
  const linkTargets = workbookTargets.size > 0 ? workbookTargets : readFallbackExternalLinkTargets(entries)
  return [...linkTargets.entries()]
    .toSorted((left, right) => left[0] - right[0])
    .flatMap(([bookIndex, path]) => {
      const xml = readPartText(path)
      if (!xml) {
        return []
      }
      const relationships = parseRelationships(readPartText(externalLinkRelationshipsPartPath(path)))
      const externalBookRelationshipId =
        readXmlAttribute(elementTags(xml, 'externalBook')[0] ?? '', 'r:id') ??
        readXmlAttribute(elementTags(xml, 'externalBook')[0] ?? '', 'id')
      const linkedWorkbookRelationship =
        (externalBookRelationshipId ? relationships.find((relationship) => relationship.id === externalBookRelationshipId) : undefined) ??
        relationships.find((relationship) => relationship.type.endsWith('/externalLinkPath'))
      const target = linkedWorkbookRelationship?.target
      const workbookName = target ? workbookNameFromExternalTarget(target) : undefined
      const sheetNames = elementTags(xml, 'sheetName').flatMap((tag) => {
        const name = readXmlAttribute(tag, 'val')
        return name ? [name] : []
      })
      const reference: WorkbookExternalWorkbookReferenceSnapshot = {
        bookIndex,
        packagePath: path,
        ...(target ? { target } : {}),
        ...(linkedWorkbookRelationship?.targetMode ? { targetMode: linkedWorkbookRelationship.targetMode } : {}),
        ...(workbookName ? { workbookName } : {}),
        ...(sheetNames.length > 0 ? { sheetNames } : {}),
      }
      return [reference]
    })
}

function readWorkbookExternalLinkTargets(workbookXml: string, workbookRelationships: readonly ParsedRelationship[]): Map<number, string> {
  const targets = new Map<number, string>()
  let bookIndex = 1
  for (const tag of elementTags(workbookXml, 'externalReference')) {
    const relationshipId = readXmlAttribute(tag, 'r:id') ?? readXmlAttribute(tag, 'id')
    const relationship = relationshipId
      ? workbookRelationships.find((candidate) => candidate.id === relationshipId && candidate.type.endsWith('/externalLink'))
      : undefined
    if (relationship) {
      targets.set(bookIndex, resolveTargetPath('xl/workbook.xml', relationship.target))
    }
    bookIndex += 1
  }
  return targets
}

function readFallbackExternalLinkTargets(entries: readonly ZipEntryInfo[]): Map<number, string> {
  const targets = new Map<number, string>()
  for (const entry of entries) {
    const match = /^xl\/externalLinks\/externalLink([1-9][0-9]*)\.xml$/u.exec(entry.path)
    if (match) {
      targets.set(Number(match[1]), entry.path)
    }
  }
  return targets
}

function parseRelationships(xml: string | null): ParsedRelationship[] {
  if (!xml) {
    return []
  }
  return elementTags(xml, 'Relationship').flatMap((tag) => {
    const id = readXmlAttribute(tag, 'Id')
    const target = readXmlAttribute(tag, 'Target')
    const type = readXmlAttribute(tag, 'Type')
    if (!id || !target || !type) {
      return []
    }
    const targetMode = readXmlAttribute(tag, 'TargetMode')
    return [{ id, target, type, ...(targetMode ? { targetMode } : {}) }]
  })
}

function externalLinkRelationshipsPartPath(partPath: string): string {
  const fileName = partPath.slice(partPath.lastIndexOf('/') + 1)
  return `xl/externalLinks/_rels/${fileName}.rels`
}

function resolveTargetPath(basePartPath: string, target: string): string {
  if (target.startsWith('/')) {
    return normalizeZipPath(target)
  }
  const parts = basePartPath.split('/')
  parts.pop()
  for (const segment of target.split('/')) {
    if (segment === '..') {
      parts.pop()
    } else if (segment !== '.' && segment.length > 0) {
      parts.push(segment)
    }
  }
  return normalizeZipPath(parts.join('/'))
}

function workbookNameFromExternalTarget(target: string): string | undefined {
  const targetPath = target.split(/[?#]/u)[0] ?? target
  const normalizedTarget = targetPath.replace(/^file:\/+/iu, '').replace(/\\/gu, '/')
  const lastSegment = normalizedTarget
    .split('/')
    .toReversed()
    .find((segment) => segment.length > 0)
  if (!lastSegment) {
    return undefined
  }
  try {
    return decodeURIComponent(lastSegment)
  } catch {
    return lastSegment
  }
}

function countZipEntries(entries: readonly WorkbookPackageEntry[], pattern: RegExp): number {
  return entries.filter((entry) => pattern.test(entry.path)).length
}

function elementTags(xml: string, elementName: string): string[] {
  const tags: string[] = []
  let index = 0
  while (index < xml.length) {
    const tagStart = xml.indexOf(`<${elementName}`, index)
    if (tagStart < 0) {
      break
    }
    if (!isElementStart(xml, tagStart, elementName)) {
      index = tagStart + elementName.length + 1
      continue
    }
    const tagEnd = xml.indexOf('>', tagStart + elementName.length + 1)
    if (tagEnd < 0) {
      break
    }
    tags.push(xml.slice(tagStart, tagEnd + 1))
    index = tagEnd + 1
  }
  return tags
}

function countElementStarts(xml: string, elementName: string): number {
  let count = 0
  let index = 0
  while (index < xml.length) {
    const tagStart = xml.indexOf(`<${elementName}`, index)
    if (tagStart < 0) {
      return count
    }
    if (isElementStart(xml, tagStart, elementName)) {
      count += 1
    }
    index = tagStart + elementName.length + 1
  }
  return count
}

function isElementStart(xml: string, index: number, elementName: string): boolean {
  if (!xml.startsWith(`<${elementName}`, index)) {
    return false
  }
  const next = xml[index + elementName.length + 1]
  return next === undefined || /\s|\/|>/u.test(next)
}

function readXmlAttribute(tag: string, name: string): string | null {
  return readXmlAttributeInRange(tag, 0, tag.length, name)
}

function readXmlAttributeInRange(source: string, start: number, end: number, name: string): string | null {
  const pattern = new RegExp(`\\s${escapeRegExp(name)}=(["'])([\\s\\S]*?)\\1`, 'u')
  const match = pattern.exec(source.slice(start, end))
  return match?.[2] ? decodeXmlEntities(match[2]) : null
}

function decodeCellAddress(address: string): { readonly row: number; readonly column: number } | null {
  const match = /^\$?([A-Z]{1,3})\$?([1-9][0-9]*)$/iu.exec(address)
  if (!match) {
    return null
  }
  let column = 0
  for (const character of match[1].toUpperCase()) {
    column = column * 26 + character.charCodeAt(0) - 64
  }
  const row = Number(match[2])
  return Number.isSafeInteger(row) ? { row: row - 1, column: column - 1 } : null
}

function expandUsedRange(current: WorksheetFootprint['usedRange'], row: number, column: number): WorksheetFootprint['usedRange'] {
  return current
    ? {
        startRow: Math.min(current.startRow, row),
        startColumn: Math.min(current.startColumn, column),
        endRow: Math.max(current.endRow, row),
        endColumn: Math.max(current.endColumn, column),
      }
    : { startRow: row, startColumn: column, endRow: row, endColumn: column }
}

function decodeXmlEntities(value: string): string {
  return value.replace(/&(#x?[0-9a-f]+|amp|apos|gt|lt|quot);/giu, (entity, raw: string) => {
    switch (raw) {
      case 'amp':
        return '&'
      case 'apos':
        return "'"
      case 'gt':
        return '>'
      case 'lt':
        return '<'
      case 'quot':
        return '"'
      default: {
        const radix = raw.toLowerCase().startsWith('#x') ? 16 : 10
        const digits = raw.replace(/^#x?/iu, '')
        const codePoint = Number.parseInt(digits, radix)
        return Number.isFinite(codePoint) ? String.fromCodePoint(codePoint) : entity
      }
    }
  })
}

function normalizeZipPath(path: string): string {
  return path.replace(/^\/+/, '')
}

function decodeBytes(bytes: Uint8Array): string {
  return decoder.decode(bytes)
}

const cElementNameBytes = asciiBytes('c')
const fElementNameBytes = asciiBytes('f')
const vElementNameBytes = asciiBytes('v')
const isElementNameBytes = asciiBytes('is')
const rAttributeNameBytes = asciiBytes('r')
const siAttributeNameBytes = asciiBytes('si')
const tAttributeNameBytes = asciiBytes('t')
const closeCellElementBytes = asciiBytes('</c>')
const closeFormulaElementBytes = asciiBytes('</f>')
const mergeCellElementNameBytes = asciiBytes('mergeCell')
const conditionalFormattingElementNameBytes = asciiBytes('conditionalFormatting')
const unsupportedLargeSimpleElementNameBytes = [asciiBytes('picture'), asciiBytes('sheetProtection')]
const countedElementTailLength =
  Math.max(
    mergeCellElementNameBytes.length,
    conditionalFormattingElementNameBytes.length,
    ...unsupportedLargeSimpleElementNameBytes.map((entry) => entry.length),
  ) + 2

function countElementStartsBytes(xml: Uint8Array, elementName: Uint8Array, start = 0, end = xml.length): number {
  let count = 0
  let index = start
  while (index < end) {
    const tagStart = indexOfElementCandidate(xml, elementName, index)
    if (tagStart < 0 || tagStart >= end) {
      return count
    }
    if (isElementStartBytes(xml, tagStart, elementName)) {
      count += 1
    }
    index = tagStart + elementName.length + 1
  }
  return count
}

function containsElementStartBytes(xml: Uint8Array, start: number, end: number, elementName: Uint8Array): boolean {
  let index = start
  while (index < end) {
    const tagStart = indexOfElementCandidate(xml, elementName, index, end)
    if (tagStart < 0) {
      return false
    }
    if (isElementStartBytes(xml, tagStart, elementName)) {
      return true
    }
    index = tagStart + elementName.length + 1
  }
  return false
}

function isUnsupportedLargeSimpleFormulaCell(xml: Uint8Array, start: number, end: number, sharedFormulaIndexes: Set<string>): boolean {
  const formulaStart = indexOfElementCandidate(xml, fElementNameBytes, start, end)
  if (formulaStart < 0 || !isElementStartBytes(xml, formulaStart, fElementNameBytes)) {
    return false
  }
  const openingEnd = indexOfByte(xml, greaterThanByte, formulaStart, end)
  if (openingEnd < 0) {
    return true
  }
  const formulaType = readXmlAttributeInBytesRange(xml, formulaStart, openingEnd + 1, tAttributeNameBytes)
  const sharedFormulaIndex =
    formulaType === 'shared' ? readXmlAttributeInBytesRange(xml, formulaStart, openingEnd + 1, siAttributeNameBytes) : null
  if (formulaType === 'array' || formulaType === 'dataTable') {
    return true
  }
  if (xml[openingEnd - 1] === slashByte) {
    return formulaType !== 'shared' || sharedFormulaIndex === null || !sharedFormulaIndexes.has(sharedFormulaIndex)
  }
  const formulaEnd = indexOfBytes(xml, closeFormulaElementBytes, openingEnd + 1, end)
  if (formulaEnd < 0) {
    return true
  }
  const formula = decodeXmlEntities(decodeBytes(xml.subarray(openingEnd + 1, formulaEnd))).trim()
  if (formula.length === 0) {
    return formulaType !== 'shared' || sharedFormulaIndex === null || !sharedFormulaIndexes.has(sharedFormulaIndex)
  }
  const unsupported =
    formula.length === 0 ||
    /(?:^|[=,+(*/\s])'?\[[^\]]+\]/u.test(formula) ||
    /\[[#@\w]/u.test(formula) ||
    /(?:^|[^A-Z0-9_.])(?:NOW|RAND|RANDBETWEEN|TODAY)\s*\(/iu.test(formula)
  if (!unsupported && sharedFormulaIndex !== null) {
    sharedFormulaIndexes.add(sharedFormulaIndex)
  }
  return unsupported
}

function indexOfElementCandidate(xml: Uint8Array, elementName: Uint8Array, start: number, end = xml.length): number {
  const maxStart = end - elementName.length - 1
  for (let index = start; index <= maxStart; index += 1) {
    if (xml[index] !== lessThanByte) {
      continue
    }
    let matches = true
    for (let nameIndex = 0; nameIndex < elementName.length; nameIndex += 1) {
      if (xml[index + nameIndex + 1] !== elementName[nameIndex]) {
        matches = false
        break
      }
    }
    if (matches) {
      return index
    }
  }
  return -1
}

function isElementStartBytes(xml: Uint8Array, index: number, elementName: Uint8Array): boolean {
  if (xml[index] !== lessThanByte) {
    return false
  }
  for (let nameIndex = 0; nameIndex < elementName.length; nameIndex += 1) {
    if (xml[index + nameIndex + 1] !== elementName[nameIndex]) {
      return false
    }
  }
  const next = xml[index + elementName.length + 1]
  return next === undefined || isElementBoundaryByte(next)
}

function isElementBoundaryByte(value: number): boolean {
  return value === slashByte || value === greaterThanByte || whitespaceBytes.has(value)
}

function readXmlAttributeInBytesRange(source: Uint8Array, start: number, end: number, name: Uint8Array): string | null {
  for (let index = start; index < end; index += 1) {
    if (!whitespaceBytes.has(source[index])) {
      continue
    }
    let nameMatches = true
    for (let nameIndex = 0; nameIndex < name.length; nameIndex += 1) {
      if (source[index + nameIndex + 1] !== name[nameIndex]) {
        nameMatches = false
        break
      }
    }
    if (!nameMatches || source[index + name.length + 1] !== equalsByte) {
      continue
    }
    const quote = source[index + name.length + 2]
    if (quote !== doubleQuoteByte && quote !== singleQuoteByte) {
      continue
    }
    const valueStart = index + name.length + 3
    const valueEnd = indexOfByte(source, quote, valueStart, end)
    if (valueEnd < 0) {
      return null
    }
    return decodeXmlEntities(decodeBytes(source.subarray(valueStart, valueEnd)))
  }
  return null
}

function indexOfByte(source: Uint8Array, byte: number, start: number, end = source.length): number {
  for (let index = start; index < end; index += 1) {
    if (source[index] === byte) {
      return index
    }
  }
  return -1
}

function indexOfBytes(source: Uint8Array, search: Uint8Array, start: number, end: number): number {
  const maxStart = end - search.length
  for (let index = start; index <= maxStart; index += 1) {
    let matches = true
    for (let searchIndex = 0; searchIndex < search.length; searchIndex += 1) {
      if (source[index + searchIndex] !== search[searchIndex]) {
        matches = false
        break
      }
    }
    if (matches) {
      return index
    }
  }
  return -1
}

function concatBytes(left: Uint8Array, right: Uint8Array): Uint8Array {
  if (left.length === 0) {
    return right
  }
  if (right.length === 0) {
    return left
  }
  const combined = new Uint8Array(left.length + right.length)
  combined.set(left, 0)
  combined.set(right, left.length)
  return combined
}

function copyBytes(bytes: Uint8Array): Uint8Array {
  if (bytes.length === 0) {
    return new Uint8Array()
  }
  const copy = new Uint8Array(bytes.length)
  copy.set(bytes)
  return copy
}

function asciiBytes(value: string): Uint8Array {
  return Uint8Array.from(value, (character) => character.charCodeAt(0))
}

function readUint16(bytes: Uint8Array, offset: number): number {
  return bytes[offset] | (bytes[offset + 1] << 8)
}

function readUint32(bytes: Uint8Array, offset: number): number {
  return (bytes[offset] | (bytes[offset + 1] << 8) | (bytes[offset + 2] << 16) | (bytes[offset + 3] << 24)) >>> 0
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/gu, '\\$&')
}
