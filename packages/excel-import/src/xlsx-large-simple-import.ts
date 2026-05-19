import type {
  CellStyleRecord,
  SheetMetadataSnapshot,
  WorkbookDefinedNameSnapshot,
  WorkbookDefinedNameValueSnapshot,
  WorkbookSnapshot,
  WorkbookTableSnapshot,
} from '@bilig/protocol'
import { createSheetPreview, normalizeWorkbookName } from './workbook-import-helpers.js'
import { XLSX_CONTENT_TYPE } from './workbook-import-content-types.js'
import { createWorkbookPreview, type ImportedWorkbookPreview } from './workbook-import-preview.js'
import {
  readImportedSheetConditionalFormatArtifactsFromWorksheetXml,
  readImportedSheetConditionalFormatsFromWorksheetXml,
} from './xlsx-conditional-formats.js'
import { readImportedWorkbookDrawingArtifacts } from './xlsx-drawing-artifacts.js'
import { readImportedSheetAutoFilters } from './xlsx-filters.js'
import { readLargeSimpleSheetHyperlinks } from './xlsx-large-simple-hyperlinks.js'
import { LargeSimpleXlsxImportPhaseRecorder, type LargeSimpleXlsxImportPhaseTelemetry } from './xlsx-large-simple-import-telemetry.js'
import { readLargeSimpleSheetPrintMetadata } from './xlsx-large-simple-printer-settings.js'
import { readAllLargeSimpleSharedStrings, readReferencedLargeSimpleSharedStrings } from './xlsx-large-simple-referenced-shared-strings.js'
import type { LargeSimpleSharedStringEntry } from './xlsx-large-simple-shared-strings.js'
import { shouldUseSharedStringlessFastPathBytes } from './xlsx-large-simple-shared-stringless-fast-path.js'
import { buildLargeSimpleStyleRanges } from './xlsx-large-simple-style-ranges.js'
import { readLargeSimpleWorkbookStylesFromChunks } from './xlsx-large-simple-styles.js'
import { ImportedWorkbookArena, ImportedWorksheetStyleIndexArena, type ImportedWorksheetCellScan } from './xlsx-large-simple-arena.js'
import {
  parseHeadlessLargeSimpleWorksheetFromChunks,
  type HeadlessLargeSimpleWorksheetScan,
} from './xlsx-large-simple-headless-worksheet-scanner.js'
import {
  hasUnsupportedLargeSimpleWorksheetTags,
  needsLargeSimpleWorksheetMetadataXml,
  readLargeSimpleWorksheetMetadataXml,
  parseLargeSimpleWorksheetCells,
} from './xlsx-large-simple-worksheet-scanner.js'
import { parseLargeSimpleWorksheetCellsFromChunks } from './xlsx-large-simple-worksheet-stream-scanner.js'
import {
  readLargeSimpleColumnMetadata,
  readLargeSimpleMergeRanges,
  readLargeSimpleRowMetadata,
  readLargeSimpleSheetFormatPr,
  type LargeSimpleWorksheetScannedMetadata,
} from './xlsx-large-simple-worksheet-metadata.js'
import { readImportedSheetTablesFromWorksheetXml } from './xlsx-tables.js'
import {
  forEachInflatedXlsxZipEntryChunk,
  getZipText,
  normalizeZipPath,
  releaseLazyXlsxZipSource,
  type XlsxZipEntries,
} from './xlsx-zip.js'

export interface LargeSimpleXlsxImportResult {
  snapshot: WorkbookSnapshot
  workbookName: string
  sheetNames: string[]
  warnings: string[]
  preview: ImportedWorkbookPreview
  stats: LargeSimpleXlsxImportStats
}

export interface LargeSimpleXlsxImportOptions {
  minByteLength?: number
  materializeCells?: boolean
  materializeMetadata?: boolean
  releaseArenaAfterMaterialization?: boolean
  releaseZipSource?: boolean
}

export interface LargeSimpleXlsxImportStats {
  readonly sheetCount: number
  readonly cellCount: number
  readonly formulaCellCount: number
  readonly valueCellCount: number
  readonly definedNameCount: number
  readonly tableCount: number
  readonly mergeCount: number
  readonly conditionalFormatCount: number
  readonly warningCount: number
  readonly dimensions: readonly LargeSimpleXlsxSheetDimension[]
  readonly phaseTelemetry: readonly LargeSimpleXlsxImportPhaseTelemetry[]
}

export interface LargeSimpleXlsxSheetDimension {
  readonly sheetName: string
  readonly rowCount: number
  readonly columnCount: number
  readonly nonEmptyCellCount: number
  readonly usedRange: ImportedWorksheetCellScan['usedRange']
}

interface WorkbookSheetEntry {
  readonly name: string
  readonly relationshipId: string
}

interface WorkbookRelationship {
  readonly id: string
  readonly type: string
  readonly target: string
}

interface ParsedWorksheet {
  readonly sheet: WorkbookSnapshot['sheets'][number]
  readonly preview: ReturnType<typeof createSheetPreview>
  readonly stats: {
    readonly cellCount: number
    readonly formulaCellCount: number
    readonly valueCellCount: number
    readonly tableCount: number
    readonly mergeCount: number
    readonly conditionalFormatCount: number
    readonly dimension: LargeSimpleXlsxSheetDimension
  }
}

interface ScannedWorksheet {
  readonly name: string
  readonly order: number
  readonly cellScan: ImportedWorksheetCellScan
  readonly worksheetXml: string | undefined
  readonly metadataScan: LargeSimpleWorksheetScannedMetadata | undefined
  readonly metadataInput: Pick<
    SheetMetadataSnapshot,
    | 'conditionalFormatArtifacts'
    | 'conditionalFormats'
    | 'drawingArtifacts'
    | 'filters'
    | 'hyperlinks'
    | 'printerSettings'
    | 'printPageSetup'
  >
}

const defaultLargeSimpleXlsxByteThreshold = 1_000_000
const workbookPath = 'xl/workbook.xml'
const workbookRelationshipsPath = 'xl/_rels/workbook.xml.rels'
const sharedStringsPath = 'xl/sharedStrings.xml'
const stylesPath = 'xl/styles.xml'
const worksheetRelationshipType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'
const definedNameElementPattern =
  /<(?:[A-Za-z_][\w.-]*:)?definedName\b(?:[^>"']|"[^"]*"|'[^']*')*\/>|<((?:[A-Za-z_][\w.-]*:)?definedName)\b(?:[^>"']|"[^"]*"|'[^']*')*>[\s\S]*?<\/\1>/gu
const unsupportedPackagePathPattern =
  /^xl\/(?:charts|chartSheets|comments|ctrlProps|externalLinks|model|pivotCache|pivotTables|threadedComments|vbaProject\.bin)/u
const maxPreservedBlankStyleCellCount = 100_000

export function tryImportLargeSimpleXlsx(
  data: Uint8Array,
  fileName: string,
  zip: XlsxZipEntries,
  options: LargeSimpleXlsxImportOptions = {},
): LargeSimpleXlsxImportResult | null {
  if (data.byteLength < (options.minByteLength ?? defaultLargeSimpleXlsxByteThreshold)) {
    return null
  }
  const phaseRecorder = new LargeSimpleXlsxImportPhaseRecorder()
  const zipSetupStart = phaseRecorder.start()
  const packagePaths = Object.keys(zip).map(normalizeZipPath)
  if (packagePaths.some((path) => unsupportedPackagePathPattern.test(path))) {
    return null
  }

  const workbookXml = getZipText(zip, workbookPath)
  const workbookRelationshipsXml = getZipText(zip, workbookRelationshipsPath)
  if (!workbookXml || !workbookRelationshipsXml) {
    return null
  }

  const workbookSheets = readWorkbookSheets(workbookXml)
  const worksheetPathsByRelationshipId = readWorksheetPathsByRelationshipId(workbookRelationshipsXml)
  if (workbookSheets.length === 0 || worksheetPathsByRelationshipId.size === 0) {
    return null
  }
  const workbookDefinedNames = readWorkbookDefinedNames(
    workbookXml,
    workbookSheets.map((entry) => entry.name),
  )
  if (workbookDefinedNames.externalWorkbookReferenceSeen) {
    return null
  }

  const worksheetEntries = workbookSheets.flatMap((entry) => {
    const path = worksheetPathsByRelationshipId.get(entry.relationshipId)
    return path ? [{ name: entry.name, relationshipId: entry.relationshipId, path }] : []
  })
  if (worksheetEntries.length !== workbookSheets.length) {
    return null
  }
  const materializeCells = options.materializeCells !== false
  const materializeMetadata = options.materializeMetadata !== false
  const hasSharedStrings = packagePaths.includes(sharedStringsPath)
  const hasStyles = packagePaths.includes(stylesPath)
  let fallbackSharedStrings: readonly LargeSimpleSharedStringEntry[] | null | undefined = hasSharedStrings ? undefined : []
  delete zip[workbookPath]
  delete zip[workbookRelationshipsPath]
  const workbookName = normalizeWorkbookName(fileName)
  const warnings = workbookDefinedNames.ignoredCount > 0 ? ['Some defined names were ignored during XLSX import.'] : []
  const hasDrawingParts = packagePaths.some((path) => path.startsWith('xl/drawings/') || path.startsWith('xl/media/'))
  const importedDrawingArtifacts =
    materializeCells && hasDrawingParts
      ? readImportedWorkbookDrawingArtifacts(
          zip,
          workbookSheets.map((entry) => entry.name),
        )
      : null
  phaseRecorder.finish('zip-setup', zipSetupStart)
  const importedTables: WorkbookTableSnapshot[] = []
  const sheets: WorkbookSnapshot['sheets'] = []
  const previewSheets: ParsedWorksheet['preview'][] = []
  const sheetStats: ParsedWorksheet['stats'][] = []
  const styleCatalog = new Map<string, CellStyleRecord>()
  const scannedWorksheets: ScannedWorksheet[] = []
  const referencedSharedStringIndexes = new Set<number>()

  for (const [order, entry] of worksheetEntries.entries()) {
    const worksheetScanStart = phaseRecorder.start()
    let streamedWorksheetXml: string | undefined
    let streamedMetadataScan: LargeSimpleWorksheetScannedMetadata | undefined
    let cellScan: ImportedWorksheetCellScan | null = null
    if (!materializeCells && !materializeMetadata) {
      const headless = parseHeadlessLargeSimpleWorksheetFromChunks(
        (onChunk) => forEachInflatedXlsxZipEntryChunk(zip, entry.path, onChunk),
        order,
        { hasSharedStrings },
      )
      if (headless && (hasSharedStrings || headless.valueCellCount > 0)) {
        cellScan = importedWorksheetCellScanFromHeadless(headless)
        delete zip[entry.path]
      }
    } else {
      const streamed = parseLargeSimpleWorksheetCellsFromChunks(
        (onChunk) => forEachInflatedXlsxZipEntryChunk(zip, entry.path, onChunk),
        order,
        {
          hasSharedStrings,
          retainCells: materializeCells,
          sharedStrings: fallbackSharedStrings ?? [],
          deferSharedStrings: materializeCells && hasSharedStrings,
          retainMetadataXml: materializeMetadata,
        },
      )
      if (streamed && (hasSharedStrings || streamed.cellScan.valueCellCount > 0)) {
        cellScan = streamed.cellScan
        streamedWorksheetXml = streamed.metadataXml
        streamedMetadataScan = streamed.metadata
        delete zip[entry.path]
      }
    }
    let worksheetBytes: Uint8Array | undefined
    if (!cellScan) {
      worksheetBytes = zip[entry.path]
      if (!worksheetBytes) {
        return null
      }
      delete zip[entry.path]
      if (!hasSharedStrings && !shouldUseSharedStringlessFastPathBytes(worksheetBytes)) {
        return null
      }
      if (hasUnsupportedLargeSimpleWorksheetTags(worksheetBytes)) {
        return null
      }
      if (hasSharedStrings && fallbackSharedStrings === undefined) {
        fallbackSharedStrings = readAllLargeSimpleSharedStrings(zip)
        if (fallbackSharedStrings === null) {
          return null
        }
      }
      cellScan = parseLargeSimpleWorksheetCells(worksheetBytes, fallbackSharedStrings ?? [], order, { retainCells: materializeCells })
    }
    if (!cellScan) {
      return null
    }
    if (materializeCells && cellScan.blankStyleCellCount > 0 && cellScan.blankStyleCellCount <= maxPreservedBlankStyleCellCount) {
      return null
    }
    cellScan.arena.collectSharedStringIndexes(referencedSharedStringIndexes)
    phaseRecorder.finish('worksheet-scan', worksheetScanStart)
    const metadataParsingStart = phaseRecorder.start()
    let worksheetXml: string | undefined
    let metadataInput: Pick<
      SheetMetadataSnapshot,
      | 'conditionalFormatArtifacts'
      | 'conditionalFormats'
      | 'drawingArtifacts'
      | 'filters'
      | 'hyperlinks'
      | 'printerSettings'
      | 'printPageSetup'
    > = {}
    const needsWorksheetXml =
      materializeMetadata &&
      (streamedWorksheetXml !== undefined || (worksheetBytes ? needsLargeSimpleWorksheetMetadataXml(worksheetBytes) : false))
    const drawingArtifacts = importedDrawingArtifacts?.sheetArtifactsByName.get(entry.name)
    if (needsWorksheetXml) {
      worksheetXml = streamedWorksheetXml ?? (worksheetBytes ? readLargeSimpleWorksheetMetadataXml(worksheetBytes) : undefined)
      if (!worksheetXml) {
        return null
      }
      const sheetTables = /<(?:[A-Za-z_][\w.-]*:)?tableParts\b/u.test(worksheetXml)
        ? readImportedSheetTablesFromWorksheetXml(zip, entry.name, entry.path, worksheetXml)
        : undefined
      if (sheetTables) {
        importedTables.push(...sheetTables)
      }
      const hasConditionalFormats = /<(?:[A-Za-z_][\w.-]*:)?conditionalFormatting\b/u.test(worksheetXml)
      const conditionalFormats = hasConditionalFormats
        ? readImportedSheetConditionalFormatsFromWorksheetXml(zip, entry.name, worksheetXml)
        : undefined
      if (materializeCells) {
        const hyperlinks = readLargeSimpleSheetHyperlinks(zip, entry.name, entry.path, worksheetXml)
        if (hyperlinks === null) {
          return null
        }
        const printMetadata = readLargeSimpleSheetPrintMetadata(zip, entry.path, worksheetXml)
        if (printMetadata === null) {
          return null
        }
        const filters = readImportedSheetAutoFilters(entry.name, worksheetXml)
        const conditionalFormatArtifacts = hasConditionalFormats
          ? readImportedSheetConditionalFormatArtifactsFromWorksheetXml(worksheetXml)
          : undefined
        metadataInput = {
          ...(drawingArtifacts ? { drawingArtifacts } : {}),
          ...(hyperlinks ? { hyperlinks } : {}),
          ...(filters.length > 0 ? { filters } : {}),
          ...(conditionalFormats ? { conditionalFormats } : {}),
          ...(conditionalFormatArtifacts ? { conditionalFormatArtifacts } : {}),
          ...printMetadata,
        }
      } else {
        metadataInput = conditionalFormats ? { conditionalFormats } : {}
      }
    } else if (drawingArtifacts) {
      metadataInput = { drawingArtifacts }
    }
    worksheetBytes = undefined
    scannedWorksheets.push({
      name: entry.name,
      order,
      cellScan,
      worksheetXml,
      metadataScan: streamedMetadataScan,
      metadataInput,
    })
    phaseRecorder.finish('metadata-parsing', metadataParsingStart)
  }
  const sharedStringResolutionStart = phaseRecorder.start()
  let sharedStrings: readonly LargeSimpleSharedStringEntry[] = fallbackSharedStrings ?? []
  if (materializeCells && hasSharedStrings && referencedSharedStringIndexes.size > 0) {
    const referencedSharedStrings = fallbackSharedStrings ?? readReferencedLargeSimpleSharedStrings(zip, referencedSharedStringIndexes)
    if (referencedSharedStrings === null) {
      return null
    }
    sharedStrings = referencedSharedStrings
  }
  delete zip[sharedStringsPath]
  for (const scanned of scannedWorksheets) {
    const resolvedRichTextCells = materializeCells && hasSharedStrings ? scanned.cellScan.arena.resolveSharedStrings(sharedStrings) : []
    if (resolvedRichTextCells === null) {
      return null
    }
    if (resolvedRichTextCells.length > 0) {
      scanned.cellScan.richTextCells.push(...resolvedRichTextCells)
    }
  }
  phaseRecorder.finish('shared-string-resolution', sharedStringResolutionStart)
  const styleParsingStart = phaseRecorder.start()
  const requiredStyleIndexes = new Set<number>()
  for (const scanned of scannedWorksheets) {
    scanned.cellScan.styleIndexes.collectRequiredStyleIndexes(requiredStyleIndexes)
  }
  const stylesByIndex =
    materializeCells && hasStyles
      ? readLargeSimpleWorkbookStylesFromChunks(
          (onChunk) => forEachInflatedXlsxZipEntryChunk(zip, stylesPath, onChunk),
          requiredStyleIndexes,
        )
      : new Map()
  if (stylesByIndex === null) {
    return null
  }
  delete zip[stylesPath]
  phaseRecorder.finish('style-parsing', styleParsingStart)
  if (options.releaseZipSource === true) {
    const zipSourceReleaseStart = phaseRecorder.start()
    releaseLazyXlsxZipSource(zip)
    phaseRecorder.finish('zip-source-release', zipSourceReleaseStart)
  }
  for (const scanned of scannedWorksheets) {
    const snapshotMaterializationStart = phaseRecorder.start()
    const parsed = buildParsedWorksheet(
      scanned.name,
      scanned.order,
      scanned.cellScan,
      scanned.worksheetXml,
      scanned.metadataScan,
      scanned.metadataInput,
      {
        materializeCells,
        releaseArenaAfterMaterialization: options.releaseArenaAfterMaterialization !== false,
        styleCatalog,
        stylesByIndex,
      },
    )
    sheets.push(parsed.sheet)
    previewSheets.push(parsed.preview)
    sheetStats.push(parsed.stats)
    phaseRecorder.finish('public-snapshot-materialization', snapshotMaterializationStart)
  }
  const sortedImportedTables =
    importedTables.length > 0 ? importedTables.toSorted((left, right) => left.name.localeCompare(right.name)) : undefined
  const stats: LargeSimpleXlsxImportStats = {
    sheetCount: sheets.length,
    cellCount: sheetStats.reduce((sum, entry) => sum + entry.cellCount, 0),
    formulaCellCount: sheetStats.reduce((sum, entry) => sum + entry.formulaCellCount, 0),
    valueCellCount: sheetStats.reduce((sum, entry) => sum + entry.valueCellCount, 0),
    definedNameCount: workbookDefinedNames.definedNames?.length ?? 0,
    tableCount: sortedImportedTables?.length ?? sheetStats.reduce((sum, entry) => sum + entry.tableCount, 0),
    mergeCount: sheetStats.reduce((sum, entry) => sum + entry.mergeCount, 0),
    conditionalFormatCount: sheetStats.reduce((sum, entry) => sum + entry.conditionalFormatCount, 0),
    warningCount: warnings.length,
    dimensions: sheetStats.map((entry) => entry.dimension),
    phaseTelemetry: phaseRecorder.entries(),
  }

  return {
    snapshot: {
      version: 1,
      workbook: {
        name: workbookName,
        ...(workbookDefinedNames.definedNames || importedDrawingArtifacts?.artifacts || sortedImportedTables || styleCatalog.size > 0
          ? {
              metadata: {
                ...(workbookDefinedNames.definedNames ? { definedNames: workbookDefinedNames.definedNames } : {}),
                ...(importedDrawingArtifacts?.artifacts ? { drawingArtifacts: importedDrawingArtifacts.artifacts } : {}),
                ...(sortedImportedTables ? { tables: sortedImportedTables } : {}),
                ...(styleCatalog.size > 0 ? { styles: [...styleCatalog.values()] } : {}),
              },
            }
          : {}),
      },
      sheets,
    },
    workbookName,
    sheetNames: workbookSheets.map((entry) => entry.name),
    warnings,
    preview: createWorkbookPreview({
      contentType: XLSX_CONTENT_TYPE,
      fileName,
      fileSizeBytes: data.byteLength,
      workbookName,
      sheets: previewSheets,
      warnings,
    }),
    stats,
  }
}

function buildParsedWorksheet(
  sheetName: string,
  order: number,
  cellScan: ImportedWorksheetCellScan,
  worksheetXml: string | undefined,
  metadataScan: LargeSimpleWorksheetScannedMetadata | undefined,
  input: Pick<
    SheetMetadataSnapshot,
    | 'conditionalFormatArtifacts'
    | 'conditionalFormats'
    | 'drawingArtifacts'
    | 'filters'
    | 'hyperlinks'
    | 'printerSettings'
    | 'printPageSetup'
  > = {},
  options: {
    readonly materializeCells: boolean
    readonly releaseArenaAfterMaterialization?: boolean
    readonly styleCatalog?: Map<string, CellStyleRecord>
    readonly stylesByIndex?: ReadonlyMap<number, Omit<CellStyleRecord, 'id'>>
  } = { materializeCells: true },
): ParsedWorksheet {
  const merges =
    metadataScan?.merges?.map((range) => ({ sheetName, ...range })) ??
    (worksheetXml ? readLargeSimpleMergeRanges(sheetName, worksheetXml) : [])
  const mergeCount = worksheetXml ? merges.length : (cellScan.mergeCount ?? 0)
  const columns = metadataScan?.columns ?? (worksheetXml ? readLargeSimpleColumnMetadata(worksheetXml) : { entries: [], metadata: [] })
  const rows = metadataScan?.rows ?? (worksheetXml ? readLargeSimpleRowMetadata(worksheetXml) : { entries: [], metadata: [] })
  const sheetFormatPr = metadataScan?.sheetFormatPr ?? (worksheetXml ? readLargeSimpleSheetFormatPr(worksheetXml) : undefined)
  const conditionalFormatCount =
    input.conditionalFormats?.length ??
    (worksheetXml ? readConditionalFormattingBlockCount(worksheetXml) : (cellScan.conditionalFormatCount ?? 0))
  const styleRanges =
    options.materializeCells && options.styleCatalog && options.stylesByIndex
      ? buildLargeSimpleStyleRanges(sheetName, cellScan, options.stylesByIndex, options.styleCatalog)
      : []
  const metadata: SheetMetadataSnapshot = {
    ...(columns.entries.length > 0 ? { columns: columns.entries } : {}),
    ...(rows.entries.length > 0 ? { rows: rows.entries } : {}),
    ...(columns.metadata.length > 0 ? { columnMetadata: columns.metadata } : {}),
    ...(rows.metadata.length > 0 ? { rowMetadata: rows.metadata } : {}),
    ...(sheetFormatPr ? { sheetFormatPr } : {}),
    ...(styleRanges.length > 0 ? { styleRanges } : {}),
    ...(merges.length > 0 ? { merges } : {}),
    ...(input.drawingArtifacts ? { drawingArtifacts: input.drawingArtifacts } : {}),
    ...(input.filters ? { filters: input.filters } : {}),
    ...(input.hyperlinks ? { hyperlinks: input.hyperlinks } : {}),
    ...(input.conditionalFormats ? { conditionalFormats: input.conditionalFormats } : {}),
    ...(input.conditionalFormatArtifacts ? { conditionalFormatArtifacts: input.conditionalFormatArtifacts } : {}),
    ...(input.printerSettings ? { printerSettings: input.printerSettings } : {}),
    ...(input.printPageSetup ? { printPageSetup: input.printPageSetup } : {}),
    ...(cellScan.richTextCells.length > 0 ? { richTextArtifacts: { cells: cellScan.richTextCells } } : {}),
  }
  const cells = options.materializeCells ? cellScan.arena.materializeSheetCells(cellScan.sheetIndex) : []
  const sheet: WorkbookSnapshot['sheets'][number] = {
    id: order + 1,
    name: sheetName,
    order,
    ...(Object.keys(metadata).length > 0 ? { metadata } : {}),
    cells,
  }
  const parsed: ParsedWorksheet = {
    sheet,
    preview: createSheetPreview({
      name: sheetName,
      rowCount: cellScan.rowCount,
      columnCount: cellScan.columnCount,
      nonEmptyCellCount: cellScan.cellCount,
      readCellText: (row, column) => cellScan.arena.readPreviewText(row, column),
    }),
    stats: {
      cellCount: cellScan.cellCount,
      formulaCellCount: cellScan.formulaCellCount,
      valueCellCount: cellScan.valueCellCount,
      tableCount: cellScan.tableCount ?? 0,
      mergeCount,
      conditionalFormatCount,
      dimension: {
        sheetName,
        rowCount: cellScan.rowCount,
        columnCount: cellScan.columnCount,
        nonEmptyCellCount: cellScan.cellCount,
        usedRange: cellScan.usedRange,
      },
    },
  }
  if (options.releaseArenaAfterMaterialization === true) {
    cellScan.arena.release()
    cellScan.styleIndexes.release()
  }
  return parsed
}

function readConditionalFormattingBlockCount(worksheetXml: string): number {
  return [...worksheetXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?conditionalFormatting\b/gu)].length
}

function importedWorksheetCellScanFromHeadless(scan: HeadlessLargeSimpleWorksheetScan): ImportedWorksheetCellScan {
  return {
    arena: new ImportedWorkbookArena(),
    sheetIndex: scan.sheetIndex,
    richTextCells: [],
    styleIndexes: new ImportedWorksheetStyleIndexArena(),
    blankStyleCellCount: 0,
    cellCount: scan.cellCount,
    valueCellCount: scan.valueCellCount,
    formulaCellCount: scan.formulaCellCount,
    mergeCount: scan.mergeCount,
    conditionalFormatCount: scan.conditionalFormatCount,
    tableCount: scan.tableCount,
    rowCount: scan.rowCount,
    columnCount: scan.columnCount,
    usedRange: scan.usedRange,
  }
}

function readWorkbookSheets(workbookXml: string): WorkbookSheetEntry[] {
  return [...workbookXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?sheet\b(?:[^>"']|"[^"]*"|'[^']*')*\/?>/gu)].flatMap((match) => {
    const tag = match[0]
    const name = readXmlAttribute(tag, 'name')
    const relationshipId = readXmlAttribute(tag, 'r:id') ?? readXmlAttribute(tag, 'id')
    return name && relationshipId ? [{ name: decodeXmlText(name), relationshipId }] : []
  })
}

function readWorksheetPathsByRelationshipId(workbookRelationshipsXml: string): Map<string, string> {
  return new Map(
    readRelationships(workbookRelationshipsXml).flatMap((relationship) => {
      if (relationship.type !== worksheetRelationshipType && !relationship.target.includes('worksheets/')) {
        return []
      }
      return [[relationship.id, normalizeZipPath(resolveTargetPath(workbookPath, relationship.target))]]
    }),
  )
}

function readRelationships(relationshipsXml: string): WorkbookRelationship[] {
  return [...relationshipsXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?Relationship\b(?:[^>"']|"[^"]*"|'[^']*')*\/?>/gu)].flatMap((match) => {
    const tag = match[0]
    const id = readXmlAttribute(tag, 'Id')
    const type = readXmlAttribute(tag, 'Type')
    const target = readXmlAttribute(tag, 'Target')
    return id && type && target ? [{ id, type, target }] : []
  })
}

function readWorkbookDefinedNames(
  workbookXml: string,
  sheetNames: readonly string[],
): {
  readonly definedNames: WorkbookDefinedNameSnapshot[] | undefined
  readonly externalWorkbookReferenceSeen: boolean
  readonly ignoredCount: number
} {
  const definedNamesByKey = new Map<string, WorkbookDefinedNameSnapshot>()
  let externalWorkbookReferenceSeen = false
  let ignoredCount = 0
  for (const match of workbookXml.matchAll(definedNameElementPattern)) {
    const xml = match[0]
    const openingTag = /<(?:[A-Za-z_][\w.-]*:)?definedName\b(?:[^>"']|"[^"]*"|'[^']*')*(?:\/>|>)/u.exec(xml)?.[0]
    const name = openingTag ? readXmlAttribute(openingTag, 'name')?.trim() : ''
    const localSheetId = openingTag ? readNonNegativeIntegerAttribute(openingTag, 'localSheetId') : null
    const scopeSheetName = localSheetId !== null ? sheetNames[localSheetId] : undefined
    const rawValue = openingTag?.endsWith('/>') ? '' : decodeXmlText(xml.replace(/^<[^>]*>/u, '').replace(/<\/[^>]*>$/u, '')).trim()
    if (!name || rawValue.length === 0 || (localSheetId !== null && scopeSheetName === undefined)) {
      ignoredCount += 1
      continue
    }
    if (definedNameReferencesExternalWorkbook(rawValue)) {
      externalWorkbookReferenceSeen = true
      continue
    }
    const value = isBuiltInPrintDefinedName(name)
      ? parsePrintDefinedNameValue(rawValue)
      : parseDefinedNameValue(rawValue, new Set(sheetNames))
    if (!value) {
      ignoredCount += 1
      continue
    }
    definedNamesByKey.set(definedNameKey(name, scopeSheetName), {
      name,
      ...(scopeSheetName !== undefined ? { scopeSheetName } : {}),
      value,
    })
  }
  const definedNames = [...definedNamesByKey.values()].toSorted(
    (left, right) => left.name.localeCompare(right.name) || (left.scopeSheetName ?? '').localeCompare(right.scopeSheetName ?? ''),
  )
  return {
    definedNames: definedNames.length > 0 ? definedNames : undefined,
    externalWorkbookReferenceSeen,
    ignoredCount,
  }
}

function definedNameReferencesExternalWorkbook(value: string): boolean {
  return /(?:^|[=,+(*/\s])'?\[[^\]]+\]/u.test(value)
}

function isBuiltInPrintDefinedName(name: string): boolean {
  const normalized = name.trim().toLocaleLowerCase('en-US')
  return normalized === '_xlnm.print_area' || normalized === '_xlnm.print_titles'
}

function parsePrintDefinedNameValue(value: string): WorkbookDefinedNameValueSnapshot | null {
  const trimmed = value.trim()
  return trimmed.length > 0 ? { kind: 'formula', formula: trimmed.startsWith('=') ? trimmed : `=${trimmed}` } : null
}

function parseDefinedNameValue(value: string, sheetNames: ReadonlySet<string>): WorkbookDefinedNameValueSnapshot | null {
  const trimmed = value.trim()
  if (trimmed.length === 0) {
    return null
  }
  const expression = trimmed.startsWith('=') ? trimmed.slice(1).trim() : trimmed
  const sheetReference = parseSheetReference(expression)
  if (sheetReference && sheetNames.has(sheetReference.sheetName)) {
    const parsedReference = parseDefinedNameReferenceValue(sheetReference.sheetName, sheetReference.reference)
    if (parsedReference) {
      return parsedReference
    }
  }
  const scalar = parseDefinedNameScalarValue(expression)
  if (scalar) {
    return scalar
  }
  return { kind: 'formula', formula: trimmed.startsWith('=') ? trimmed : `=${trimmed}` }
}

function parseDefinedNameReferenceValue(sheetName: string, reference: string): WorkbookDefinedNameValueSnapshot | null {
  const parts = reference.split(':')
  if (parts.length === 1) {
    const address = normalizeDefinedNameCellAddress(parts[0] ?? '')
    return address ? { kind: 'cell-ref', sheetName, address } : null
  }
  if (parts.length === 2) {
    const startAddress = normalizeDefinedNameCellAddress(parts[0] ?? '')
    const endAddress = normalizeDefinedNameCellAddress(parts[1] ?? '')
    return startAddress && endAddress ? { kind: 'range-ref', sheetName, startAddress, endAddress } : null
  }
  return null
}

function parseDefinedNameScalarValue(value: string): WorkbookDefinedNameValueSnapshot | null {
  const trimmed = value.trim()
  if (/^[+-]?(?:(?:[0-9]+(?:\.[0-9]*)?)|(?:\.[0-9]+))(?:[eE][+-]?[0-9]+)?$/u.test(trimmed)) {
    const numberValue = Number(trimmed)
    return Number.isFinite(numberValue) ? { kind: 'scalar', value: numberValue } : null
  }
  if (/^TRUE$/iu.test(trimmed)) {
    return { kind: 'scalar', value: true }
  }
  if (/^FALSE$/iu.test(trimmed)) {
    return { kind: 'scalar', value: false }
  }
  if (trimmed.startsWith('"') && trimmed.endsWith('"')) {
    return { kind: 'scalar', value: trimmed.slice(1, -1).replace(/""/gu, '"') }
  }
  return null
}

function parseSheetReference(value: string): { readonly sheetName: string; readonly reference: string } | null {
  const quoted = parseQuotedSheetReference(value)
  if (quoted) {
    return quoted
  }
  const separatorIndex = value.indexOf('!')
  if (separatorIndex <= 0) {
    return null
  }
  const sheetName = value.slice(0, separatorIndex).trim()
  const reference = value.slice(separatorIndex + 1).trim()
  return sheetName.length > 0 && reference.length > 0 ? { sheetName, reference } : null
}

function parseQuotedSheetReference(value: string): { readonly sheetName: string; readonly reference: string } | null {
  if (!value.startsWith("'")) {
    return null
  }
  let sheetName = ''
  for (let index = 1; index < value.length; index += 1) {
    const character = value[index]
    if (character === "'" && value[index + 1] === "'") {
      sheetName += "'"
      index += 1
      continue
    }
    if (character === "'" && value[index + 1] === '!') {
      const reference = value.slice(index + 2).trim()
      return sheetName.trim().length > 0 && reference.length > 0 ? { sheetName, reference } : null
    }
    sheetName += character
  }
  return null
}

function normalizeDefinedNameCellAddress(value: string): string | null {
  const normalized = value.trim().replaceAll('$', '').toUpperCase()
  const match = /^([A-Z]{1,3})([1-9][0-9]*)$/u.exec(normalized)
  return match && decodeCellAddress(normalized) ? normalized : null
}

function definedNameKey(name: string, scopeSheetName: string | undefined): string {
  return `${scopeSheetName ?? '<workbook>'}\u0000${name.toUpperCase()}`
}

function readXmlAttribute(xml: string, attributeName: string): string | null {
  return new RegExp(`\\s${attributeName}=("|')([\\s\\S]*?)\\1`, 'u').exec(xml)?.[2] ?? null
}

function readNumberAttribute(xml: string, attributeName: string): number | null {
  const raw = readXmlAttribute(xml, attributeName)
  if (raw === null || raw.trim().length === 0) {
    return null
  }
  const value = Number(raw)
  return Number.isFinite(value) ? value : null
}

function readNonNegativeIntegerAttribute(xml: string, attributeName: string): number | null {
  const value = readNumberAttribute(xml, attributeName)
  return Number.isInteger(value) && value !== null && value >= 0 ? value : null
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
  if (!Number.isSafeInteger(row) || row <= 0 || column <= 0) {
    return null
  }
  return { row: row - 1, column: column - 1 }
}

function resolveTargetPath(basePath: string, target: string): string {
  if (target.startsWith('/')) {
    return target.slice(1)
  }
  const parts = basePath.split('/')
  parts.pop()
  for (const segment of target.split('/')) {
    if (segment === '..') {
      parts.pop()
    } else if (segment !== '.' && segment.length > 0) {
      parts.push(segment)
    }
  }
  return parts.join('/')
}
