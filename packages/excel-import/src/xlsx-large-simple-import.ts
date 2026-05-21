import type { CellStyleRecord, SheetMetadataSnapshot, WorkbookSnapshot, WorkbookTableSnapshot } from '@bilig/protocol'
import { attachRuntimeImage } from '@bilig/core'
import { createSheetPreview, normalizeWorkbookName } from './workbook-import-helpers.js'
import { XLSX_CONTENT_TYPE } from './workbook-import-content-types.js'
import { createWorkbookPreview, type ImportedWorkbookPreview } from './workbook-import-preview.js'
import {
  readImportedSheetConditionalFormatArtifactsFromElementXml,
  readImportedSheetConditionalFormatArtifactsFromWorksheetXml,
  readImportedSheetConditionalFormatsFromElementXml,
  readImportedSheetConditionalFormatsFromWorksheetXml,
} from './xlsx-conditional-formats.js'
import { readImportedWorkbookCellMetadataPart } from './xlsx-cell-metadata.js'
import { legacyCommentThreadSignature, readImportedWorkbookLegacyCommentVmlFromSheetSources } from './xlsx-comment-vml.js'
import { readImportedWorkbookControlArtifactsFromSheetSources } from './xlsx-control-artifacts.js'
import { readImportedWorkbookDataModelArtifacts } from './xlsx-data-model-artifacts.js'
import { readImportedWorkbookDrawingArtifactsFromWorksheetRelationships } from './xlsx-drawing-artifacts.js'
import { readImportedWorkbookExternalLinkArtifacts } from './xlsx-external-link-artifacts.js'
import { readImportedSheetAutoFilters } from './xlsx-filters.js'
import { readImportedWorkbookChartDrawingArtifacts } from './xlsx-import-chart-drawing-artifacts.js'
import { buildLargeSimpleCellMetadataReferenceSnapshots } from './xlsx-large-simple-cell-metadata.js'
import { readWorkbookDefinedNames } from './xlsx-large-simple-defined-names.js'
import { readLargeSimpleSheetHyperlinks, resolveLargeSimpleSheetHyperlinks } from './xlsx-large-simple-hyperlinks.js'
import { LargeSimpleXlsxImportPhaseRecorder, type LargeSimpleXlsxImportPhaseTelemetry } from './xlsx-large-simple-import-telemetry.js'
import {
  appendLargeSimpleConditionalFormats,
  normalizeLargeSimpleConditionalFormatIds,
  readLargeSimpleConditionalFormattingBlockCount,
} from './xlsx-large-simple-conditional-format-helpers.js'
import { internLargeSimpleWorksheetMetadata } from './xlsx-large-simple-metadata-interning.js'
import { prepareLargeSimplePackageArtifactsForZipRelease } from './xlsx-large-simple-package-artifact-release.js'
import { readLargeSimpleSheetPrintMetadata, readLargeSimpleSheetPrintPageSetup } from './xlsx-large-simple-printer-settings.js'
import { readAllLargeSimpleSharedStrings, readReferencedLargeSimpleSharedStrings } from './xlsx-large-simple-referenced-shared-strings.js'
import {
  createLargeSimpleSharedStringSubset,
  hasReferencedLargeSimpleRichSharedStrings,
  type LargeSimpleSharedStrings,
} from './xlsx-large-simple-shared-strings.js'
import { shouldUseSharedStringlessFastPathBytes } from './xlsx-large-simple-shared-stringless-fast-path.js'
import { buildLargeSimpleStyleRanges } from './xlsx-large-simple-style-ranges.js'
import { readLargeSimpleWorkbookStylesFromChunks } from './xlsx-large-simple-styles.js'
import {
  maxPreallocatedWorksheetCells,
  prepareLargeSimpleStyleIndexes,
  shouldDeferLargeSimpleStyleCoordinates,
} from './xlsx-large-simple-style-coordinate-rescan.js'
import { collectLargeSimpleImportGarbage } from './xlsx-large-simple-garbage.js'
import { mergeWorkbookRichTextCells } from './xlsx-large-simple-lazy-rich-text-cells.js'
import { ImportedWorkbookStringPool } from './xlsx-large-simple-string-pool.js'
import { readWorkbookSheets, readWorksheetPathsByRelationshipId } from './xlsx-large-simple-workbook-metadata.js'
import type { ImportedWorksheetCellScan } from './xlsx-large-simple-arena.js'
import {
  largeSimpleControlArtifactSheetSources,
  largeSimpleLegacyCommentVmlSheetSources,
  largeSimpleSlicerConnectionSheetSources,
} from './xlsx-large-simple-package-artifact-sources.js'
import { parseHeadlessLargeSimpleWorksheetFromChunks } from './xlsx-large-simple-headless-worksheet-scanner.js'
import { importedWorksheetCellScanFromHeadless } from './xlsx-large-simple-headless-cell-scan.js'
import {
  hasUnsupportedLargeSimpleWorksheetTags,
  needsLargeSimpleWorksheetMetadataXml,
  readLargeSimpleWorksheetMetadataXml,
  parseLargeSimpleWorksheetCells,
} from './xlsx-large-simple-worksheet-scanner.js'
import { parseLargeSimpleWorksheetCellsFromChunks } from './xlsx-large-simple-worksheet-stream-scanner.js'
import {
  readLargeSimpleColumnMetadata,
  readLargeSimpleDrawingRelationshipId,
  readLargeSimpleMergeRanges,
  readLargeSimpleRowMetadata,
  readLargeSimpleSheetFormatPr,
  withoutLargeSimpleConditionalFormattingXml,
  type LargeSimpleWorksheetScannedMetadata,
} from './xlsx-large-simple-worksheet-metadata.js'
import { readImportedPivotArtifacts } from './xlsx-pivot-artifacts.js'
import { readImportedWorkbookSlicerConnectionArtifactsFromSheets } from './xlsx-slicer-connection-artifacts.js'
import { readImportedSheetTablesFromRelationshipIds, readImportedSheetTablesFromWorksheetXml } from './xlsx-tables.js'
import { readImportedWorkbookDocumentPropertiesArtifacts, readImportedWorkbookProperties } from './xlsx-workbook-properties.js'
import {
  forEachInflatedXlsxZipEntryChunk,
  getZipText,
  normalizeZipPath,
  readLazyXlsxZipSourceByteLength,
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
  allowUnsupportedFormulaText?: boolean
  allowUnsupportedCellMetadata?: boolean
  maxMaterializedLazyPackageArtifactBytes?: number
  releaseOwnedSourceBytes?: () => LargeSimpleXlsxOwnedSourceReleaseEvidence | undefined
}

export interface LargeSimpleXlsxImportSource {
  readonly byteLength: number
}

export interface LargeSimpleXlsxOwnedSourceReleaseEvidence {
  readonly ownedSourceBytesBeforeRelease?: number
  readonly ownedSourceBytesAfterRelease?: number
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
  readonly dataValidationCount: number
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
    readonly dataValidationCount: number
    readonly dimension: LargeSimpleXlsxSheetDimension
  }
}

interface ScannedWorksheet {
  readonly name: string
  readonly order: number
  readonly cellScan: ImportedWorksheetCellScan
  readonly worksheetXml: string | undefined
  readonly metadataScan: LargeSimpleWorksheetScannedMetadata | undefined
  readonly metadataInput: LargeSimpleSheetMetadataInput
  readonly sharedStringIndexes: ReadonlySet<number>
  readonly sharedStrings?: LargeSimpleSharedStrings
}

type LargeSimpleSheetMetadataInput = Pick<
  SheetMetadataSnapshot,
  | 'conditionalFormatArtifacts'
  | 'conditionalFormats'
  | 'controlArtifacts'
  | 'validations'
  | 'drawingArtifacts'
  | 'filters'
  | 'hyperlinks'
  | 'legacyCommentVml'
  | 'pivotArtifacts'
  | 'printerSettings'
  | 'printPageSetup'
  | 'sheetProtection'
>

const defaultLargeSimpleXlsxByteThreshold = 1_000_000
const lazySheetCellMaterializationThreshold = 100_000
const workbookPath = 'xl/workbook.xml'
const workbookRelationshipsPath = 'xl/_rels/workbook.xml.rels'
const sharedStringsPath = 'xl/sharedStrings.xml'
const stylesPath = 'xl/styles.xml'
const unsupportedPackagePathPattern = /^xl\/(?:ctrlProps|threadedComments|vbaProject\.bin)/u
const emptySharedStringIndexes: ReadonlySet<number> = new Set()
export function tryImportLargeSimpleXlsx(
  source: LargeSimpleXlsxImportSource,
  fileName: string,
  zip: XlsxZipEntries,
  options: LargeSimpleXlsxImportOptions = {},
): LargeSimpleXlsxImportResult | null {
  if (source.byteLength < (options.minByteLength ?? defaultLargeSimpleXlsxByteThreshold)) {
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

  const stringPool = new ImportedWorkbookStringPool()
  const workbookSheets = readWorkbookSheets(workbookXml, stringPool)
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
  const hasDrawingParts = packagePaths.some((path) => path.startsWith('xl/drawings/') || path.startsWith('xl/media/'))
  const hasChartParts = packagePaths.some((path) => path.startsWith('xl/charts/') || path.startsWith('xl/chartSheets/'))
  const hasPivotParts = packagePaths.some((path) => path.startsWith('xl/pivotTables/') || path.startsWith('xl/pivotCache/'))
  const hasExternalLinkParts = packagePaths.some((path) => path.startsWith('xl/externalLinks/'))
  const hasLegacyCommentParts = packagePaths.some((path) => path.startsWith('xl/comments') || path.endsWith('.vml'))
  const hasDataModelParts = packagePaths.some(
    (path) => path.startsWith('xl/model/') || path.startsWith('xl/customData/') || path.startsWith('customXml/'),
  )
  const hasSlicerConnectionParts = packagePaths.some(
    (path) => path === 'xl/connections.xml' || path.startsWith('xl/slicerCaches/') || path.startsWith('xl/slicers/'),
  )
  const importedExternalLinkArtifacts =
    materializeCells && hasExternalLinkParts ? readImportedWorkbookExternalLinkArtifacts(zip) : undefined
  const importedDataModelArtifacts = materializeCells && hasDataModelParts ? readImportedWorkbookDataModelArtifacts(zip) : undefined
  const importedWorkbookProperties = materializeCells ? readImportedWorkbookProperties(zip) : undefined
  const importedWorkbookDocumentProperties = materializeCells ? readImportedWorkbookDocumentPropertiesArtifacts(zip) : undefined
  const importedWorkbookCellMetadata = materializeCells ? readImportedWorkbookCellMetadataPart(zip) : undefined
  const importedPivotArtifacts =
    materializeCells && hasPivotParts
      ? readImportedPivotArtifacts(
          zip,
          workbookSheets.map((entry) => entry.name),
          { readWorksheetPivotTableDefinitionsXml: false },
        )
      : null
  const importedChartDrawingArtifacts =
    materializeCells && hasChartParts
      ? readImportedWorkbookChartDrawingArtifacts(
          zip,
          workbookSheets.map((entry) => entry.name),
        )
      : null
  const deduplicateInlineStrings = hasSharedStrings ? true : 'bounded'
  const deduplicateFormulaStrings = 'bounded'
  let fallbackSharedStrings: LargeSimpleSharedStrings | null | undefined = hasSharedStrings ? undefined : []
  delete zip[workbookPath]
  delete zip[workbookRelationshipsPath]
  const workbookName = stringPool.intern(normalizeWorkbookName(fileName))
  const warnings = workbookDefinedNames.ignoredCount > 0 ? ['Some defined names were ignored during XLSX import.'] : []
  phaseRecorder.finish('zip-setup', zipSetupStart)
  const importedTables: WorkbookTableSnapshot[] = []
  const sheets: WorkbookSnapshot['sheets'] = []
  const previewSheets: ParsedWorksheet['preview'][] = []
  const sheetStats: ParsedWorksheet['stats'][] = []
  const styleCatalog = new Map<string, CellStyleRecord>()
  const scannedWorksheets: (ScannedWorksheet | undefined)[] = []
  const referencedSharedStringIndexes = new Set<number>()
  const materializeSheetsImmediately =
    materializeCells &&
    options.releaseZipSource !== true &&
    !hasSharedStrings &&
    !hasStyles &&
    !hasDrawingParts &&
    !hasSlicerConnectionParts
  const emptyStylesByIndex = new Map<number, Omit<CellStyleRecord, 'id'>>()
  const appendParsedWorksheet = (parsed: ParsedWorksheet): void => {
    sheets.push(parsed.sheet)
    previewSheets.push(parsed.preview)
    sheetStats.push(parsed.stats)
  }

  for (const [order, entry] of worksheetEntries.entries()) {
    const worksheetScanStart = phaseRecorder.start()
    let streamedWorksheetXml: string | undefined
    let streamedMetadataScan: LargeSimpleWorksheetScannedMetadata | undefined
    let retainedMetadataScan: LargeSimpleWorksheetScannedMetadata | undefined
    let cellScan: ImportedWorksheetCellScan | null = null
    if (!materializeCells && !materializeMetadata) {
      const headless = parseHeadlessLargeSimpleWorksheetFromChunks(
        (onChunk) => forEachInflatedXlsxZipEntryChunk(zip, entry.path, onChunk),
        order,
        { hasSharedStrings },
      )
      if (!headless) {
        return null
      }
      if (hasSharedStrings || headless.valueCellCount > 0) {
        cellScan = importedWorksheetCellScanFromHeadless(headless)
        delete zip[entry.path]
      }
    } else {
      const deferStyleCoordinates = shouldDeferLargeSimpleStyleCoordinates(zip, entry.path, { materializeCells, hasStyles })
      const streamed = parseLargeSimpleWorksheetCellsFromChunks(
        (onChunk) => forEachInflatedXlsxZipEntryChunk(zip, entry.path, onChunk),
        order,
        {
          hasSharedStrings,
          retainCells: materializeCells,
          retainStyleIndexes: materializeCells && hasStyles,
          retainStyleCoordinates: materializeCells && hasStyles && !deferStyleCoordinates,
          sharedStrings: fallbackSharedStrings ?? [],
          deferSharedStrings: materializeCells && hasSharedStrings,
          retainMetadataXml: materializeMetadata,
          sheetName: entry.name,
          stringPool,
          deduplicateStrings: deduplicateInlineStrings,
          deduplicateFormulas: deduplicateFormulaStrings,
          ...(options.allowUnsupportedFormulaText === undefined
            ? {}
            : { allowUnsupportedFormulaText: options.allowUnsupportedFormulaText }),
          ...(options.allowUnsupportedCellMetadata === undefined
            ? {}
            : { allowUnsupportedCellMetadata: options.allowUnsupportedCellMetadata }),
          maxDimensionCellPreallocation: maxPreallocatedWorksheetCells(zip, entry.path),
        },
      )
      if (!streamed) {
        return null
      }
      if (hasSharedStrings || streamed.cellScan.valueCellCount > 0) {
        cellScan = streamed.cellScan
        streamedWorksheetXml = streamed.metadataXml
        streamedMetadataScan = internLargeSimpleWorksheetMetadata(streamed.metadata, stringPool)
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
      cellScan = parseLargeSimpleWorksheetCells(worksheetBytes, fallbackSharedStrings ?? [], order, {
        retainCells: materializeCells,
        stringPool,
        deduplicateStrings: deduplicateInlineStrings,
        deduplicateFormulas: deduplicateFormulaStrings,
        ...(options.allowUnsupportedFormulaText === undefined ? {} : { allowUnsupportedFormulaText: options.allowUnsupportedFormulaText }),
      })
    }
    if (!cellScan) {
      return null
    }
    retainedMetadataScan = streamedMetadataScan
    const sharedStringIndexes = new Set<number>()
    if (hasSharedStrings) {
      cellScan.arena.collectSharedStringIndexes(sharedStringIndexes)
      for (const index of sharedStringIndexes) {
        referencedSharedStringIndexes.add(index)
      }
    }
    phaseRecorder.finish('worksheet-scan', worksheetScanStart)
    collectLargeSimpleImportGarbage()
    const metadataParsingStart = phaseRecorder.start()
    let worksheetXml: string | undefined
    let metadataInput: LargeSimpleSheetMetadataInput = {}
    const needsWorksheetXml =
      materializeMetadata &&
      (streamedWorksheetXml !== undefined || (worksheetBytes ? needsLargeSimpleWorksheetMetadataXml(worksheetBytes) : false))
    if (needsWorksheetXml) {
      worksheetXml = streamedWorksheetXml ?? (worksheetBytes ? readLargeSimpleWorksheetMetadataXml(worksheetBytes) : undefined)
      if (!worksheetXml) {
        return null
      }
      const sheetTables = streamedMetadataScan?.tableRelationshipIds
        ? undefined
        : /<(?:[A-Za-z_][\w.-]*:)?tableParts\b/u.test(worksheetXml)
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
        const hyperlinks = streamedMetadataScan?.hyperlinks
          ? undefined
          : readLargeSimpleSheetHyperlinks(zip, entry.name, entry.path, worksheetXml)
        if (hyperlinks === null) {
          return null
        }
        const filters = streamedMetadataScan?.filters ? [] : readImportedSheetAutoFilters(entry.name, worksheetXml)
        const conditionalFormatArtifacts = hasConditionalFormats
          ? readImportedSheetConditionalFormatArtifactsFromWorksheetXml(worksheetXml)
          : undefined
        metadataInput = appendLargeSimpleConditionalFormats(
          {
            ...(hyperlinks ? { hyperlinks } : {}),
            ...(filters.length > 0 ? { filters } : {}),
            ...(conditionalFormatArtifacts ? { conditionalFormatArtifacts } : {}),
          },
          conditionalFormats,
        )
      } else {
        metadataInput = appendLargeSimpleConditionalFormats(metadataInput, conditionalFormats)
      }
    }
    if (streamedMetadataScan?.conditionalFormats && streamedMetadataScan.conditionalFormats.length > 0) {
      metadataInput = appendLargeSimpleConditionalFormats(metadataInput, streamedMetadataScan.conditionalFormats)
    }
    if (materializeMetadata && streamedMetadataScan?.conditionalFormattingXml && streamedMetadataScan.conditionalFormattingXml.length > 0) {
      const conditionalFormats = readImportedSheetConditionalFormatsFromElementXml(
        zip,
        entry.name,
        streamedMetadataScan.conditionalFormattingXml,
      )
      const conditionalFormatArtifacts = materializeCells
        ? readImportedSheetConditionalFormatArtifactsFromElementXml(streamedMetadataScan.conditionalFormattingXml)
        : undefined
      metadataInput = appendLargeSimpleConditionalFormats(
        {
          ...metadataInput,
          ...(conditionalFormatArtifacts ? { conditionalFormatArtifacts } : {}),
        },
        conditionalFormats,
      )
      retainedMetadataScan = withoutLargeSimpleConditionalFormattingXml(streamedMetadataScan)
    }
    if (materializeMetadata && streamedMetadataScan?.dataValidations && streamedMetadataScan.dataValidations.length > 0) {
      metadataInput = {
        ...metadataInput,
        validations: [...(metadataInput.validations ?? []), ...streamedMetadataScan.dataValidations],
      }
    }
    if (materializeMetadata && streamedMetadataScan?.tableRelationshipIds && streamedMetadataScan.tableRelationshipIds.length > 0) {
      const sheetTables = readImportedSheetTablesFromRelationshipIds(zip, entry.name, entry.path, streamedMetadataScan.tableRelationshipIds)
      if (sheetTables) {
        importedTables.push(...sheetTables)
      }
    }
    if (materializeCells) {
      const printPageSetup =
        streamedMetadataScan?.printPageSetup ?? (worksheetXml ? readLargeSimpleSheetPrintPageSetup(worksheetXml) : undefined)
      const printMetadata = readLargeSimpleSheetPrintMetadata(zip, entry.path, printPageSetup)
      if (printMetadata === null) {
        return null
      }
      metadataInput = { ...metadataInput, ...printMetadata }
    }
    const streamedHyperlinks =
      materializeCells && streamedMetadataScan?.hyperlinks
        ? resolveLargeSimpleSheetHyperlinks(zip, entry.name, entry.path, streamedMetadataScan.hyperlinks)
        : undefined
    if (streamedHyperlinks === null) {
      return null
    }
    if (streamedHyperlinks) {
      metadataInput = { ...metadataInput, hyperlinks: streamedHyperlinks }
    }
    if (materializeCells && streamedMetadataScan?.filters && streamedMetadataScan.filters.length > 0) {
      metadataInput = { ...metadataInput, filters: [...streamedMetadataScan.filters] }
    }
    worksheetBytes = undefined
    if (materializeSheetsImmediately && !retainedMetadataScan?.controlArtifacts) {
      phaseRecorder.finish('metadata-parsing', metadataParsingStart)
      const snapshotMaterializationStart = phaseRecorder.start()
      appendParsedWorksheet(
        buildParsedWorksheet(entry.name, order, cellScan, worksheetXml, retainedMetadataScan, metadataInput, {
          materializeCells,
          releaseArenaAfterMaterialization: options.releaseArenaAfterMaterialization !== false,
          styleCatalog,
          stylesByIndex: emptyStylesByIndex,
        }),
      )
      phaseRecorder.finish('public-snapshot-materialization', snapshotMaterializationStart)
      continue
    }
    scannedWorksheets.push({
      name: entry.name,
      order,
      cellScan,
      worksheetXml,
      metadataScan: retainedMetadataScan,
      metadataInput,
      sharedStringIndexes,
    })
    phaseRecorder.finish('metadata-parsing', metadataParsingStart)
  }
  const sharedStringResolutionStart = phaseRecorder.start()
  let sharedStrings: LargeSimpleSharedStrings = fallbackSharedStrings ?? []
  if (materializeCells && hasSharedStrings && referencedSharedStringIndexes.size > 0) {
    const referencedSharedStrings = fallbackSharedStrings ?? readReferencedLargeSimpleSharedStrings(zip, referencedSharedStringIndexes)
    if (referencedSharedStrings === null) {
      return null
    }
    sharedStrings = referencedSharedStrings
  }
  delete zip[sharedStringsPath]
  if (materializeCells && hasSharedStrings && referencedSharedStringIndexes.size > 0) {
    for (const [index, scanned] of scannedWorksheets.entries()) {
      if (!scanned || scanned.sharedStringIndexes.size === 0) {
        continue
      }
      if (!hasReferencedLargeSimpleRichSharedStrings(sharedStrings, scanned.sharedStringIndexes)) {
        if (scanned.cellScan.arena.resolveSharedStrings(sharedStrings) === null) {
          return null
        }
        scannedWorksheets[index] = {
          ...scanned,
          sharedStringIndexes: emptySharedStringIndexes,
        }
        continue
      }
      const sheetSharedStrings = createLargeSimpleSharedStringSubset(sharedStrings, scanned.sharedStringIndexes)
      if (sheetSharedStrings === null) {
        return null
      }
      scannedWorksheets[index] = {
        ...scanned,
        sharedStrings: sheetSharedStrings,
      }
    }
    sharedStrings = []
  }
  referencedSharedStringIndexes.clear()
  fallbackSharedStrings = null
  phaseRecorder.finish('shared-string-resolution', sharedStringResolutionStart)
  collectLargeSimpleImportGarbage()
  const styleParsingStart = phaseRecorder.start()
  const requiredStyleIndexes = new Set<number>()
  for (const scanned of scannedWorksheets) {
    if (!scanned) {
      continue
    }
    scanned.cellScan.styleIndexes.collectRequiredStyleIndexes(requiredStyleIndexes)
  }
  const parsedStylesByIndex =
    materializeCells && hasStyles
      ? readLargeSimpleWorkbookStylesFromChunks(
          (onChunk) => forEachInflatedXlsxZipEntryChunk(zip, stylesPath, onChunk),
          requiredStyleIndexes,
        )
      : new Map()
  const stylesByIndex = parsedStylesByIndex ?? new Map()
  if (parsedStylesByIndex === null) {
    warnings.push('Some cell styles were ignored during XLSX import.')
  }
  requiredStyleIndexes.clear()
  if (
    !prepareLargeSimpleStyleIndexes(zip, worksheetEntries, scannedWorksheets, stylesByIndex, {
      hasSharedStrings,
      ...(options.allowUnsupportedFormulaText === undefined ? {} : { allowUnsupportedFormulaText: options.allowUnsupportedFormulaText }),
      ...(options.allowUnsupportedCellMetadata === undefined ? {} : { allowUnsupportedCellMetadata: options.allowUnsupportedCellMetadata }),
    })
  ) {
    return null
  }
  delete zip[stylesPath]
  phaseRecorder.finish('style-parsing', styleParsingStart)
  const importedDrawingArtifacts =
    materializeCells && hasDrawingParts
      ? readImportedWorkbookDrawingArtifactsFromWorksheetRelationships(
          zip,
          scannedWorksheets.flatMap((scanned) => {
            if (!scanned) {
              return []
            }
            const drawingRelationshipId = drawingRelationshipIdForScannedWorksheet(scanned)
            return [
              {
                name: scanned.name,
                path: worksheetEntries[scanned.order]?.path ?? '',
                ...(drawingRelationshipId ? { drawingRelationshipId } : {}),
              },
            ]
          }),
        )
      : null
  const importedSlicerConnectionArtifacts =
    materializeCells && hasSlicerConnectionParts
      ? readImportedWorkbookSlicerConnectionArtifactsFromSheets(
          zip,
          largeSimpleSlicerConnectionSheetSources(scannedWorksheets, worksheetEntries),
          {
            workbookXml,
            workbookRelationshipsXml,
          },
        )
      : undefined
  const importedControlArtifacts = materializeCells
    ? readImportedWorkbookControlArtifactsFromSheetSources(zip, largeSimpleControlArtifactSheetSources(scannedWorksheets, worksheetEntries))
    : undefined
  const importedLegacyCommentVmlBySheet =
    materializeCells && hasLegacyCommentParts
      ? readImportedWorkbookLegacyCommentVmlFromSheetSources(
          zip,
          largeSimpleLegacyCommentVmlSheetSources(scannedWorksheets, worksheetEntries),
        )
      : null
  if (options.releaseZipSource === true) {
    const zipSourceReleaseStart = phaseRecorder.start()
    const zipSourceBytesBeforeRelease = readLazyXlsxZipSourceByteLength(zip)
    const artifactReleasePlan = prepareLargeSimplePackageArtifactsForZipRelease({
      ...(options.maxMaterializedLazyPackageArtifactBytes !== undefined
        ? { maxMaterializedBytes: options.maxMaterializedLazyPackageArtifactBytes }
        : {}),
      preservedArtifacts: [
        importedDataModelArtifacts,
        importedSlicerConnectionArtifacts,
        importedDrawingArtifacts?.artifacts,
        importedChartDrawingArtifacts?.drawingArtifacts.artifacts,
        importedChartDrawingArtifacts?.chartArtifacts.artifacts,
      ],
      opaqueArtifacts: [importedPivotArtifacts?.artifacts],
    })
    const retainZipSourceForLazyPackageArtifacts = zipSourceBytesBeforeRelease !== undefined && artifactReleasePlan.retainZipSource
    if (!retainZipSourceForLazyPackageArtifacts) {
      releaseLazyXlsxZipSource(zip)
    }
    const ownedSourceReleaseEvidence = retainZipSourceForLazyPackageArtifacts ? undefined : options.releaseOwnedSourceBytes?.()
    phaseRecorder.finish('zip-source-release', zipSourceReleaseStart, {
      ...(zipSourceBytesBeforeRelease !== undefined ? { zipSourceBytesBeforeRelease } : {}),
      ...(zipSourceBytesBeforeRelease !== undefined ? { zipSourceBytesAfterRelease: readLazyXlsxZipSourceByteLength(zip) ?? 0 } : {}),
      ...ownedSourceReleaseEvidence,
    })
  }
  for (const [index, scanned] of scannedWorksheets.entries()) {
    if (!scanned) {
      continue
    }
    const snapshotMaterializationStart = phaseRecorder.start()
    const resolvedRichTextCells =
      materializeCells && hasSharedStrings && scanned.sharedStringIndexes.size > 0
        ? scanned.cellScan.arena.retainSharedStringReferences(scanned.sharedStrings ?? sharedStrings)
        : []
    if (resolvedRichTextCells === null) {
      return null
    }
    const cellScan = {
      ...scanned.cellScan,
      richTextCells: mergeWorkbookRichTextCells(scanned.cellScan.richTextCells, resolvedRichTextCells),
    }
    const drawingArtifacts =
      importedChartDrawingArtifacts?.drawingArtifacts.sheetArtifactsByName.get(scanned.name) ??
      importedDrawingArtifacts?.sheetArtifactsByName.get(scanned.name)
    const controlArtifacts = importedControlArtifacts?.sheetArtifactsByName.get(scanned.name)
    const pivotArtifacts = sheetPivotArtifactsWithStreamedDefinitions(
      importedPivotArtifacts?.sheetArtifactsByName.get(scanned.name),
      scanned.metadataScan?.pivotTableDefinitionsXml,
    )
    const legacyCommentVml = importedLegacyCommentVmlBySheet?.get(scanned.name)
    const parsed = buildParsedWorksheet(
      scanned.name,
      scanned.order,
      cellScan,
      scanned.worksheetXml,
      scanned.metadataScan,
      {
        ...scanned.metadataInput,
        ...(drawingArtifacts ? { drawingArtifacts } : {}),
        ...(controlArtifacts ? { controlArtifacts } : {}),
        ...(pivotArtifacts ? { pivotArtifacts } : {}),
        ...(legacyCommentVml
          ? {
              legacyCommentVml: {
                ...legacyCommentVml,
                commentSignature: legacyCommentThreadSignature(undefined),
              },
            }
          : {}),
      },
      {
        materializeCells,
        releaseArenaAfterMaterialization: options.releaseArenaAfterMaterialization !== false,
        styleCatalog,
        stylesByIndex,
      },
    )
    appendParsedWorksheet(parsed)
    scannedWorksheets[index] = undefined
    phaseRecorder.finish('public-snapshot-materialization', snapshotMaterializationStart)
  }
  sharedStrings = []
  stringPool.release()
  const sortedImportedTables =
    importedTables.length > 0 ? importedTables.toSorted((left, right) => left.name.localeCompare(right.name)) : undefined
  const hasFormulaCells = sheetStats.some((entry) => entry.formulaCellCount > 0)
  const workbookMetadata =
    workbookDefinedNames.definedNames ||
    importedWorkbookProperties ||
    importedWorkbookDocumentProperties ||
    importedChartDrawingArtifacts?.drawingArtifacts.artifacts ||
    importedDrawingArtifacts?.artifacts ||
    importedChartDrawingArtifacts?.chartArtifacts.artifacts ||
    importedChartDrawingArtifacts?.chartArtifacts.chartSheetArtifacts ||
    importedChartDrawingArtifacts?.charts ||
    importedPivotArtifacts?.artifacts ||
    importedControlArtifacts?.artifacts ||
    sortedImportedTables ||
    styleCatalog.size > 0 ||
    importedDataModelArtifacts ||
    importedExternalLinkArtifacts ||
    importedSlicerConnectionArtifacts ||
    importedWorkbookCellMetadata ||
    hasFormulaCells
      ? {
          ...(importedWorkbookProperties ? { properties: importedWorkbookProperties } : {}),
          ...(importedWorkbookDocumentProperties ? { documentPropertyArtifacts: importedWorkbookDocumentProperties } : {}),
          ...(workbookDefinedNames.definedNames ? { definedNames: workbookDefinedNames.definedNames } : {}),
          ...(importedChartDrawingArtifacts?.drawingArtifacts.artifacts
            ? { drawingArtifacts: importedChartDrawingArtifacts.drawingArtifacts.artifacts }
            : importedDrawingArtifacts?.artifacts
              ? { drawingArtifacts: importedDrawingArtifacts.artifacts }
              : {}),
          ...(importedChartDrawingArtifacts?.chartArtifacts.artifacts
            ? { chartArtifacts: importedChartDrawingArtifacts.chartArtifacts.artifacts }
            : {}),
          ...(importedChartDrawingArtifacts?.chartArtifacts.chartSheetArtifacts
            ? { chartSheetArtifacts: importedChartDrawingArtifacts.chartArtifacts.chartSheetArtifacts }
            : {}),
          ...(importedChartDrawingArtifacts?.charts ? { charts: importedChartDrawingArtifacts.charts } : {}),
          ...(importedPivotArtifacts?.artifacts ? { pivotArtifacts: importedPivotArtifacts.artifacts } : {}),
          ...(importedControlArtifacts?.artifacts ? { controlArtifacts: importedControlArtifacts.artifacts } : {}),
          ...(sortedImportedTables ? { tables: sortedImportedTables } : {}),
          ...(styleCatalog.size > 0 ? { styles: [...styleCatalog.values()] } : {}),
          ...(importedDataModelArtifacts ? { dataModelArtifacts: importedDataModelArtifacts } : {}),
          ...(importedExternalLinkArtifacts ? { externalLinkArtifacts: importedExternalLinkArtifacts } : {}),
          ...(importedSlicerConnectionArtifacts ? { slicerConnectionArtifacts: importedSlicerConnectionArtifacts } : {}),
          ...(importedWorkbookCellMetadata ? { cellMetadata: importedWorkbookCellMetadata } : {}),
          ...(hasFormulaCells
            ? {
                calculationSettings: {
                  mode: 'automatic' as const,
                  compatibilityMode: 'excel-modern' as const,
                  fullCalcOnLoad: false,
                  forceFullCalc: false,
                },
              }
            : {}),
        }
      : undefined
  const runtimeSheetCells = sheetStats.flatMap((entry, index) => {
    const sheet = sheets[index]
    const usedRange = entry.dimension.usedRange
    if (
      !sheet ||
      usedRange === null ||
      usedRange.startRow !== 0 ||
      usedRange.startColumn !== 0 ||
      entry.cellCount !== sheet.cells.length ||
      entry.cellCount !== entry.dimension.rowCount * entry.dimension.columnCount
    ) {
      return []
    }
    return [
      {
        sheetName: sheet.name,
        coords: [],
        coordinateOrder: 'dense-row-major' as const,
        dimensions: {
          width: entry.dimension.columnCount,
          height: entry.dimension.rowCount,
        },
        cellCount: entry.cellCount,
      },
    ]
  })
  const snapshot: WorkbookSnapshot = {
    version: 1,
    workbook: {
      name: workbookName,
      ...(workbookMetadata ? { metadata: workbookMetadata } : {}),
    },
    sheets,
  }
  const stats: LargeSimpleXlsxImportStats = {
    sheetCount: sheets.length,
    cellCount: sheetStats.reduce((sum, entry) => sum + entry.cellCount, 0),
    formulaCellCount: sheetStats.reduce((sum, entry) => sum + entry.formulaCellCount, 0),
    valueCellCount: sheetStats.reduce((sum, entry) => sum + entry.valueCellCount, 0),
    definedNameCount: workbookDefinedNames.definedNames?.length ?? 0,
    tableCount: sortedImportedTables?.length ?? sheetStats.reduce((sum, entry) => sum + entry.tableCount, 0),
    mergeCount: sheetStats.reduce((sum, entry) => sum + entry.mergeCount, 0),
    conditionalFormatCount: sheetStats.reduce((sum, entry) => sum + entry.conditionalFormatCount, 0),
    dataValidationCount: sheetStats.reduce((sum, entry) => sum + entry.dataValidationCount, 0),
    warningCount: warnings.length,
    dimensions: sheetStats.map((entry) => entry.dimension),
    phaseTelemetry: phaseRecorder.entries(),
  }

  return {
    snapshot:
      runtimeSheetCells.length > 0
        ? attachRuntimeImage(snapshot, {
            version: 1,
            templateBank: [],
            formulaInstances: [],
            formulaValues: [],
            sheetCells: runtimeSheetCells,
          })
        : snapshot,
    workbookName,
    sheetNames: workbookSheets.map((entry) => entry.name),
    warnings,
    preview: createWorkbookPreview({
      contentType: XLSX_CONTENT_TYPE,
      fileName,
      fileSizeBytes: source.byteLength,
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
  input: LargeSimpleSheetMetadataInput = {},
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
    (worksheetXml ? readLargeSimpleConditionalFormattingBlockCount(worksheetXml) : (cellScan.conditionalFormatCount ?? 0))
  const conditionalFormats = normalizeLargeSimpleConditionalFormatIds(sheetName, input.conditionalFormats)
  const dataValidationCount = input.validations?.length ?? cellScan.dataValidationCount ?? 0
  const styleRanges =
    options.materializeCells && options.styleCatalog && options.stylesByIndex
      ? buildLargeSimpleStyleRanges(sheetName, cellScan, options.stylesByIndex, options.styleCatalog)
      : []
  const preview = createSheetPreview({
    name: sheetName,
    rowCount: cellScan.rowCount,
    columnCount: cellScan.columnCount,
    nonEmptyCellCount: cellScan.cellCount,
    readCellText: (row, column) => cellScan.arena.readPreviewText(row, column),
  })
  const useLazyCells = options.materializeCells && cellScan.cellCount > lazySheetCellMaterializationThreshold
  const cells = options.materializeCells
    ? useLazyCells
      ? cellScan.arena.createLazySheetCells(cellScan.sheetIndex)
      : cellScan.arena.materializeSheetCells(cellScan.sheetIndex)
    : []
  const cellMetadataRefs = buildLargeSimpleCellMetadataReferenceSnapshots(metadataScan?.cellMetadataRefs, cells, cellScan, useLazyCells)
  const metadata: SheetMetadataSnapshot = {
    ...(columns.entries.length > 0 ? { columns: columns.entries } : {}),
    ...(rows.entries.length > 0 ? { rows: rows.entries } : {}),
    ...(columns.metadata.length > 0 ? { columnMetadata: columns.metadata } : {}),
    ...(rows.metadata.length > 0 ? { rowMetadata: rows.metadata } : {}),
    ...(sheetFormatPr ? { sheetFormatPr } : {}),
    ...(styleRanges.length > 0 ? { styleRanges } : {}),
    ...(merges.length > 0 ? { merges } : {}),
    ...(input.drawingArtifacts ? { drawingArtifacts: input.drawingArtifacts } : {}),
    ...(input.controlArtifacts ? { controlArtifacts: input.controlArtifacts } : {}),
    ...(input.pivotArtifacts ? { pivotArtifacts: input.pivotArtifacts } : {}),
    ...(input.legacyCommentVml ? { legacyCommentVml: input.legacyCommentVml } : {}),
    ...(input.sheetProtection ? { sheetProtection: input.sheetProtection } : {}),
    ...(input.filters ? { filters: input.filters } : {}),
    ...(input.hyperlinks ? { hyperlinks: input.hyperlinks } : {}),
    ...(input.validations ? { validations: input.validations } : {}),
    ...(conditionalFormats ? { conditionalFormats } : {}),
    ...(input.conditionalFormatArtifacts ? { conditionalFormatArtifacts: input.conditionalFormatArtifacts } : {}),
    ...(cellMetadataRefs ? { cellMetadataRefs } : {}),
    ...(input.printerSettings ? { printerSettings: input.printerSettings } : {}),
    ...(input.printPageSetup ? { printPageSetup: input.printPageSetup } : {}),
    ...(cellScan.richTextCells.length > 0 ? { richTextArtifacts: { cells: cellScan.richTextCells } } : {}),
  }
  const sheet: WorkbookSnapshot['sheets'][number] = {
    id: order + 1,
    name: sheetName,
    order,
    ...(Object.keys(metadata).length > 0 ? { metadata } : {}),
    cells,
  }
  const parsed: ParsedWorksheet = {
    sheet,
    preview,
    stats: {
      cellCount: cellScan.cellCount,
      formulaCellCount: cellScan.formulaCellCount,
      valueCellCount: cellScan.valueCellCount,
      tableCount: cellScan.tableCount ?? 0,
      mergeCount,
      conditionalFormatCount,
      dataValidationCount,
      dimension: {
        sheetName,
        rowCount: cellScan.rowCount,
        columnCount: cellScan.columnCount,
        nonEmptyCellCount: cellScan.cellCount,
        usedRange: cellScan.usedRange,
      },
    },
  }
  if (options.releaseArenaAfterMaterialization === true && !useLazyCells) {
    cellScan.arena.release()
    cellScan.styleIndexes.release()
  } else if (options.releaseArenaAfterMaterialization === true) {
    cellScan.arena.releaseMaterializationScratch()
    cellScan.styleIndexes.release()
  }
  return parsed
}

function sheetPivotArtifactsWithStreamedDefinitions(
  artifacts: SheetMetadataSnapshot['pivotArtifacts'],
  pivotTableDefinitionsXml: string | undefined,
): SheetMetadataSnapshot['pivotArtifacts'] {
  if (!pivotTableDefinitionsXml) {
    return artifacts
  }
  return {
    relationships: artifacts?.relationships ?? [],
    pivotTableDefinitionsXml,
  }
}

function drawingRelationshipIdForScannedWorksheet(scanned: ScannedWorksheet): string | undefined {
  return (
    scanned.metadataScan?.drawingRelationshipId ??
    (scanned.worksheetXml ? readLargeSimpleDrawingRelationshipId(scanned.worksheetXml) : undefined)
  )
}
