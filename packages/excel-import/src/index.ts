import * as XLSX from 'xlsx'
import type { Unzipped } from 'fflate'
import type { CsvParseOptions } from '@bilig/core'
import type {
  CellStyleRecord,
  WorkbookCommentThreadSnapshot,
  WorkbookLegacyCommentVmlSnapshot,
  WorkbookMetadataSnapshot,
  WorkbookSnapshot,
} from '@bilig/protocol'
import { readImportedArrayFormulaSpills, readImportedWorkbookArrayFormulas } from './xlsx-array-formulas.js'
import { buildColumnEntries, buildRowEntries } from './xlsx-axis-entries.js'
import { readImportedWorkbookCalculationSettings, readImportedWorkbookCalculationWarnings } from './xlsx-calculation-settings.js'
import { shouldUseCachedFormulaOpenMode } from './xlsx-cached-formula-open-mode.js'
import { buildImportedCellMetadataReferenceSnapshots, readImportedWorkbookCellMetadata } from './xlsx-cell-metadata.js'
import { readImportedWorkbookChartDrawingArtifacts } from './xlsx-import-chart-drawing-artifacts.js'
import { legacyCommentThreadSignature, readImportedWorkbookLegacyCommentVml, type ImportedLegacyCommentVml } from './xlsx-comment-vml.js'
import { readImportedSheetComments } from './xlsx-comments.js'
import { readImportedWorkbookConditionalFormatArtifacts, readImportedWorkbookConditionalFormats } from './xlsx-conditional-formats.js'
import { readImportedWorkbookControlArtifacts } from './xlsx-control-artifacts.js'
import { readImportedWorkbookDataModelArtifacts } from './xlsx-data-model-artifacts.js'
import { readImportedWorkbookDataTableFormulas } from './xlsx-data-table-formulas.js'
import { readImportedDefinedNames } from './xlsx-defined-names.js'
import { readImportedWorkbookExternalLinkArtifacts } from './xlsx-external-link-artifacts.js'
import { readImportedWorkbookExternalConnections } from './xlsx-external-connections.js'
import { readImportedWorkbookFilters } from './xlsx-filters.js'
import { readImportedWorksheetFormulaManifests } from './xlsx-formulas.js'
import { readImportedWorkbookFormulaAudit } from './xlsx-formula-audit.js'
import { readImportedWorkbookFreezePanes } from './xlsx-freeze-panes.js'
import { readImportedWorkbookIgnoredErrors } from './xlsx-ignored-errors.js'
import { buildMergeEntries } from './xlsx-merge-entries.js'
import { readImportedWorkbookPivots } from './xlsx-pivots.js'
import { readImportedWorkbookPrintPageSetup } from './xlsx-print-page-setup.js'
import { readImportedWorkbookPrinterSettings } from './xlsx-printer-settings.js'
import { readImportedWorkbookProtectedRanges } from './xlsx-protected-ranges.js'
import { readImportedWorkbookRichTextArtifacts } from './xlsx-rich-text-artifacts.js'
import { readImportedWorkbookFileNumberFormats } from './xlsx-number-formats.js'
import { readImportedWorkbookSheetProperties } from './xlsx-sheet-properties.js'
import { readImportedWorkbookSheetProtections } from './xlsx-sheet-protection.js'
import { readImportedWorkbookSheetVisibilities } from './xlsx-sheet-visibility.js'
import { readImportedWorkbookSlicerConnectionArtifacts } from './xlsx-slicer-connection-artifacts.js'
import { readImportedWorkbookSorts } from './xlsx-sorts.js'
import { readImportedWorkbookSparklines } from './xlsx-sparklines.js'
import { prepareSheetJsParserXlsxBytes } from './xlsx-style-only-blank-cells.js'
import { mergeStyleRuns, styleRunsToRanges, type HorizontalStyleRun, type RectangularStyleRun } from './xlsx-style-runs.js'
import { readImportedWorkbookFileStyles, readImportedWorkbookSheetDimensions, readImportedWorkbookStyleArtifacts } from './xlsx-styles.js'
import { readImportedWorkbookSheetTabColors } from './xlsx-tab-colors.js'
import { readImportedWorkbookTables } from './xlsx-tables.js'
import { readImportedWorkbookThreadedCommentArtifacts } from './xlsx-threaded-comment-artifacts.js'
import { readImportedWorkbookDataValidations } from './xlsx-validations.js'
import { readImportedWorkbookViewState } from './xlsx-view-state.js'
import { readImportedWorkbookProtection } from './xlsx-workbook-protection.js'
import { workbookDirectorySheetPaths, workbookSheetPathsByName } from './xlsx-workbook-sheet-paths.js'
import { readImportedWorkbookDocumentPropertiesArtifacts, readImportedWorkbookProperties } from './xlsx-workbook-properties.js'
import { createSheetPreview, normalizeWorkbookName, toDisplayText, type ImportedWorkbookSheetPreview } from './workbook-import-helpers.js'
import {
  CSV_CONTENT_TYPE,
  LEGACY_XLS_CONTENT_TYPE,
  XLSB_CONTENT_TYPE,
  XLSM_CONTENT_TYPE,
  XLSX_CONTENT_TYPE,
  normalizeWorkbookImportContentType,
  type ExcelWorkbookImportContentType,
} from './workbook-import-content-types.js'
import { createWorkbookPreview } from './workbook-import-preview.js'
import type { ImportedWorkbook } from './workbook-import-result.js'
import { readImportedExternalLinkCaches, readImportedExternalWorkbookReferences } from './xlsx-external-references.js'
import { shouldUseDenseSheetJsParse } from './xlsx-dense-sheetjs-parse.js'
import { normalizeImportedFormulaSource } from './xlsx-formula-translation.js'
import { readImportedSheetHyperlinks } from './xlsx-hyperlinks.js'
import { collectStyleCandidateAddresses, internImportedStyle, readImportedXlsxCellStyle } from './xlsx-import-cell-styles.js'
import { buildImportedFormulaSnapshotCell } from './xlsx-import-formula-cells.js'
import { buildImportedSheetMetadata } from './xlsx-import-sheet-metadata.js'
import { buildImportedWorkbookMetadata } from './xlsx-import-workbook-metadata.js'
import {
  addWorkbookWarnings,
  dataTableFormulasWarning,
  externalPivotCachesWarning,
  externalWorkbookReferencesWarning,
  volatileFormulasWarning,
  workbookDefinedNamesReferenceExternalWorkbook,
} from './xlsx-import-warnings.js'
import { compareCellAddresses, readImportedLiteralCellValue, readImportedNumberFormat } from './xlsx-import-cell-values.js'
import {
  assertXlsxInspectionWithinMaterializationLimits,
  denseSheetJsByteThreshold,
  largeCalcChainStreamingFormulaThreshold,
  resolveXlsxImportLimits,
  shouldRetryDataOnlyLargeSimpleImport,
  type XlsxImportOptions,
} from './xlsx-import-limits.js'
import { tryInspectLargeSimpleXlsxHeadless, type LargeSimpleXlsxHeadlessInspectResult } from './xlsx-large-simple-headless-inspect.js'
import { tryImportLargeSimpleXlsx } from './xlsx-large-simple-import.js'
import {
  hasFullImporterOnlyPackageMetadata,
  shouldBypassLargeSimpleByteThresholdForPackageArtifacts,
} from './xlsx-large-simple-package-artifact-threshold.js'
import { createPreservedVbaProjectPayload, type PreservedVbaProjectCodeNames } from './xlsx-macros.js'
import { releaseOwnedXlsxSourceBytes, type OwnedXlsxSourceBytes } from './xlsx-owned-source-release.js'
import { attachImportedXlsxSourceBytes } from './xlsx-source-bytes.js'
import { worksheetCellAt, worksheetCellEntries, worksheetCellEntriesAtAddresses } from './xlsx-worksheet-cells.js'
import { readImportedWorksheetTextValues } from './xlsx-worksheet-text-values.js'
import { readLazyXlsxZipSource, readXlsxZipEntries, readXlsxZipEntriesLazy } from './xlsx-zip.js'
import { importCsv } from './csv-import.js'
import {
  attachImportedRuntimeCoordinates,
  createImportedRuntimeSheetCells,
  pushImportedSnapshotCell,
  type ImportedRuntimeCellCoordinate,
  type ImportedRuntimeSheetCells,
} from './imported-runtime-coordinates.js'

export { exportXlsx } from './xlsx-export.js'
export { importCsv } from './csv-import.js'
export { manualCalculationModeWarning, precisionAsDisplayedCalculationWarning } from './xlsx-calculation-settings.js'
export {
  dataTableFormulasWarning,
  externalPivotCachesWarning,
  externalWorkbookReferencesWarning,
  macroExecutionDeclinedWarning,
  volatileFormulasWarning,
} from './xlsx-import-warnings.js'
export { readImportedXlsxCellStyle } from './xlsx-import-cell-styles.js'
export { XlsxImportSizeLimitExceededError } from './xlsx-import-limits.js'
export type { ImportedWorkbookSheetPreview } from './workbook-import-helpers.js'
export type { ImportedWorkbookPreview } from './workbook-import-preview.js'
export type { ImportedWorkbook } from './workbook-import-result.js'
export type { XlsxImportLimits, XlsxImportOptions } from './xlsx-import-limits.js'
export {
  CSV_CONTENT_TYPE,
  EXCEL_WORKBOOK_IMPORT_CONTENT_TYPES,
  LEGACY_XLS_CONTENT_TYPE,
  WORKBOOK_IMPORT_CONTENT_TYPES,
  XLSB_CONTENT_TYPE,
  XLSM_CONTENT_TYPE,
  XLSX_CONTENT_TYPE,
  normalizeWorkbookImportContentType,
} from './workbook-import-content-types.js'
export type { ExcelWorkbookImportContentType, WorkbookImportContentType } from './workbook-import-content-types.js'

const largeWorkbookStyleCandidateThreshold = 100_000
const largeCalcChainStreamingByteThreshold = 5_000_000
const sheetJsBlankStyleStripMinCellCount = 1_000
const denseSheetJsMaxColumnCount = 128

export type CsvImportOptions = CsvParseOptions
export type XlsxHeadlessInspectResult = LargeSimpleXlsxHeadlessInspectResult

export interface WorkbookImportFileOptions {
  csv?: CsvImportOptions
  xlsx?: XlsxImportOptions
}

export class InvalidXlsxZipContainerError extends Error {
  constructor() {
    super('Invalid or corrupt XLSX zip container')
    this.name = 'InvalidXlsxZipContainerError'
  }
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function toUint8Array(value: unknown): Uint8Array | null {
  if (value instanceof Uint8Array) {
    return new Uint8Array(value)
  }
  if (value instanceof ArrayBuffer) {
    return new Uint8Array(value)
  }
  return null
}

function readNonEmptyString(value: unknown): string | undefined {
  if (typeof value !== 'string') {
    return undefined
  }
  const trimmed = value.trim()
  return trimmed.length > 0 ? trimmed : undefined
}

function addCandidateAddress(addressesBySheet: Map<string, Set<string>>, sheetName: string, address: string): boolean {
  const addresses = addressesBySheet.get(sheetName) ?? new Set<string>()
  const previousSize = addresses.size
  addresses.add(address)
  if (addresses.size > previousSize) {
    addressesBySheet.set(sheetName, addresses)
    return true
  }
  return false
}

function addStyleArtifactCandidateAddresses(
  candidates: ReturnType<typeof collectStyleCandidateAddresses>,
  importedStyleArtifacts: ReturnType<typeof readImportedWorkbookStyleArtifacts>,
  maxCandidateCount: number,
): ReturnType<typeof collectStyleCandidateAddresses> {
  let count = candidates.count
  const addressesBySheet = new Map([...candidates.addressesBySheet].map(([sheetName, addresses]) => [sheetName, new Set(addresses)]))
  for (const [sheetName, artifacts] of importedStyleArtifacts.sheetArtifactsByName) {
    for (const entry of artifacts.cellStyleIndexes) {
      if (addCandidateAddress(addressesBySheet, sheetName, entry.address)) {
        count += 1
      }
      if (count > maxCandidateCount) {
        return candidates
      }
    }
    for (const address of artifacts.blankCellAddresses ?? []) {
      if (addCandidateAddress(addressesBySheet, sheetName, address)) {
        count += 1
      }
      if (count > maxCandidateCount) {
        return candidates
      }
    }
  }
  return { addressesBySheet, count }
}

function readImportedMacroCodeNames(workbook: XLSX.WorkBook): PreservedVbaProjectCodeNames {
  const workbookMetadata = isRecord(workbook.Workbook) ? workbook.Workbook : undefined
  const workbookProperties = isRecord(workbookMetadata?.['WBProps']) ? workbookMetadata['WBProps'] : undefined
  const workbookCodeName = readNonEmptyString(workbookProperties?.['CodeName'])
  const workbookSheets = workbookMetadata?.['Sheets']
  const sheetCodeNames = Array.isArray(workbookSheets)
    ? workbookSheets.flatMap((entry) => {
        if (!isRecord(entry)) {
          return []
        }
        const sheetName = readNonEmptyString(entry['name'])
        const codeName = readNonEmptyString(entry['CodeName'])
        return sheetName && codeName ? [{ sheetName, codeName }] : []
      })
    : []
  return {
    ...(workbookCodeName ? { workbookCodeName } : {}),
    ...(sheetCodeNames.length > 0 ? { sheetCodeNames } : {}),
  }
}

function buildImportedLegacyCommentVmlSnapshot(
  imported: ImportedLegacyCommentVml | undefined,
  commentThreads: readonly WorkbookCommentThreadSnapshot[],
): WorkbookLegacyCommentVmlSnapshot | undefined {
  if (!imported) {
    return undefined
  }
  return {
    ...imported,
    commentSignature: legacyCommentThreadSignature(commentThreads),
  }
}

function readValidXlsxZipContainer(bytes: Uint8Array, mode: 'eager' | 'lazy' = 'eager'): Unzipped {
  try {
    const zip = mode === 'lazy' ? readXlsxZipEntriesLazy(bytes) : readXlsxZipEntries(bytes)
    void zip['xl/workbook.xml']
    return zip
  } catch {
    throw new InvalidXlsxZipContainerError()
  }
}

function inspectLargeSimpleXlsxSource(
  data: Uint8Array,
  fileName: string,
  options: { readonly minByteLength?: number } = {},
): LargeSimpleXlsxHeadlessInspectResult | null {
  const inspectionZip = readValidXlsxZipContainer(data, 'lazy')
  return tryInspectLargeSimpleXlsxHeadless({ byteLength: data.byteLength }, fileName, inspectionZip, {
    allowUnsupportedWorksheetFeaturesForMetrics: true,
    ...(options.minByteLength !== undefined ? { minByteLength: options.minByteLength } : {}),
    releaseZipSource: true,
  })
}

function importSheetJsWorkbook(
  data: Uint8Array,
  fileName: string,
  contentType: ExcelWorkbookImportContentType,
  workbookZip: Unzipped | null,
  sourceBytesForUntouchedExport?: Uint8Array,
): ImportedWorkbook {
  const denseSheetJsParse = shouldUseDenseSheetJsParse(data, workbookZip, {
    maxColumnCount: denseSheetJsMaxColumnCount,
    minByteLength: denseSheetJsByteThreshold,
  })
  const parserData = workbookZip
    ? prepareSheetJsParserXlsxBytes(data, workbookZip, {
        minBlankCellCount: sheetJsBlankStyleStripMinCellCount,
        omitParserIgnoredPackageParts: true,
      })
    : data
  const workbook = XLSX.read(parserData, {
    type: 'array',
    cellFormula: true,
    cellNF: true,
    cellStyles: false,
    cellText: false,
    cellDates: false,
    bookFiles: true,
    bookVBA: true,
    dense: denseSheetJsParse,
  })
  const workbookName = normalizeWorkbookName(fileName)
  const sheetPathsByName = workbookSheetPathsByName(workbook)
  const fallbackSheetPaths = workbookDirectorySheetPaths(workbook)
  const warnings: string[] = []
  const importedDefinedNames = readImportedDefinedNames(workbook)
  addWorkbookWarnings(workbook, warnings, importedDefinedNames.ignoredCount)
  if (workbookDefinedNamesReferenceExternalWorkbook(workbook)) {
    warnings.push(externalWorkbookReferencesWarning)
  }
  const styleArtifactSource =
    contentType === XLSX_CONTENT_TYPE || contentType === XLSM_CONTENT_TYPE ? (workbookZip ?? parserData) : undefined
  const importedStyleArtifacts = readImportedWorkbookStyleArtifacts(workbook, workbook.SheetNames, styleArtifactSource)
  const baseStyleCandidates = collectStyleCandidateAddresses(workbook, workbook.SheetNames, largeWorkbookStyleCandidateThreshold)
  const styleCandidates = addStyleArtifactCandidateAddresses(
    baseStyleCandidates,
    importedStyleArtifacts,
    largeWorkbookStyleCandidateThreshold,
  )
  const importedWorkbookStyles =
    styleCandidates.count === 0 || styleCandidates.count > largeWorkbookStyleCandidateThreshold
      ? new Map<string, Map<string, Omit<CellStyleRecord, 'id'>>>()
      : readImportedWorkbookFileStyles(
          workbook,
          workbook.SheetNames,
          {
            styleCandidateAddressesBySheet: styleCandidates.addressesBySheet,
          },
          styleArtifactSource,
        )
  const importedWorkbookNumberFormats =
    styleCandidates.count === 0 || styleCandidates.count > largeWorkbookStyleCandidateThreshold
      ? new Map<string, Map<string, string>>()
      : readImportedWorkbookFileNumberFormats(
          workbook,
          workbook.SheetNames,
          {
            formatCandidateAddressesBySheet: styleCandidates.addressesBySheet,
          },
          styleArtifactSource,
        )
  const importedWorkbookSheetDimensions = readImportedWorkbookSheetDimensions(workbook, workbook.SheetNames, styleArtifactSource)
  const importedWorkbookProperties = workbookZip ? readImportedWorkbookProperties(workbookZip) : undefined
  const importedWorkbookDocumentProperties = workbookZip ? readImportedWorkbookDocumentPropertiesArtifacts(workbookZip) : undefined
  const importedWorkbookProtection = workbookZip ? readImportedWorkbookProtection(workbookZip) : undefined
  const importedCalculationSettings = workbookZip ? readImportedWorkbookCalculationSettings(workbookZip) : undefined
  const importedMacroPayload = toUint8Array(workbook.vbaraw)
  const importedMacroCodeNames = importedMacroPayload ? readImportedMacroCodeNames(workbook) : undefined
  const importedCellMetadata = workbookZip ? readImportedWorkbookCellMetadata(workbookZip, workbook.SheetNames) : undefined
  const importedChartDrawingArtifacts = workbookZip
    ? readImportedWorkbookChartDrawingArtifacts(workbookZip, workbook.SheetNames)
    : undefined
  const importedCharts = importedChartDrawingArtifacts?.charts
  const importedTables = workbookZip ? readImportedWorkbookTables(workbookZip, workbook.SheetNames) : undefined
  const importedControlArtifacts = workbookZip ? readImportedWorkbookControlArtifacts(workbookZip, workbook.SheetNames) : undefined
  const importedDataModelArtifacts = workbookZip ? readImportedWorkbookDataModelArtifacts(workbookZip) : undefined
  const importedExternalLinkArtifacts = workbookZip ? readImportedWorkbookExternalLinkArtifacts(workbookZip) : undefined
  const importedSlicerConnectionArtifacts = workbookZip
    ? readImportedWorkbookSlicerConnectionArtifacts(workbookZip, workbook.SheetNames)
    : undefined
  const importedArrayFormulasBySheet = workbookZip ? readImportedWorkbookArrayFormulas(workbookZip, workbook.SheetNames) : new Map()
  const importedDataTableFormulasBySheet = workbookZip ? readImportedWorkbookDataTableFormulas(workbookZip, workbook.SheetNames) : new Map()
  if (importedDataTableFormulasBySheet.size > 0) {
    warnings.push(dataTableFormulasWarning)
  }
  const importedPivots = workbookZip
    ? readImportedWorkbookPivots(workbookZip, workbook.SheetNames, importedTables, importedDefinedNames.definedNames)
    : undefined
  if (importedPivots?.hasExternalPivotCaches) {
    warnings.push(externalPivotCachesWarning)
  }
  const importedLegacyCommentVmlBySheet = workbookZip ? readImportedWorkbookLegacyCommentVml(workbookZip, workbook.SheetNames) : new Map()
  const importedPrinterSettingsBySheet = workbookZip ? readImportedWorkbookPrinterSettings(workbookZip, workbook.SheetNames) : new Map()
  const importedPrintPageSetupBySheet = workbookZip ? readImportedWorkbookPrintPageSetup(workbookZip, workbook.SheetNames) : new Map()
  const importedFiltersBySheet = workbookZip ? readImportedWorkbookFilters(workbookZip, workbook.SheetNames) : new Map()
  const importedFreezePanesBySheet = workbookZip ? readImportedWorkbookFreezePanes(workbookZip, workbook.SheetNames) : new Map()
  const importedSheetTabColorsBySheet = workbookZip ? readImportedWorkbookSheetTabColors(workbookZip, workbook.SheetNames) : new Map()
  const importedSheetPropertiesBySheet = workbookZip ? readImportedWorkbookSheetProperties(workbookZip, workbook.SheetNames) : new Map()
  const importedIgnoredErrorsBySheet = workbookZip ? readImportedWorkbookIgnoredErrors(workbookZip, workbook.SheetNames) : new Map()
  const importedSparklinesBySheet = workbookZip ? readImportedWorkbookSparklines(workbookZip, workbook.SheetNames) : new Map()
  const importedSheetVisibilitiesBySheet = readImportedWorkbookSheetVisibilities(workbook, workbook.SheetNames)
  const importedSheetProtectionsBySheet = workbookZip ? readImportedWorkbookSheetProtections(workbookZip, workbook.SheetNames) : new Map()
  const importedProtectedRangesBySheet = workbookZip ? readImportedWorkbookProtectedRanges(workbookZip, workbook.SheetNames) : new Map()
  const importedSortsBySheet = workbookZip ? readImportedWorkbookSorts(workbookZip, workbook.SheetNames) : new Map()
  const importedValidationsBySheet = workbookZip ? readImportedWorkbookDataValidations(workbookZip, workbook.SheetNames) : new Map()
  const conditionalFormatArtifactSource =
    contentType === XLSX_CONTENT_TYPE || contentType === XLSM_CONTENT_TYPE ? (workbookZip ?? data) : workbookZip
  const importedConditionalFormatsBySheet = conditionalFormatArtifactSource
    ? readImportedWorkbookConditionalFormats(conditionalFormatArtifactSource, workbook.SheetNames)
    : new Map()
  const importedConditionalFormatArtifactsBySheet = conditionalFormatArtifactSource
    ? readImportedWorkbookConditionalFormatArtifacts(conditionalFormatArtifactSource, workbook.SheetNames)
    : new Map()
  const importedExternalLinkCaches = workbookZip ? readImportedExternalLinkCaches(workbookZip) : new Map()
  const importedExternalWorkbookReferences = workbookZip ? readImportedExternalWorkbookReferences(workbookZip) : new Map()
  const importedExternalConnections = workbookZip ? readImportedWorkbookExternalConnections(workbookZip) : undefined
  const importedRichTextArtifactsBySheet = workbookZip ? readImportedWorkbookRichTextArtifacts(workbookZip, workbook.SheetNames) : new Map()
  const importedWorksheetTextValuesBySheet = workbookZip
    ? readImportedWorksheetTextValues(workbookZip, workbook.SheetNames, sheetPathsByName, fallbackSheetPaths)
    : new Map()
  const importedWorksheetFormulaManifestsBySheet = workbookZip
    ? readImportedWorksheetFormulaManifests(workbookZip, workbook.SheetNames, sheetPathsByName, fallbackSheetPaths)
    : new Map()
  const importedThreadedCommentArtifacts = workbookZip
    ? readImportedWorkbookThreadedCommentArtifacts(workbookZip, workbook.SheetNames)
    : undefined
  const importedViewState = workbookZip ? readImportedWorkbookViewState(workbookZip, workbook.SheetNames) : undefined
  const chartSheetNames = new Set(
    (importedChartDrawingArtifacts?.chartArtifacts.chartSheetArtifacts ?? []).map((artifact) => artifact.name),
  )

  let ignoredCommentsSeen = false
  let externalWorkbookReferenceWarningSeen = warnings.includes(externalWorkbookReferencesWarning)
  let volatileFormulaWarningSeen = false
  let formulaCellCount = 0
  let cachedFormulaValueCount = 0
  const styleCatalog = new Map<string, CellStyleRecord>()
  const unsupportedFormulaDependencies: NonNullable<WorkbookMetadataSnapshot['unsupportedFormulaDependencies']> = []
  const importedArrayFormulaSpills: NonNullable<WorkbookMetadataSnapshot['spills']> = []
  const previewSheets: ImportedWorkbookSheetPreview[] = []
  const runtimeSheetCells: ImportedRuntimeSheetCells[] = []
  const sheets = workbook.SheetNames.map((sheetName, order) => {
    const sheet = chartSheetNames.has(sheetName) ? undefined : workbook.Sheets[sheetName]
    if (!sheet) {
      runtimeSheetCells.push(createImportedRuntimeSheetCells({ sheetName, coords: [], width: 0, height: 0 }))
      previewSheets.push(
        createSheetPreview({
          name: sheetName,
          rowCount: 0,
          columnCount: 0,
          nonEmptyCellCount: 0,
          readCellText: () => '',
        }),
      )
      return {
        id: order + 1,
        name: sheetName,
        order,
        cells: [],
      }
    }

    const importedComments = readImportedSheetComments(sheetName, sheet)
    const importedWorksheetTextValues = importedWorksheetTextValuesBySheet.get(sheetName)
    const importedWorksheetFormulaManifests = importedWorksheetFormulaManifestsBySheet.get(sheetName)
    const importedHyperlinks = readImportedSheetHyperlinks(sheetName, sheet)
    if (importedComments.ignoredCount > 0 && !ignoredCommentsSeen) {
      ignoredCommentsSeen = true
      warnings.push('Some cell comments were ignored during XLSX import.')
    }
    const importedLegacyCommentVml = importedComments.commentThreads
      ? buildImportedLegacyCommentVmlSnapshot(importedLegacyCommentVmlBySheet.get(sheetName), importedComments.commentThreads)
      : undefined
    const importedArrayFormulaSheetSpills = readImportedArrayFormulaSpills(sheetName, sheet)
    if (importedArrayFormulaSheetSpills) {
      importedArrayFormulaSpills.push(...importedArrayFormulaSheetSpills)
    }
    const range = sheet['!ref'] ? XLSX.utils.decode_range(sheet['!ref']) : null
    const cells: WorkbookSnapshot['sheets'][number]['cells'] = []
    const runtimeCellCoords: ImportedRuntimeCellCoordinate[] = []
    const styleRuns: RectangularStyleRun[] = []
    let openStyleRunsByKey = new Map<string, RectangularStyleRun>()
    let activeStyleRow: number | null = null
    let activeStyleRun: HorizontalStyleRun | null = null
    let activeStyleRowRuns: HorizontalStyleRun[] = []
    const importedStylesByAddress = importedWorkbookStyles.get(sheetName)
    const importedFormatsByAddress = importedWorkbookNumberFormats.get(sheetName)
    const flushActiveStyleRun = () => {
      if (activeStyleRun) {
        activeStyleRowRuns.push(activeStyleRun)
        activeStyleRun = null
      }
    }
    const flushActiveStyleRow = () => {
      if (activeStyleRow === null) {
        return
      }
      flushActiveStyleRun()
      openStyleRunsByKey = mergeStyleRuns(activeStyleRow, activeStyleRowRuns, openStyleRunsByKey, styleRuns)
      activeStyleRowRuns = []
      activeStyleRow = null
    }
    const addStyleCell = (row: number, column: number, styleId: string) => {
      if (activeStyleRow === null) {
        activeStyleRow = row
      } else if (activeStyleRow !== row) {
        flushActiveStyleRow()
        activeStyleRow = row
      }
      if (activeStyleRun && activeStyleRun.styleId === styleId && activeStyleRun.endColumn + 1 === column) {
        activeStyleRun.endColumn = column
        return
      }
      flushActiveStyleRun()
      activeStyleRun = {
        styleId,
        startColumn: column,
        endColumn: column,
      }
    }
    const recordImportedFormulaDiagnostics = (result: NonNullable<ReturnType<typeof buildImportedFormulaSnapshotCell>>) => {
      formulaCellCount += 1
      if (result.hasCachedLiteral) {
        cachedFormulaValueCount += 1
      }
      if (!externalWorkbookReferenceWarningSeen && result.hasExternalWorkbookDependency) {
        externalWorkbookReferenceWarningSeen = true
        warnings.push(externalWorkbookReferencesWarning)
      }
      if (!volatileFormulaWarningSeen && result.hasVolatileFormula) {
        volatileFormulaWarningSeen = true
        warnings.push(volatileFormulasWarning)
      }
      if (result.unsupportedFormulaDependency) {
        unsupportedFormulaDependencies.push(result.unsupportedFormulaDependency)
      }
    }
    const rowCount = range ? range.e.r + 1 : 0
    const columnCount = range ? range.e.c + 1 : 0
    const importableAddresses =
      styleCandidates.count <= largeWorkbookStyleCandidateThreshold ? styleCandidates.addressesBySheet.get(sheetName) : undefined
    const sheetCellEntries = range
      ? importableAddresses
        ? worksheetCellEntriesAtAddresses(sheet, importableAddresses)
        : worksheetCellEntries(sheet)
      : []
    const seenCellAddresses = new Set<string>()
    for (const { address, cell, row, column } of sheetCellEntries) {
      seenCellAddresses.add(address)
      const nextCell: WorkbookSnapshot['sheets'][number]['cells'][number] = { address }
      const formulaManifest = importedWorksheetFormulaManifests?.get(address)
      const manifestFormula = formulaManifest?.formula
      const formula = typeof manifestFormula === 'string' && manifestFormula.trim().length > 0 ? manifestFormula : cell['f']
      const xmlTextValue = importedWorksheetTextValues?.get(address)
      if (typeof formula === 'string' && formula.trim().length > 0) {
        const formulaResult = buildImportedFormulaSnapshotCell({
          sheetName,
          address,
          formula,
          formulaManifest,
          cachedLiteral: xmlTextValue ?? readImportedLiteralCellValue(cell),
          tables: importedTables,
          externalLinkCaches: importedExternalLinkCaches,
          externalWorkbookReferences: importedExternalWorkbookReferences,
        })
        if (formulaResult) {
          Object.assign(nextCell, formulaResult.formulaCell)
          recordImportedFormulaDiagnostics(formulaResult)
        }
      } else {
        const literal = xmlTextValue ?? readImportedLiteralCellValue(cell)
        if (literal !== undefined) {
          nextCell.value = literal
        }
      }
      const importedFormat = readImportedNumberFormat(cell['z']) ?? importedFormatsByAddress?.get(address)
      if (importedFormat !== undefined) {
        nextCell.format = importedFormat
      }
      const importedStyle = importedStylesByAddress?.get(address) ?? readImportedXlsxCellStyle(cell['s'])
      if (importedStyle) {
        addStyleCell(row, column, internImportedStyle(importedStyle, styleCatalog))
      } else if (activeStyleRow === row) {
        flushActiveStyleRun()
      }
      if (nextCell.value !== undefined || nextCell.formula !== undefined || nextCell.format !== undefined) {
        pushImportedSnapshotCell(cells, runtimeCellCoords, nextCell, row, column)
      }
    }
    for (const [address, formulaManifest] of [...(importedWorksheetFormulaManifests?.entries() ?? [])]
      .filter(([candidateAddress, candidate]) => !seenCellAddresses.has(candidateAddress) && candidate.formula.trim().length > 0)
      .toSorted((left, right) => compareCellAddresses(left[0], right[0]))) {
      const decoded = XLSX.utils.decode_cell(address)
      seenCellAddresses.add(address)
      const formulaResult = buildImportedFormulaSnapshotCell({
        sheetName,
        address,
        formula: formulaManifest.formula,
        formulaManifest,
        cachedLiteral: importedWorksheetTextValues?.get(address),
        tables: importedTables,
        externalLinkCaches: importedExternalLinkCaches,
        externalWorkbookReferences: importedExternalWorkbookReferences,
      })
      if (!formulaResult) {
        continue
      }
      const nextCell: WorkbookSnapshot['sheets'][number]['cells'][number] = { ...formulaResult.formulaCell }
      const importedFormat = importedFormatsByAddress?.get(address)
      if (importedFormat !== undefined) {
        nextCell.format = importedFormat
      }
      const importedStyle = importedStylesByAddress?.get(address)
      if (importedStyle) {
        addStyleCell(decoded.r, decoded.c, internImportedStyle(importedStyle, styleCatalog))
      } else if (activeStyleRow === decoded.r) {
        flushActiveStyleRun()
      }
      recordImportedFormulaDiagnostics(formulaResult)
      pushImportedSnapshotCell(cells, runtimeCellCoords, nextCell, decoded.r, decoded.c)
    }
    const missingStyledAddresses = new Set([...(importedStylesByAddress?.keys() ?? []), ...(importedFormatsByAddress?.keys() ?? [])])
    for (const missingAddress of [...missingStyledAddresses]
      .filter((candidateAddress) => !seenCellAddresses.has(candidateAddress))
      .toSorted(compareCellAddresses)) {
      const decoded = XLSX.utils.decode_cell(missingAddress)
      seenCellAddresses.add(missingAddress)
      const importedStyle = importedStylesByAddress?.get(missingAddress)
      if (importedStyle) {
        addStyleCell(decoded.r, decoded.c, internImportedStyle(importedStyle, styleCatalog))
      } else if (activeStyleRow === decoded.r) {
        flushActiveStyleRun()
      }
      const importedFormat = importedFormatsByAddress?.get(missingAddress)
      if (importedFormat !== undefined) {
        pushImportedSnapshotCell(cells, runtimeCellCoords, { address: missingAddress, format: importedFormat }, decoded.r, decoded.c)
      }
    }
    for (const [address, value] of importedWorksheetTextValues ?? []) {
      if (!seenCellAddresses.has(address)) {
        const decoded = XLSX.utils.decode_cell(address)
        pushImportedSnapshotCell(cells, runtimeCellCoords, { address, value }, decoded.r, decoded.c)
      }
    }
    flushActiveStyleRow()
    const styleRanges = styleRunsToRanges(sheetName, styleRuns)
    runtimeSheetCells.push(
      createImportedRuntimeSheetCells({
        sheetName,
        coords: runtimeCellCoords,
        width: columnCount,
        height: rowCount,
      }),
    )

    previewSheets.push(
      createSheetPreview({
        name: sheetName,
        rowCount,
        columnCount,
        nonEmptyCellCount: cells.length,
        readCellText: (row, col) => {
          const address = XLSX.utils.encode_cell({ r: row, c: col })
          const cell = worksheetCellAt(sheet, row, col)
          const manifestFormula = importedWorksheetFormulaManifests?.get(address)?.formula
          const formula = typeof manifestFormula === 'string' && manifestFormula.trim().length > 0 ? manifestFormula : cell?.['f']
          if (typeof formula === 'string' && formula.trim().length > 0) {
            return `=${normalizeImportedFormulaSource(formula)}`
          }
          if (!cell) {
            return toDisplayText(importedWorksheetTextValues?.get(address))
          }
          return toDisplayText(readImportedLiteralCellValue(cell) ?? importedWorksheetTextValues?.get(address))
        },
      }),
    )

    const importedSheetDimensions = importedWorkbookSheetDimensions.get(sheetName)
    const rows = importedSheetDimensions?.rows ?? buildRowEntries(sheet['!rows'])
    const columns = importedSheetDimensions ? importedSheetDimensions.columns : buildColumnEntries(sheet['!cols'])
    const rowMetadata = importedSheetDimensions?.rowMetadata
    const columnMetadata = importedSheetDimensions?.columnMetadata
    const sheetFormatPr = importedSheetDimensions?.sheetFormatPr
    const importedFreezePane = importedFreezePanesBySheet.get(sheetName)
    const importedSheetTabColor = importedSheetTabColorsBySheet.get(sheetName)
    const importedSheetPr = importedSheetPropertiesBySheet.get(sheetName)
    const importedIgnoredErrors = importedIgnoredErrorsBySheet.get(sheetName)
    const importedSparklines = importedSparklinesBySheet.get(sheetName)
    const importedStyleArtifactsForSheet = importedStyleArtifacts.sheetArtifactsByName.get(sheetName)
    const importedPivotArtifacts = importedPivots?.sheetArtifactsByName.get(sheetName)
    const importedDrawingArtifactsForSheet = importedChartDrawingArtifacts?.drawingArtifacts.sheetArtifactsByName.get(sheetName)
    const importedControlArtifactsForSheet = importedControlArtifacts?.sheetArtifactsByName.get(sheetName)
    const importedArrayFormulasForSheet = importedArrayFormulasBySheet.get(sheetName)
    const importedDataTableFormulasForSheet = importedDataTableFormulasBySheet.get(sheetName)
    const importedSheetVisibility = importedSheetVisibilitiesBySheet.get(sheetName)
    const merges = buildMergeEntries(sheetName, sheet['!merges'])
    const importedSheetProtection = importedSheetProtectionsBySheet.get(sheetName)
    const importedProtectedRanges = importedProtectedRangesBySheet.get(sheetName)
    const importedSorts = importedSortsBySheet.get(sheetName)
    const importedFilters = importedFiltersBySheet.get(sheetName)
    const importedValidations = importedValidationsBySheet.get(sheetName)
    const importedConditionalFormats = importedConditionalFormatsBySheet.get(sheetName)
    const importedConditionalFormatArtifacts = importedConditionalFormatArtifactsBySheet.get(sheetName)
    const importedPrinterSettings = importedPrinterSettingsBySheet.get(sheetName)
    const importedPrintPageSetup = importedPrintPageSetupBySheet.get(sheetName)
    const importedCellMetadataRefs = buildImportedCellMetadataReferenceSnapshots(importedCellMetadata?.refsBySheet.get(sheetName), cells)
    const importedRichTextArtifacts = importedRichTextArtifactsBySheet.get(sheetName)
    const importedThreadedCommentArtifactsForSheet = importedThreadedCommentArtifacts?.sheetArtifactsByName.get(sheetName)
    const importedViewStateForSheet = importedViewState?.sheetViewStateByName.get(sheetName)
    const metadata = buildImportedSheetMetadata({
      rows,
      columns,
      rowMetadata,
      columnMetadata,
      sheetFormatPr,
      ...(styleRanges.length > 0 ? { styleRanges } : {}),
      freezePane: importedFreezePane,
      tabColor: importedSheetTabColor,
      sheetPr: importedSheetPr,
      ignoredErrors: importedIgnoredErrors,
      sparklines: importedSparklines,
      styleArtifacts: importedStyleArtifactsForSheet,
      pivotArtifacts: importedPivotArtifacts,
      drawingArtifacts: importedDrawingArtifactsForSheet,
      controlArtifacts: importedControlArtifactsForSheet,
      arrayFormulas: importedArrayFormulasForSheet,
      dataTableFormulas: importedDataTableFormulasForSheet,
      visibility: importedSheetVisibility,
      merges,
      sheetProtection: importedSheetProtection,
      protectedRanges: importedProtectedRanges,
      sorts: importedSorts,
      filters: importedFilters,
      validations: importedValidations,
      conditionalFormats: importedConditionalFormats,
      conditionalFormatArtifacts: importedConditionalFormatArtifacts,
      commentThreads: importedComments.commentThreads,
      legacyCommentVml: importedLegacyCommentVml,
      hyperlinks: importedHyperlinks,
      printerSettings: importedPrinterSettings,
      printPageSetup: importedPrintPageSetup,
      cellMetadataRefs: importedCellMetadataRefs,
      richTextArtifacts: importedRichTextArtifacts,
      threadedCommentArtifacts: importedThreadedCommentArtifactsForSheet,
      viewState: importedViewStateForSheet,
    })

    return {
      id: order + 1,
      name: sheetName,
      order,
      ...(metadata ? { metadata } : {}),
      cells,
    }
  })

  if (workbookZip) {
    warnings.push(...readImportedWorkbookCalculationWarnings(workbookZip, { hasFormulaCells: formulaCellCount > 0 }))
  }

  const importedFormulaAudit = workbookZip
    ? readImportedWorkbookFormulaAudit({
        source: workbookZip,
        sheetNames: workbook.SheetNames,
        sheetPathsByName,
        fallbackSheetPaths,
        worksheetFormulasBySheet: importedWorksheetFormulaManifestsBySheet,
        definedNames: importedDefinedNames.definedNames ?? [],
        ...(importedCalculationSettings ? { calculationSettings: importedCalculationSettings } : {}),
      })
    : undefined

  const shouldUseCachedFormulaOpenModeForImportedWorkbook = shouldUseCachedFormulaOpenMode({
    cachedFormulaValueCount,
    formulaCellCount,
    calculationSettings: importedCalculationSettings,
    formulaAudit: importedFormulaAudit,
  })

  const workbookMetadata = buildImportedWorkbookMetadata({
    properties: importedWorkbookProperties,
    documentPropertyArtifacts: importedWorkbookDocumentProperties,
    workbookProtection: importedWorkbookProtection,
    calculationSettings: importedCalculationSettings,
    macroPayloads: importedMacroPayload ? [createPreservedVbaProjectPayload(importedMacroPayload, importedMacroCodeNames)] : undefined,
    styles: styleCatalog.size > 0 ? [...styleCatalog.values()] : undefined,
    definedNames: importedDefinedNames.definedNames,
    tables: importedTables,
    spills: importedArrayFormulaSpills.length > 0 ? importedArrayFormulaSpills : undefined,
    pivots: importedPivots?.pivots,
    externalWorkbookReferences: importedExternalWorkbookReferences.size > 0 ? [...importedExternalWorkbookReferences.values()] : undefined,
    unsupportedFormulaDependencies:
      unsupportedFormulaDependencies.length > 0
        ? unsupportedFormulaDependencies.toSorted((left, right) =>
            `${left.sheetName}:${left.address}`.localeCompare(`${right.sheetName}:${right.address}`),
          )
        : undefined,
    unsupportedPivots: importedPivots?.unsupportedPivots,
    formulaAudit: importedFormulaAudit,
    externalConnections: importedExternalConnections,
    pivotArtifacts: importedPivots?.artifacts,
    drawingArtifacts: importedChartDrawingArtifacts?.drawingArtifacts.artifacts,
    chartArtifacts: importedChartDrawingArtifacts?.chartArtifacts.artifacts,
    chartSheetArtifacts: importedChartDrawingArtifacts?.chartArtifacts.chartSheetArtifacts,
    controlArtifacts: importedControlArtifacts?.artifacts,
    dataModelArtifacts: importedDataModelArtifacts,
    externalLinkArtifacts: importedExternalLinkArtifacts,
    slicerConnectionArtifacts: importedSlicerConnectionArtifacts,
    threadedCommentArtifacts: importedThreadedCommentArtifacts?.artifacts,
    viewState: importedViewState?.workbookViewState,
    charts: importedCharts,
    styleArtifacts: importedStyleArtifacts.workbookArtifacts,
    cellMetadata: importedCellMetadata?.workbookMetadata,
  })

  const snapshot: WorkbookSnapshot = {
    version: 1,
    workbook: {
      name: workbookName,
      ...(workbookMetadata ? { metadata: workbookMetadata } : {}),
    },
    sheets,
  }

  const restoredSnapshot = attachImportedRuntimeCoordinates(snapshot, runtimeSheetCells)
  if (
    shouldUseCachedFormulaOpenModeForImportedWorkbook &&
    sourceBytesForUntouchedExport !== undefined &&
    (contentType === XLSX_CONTENT_TYPE || contentType === XLSM_CONTENT_TYPE)
  ) {
    attachImportedXlsxSourceBytes(restoredSnapshot, sourceBytesForUntouchedExport)
  }

  return {
    snapshot: restoredSnapshot,
    workbookName,
    sheetNames: workbook.SheetNames,
    warnings,
    preview: createWorkbookPreview({
      contentType,
      fileName,
      fileSizeBytes: data.byteLength,
      workbookName,
      sheets: previewSheets,
      warnings,
    }),
  }
}

export function inspectXlsx(bytes: Uint8Array | ArrayBuffer, fileName: string): XlsxHeadlessInspectResult | null {
  const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)
  return inspectLargeSimpleXlsxSource(data, fileName, { minByteLength: 0 })
}

export function importXlsx(bytes: Uint8Array | ArrayBuffer, fileName: string, options: XlsxImportOptions = {}): ImportedWorkbook {
  const ownedSource: OwnedXlsxSourceBytes = { bytes: bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes) }
  const sourceByteLength = ownedSource.bytes.byteLength
  const limits = resolveXlsxImportLimits(options)
  const inspectionOptions = options.limits ? { minByteLength: 0 } : undefined
  const workbookZip = readValidXlsxZipContainer(ownedSource.bytes, 'lazy')
  const hasCalcChain = Object.hasOwn(workbookZip, 'xl/calcChain.xml')
  const bypassLargeSimpleByteThreshold =
    shouldBypassLargeSimpleByteThresholdForPackageArtifacts(workbookZip) && !hasFullImporterOnlyPackageMetadata(workbookZip)
  const inspection =
    limits || (hasCalcChain && sourceByteLength >= denseSheetJsByteThreshold)
      ? inspectLargeSimpleXlsxSource(ownedSource.bytes, fileName, inspectionOptions)
      : null
  assertXlsxInspectionWithinMaterializationLimits(inspection, limits)
  const hasLargeCalcChainFormulaSet = hasCalcChain && (inspection?.stats.formulaCellCount ?? 0) >= largeCalcChainStreamingFormulaThreshold
  const allowCachedUnsupportedFormulaText =
    hasCalcChain && (sourceByteLength >= largeCalcChainStreamingByteThreshold || hasLargeCalcChainFormulaSet)
  const shouldTryLargeSimpleImport =
    !hasCalcChain ||
    sourceByteLength >= largeCalcChainStreamingByteThreshold ||
    hasLargeCalcChainFormulaSet ||
    bypassLargeSimpleByteThreshold
  const releaseOwnedSourceBytesForMaterializedPackageArtifacts =
    bypassLargeSimpleByteThreshold && sourceByteLength < denseSheetJsByteThreshold
      ? () => releaseOwnedXlsxSourceBytes(ownedSource, (releasedBytes) => (bytes = releasedBytes))
      : undefined
  const largeSimpleImportOptions = {
    ...(options.limits || bypassLargeSimpleByteThreshold ? { minByteLength: 0 } : {}),
    allowUnsupportedFormulaText: allowCachedUnsupportedFormulaText,
    allowUnsupportedCellMetadata: allowCachedUnsupportedFormulaText,
    releaseArenaAfterMaterialization: true,
    releaseZipSource: true,
    maxMaterializedLazyPackageArtifactBytes: 8 * 1024 * 1024,
    ...(releaseOwnedSourceBytesForMaterializedPackageArtifacts
      ? { releaseOwnedSourceBytes: releaseOwnedSourceBytesForMaterializedPackageArtifacts }
      : {}),
  }
  let largeSimpleImport = shouldTryLargeSimpleImport
    ? tryImportLargeSimpleXlsx({ byteLength: sourceByteLength }, fileName, workbookZip, largeSimpleImportOptions)
    : null
  if (!largeSimpleImport && shouldRetryDataOnlyLargeSimpleImport(inspection, sourceByteLength, allowCachedUnsupportedFormulaText)) {
    largeSimpleImport = tryImportLargeSimpleXlsx(
      { byteLength: sourceByteLength },
      fileName,
      readValidXlsxZipContainer(ownedSource.bytes, 'lazy'),
      {
        ...largeSimpleImportOptions,
        materializeMetadata: false,
      },
    )
  }
  if (largeSimpleImport) {
    if (ownedSource.bytes.byteLength > 0) attachImportedXlsxSourceBytes(largeSimpleImport.snapshot, ownedSource.bytes)
    return largeSimpleImport
  }
  const fallbackData = ownedSource.bytes.byteLength > 0 ? ownedSource.bytes : readLazyXlsxZipSource(workbookZip)
  if (!fallbackData) {
    throw new InvalidXlsxZipContainerError()
  }
  const imported = importSheetJsWorkbook(
    fallbackData,
    fileName,
    XLSX_CONTENT_TYPE,
    readValidXlsxZipContainer(fallbackData, 'lazy'),
    ownedSource.bytes,
  )
  return imported
}

export function importXlsm(bytes: Uint8Array | ArrayBuffer, fileName: string): ImportedWorkbook {
  const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)
  const workbookZip = readValidXlsxZipContainer(data, 'lazy')
  return importSheetJsWorkbook(data, fileName, XLSM_CONTENT_TYPE, workbookZip, data)
}

export function importXlsb(bytes: Uint8Array | ArrayBuffer, fileName: string): ImportedWorkbook {
  const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)
  return importSheetJsWorkbook(data, fileName, XLSB_CONTENT_TYPE, null)
}

export function importXls(bytes: Uint8Array | ArrayBuffer, fileName: string): ImportedWorkbook {
  const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)
  return importSheetJsWorkbook(data, fileName, LEGACY_XLS_CONTENT_TYPE, null)
}

export function importWorkbookFile(
  bytes: Uint8Array | ArrayBuffer,
  fileName: string,
  contentType: string,
  options: WorkbookImportFileOptions = {},
): ImportedWorkbook {
  const normalizedContentType = normalizeWorkbookImportContentType(contentType)
  if (normalizedContentType === XLSX_CONTENT_TYPE) {
    return importXlsx(bytes, fileName, options.xlsx)
  }
  if (normalizedContentType === XLSM_CONTENT_TYPE) {
    return importXlsm(bytes, fileName)
  }
  if (normalizedContentType === XLSB_CONTENT_TYPE) {
    return importXlsb(bytes, fileName)
  }
  if (normalizedContentType === LEGACY_XLS_CONTENT_TYPE) {
    return importXls(bytes, fileName)
  }
  if (normalizedContentType === CSV_CONTENT_TYPE) {
    const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)
    return importCsv(new TextDecoder().decode(data), fileName, options.csv)
  }
  throw new Error('Unsupported workbook import content type')
}
