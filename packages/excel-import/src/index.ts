import { importBiligBackup } from './bilig-backup.js'
import { createRequire } from 'node:module'
import type { Unzipped } from 'fflate'
import type { CsvParseOptions } from '@bilig/core'
import type { WorkbookSnapshot } from '@bilig/protocol'
import {
  BILIG_CONTENT_TYPE,
  CSV_CONTENT_TYPE,
  LEGACY_XLS_CONTENT_TYPE,
  XLSB_CONTENT_TYPE,
  XLSM_CONTENT_TYPE,
  XLSX_CONTENT_TYPE,
  normalizeWorkbookImportContentType,
  type ExcelWorkbookImportContentType,
} from './workbook-import-content-types.js'
import type { ImportedWorkbook } from './workbook-import-result.js'
import {
  assertXlsxInspectionWithinMaterializationLimits,
  assertXlsxByteInputApiWithinLimit,
  assertXlsxSheetJsFallbackWithinMaterializationLimits,
  assertXlsxSourceWithinMaterializationLimits,
  assertXlsxZipWithinMaterializationLimits,
  denseSheetJsByteThreshold,
  largeSimpleInMemoryUntouchedExportSourceLimit,
  planXlsxImportRoute,
  resolveXlsxImportLimits,
  shouldAllowLegacyLargeSheetJsFallback,
  xlsxByteInputApiLimit,
  type XlsxImportOptions,
} from './xlsx-import-limits.js'
import type { LargeSimpleXlsxHeadlessInspectResult, tryInspectLargeSimpleXlsxHeadless } from './xlsx-large-simple-headless-inspect.js'
import type { tryImportLargeSimpleXlsx } from './xlsx-large-simple-import.js'
import { releaseOwnedXlsxSourceBytes, type OwnedXlsxSourceBytes } from './xlsx-owned-source-release.js'
import { importXlsxFromZipByteSource } from './xlsx-byte-source-import.js'
import {
  attachImportedXlsxSourceBytes,
  attachImportedXlsxSourceReader,
  createTempFileImportedXlsxSourceReader,
  type ImportedXlsxSourceReference,
  type ImportedXlsxSourceReader,
} from './xlsx-source-bytes.js'
import {
  readLazyXlsxZipSource,
  readXlsxZipEntries,
  readXlsxZipEntriesLazy,
  readXlsxZipEntriesLazyFromByteSource,
  type XlsxZipByteSource,
} from './xlsx-zip.js'
import type {
  XlsxSourceLiteralPatchExportInput,
  XlsxSourceLiteralPatchFileExportInput,
  XlsxSourceLiteralPatchFileExportResult,
} from './xlsx-source-preserving-export.js'

export { manualCalculationModeWarning, precisionAsDisplayedCalculationWarning } from './xlsx-calculation-settings.js'
export {
  dataTableFormulasWarning,
  definedNameFormulaCachesWarning,
  externalWorkbookCompanionAmbiguousMatchWarning,
  externalWorkbookCompanionNoMatchWarning,
  externalPivotCachesWarning,
  externalWorkbookReferencesWarning,
  macroExecutionDeclinedWarning,
  unsupportedCellStylesWarning,
  unsupportedFormulaCachesWarning,
  volatileFormulasWarning,
} from './xlsx-import-warnings.js'
export { readImportedXlsxCellStyle } from './xlsx-import-cell-styles.js'
export { XlsxImportSizeLimitExceededError } from './xlsx-import-limits.js'
export type { ImportedWorkbookSheetPreview } from './workbook-import-helpers.js'
export type { ImportedWorkbookPreview } from './workbook-import-preview.js'
export type { ImportedWorkbook } from './workbook-import-result.js'
export type {
  ImportedWorkbookDiagnostics,
  XlsxExternalWorkbookHydrationDiagnostics,
  XlsxExternalWorkbookHydrationMatchKind,
  XlsxExternalWorkbookHydrationReferenceDiagnostic,
  XlsxExternalWorkbookHydrationStatus,
  XlsxExternalWorkbookInput,
  XlsxImportLimits,
  XlsxImportOptions,
} from './xlsx-import-limits.js'
export {
  BILIG_CONTENT_TYPE,
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
export {
  createFileImportedXlsxSourceReader,
  createTempFileImportedXlsxSourceReader,
  defaultImportedXlsxSourceReadBytesLimit,
} from './xlsx-source-bytes.js'
export type { ImportedXlsxSourceReader, ImportedXlsxSourceReaderOptions } from './xlsx-source-bytes.js'
export { importXlsxFromZipByteSource } from './xlsx-byte-source-import.js'
export type { XlsxByteSourceImportOptions } from './xlsx-byte-source-import.js'

const requireModule = createRequire(import.meta.url)
const bundledLocalModules = readBundledLocalModules()

type ImportMetaGlob = (patterns: readonly string[], options: { readonly eager: true }) => Readonly<Record<string, unknown>>

declare global {
  interface ImportMeta {
    readonly glob: ImportMetaGlob
  }
}

interface XlsxExportModule {
  readonly exportXlsx: (snapshot: WorkbookSnapshot) => Uint8Array
  readonly exportXlsxToFile: (snapshot: WorkbookSnapshot, outputPath: string) => XlsxSourceLiteralPatchFileExportResult
}

interface XlsxSourcePreservingExportModule {
  readonly exportXlsxSourceLiteralPatches: (input: XlsxSourceLiteralPatchExportInput) => Uint8Array
  readonly exportXlsxSourceLiteralPatchesToFile: (input: XlsxSourceLiteralPatchFileExportInput) => XlsxSourceLiteralPatchFileExportResult
  readonly exportXlsxSourceLiteralPatchesToFileAsync: (
    input: XlsxSourceLiteralPatchFileExportInput,
  ) => Promise<XlsxSourceLiteralPatchFileExportResult>
}

interface CsvImportModule {
  readonly importCsv: (csv: string, fileName: string, options?: CsvParseOptions) => ImportedWorkbook
}

interface SheetJsImportModule {
  readonly importSheetJsWorkbook: (
    data: Uint8Array,
    fileName: string,
    contentType: ExcelWorkbookImportContentType,
    workbookZip: Unzipped | null,
    sourceForUntouchedExport?: ImportedXlsxSourceReference,
    options?: XlsxImportOptions,
  ) => ImportedWorkbook
  readonly importXlsxFromPreparedSheetJsParserData: (
    parserData: Uint8Array,
    fileName: string,
    contentType: ExcelWorkbookImportContentType,
    workbookZip: Unzipped | null,
    sourceFileSizeBytes: number,
    options?: XlsxImportOptions,
  ) => ImportedWorkbook
}

type TryImportLargeSimpleXlsx = typeof tryImportLargeSimpleXlsx
type TryInspectLargeSimpleXlsxHeadless = typeof tryInspectLargeSimpleXlsxHeadless

interface LargeSimpleImportModule {
  readonly tryImportLargeSimpleXlsx: TryImportLargeSimpleXlsx
}

interface LargeSimpleInspectModule {
  readonly tryInspectLargeSimpleXlsxHeadless: TryInspectLargeSimpleXlsxHeadless
}

let xlsxExportModule: XlsxExportModule | undefined
let xlsxSourcePreservingExportModule: XlsxSourcePreservingExportModule | undefined
let csvImportModule: CsvImportModule | undefined
let sheetJsImportModule: SheetJsImportModule | undefined
let largeSimpleImportModule: LargeSimpleImportModule | undefined
let largeSimpleInspectModule: LargeSimpleInspectModule | undefined

function loadXlsxExportModule(): XlsxExportModule {
  xlsxExportModule ??= readXlsxExportModule(readBundledLocalModule('./xlsx-export.js') ?? requireLocalModule('./xlsx-export.js'))
  return xlsxExportModule
}

function loadXlsxSourcePreservingExportModule(): XlsxSourcePreservingExportModule {
  xlsxSourcePreservingExportModule ??= readXlsxSourcePreservingExportModule(
    readBundledLocalModule('./xlsx-source-preserving-export.js') ?? requireLocalModule('./xlsx-source-preserving-export.js'),
  )
  return xlsxSourcePreservingExportModule
}

function loadCsvImportModule(): CsvImportModule {
  csvImportModule ??= readCsvImportModule(readBundledLocalModule('./csv-import.js') ?? requireLocalModule('./csv-import.js'))
  return csvImportModule
}

function loadSheetJsImportModule(): SheetJsImportModule {
  sheetJsImportModule ??= readSheetJsImportModule(
    readBundledLocalModule('./xlsx-sheetjs-import.js') ?? requireLocalModule('./xlsx-sheetjs-import.js'),
  )
  return sheetJsImportModule
}

function loadLargeSimpleImportModule(): LargeSimpleImportModule {
  largeSimpleImportModule ??= readLargeSimpleImportModule(
    readBundledLocalModule('./xlsx-large-simple-import.js') ?? requireLocalModule('./xlsx-large-simple-import.js'),
  )
  return largeSimpleImportModule
}

function loadLargeSimpleInspectModule(): LargeSimpleInspectModule {
  largeSimpleInspectModule ??= readLargeSimpleInspectModule(
    readBundledLocalModule('./xlsx-large-simple-headless-inspect.js') ?? requireLocalModule('./xlsx-large-simple-headless-inspect.js'),
  )
  return largeSimpleInspectModule
}

function readBundledLocalModules(): Readonly<Record<string, unknown>> {
  try {
    return import.meta.glob(
      [
        './xlsx-export.ts',
        './xlsx-source-preserving-export.ts',
        './csv-import.ts',
        './xlsx-sheetjs-import.ts',
        './xlsx-large-simple-import.ts',
        './xlsx-large-simple-headless-inspect.ts',
      ],
      { eager: true },
    )
  } catch {
    return {}
  }
}

function readBundledLocalModule(path: string): unknown {
  return bundledLocalModules[path] ?? bundledLocalModules[path.replace(/\.js$/u, '.ts')]
}

function requireLocalModule(jsPath: string): unknown {
  try {
    return requireModule(jsPath)
  } catch (error) {
    if (!isModuleNotFoundError(error)) {
      throw error
    }
    return requireModule(`../dist/${jsPath.slice(2)}`)
  }
}

function isModuleNotFoundError(error: unknown): boolean {
  return isRecord(error) && error['code'] === 'MODULE_NOT_FOUND'
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function isXlsxExportFunction(value: unknown): value is XlsxExportModule['exportXlsx'] {
  return typeof value === 'function'
}

function isXlsxExportToFileFunction(value: unknown): value is XlsxExportModule['exportXlsxToFile'] {
  return typeof value === 'function'
}

function isXlsxSourceLiteralPatchExportFunction(
  value: unknown,
): value is XlsxSourcePreservingExportModule['exportXlsxSourceLiteralPatches'] {
  return typeof value === 'function'
}

function isXlsxSourceLiteralPatchFileExportFunction(
  value: unknown,
): value is XlsxSourcePreservingExportModule['exportXlsxSourceLiteralPatchesToFile'] {
  return typeof value === 'function'
}

function isXlsxSourceLiteralPatchAsyncFileExportFunction(
  value: unknown,
): value is XlsxSourcePreservingExportModule['exportXlsxSourceLiteralPatchesToFileAsync'] {
  return typeof value === 'function'
}

function isCsvImportFunction(value: unknown): value is CsvImportModule['importCsv'] {
  return typeof value === 'function'
}

function isSheetJsImportFunction(value: unknown): value is SheetJsImportModule['importSheetJsWorkbook'] {
  return typeof value === 'function'
}

function isPreparedSheetJsImportFunction(value: unknown): value is SheetJsImportModule['importXlsxFromPreparedSheetJsParserData'] {
  return typeof value === 'function'
}

function isLargeSimpleImportFunction(value: unknown): value is TryImportLargeSimpleXlsx {
  return typeof value === 'function'
}

function isLargeSimpleInspectFunction(value: unknown): value is TryInspectLargeSimpleXlsxHeadless {
  return typeof value === 'function'
}

function readXlsxExportModule(value: unknown): XlsxExportModule {
  const loadedExportXlsx = isRecord(value) ? value['exportXlsx'] : undefined
  const loadedExportXlsxToFile = isRecord(value) ? value['exportXlsxToFile'] : undefined
  if (isXlsxExportFunction(loadedExportXlsx) && isXlsxExportToFileFunction(loadedExportXlsxToFile)) {
    return {
      exportXlsx: loadedExportXlsx,
      exportXlsxToFile: loadedExportXlsxToFile,
    }
  }
  throw new Error('XLSX export module is missing required exports')
}

function readXlsxSourcePreservingExportModule(value: unknown): XlsxSourcePreservingExportModule {
  const loadedExportXlsxSourceLiteralPatches = isRecord(value) ? value['exportXlsxSourceLiteralPatches'] : undefined
  const loadedExportXlsxSourceLiteralPatchesToFile = isRecord(value) ? value['exportXlsxSourceLiteralPatchesToFile'] : undefined
  const loadedExportXlsxSourceLiteralPatchesToFileAsync = isRecord(value) ? value['exportXlsxSourceLiteralPatchesToFileAsync'] : undefined
  if (
    isXlsxSourceLiteralPatchExportFunction(loadedExportXlsxSourceLiteralPatches) &&
    isXlsxSourceLiteralPatchFileExportFunction(loadedExportXlsxSourceLiteralPatchesToFile) &&
    isXlsxSourceLiteralPatchAsyncFileExportFunction(loadedExportXlsxSourceLiteralPatchesToFileAsync)
  ) {
    return {
      exportXlsxSourceLiteralPatches: loadedExportXlsxSourceLiteralPatches,
      exportXlsxSourceLiteralPatchesToFile: loadedExportXlsxSourceLiteralPatchesToFile,
      exportXlsxSourceLiteralPatchesToFileAsync: loadedExportXlsxSourceLiteralPatchesToFileAsync,
    }
  }
  throw new Error('XLSX export module is missing required exports')
}

function readCsvImportModule(value: unknown): CsvImportModule {
  const loadedImportCsv = isRecord(value) ? value['importCsv'] : undefined
  if (isCsvImportFunction(loadedImportCsv)) {
    return { importCsv: loadedImportCsv }
  }
  throw new Error('CSV import module is missing required exports')
}

function readSheetJsImportModule(value: unknown): SheetJsImportModule {
  const loadedImportSheetJsWorkbook = isRecord(value) ? value['importSheetJsWorkbook'] : undefined
  const loadedPreparedImport = isRecord(value) ? value['importXlsxFromPreparedSheetJsParserData'] : undefined
  if (isSheetJsImportFunction(loadedImportSheetJsWorkbook) && isPreparedSheetJsImportFunction(loadedPreparedImport)) {
    return {
      importSheetJsWorkbook: loadedImportSheetJsWorkbook,
      importXlsxFromPreparedSheetJsParserData: loadedPreparedImport,
    }
  }
  throw new Error('SheetJS XLSX import module is missing required exports')
}

function readLargeSimpleImportModule(value: unknown): LargeSimpleImportModule {
  const loadedTryImportLargeSimpleXlsx = isRecord(value) ? value['tryImportLargeSimpleXlsx'] : undefined
  if (isLargeSimpleImportFunction(loadedTryImportLargeSimpleXlsx)) {
    return { tryImportLargeSimpleXlsx: loadedTryImportLargeSimpleXlsx }
  }
  throw new Error('Large-simple XLSX import module is missing required exports')
}

function readLargeSimpleInspectModule(value: unknown): LargeSimpleInspectModule {
  const loadedTryInspectLargeSimpleXlsxHeadless = isRecord(value) ? value['tryInspectLargeSimpleXlsxHeadless'] : undefined
  if (isLargeSimpleInspectFunction(loadedTryInspectLargeSimpleXlsxHeadless)) {
    return { tryInspectLargeSimpleXlsxHeadless: loadedTryInspectLargeSimpleXlsxHeadless }
  }
  throw new Error('Large-simple XLSX inspect module is missing required exports')
}

export function exportXlsx(snapshot: WorkbookSnapshot): Uint8Array {
  return loadXlsxExportModule().exportXlsx(snapshot)
}

export function exportXlsxToFile(snapshot: WorkbookSnapshot, outputPath: string): XlsxSourceLiteralPatchFileExportResult {
  return loadXlsxExportModule().exportXlsxToFile(snapshot, outputPath)
}

export function exportXlsxSourceLiteralPatches(input: XlsxSourceLiteralPatchExportInput): Uint8Array {
  return loadXlsxSourcePreservingExportModule().exportXlsxSourceLiteralPatches(input)
}

export function exportXlsxSourceLiteralPatchesToFile(input: XlsxSourceLiteralPatchFileExportInput): XlsxSourceLiteralPatchFileExportResult {
  return loadXlsxSourcePreservingExportModule().exportXlsxSourceLiteralPatchesToFile(input)
}

export function exportXlsxSourceLiteralPatchesToFileAsync(
  input: XlsxSourceLiteralPatchFileExportInput,
): Promise<XlsxSourceLiteralPatchFileExportResult> {
  return loadXlsxSourcePreservingExportModule().exportXlsxSourceLiteralPatchesToFileAsync(input)
}

export function importCsv(csv: string, fileName: string, options?: CsvParseOptions): ImportedWorkbook {
  return loadCsvImportModule().importCsv(csv, fileName, options)
}

function importSheetJsWorkbook(
  data: Uint8Array,
  fileName: string,
  contentType: ExcelWorkbookImportContentType,
  workbookZip: Unzipped | null,
  sourceForUntouchedExport?: ImportedXlsxSourceReference,
  options?: XlsxImportOptions,
): ImportedWorkbook {
  return loadSheetJsImportModule().importSheetJsWorkbook(data, fileName, contentType, workbookZip, sourceForUntouchedExport, options)
}

export function importXlsxFromPreparedSheetJsParserData(
  parserData: Uint8Array,
  fileName: string,
  contentType: ExcelWorkbookImportContentType,
  workbookZip: Unzipped | null,
  sourceFileSizeBytes: number,
  options: XlsxImportOptions = {},
): ImportedWorkbook {
  return loadSheetJsImportModule().importXlsxFromPreparedSheetJsParserData(
    parserData,
    fileName,
    contentType,
    workbookZip,
    sourceFileSizeBytes,
    options,
  )
}

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

function readValidXlsxZipContainer(bytes: Uint8Array, mode: 'eager' | 'lazy' = 'eager'): Unzipped {
  try {
    const zip = mode === 'lazy' ? readXlsxZipEntriesLazy(bytes) : readXlsxZipEntries(bytes)
    void zip['xl/workbook.xml']
    return zip
  } catch {
    throw new InvalidXlsxZipContainerError()
  }
}

function readValidXlsxZipContainerFromByteSource(source: XlsxZipByteSource): Unzipped {
  try {
    const zip = readXlsxZipEntriesLazyFromByteSource(borrowXlsxZipByteSource(source))
    if (!zip) {
      throw new Error('Unsupported XLSX zip byte source')
    }
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
  return loadLargeSimpleInspectModule().tryInspectLargeSimpleXlsxHeadless({ byteLength: data.byteLength }, fileName, inspectionZip, {
    allowUnsupportedWorksheetFeaturesForMetrics: true,
    ...(options.minByteLength !== undefined ? { minByteLength: options.minByteLength } : {}),
    releaseZipSource: true,
  })
}

function inspectLargeSimpleXlsxByteSource(
  source: XlsxZipByteSource,
  fileName: string,
  options: { readonly minByteLength?: number } = {},
): LargeSimpleXlsxHeadlessInspectResult | null {
  const inspectionZip = readValidXlsxZipContainerFromByteSource(source)
  return loadLargeSimpleInspectModule().tryInspectLargeSimpleXlsxHeadless({ byteLength: source.byteLength }, fileName, inspectionZip, {
    allowUnsupportedWorksheetFeaturesForMetrics: true,
    ...(options.minByteLength !== undefined ? { minByteLength: options.minByteLength } : {}),
    releaseZipSource: true,
  })
}

export function inspectXlsx(bytes: Uint8Array | ArrayBuffer, fileName: string): XlsxHeadlessInspectResult | null {
  const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)
  assertXlsxByteInputApiWithinLimit(data.byteLength, 'inspectXlsx')
  return inspectLargeSimpleXlsxSource(data, fileName, { minByteLength: 0 })
}

function borrowXlsxZipByteSource(source: XlsxZipByteSource): XlsxZipByteSource {
  return {
    byteLength: source.byteLength,
    readRange: (start, end) => source.readRange(start, end),
    ...(source.readRangeInto
      ? {
          readRangeInto: (start: number, end: number, target: Uint8Array) => source.readRangeInto!(start, end, target),
        }
      : {}),
    ...(source.inflateRawRange ? { inflateRawRange: (start: number, end: number) => source.inflateRawRange!(start, end) } : {}),
  }
}

function xlsxZipByteSourceFromBytes(bytes: Uint8Array): XlsxZipByteSource {
  return {
    byteLength: bytes.byteLength,
    readRange(start, end) {
      return bytes.subarray(start, end)
    },
    readRangeInto(start, end, target) {
      const length = end - start
      if (target.byteLength < length) {
        throw new Error('XLSX byte-source import target is too small')
      }
      target.set(bytes.subarray(start, end), 0)
      return target.subarray(0, length)
    },
  }
}

export function importXlsx(bytes: Uint8Array | ArrayBuffer, fileName: string, options: XlsxImportOptions = {}): ImportedWorkbook {
  const ownedSource: OwnedXlsxSourceBytes = { bytes: bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes) }
  const sourceByteLength = ownedSource.bytes.byteLength
  assertXlsxByteInputApiWithinLimit(sourceByteLength, 'importXlsx')
  let spooledUntouchedExportSource: (ImportedXlsxSourceReader & XlsxZipByteSource) | undefined =
    sourceByteLength > largeSimpleInMemoryUntouchedExportSourceLimit
      ? createTempFileImportedXlsxSourceReader(ownedSource.bytes, { maxReadBytes: spooledSourceReadBytesLimitFor(options) })
      : undefined
  let preReleasedOwnedSourceEvidence: ReturnType<typeof releaseOwnedXlsxSourceBytes> | undefined
  const workbookZip = spooledUntouchedExportSource
    ? readValidXlsxZipContainerFromByteSource(spooledUntouchedExportSource)
    : readValidXlsxZipContainer(ownedSource.bytes, 'lazy')
  if (spooledUntouchedExportSource) {
    preReleasedOwnedSourceEvidence = releaseOwnedXlsxSourceBytes(ownedSource, (releasedBytes) => (bytes = releasedBytes))
  }
  const limits = resolveXlsxImportLimits(options)
  assertXlsxZipWithinMaterializationLimits(workbookZip, limits)
  let route = planXlsxImportRoute({ workbookZip, sourceByteLength, options, inspection: null })
  let inspection = route.shouldInspectBeforeLargeSimpleRouting
    ? spooledUntouchedExportSource
      ? inspectLargeSimpleXlsxByteSource(spooledUntouchedExportSource, fileName, route.inspectionOptions)
      : inspectLargeSimpleXlsxSource(ownedSource.bytes, fileName, route.inspectionOptions)
    : null
  route = planXlsxImportRoute({ workbookZip, sourceByteLength, options, inspection })
  assertXlsxInspectionWithinMaterializationLimits(inspection, limits)
  const releaseOwnedSourceBytesForLargeSimpleImport =
    spooledUntouchedExportSource || (route.bypassLargeSimpleByteThreshold && sourceByteLength < denseSheetJsByteThreshold)
      ? () => {
          if (preReleasedOwnedSourceEvidence) {
            const evidence = preReleasedOwnedSourceEvidence
            preReleasedOwnedSourceEvidence = undefined
            return evidence
          }
          return releaseOwnedXlsxSourceBytes(ownedSource, (releasedBytes) => (bytes = releasedBytes))
        }
      : undefined
  const largeSimpleImportOptions = route.createLargeSimpleImportOptions({
    allowPreReleaseSheetFinalization:
      releaseOwnedSourceBytesForLargeSimpleImport === undefined || spooledUntouchedExportSource !== undefined,
    ...(spooledUntouchedExportSource ? { allowPreReleaseSheetFinalizationWithOwnedSourceRelease: true } : {}),
    ...(spooledUntouchedExportSource ? { replacementZipSource: spooledUntouchedExportSource } : {}),
    ...(releaseOwnedSourceBytesForLargeSimpleImport ? { releaseOwnedSourceBytes: releaseOwnedSourceBytesForLargeSimpleImport } : {}),
  })
  let largeSimpleImport = route.shouldTryLargeSimpleImport
    ? loadLargeSimpleImportModule().tryImportLargeSimpleXlsx(
        { byteLength: sourceByteLength },
        fileName,
        workbookZip,
        largeSimpleImportOptions,
      )
    : null
  if (!largeSimpleImport && route.shouldRetryDataOnlyLargeSimpleImport) {
    const retryZip = spooledUntouchedExportSource
      ? readValidXlsxZipContainerFromByteSource(spooledUntouchedExportSource)
      : readValidXlsxZipContainer(ownedSource.bytes, 'lazy')
    largeSimpleImport = loadLargeSimpleImportModule().tryImportLargeSimpleXlsx({ byteLength: sourceByteLength }, fileName, retryZip, {
      ...largeSimpleImportOptions,
      materializeMetadata: false,
    })
  }
  if (largeSimpleImport) {
    if (ownedSource.bytes.byteLength > 0) {
      spooledUntouchedExportSource?.release?.()
      spooledUntouchedExportSource = undefined
      attachImportedXlsxSourceBytes(largeSimpleImport.snapshot, ownedSource.bytes)
    } else if (spooledUntouchedExportSource) {
      attachImportedXlsxSourceReader(largeSimpleImport.snapshot, spooledUntouchedExportSource)
      spooledUntouchedExportSource = undefined
    }
    return largeSimpleImport
  }
  if (options.nativeOnly === true) {
    spooledUntouchedExportSource?.release?.()
    spooledUntouchedExportSource = undefined
    throw new Error('Native XLSX import could not materialize this workbook without SheetJS fallback.')
  }
  if (!inspection && route.shouldInspectBeforeSheetJsFallback) {
    inspection = spooledUntouchedExportSource
      ? inspectLargeSimpleXlsxByteSource(spooledUntouchedExportSource, fileName, route.inspectionOptions)
      : inspectLargeSimpleXlsxSource(ownedSource.bytes, fileName, route.inspectionOptions)
  }
  assertXlsxSheetJsFallbackWithinMaterializationLimits(inspection, options, sourceByteLength)
  const fallbackData =
    ownedSource.bytes.byteLength > 0
      ? ownedSource.bytes
      : spooledUntouchedExportSource
        ? spooledUntouchedExportSource.readBytes()
        : readLazyXlsxZipSource(workbookZip)
  if (!fallbackData || fallbackData.byteLength === 0) {
    spooledUntouchedExportSource?.release?.()
    spooledUntouchedExportSource = undefined
    throw new InvalidXlsxZipContainerError()
  }
  const shouldRetainSourceForUntouchedExport = (options.externalWorkbooks?.length ?? 0) === 0
  const sourceForUntouchedExport: ImportedXlsxSourceReference | undefined = shouldRetainSourceForUntouchedExport
    ? ownedSource.bytes.byteLength > 0
      ? ownedSource.bytes
      : (spooledUntouchedExportSource ?? fallbackData)
    : undefined
  const spooledSourceRetainedForSnapshot =
    spooledUntouchedExportSource !== undefined && sourceForUntouchedExport === spooledUntouchedExportSource
  try {
    const imported = importSheetJsWorkbook(
      fallbackData,
      fileName,
      XLSX_CONTENT_TYPE,
      readValidXlsxZipContainer(fallbackData, 'lazy'),
      sourceForUntouchedExport,
      options,
    )
    if (!spooledSourceRetainedForSnapshot) {
      spooledUntouchedExportSource?.release?.()
    }
    spooledUntouchedExportSource = undefined
    return imported
  } catch (error) {
    spooledUntouchedExportSource?.release?.()
    spooledUntouchedExportSource = undefined
    throw error
  }
}

function spooledSourceReadBytesLimitFor(options: XlsxImportOptions): number | false {
  if (shouldAllowLegacyLargeSheetJsFallback(options)) {
    return false
  }
  if (options.limits === false || options.limits === undefined) {
    return denseSheetJsByteThreshold
  }
  const maxSourceBytes = options.limits.maxMaterializedSourceBytes
  return maxSourceBytes === undefined || !Number.isFinite(maxSourceBytes) ? false : maxSourceBytes
}

export function importXlsm(bytes: Uint8Array | ArrayBuffer, fileName: string, options: XlsxImportOptions = {}): ImportedWorkbook {
  const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)
  const limits = resolveXlsxImportLimits(options)
  if (options.limits === undefined) {
    assertXlsxByteInputApiWithinLimit(data.byteLength, 'importXlsm')
  } else {
    assertXlsxSourceWithinMaterializationLimits(data.byteLength, limits)
  }
  const workbookZip = readValidXlsxZipContainer(data, 'lazy')
  assertXlsxZipWithinMaterializationLimits(workbookZip, limits)
  return importSheetJsWorkbook(data, fileName, XLSM_CONTENT_TYPE, workbookZip, data, options)
}

export function importXlsb(bytes: Uint8Array | ArrayBuffer, fileName: string, options: XlsxImportOptions = {}): ImportedWorkbook {
  const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)
  assertXlsxSourceWithinMaterializationLimits(data.byteLength, resolveXlsxImportLimits(options))
  return importSheetJsWorkbook(data, fileName, XLSB_CONTENT_TYPE, null, undefined, options)
}

export function importXls(bytes: Uint8Array | ArrayBuffer, fileName: string, options: XlsxImportOptions = {}): ImportedWorkbook {
  const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)
  assertXlsxSourceWithinMaterializationLimits(data.byteLength, resolveXlsxImportLimits(options))
  return importSheetJsWorkbook(data, fileName, LEGACY_XLS_CONTENT_TYPE, null, undefined, options)
}

export function importWorkbookFile(
  bytes: Uint8Array | ArrayBuffer,
  fileName: string,
  contentType: string,
  options: WorkbookImportFileOptions = {},
): ImportedWorkbook {
  const normalizedContentType = normalizeWorkbookImportContentType(contentType)
  if (normalizedContentType === XLSX_CONTENT_TYPE) {
    const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)
    return data.byteLength > xlsxByteInputApiLimit
      ? importXlsxFromZipByteSource(xlsxZipByteSourceFromBytes(data), fileName, options.xlsx)
      : importXlsx(data, fileName, options.xlsx)
  }
  if (normalizedContentType === XLSM_CONTENT_TYPE) {
    return importXlsm(bytes, fileName, options.xlsx)
  }
  if (normalizedContentType === XLSB_CONTENT_TYPE) {
    return importXlsb(bytes, fileName, options.xlsx)
  }
  if (normalizedContentType === LEGACY_XLS_CONTENT_TYPE) {
    return importXls(bytes, fileName, options.xlsx)
  }
  if (normalizedContentType === BILIG_CONTENT_TYPE) {
    return importBiligBackup(bytes, fileName)
  }
  if (normalizedContentType === CSV_CONTENT_TYPE) {
    const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)
    return importCsv(new TextDecoder().decode(data), fileName, options.csv)
  }
  throw new Error('Unsupported workbook import content type')
}
