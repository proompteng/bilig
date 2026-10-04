import { importBiligBackup } from './bilig-backup.js'
import type { CsvParseOptions } from '@bilig/core'
import type { WorkbookSnapshot } from '@bilig/protocol'
import { importCsv } from './csv-import.js'
import {
  BILIG_CONTENT_TYPE,
  CSV_CONTENT_TYPE,
  LEGACY_XLS_CONTENT_TYPE,
  XLSB_CONTENT_TYPE,
  XLSM_CONTENT_TYPE,
  XLSX_CONTENT_TYPE,
  normalizeWorkbookImportContentType,
} from './workbook-import-content-types.js'
import type { ImportedWorkbook } from './workbook-import-result.js'
import {
  assertXlsxByteInputApiWithinLimit,
  assertXlsxInspectionWithinMaterializationLimits,
  planXlsxImportRoute,
  resolveXlsxImportLimits,
  xlsxByteInputApiLimit,
  type XlsxImportOptions,
} from './xlsx-import-limits.js'
import { tryInspectLargeSimpleXlsxHeadless, type LargeSimpleXlsxHeadlessInspectResult } from './xlsx-large-simple-headless-inspect.js'
import { tryImportLargeSimpleXlsx } from './xlsx-large-simple-import.js'
import { readXlsxZipEntriesLazy, readXlsxZipEntriesLazyFromByteSource, type XlsxZipByteSource, type XlsxZipEntries } from './xlsx-zip.js'

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
export type { ImportedWorkbookSheetPreview } from './workbook-import-helpers.js'
export type { ImportedWorkbookPreview } from './workbook-import-preview.js'
export type { ImportedWorkbook } from './workbook-import-result.js'
export { manualCalculationModeWarning, precisionAsDisplayedCalculationWarning } from './xlsx-calculation-settings.js'
export { XlsxImportSizeLimitExceededError } from './xlsx-import-limits.js'
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
  dataTableFormulasWarning,
  definedNameFormulaCachesWarning,
  externalPivotCachesWarning,
  externalWorkbookCompanionAmbiguousMatchWarning,
  externalWorkbookCompanionNoMatchWarning,
  externalWorkbookReferencesWarning,
  macroExecutionDeclinedWarning,
  unsupportedCellStylesWarning,
  unsupportedFormulaCachesWarning,
  volatileFormulasWarning,
} from './xlsx-import-warnings.js'

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

function readValidXlsxZipContainer(bytes: Uint8Array): XlsxZipEntries {
  try {
    const zip = readXlsxZipEntriesLazy(bytes)
    void zip['xl/workbook.xml']
    return zip
  } catch {
    throw new InvalidXlsxZipContainerError()
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

function readValidXlsxZipContainerFromByteSource(source: XlsxZipByteSource): XlsxZipEntries {
  try {
    const zip = readXlsxZipEntriesLazyFromByteSource(source)
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
  source: { readonly byteLength: number },
  fileName: string,
  zip: XlsxZipEntries,
  options: { readonly minByteLength?: number } = {},
): LargeSimpleXlsxHeadlessInspectResult | null {
  return tryInspectLargeSimpleXlsxHeadless(source, fileName, zip, {
    allowUnsupportedWorksheetFeaturesForMetrics: true,
    ...(options.minByteLength !== undefined ? { minByteLength: options.minByteLength } : {}),
    releaseZipSource: true,
  })
}

export function inspectXlsx(bytes: Uint8Array | ArrayBuffer, fileName: string): XlsxHeadlessInspectResult | null {
  const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)
  assertXlsxByteInputApiWithinLimit(data.byteLength, 'inspectXlsx')
  return inspectLargeSimpleXlsxSource({ byteLength: data.byteLength }, fileName, readValidXlsxZipContainer(data), {
    minByteLength: 0,
  })
}

export function importXlsx(bytes: Uint8Array | ArrayBuffer, fileName: string, options: XlsxImportOptions = {}): ImportedWorkbook {
  const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)
  const sourceByteLength = data.byteLength
  assertXlsxByteInputApiWithinLimit(sourceByteLength, 'importXlsx')
  const workbookZip = readValidXlsxZipContainer(data)
  const limits = resolveXlsxImportLimits(options)
  let route = planXlsxImportRoute({ workbookZip, sourceByteLength, options, inspection: null })
  let inspection = route.shouldInspectBeforeLargeSimpleRouting
    ? inspectLargeSimpleXlsxSource({ byteLength: sourceByteLength }, fileName, workbookZip, route.inspectionOptions)
    : null
  route = planXlsxImportRoute({ workbookZip, sourceByteLength, options, inspection })
  assertXlsxInspectionWithinMaterializationLimits(inspection, limits)
  if (!route.shouldTryLargeSimpleImport) {
    throw new Error('Browser XLSX import only supports native simple workbooks.')
  }
  const imported = tryImportLargeSimpleXlsx({ byteLength: sourceByteLength }, fileName, workbookZip, route.createLargeSimpleImportOptions())
  if (imported) {
    return imported
  }
  if (route.shouldRetryDataOnlyLargeSimpleImport) {
    const retryZip = readValidXlsxZipContainer(data)
    const retryImport = tryImportLargeSimpleXlsx(
      { byteLength: sourceByteLength },
      fileName,
      retryZip,
      route.createLargeSimpleImportOptions({ materializeMetadata: false }),
    )
    if (retryImport) {
      return retryImport
    }
  }
  throw new Error(
    options.nativeOnly === true
      ? 'Native XLSX import could not materialize this workbook in the browser.'
      : 'Browser XLSX import requires native simple workbook support.',
  )
}

function importXlsxFromBrowserByteSource(
  source: XlsxZipByteSource,
  fileName: string,
  options: XlsxImportOptions | undefined,
): ImportedWorkbook {
  const workbookZip = readValidXlsxZipContainerFromByteSource(source)
  const sourceByteLength = source.byteLength
  const limits = resolveXlsxImportLimits(options ?? {})
  let route = planXlsxImportRoute({ workbookZip, sourceByteLength, options: options ?? {}, inspection: null })
  let inspection = route.shouldInspectBeforeLargeSimpleRouting
    ? inspectLargeSimpleXlsxSource(source, fileName, workbookZip, route.inspectionOptions)
    : null
  route = planXlsxImportRoute({ workbookZip, sourceByteLength, options: options ?? {}, inspection })
  assertXlsxInspectionWithinMaterializationLimits(inspection, limits)
  const imported = route.shouldTryLargeSimpleImport
    ? tryImportLargeSimpleXlsx({ byteLength: sourceByteLength }, fileName, workbookZip, route.createLargeSimpleImportOptions())
    : null
  if (imported) {
    return imported
  }
  throw new Error('Browser XLSX import could not materialize this workbook without a Node file-backed importer.')
}

export function importXlsm(bytes: Uint8Array | ArrayBuffer, fileName: string, options: XlsxImportOptions = {}): ImportedWorkbook {
  return importXlsx(bytes, fileName, { ...options, nativeOnly: true })
}

export function importXlsb(_bytes: Uint8Array | ArrayBuffer, _fileName: string): ImportedWorkbook {
  throw new Error('Browser XLSB import is not supported without the Node workbook importer.')
}

export function importXls(_bytes: Uint8Array | ArrayBuffer, _fileName: string): ImportedWorkbook {
  throw new Error('Browser legacy XLS import is not supported without the Node workbook importer.')
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
      ? importXlsxFromBrowserByteSource(xlsxZipByteSourceFromBytes(data), fileName, options.xlsx)
      : importXlsx(data, fileName, options.xlsx)
  }
  if (normalizedContentType === XLSM_CONTENT_TYPE) {
    return importXlsm(bytes, fileName, options.xlsx)
  }
  if (normalizedContentType === XLSB_CONTENT_TYPE) {
    return importXlsb(bytes, fileName)
  }
  if (normalizedContentType === LEGACY_XLS_CONTENT_TYPE) {
    return importXls(bytes, fileName)
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

export function exportXlsx(_snapshot: WorkbookSnapshot): Uint8Array {
  throw new Error('Browser XLSX export is not exposed from @bilig/excel-import/browser.')
}

export function exportXlsxToFile(_snapshot: WorkbookSnapshot, _outputPath: string): never {
  throw new Error('Browser XLSX file export is only available from the Node @bilig/excel-import entrypoint.')
}
