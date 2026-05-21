import type { ImportedWorkbook } from './workbook-import-result.js'
import { importXlsx } from './index.js'
import {
  assertXlsxInspectionWithinMaterializationLimits,
  denseSheetJsByteThreshold,
  largeCalcChainStreamingFormulaThreshold,
  resolveXlsxImportLimits,
  shouldRetryDataOnlyLargeSimpleImport,
  type XlsxImportOptions,
} from './xlsx-import-limits.js'
import { tryInspectLargeSimpleXlsxHeadless } from './xlsx-large-simple-headless-inspect.js'
import { tryImportLargeSimpleXlsx } from './xlsx-large-simple-import.js'
import {
  hasFullImporterOnlyPackageMetadata,
  shouldBypassLargeSimpleByteThresholdForPackageArtifacts,
} from './xlsx-large-simple-package-artifact-threshold.js'
import { attachImportedXlsxSourceReader, detachImportedXlsxSourceBytes } from './xlsx-source-bytes.js'
import { readXlsxZipEntriesLazyFromByteSource, type XlsxZipByteSource } from './xlsx-zip.js'

const largeCalcChainStreamingByteThreshold = 5_000_000

export interface XlsxByteSourceImportOptions extends XlsxImportOptions {
  readonly attachSourceReaderForUntouchedExport?: boolean
}

export function importXlsxFromZipByteSource(
  source: XlsxZipByteSource,
  fileName: string,
  options: XlsxByteSourceImportOptions = {},
): ImportedWorkbook {
  const workbookZip = readXlsxZipEntriesLazyFromByteSource(borrowXlsxZipByteSource(source))
  if (!workbookZip) {
    return importXlsxFromMaterializedSource(source, fileName, options)
  }
  const limits = resolveXlsxImportLimits(options)
  const hasCalcChain = Object.hasOwn(workbookZip, 'xl/calcChain.xml')
  const bypassLargeSimpleByteThreshold =
    shouldBypassLargeSimpleByteThresholdForPackageArtifacts(workbookZip) && !hasFullImporterOnlyPackageMetadata(workbookZip)
  const needsCalcChainFormulaCountInspection =
    hasCalcChain && source.byteLength >= denseSheetJsByteThreshold && source.byteLength < largeCalcChainStreamingByteThreshold
  const inspection =
    limits || needsCalcChainFormulaCountInspection
      ? inspectLargeSimpleXlsxSource(source, fileName, options.limits ? { minByteLength: 0 } : undefined)
      : null
  assertXlsxInspectionWithinMaterializationLimits(inspection, limits)
  const hasLargeCalcChainFormulaSet = hasCalcChain && (inspection?.stats.formulaCellCount ?? 0) >= largeCalcChainStreamingFormulaThreshold
  const allowCachedUnsupportedFormulaText =
    hasCalcChain && (source.byteLength >= largeCalcChainStreamingByteThreshold || hasLargeCalcChainFormulaSet)
  const shouldTryLargeSimpleImport =
    !hasCalcChain ||
    source.byteLength >= largeCalcChainStreamingByteThreshold ||
    hasLargeCalcChainFormulaSet ||
    bypassLargeSimpleByteThreshold
  const largeSimpleImportOptions = {
    ...(options.limits || bypassLargeSimpleByteThreshold ? { minByteLength: 0 } : {}),
    allowUnsupportedFormulaText: allowCachedUnsupportedFormulaText,
    allowUnsupportedCellMetadata: allowCachedUnsupportedFormulaText,
    releaseArenaAfterMaterialization: true,
    releaseZipSource: true,
    maxMaterializedLazyPackageArtifactBytes: 8 * 1024 * 1024,
  }
  let largeSimpleImport = !shouldTryLargeSimpleImport
    ? null
    : tryImportLargeSimpleXlsx({ byteLength: source.byteLength }, fileName, workbookZip, largeSimpleImportOptions)
  if (!largeSimpleImport && shouldRetryDataOnlyLargeSimpleImport(inspection, source.byteLength, allowCachedUnsupportedFormulaText)) {
    const retryZip = readXlsxZipEntriesLazyFromByteSource(borrowXlsxZipByteSource(source))
    largeSimpleImport = retryZip
      ? tryImportLargeSimpleXlsx({ byteLength: source.byteLength }, fileName, retryZip, {
          ...largeSimpleImportOptions,
          materializeMetadata: false,
        })
      : null
  }
  if (largeSimpleImport) {
    if (options.attachSourceReaderForUntouchedExport !== false) {
      attachImportedXlsxSourceReader(largeSimpleImport.snapshot, {
        byteLength: source.byteLength,
        readBytes: () => readAllSourceBytes(source),
      })
    }
    return largeSimpleImport
  }
  return importXlsxFromMaterializedSource(source, fileName, options)
}

function importXlsxFromMaterializedSource(
  source: XlsxZipByteSource,
  fileName: string,
  options: XlsxByteSourceImportOptions,
): ImportedWorkbook {
  const imported = importXlsx(readAllSourceBytes(source), fileName, options)
  if (options.attachSourceReaderForUntouchedExport === false) {
    detachImportedXlsxSourceBytes(imported.snapshot)
  }
  return imported
}

function inspectLargeSimpleXlsxSource(
  source: XlsxZipByteSource,
  fileName: string,
  options: { readonly minByteLength?: number } = {},
): ReturnType<typeof tryInspectLargeSimpleXlsxHeadless> {
  const inspectionZip = readXlsxZipEntriesLazyFromByteSource(borrowXlsxZipByteSource(source))
  return inspectionZip
    ? tryInspectLargeSimpleXlsxHeadless({ byteLength: source.byteLength }, fileName, inspectionZip, {
        allowUnsupportedWorksheetFeaturesForMetrics: true,
        ...(options.minByteLength !== undefined ? { minByteLength: options.minByteLength } : {}),
        releaseZipSource: true,
      })
    : null
}

function readAllSourceBytes(source: XlsxZipByteSource): Uint8Array {
  return source.readRange(0, source.byteLength)
}

function borrowXlsxZipByteSource(source: XlsxZipByteSource): XlsxZipByteSource {
  return {
    byteLength: source.byteLength,
    readRange: (start, end) => source.readRange(start, end),
  }
}
