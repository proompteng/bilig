import type { ImportedWorkbook } from './workbook-import-result.js'
import { importXlsx, XlsxImportSizeLimitExceededError, type XlsxImportLimits, type XlsxImportOptions } from './index.js'
import { tryInspectLargeSimpleXlsxHeadless } from './xlsx-large-simple-headless-inspect.js'
import { tryImportLargeSimpleXlsx } from './xlsx-large-simple-import.js'
import { attachImportedXlsxSourceReader } from './xlsx-source-bytes.js'
import { readXlsxZipEntriesLazyFromByteSource, type XlsxZipByteSource } from './xlsx-zip.js'

const denseSheetJsByteThreshold = 1_000_000
const largeCalcChainStreamingByteThreshold = 5_000_000
const largeCalcChainStreamingFormulaThreshold = 50_000

export function importXlsxFromZipByteSource(
  source: XlsxZipByteSource,
  fileName: string,
  options: XlsxImportOptions = {},
): ImportedWorkbook {
  const workbookZip = readXlsxZipEntriesLazyFromByteSource(borrowXlsxZipByteSource(source))
  if (!workbookZip) {
    return importXlsx(readAllSourceBytes(source), fileName, options)
  }
  const limits = resolveXlsxImportLimits(options)
  const hasCalcChain = Object.hasOwn(workbookZip, 'xl/calcChain.xml')
  const inspection =
    limits || (hasCalcChain && source.byteLength >= denseSheetJsByteThreshold)
      ? inspectLargeSimpleXlsxSource(source, fileName, options.limits ? { minByteLength: 0 } : undefined)
      : null
  assertXlsxInspectionWithinMaterializationLimits(inspection, limits)
  const hasLargeCalcChainFormulaSet = hasCalcChain && (inspection?.stats.formulaCellCount ?? 0) >= largeCalcChainStreamingFormulaThreshold
  const allowCachedUnsupportedFormulaText =
    hasCalcChain && (source.byteLength >= largeCalcChainStreamingByteThreshold || hasLargeCalcChainFormulaSet)
  const shouldTryLargeSimpleImport =
    !hasCalcChain || source.byteLength >= largeCalcChainStreamingByteThreshold || hasLargeCalcChainFormulaSet
  const largeSimpleImport = !shouldTryLargeSimpleImport
    ? null
    : tryImportLargeSimpleXlsx({ byteLength: source.byteLength }, fileName, workbookZip, {
        ...(options.limits ? { minByteLength: 0 } : {}),
        allowUnsupportedFormulaText: allowCachedUnsupportedFormulaText,
        allowUnsupportedCellMetadata: allowCachedUnsupportedFormulaText,
        releaseArenaAfterMaterialization: true,
        releaseZipSource: true,
      })
  if (largeSimpleImport) {
    attachImportedXlsxSourceReader(largeSimpleImport.snapshot, {
      byteLength: source.byteLength,
      readBytes: () => readAllSourceBytes(source),
    })
    return largeSimpleImport
  }
  return importXlsx(readAllSourceBytes(source), fileName, options)
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

function resolveXlsxImportLimits(options: XlsxImportOptions): Required<XlsxImportLimits> | null {
  if (options.limits === false || options.limits === undefined) {
    return null
  }
  return {
    maxMaterializedCells: options.limits.maxMaterializedCells ?? Number.POSITIVE_INFINITY,
    maxMaterializedFormulaCells: options.limits.maxMaterializedFormulaCells ?? Number.POSITIVE_INFINITY,
  }
}

function assertXlsxInspectionWithinMaterializationLimits(
  inspection: ReturnType<typeof tryInspectLargeSimpleXlsxHeadless>,
  limits: Required<XlsxImportLimits> | null,
): void {
  if (!inspection || !limits) {
    return
  }
  if (inspection.stats.cellCount > limits.maxMaterializedCells) {
    throw new XlsxImportSizeLimitExceededError({ reason: 'cell-count', limits, stats: inspection.stats })
  }
  if (inspection.stats.formulaCellCount > limits.maxMaterializedFormulaCells) {
    throw new XlsxImportSizeLimitExceededError({ reason: 'formula-cell-count', limits, stats: inspection.stats })
  }
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
