import type { LargeSimpleXlsxHeadlessInspectResult } from './xlsx-large-simple-headless-inspect.js'

export const denseSheetJsByteThreshold = 1_000_000
export const largeCalcChainStreamingFormulaThreshold = 50_000

export interface XlsxImportLimits {
  maxMaterializedCells?: number
  maxMaterializedFormulaCells?: number
}

export interface XlsxImportOptions {
  limits?: XlsxImportLimits | false
}

export class XlsxImportSizeLimitExceededError extends Error {
  readonly limits: Required<XlsxImportLimits>
  readonly stats: LargeSimpleXlsxHeadlessInspectResult['stats']
  readonly reason: 'cell-count' | 'formula-cell-count'

  constructor(args: {
    reason: 'cell-count' | 'formula-cell-count'
    limits: Required<XlsxImportLimits>
    stats: LargeSimpleXlsxHeadlessInspectResult['stats']
  }) {
    const observed = args.reason === 'cell-count' ? args.stats.cellCount : args.stats.formulaCellCount
    const limit = args.reason === 'cell-count' ? args.limits.maxMaterializedCells : args.limits.maxMaterializedFormulaCells
    super(
      `XLSX import exceeds the materialized ${args.reason === 'cell-count' ? 'cell' : 'formula cell'} limit ` +
        `(${observed.toLocaleString('en-US')} > ${limit.toLocaleString('en-US')}). ` +
        'Use inspectXlsx() for bounded metadata, raise importXlsx limits explicitly, or split the workbook before materializing it.',
    )
    this.name = 'XlsxImportSizeLimitExceededError'
    this.reason = args.reason
    this.limits = args.limits
    this.stats = args.stats
  }
}

export function resolveXlsxImportLimits(options: XlsxImportOptions): Required<XlsxImportLimits> | null {
  if (options.limits === false || options.limits === undefined) {
    return null
  }
  return {
    maxMaterializedCells: options.limits.maxMaterializedCells ?? Number.POSITIVE_INFINITY,
    maxMaterializedFormulaCells: options.limits.maxMaterializedFormulaCells ?? Number.POSITIVE_INFINITY,
  }
}

export function assertXlsxInspectionWithinMaterializationLimits(
  inspection: LargeSimpleXlsxHeadlessInspectResult | null,
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

export function shouldRetryDataOnlyLargeSimpleImport(
  inspection: LargeSimpleXlsxHeadlessInspectResult | null,
  sourceByteLength: number,
  allowCachedUnsupportedFormulaText: boolean,
): boolean {
  if (sourceByteLength >= denseSheetJsByteThreshold) {
    return true
  }
  const stats = inspection?.stats
  return (
    allowCachedUnsupportedFormulaText &&
    stats !== undefined &&
    (stats.cellCount >= denseSheetJsByteThreshold || stats.formulaCellCount >= largeCalcChainStreamingFormulaThreshold)
  )
}
