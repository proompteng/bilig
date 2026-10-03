export * from './xlsx-runtime.js'
export {
  WorkPaper,
  inspectXlsxCache,
  inspectXlsxCacheFileStreamingNative,
  parseQualifiedA1,
  parseQualifiedCellTarget,
  recalculateSheetjsWorkbook,
  recalculateXlsx,
  recalculateXlsxFileToFile,
  recalculateXlsxToFile,
  StreamingNativeXlsxRecalcError,
  xlsxCacheDoctorSchemaVersion,
} from './xlsx-recalc.js'
export type {
  StreamingNativeFormulaCounts,
  StreamingNativeXlsxCacheFormulaInspection,
  StreamingNativeXlsxCacheInspectionLimit,
  StreamingNativeXlsxCacheInspectionResult,
  StreamingNativeXlsxCacheLiteral,
  StreamingNativeXlsxCacheStatus,
  StreamingNativeXlsxCacheStatusSummary,
  XlsxCacheFormulaInspection,
  XlsxCacheInspectionLimit,
  XlsxCacheInspectionOptions,
  XlsxCacheInspectionResult,
  XlsxCacheStatus,
  XlsxCacheStatusSummary,
  XlsxFormulaRecalcCellValue,
  XlsxFormulaRecalcChange,
  XlsxFormulaRecalcDiagnostics,
  XlsxFormulaRecalcEdit,
  XlsxFormulaRecalcFileToFileEngine,
  XlsxFormulaRecalcFileToFileFallbackPolicy,
  XlsxFormulaRecalcFileToFileOptions,
  XlsxFormulaRecalcFileOptions,
  XlsxFormulaRecalcFileResult,
  XlsxFormulaRecalcNativeDiagnostics,
  XlsxFormulaRecalcOptions,
  XlsxFormulaRecalcPhaseRss,
  XlsxFormulaRecalcResult,
  XlsxFormulaRecalcWorkPaperConfig,
  XlsxFormulaRecalcWorkPaperEngine,
  XlsxFormulaRecalcWorkPaperFallbackPolicy,
} from './xlsx-recalc.js'

import { importXlsx as importHeadlessXlsx, type XlsxImportOptions } from './xlsx-runtime.js'

export function importXlsx(bytes: Uint8Array | ArrayBuffer, fileName: string, options: XlsxImportOptions = {}) {
  return importHeadlessXlsx(bytes, fileName, {
    preferNativeSimpleImport: true,
    ...options,
  })
}
