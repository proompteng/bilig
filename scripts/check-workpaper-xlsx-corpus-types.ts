import type { WorkPaperSheet, WorkPaperSheets } from '@bilig/workpaper'

export type CellContent = WorkPaperSheet[number][number]

export type CachedFormulaValue =
  | { kind: 'blank' }
  | { kind: 'boolean'; value: boolean }
  | { kind: 'error'; value: string }
  | { kind: 'number'; value: number }
  | { kind: 'string'; value: string }

export type WorkPaperXlsxFormulaSkipReason =
  | 'missing-cached-result'
  | 'stale-cached-result'
  | 'stale-cached-name-error'
  | 'unsupported-cached-result-type'
  | 'volatile-or-environment-dependent-formula'

export interface WorkPaperXlsxCorpusOptions {
  readonly allowLargeWorkPaperMaterialization?: boolean
  readonly childProcessTimeoutMs?: number
  readonly evaluationTimeoutMs?: number
  readonly maxFileBytes?: number
  readonly mismatchSampleLimit?: number
}

export interface WorkPaperXlsxCorpusResult {
  readonly summary: WorkPaperXlsxCorpusSummary
  readonly files: readonly WorkPaperXlsxCorpusFileResult[]
  readonly mismatches: readonly WorkPaperXlsxCorpusMismatch[]
  readonly skippedByReason: Readonly<Record<WorkPaperXlsxFormulaSkipReason, number>>
}

export interface WorkPaperXlsxCorpusCompatibilitySummary {
  readonly formulaContextDiagnosticCount: number
  readonly staleCacheRiskFormulaCount: number
  readonly externalCacheOnlyPivotCount: number
  readonly unsupportedRefreshFeatureCount: number
}

export interface WorkPaperXlsxCorpusSummary {
  readonly totalFiles: number
  readonly filesProcessed: number
  readonly ok: number
  readonly failedTimeouts: number
  readonly failedErrors: number
  readonly formulaCells: number
  readonly comparableFormulaCells: number
  readonly matchingFormulaCells: number
  readonly mismatchedFormulaCells: number
  readonly skippedFormulaCells: number
  readonly matchRate: number
  readonly compatibility: WorkPaperXlsxCorpusCompatibilitySummary
  readonly elapsedMs: number
}

export interface WorkPaperXlsxCorpusFileResult {
  readonly path: string
  readonly fileName: string
  readonly status: 'error' | 'mismatched' | 'ok' | 'timeout'
  readonly formulaCells: number
  readonly comparableFormulaCells: number
  readonly matchingFormulaCells: number
  readonly mismatchedFormulaCells: number
  readonly skippedFormulaCells: number
  readonly matchRate: number
  readonly compatibility: WorkPaperXlsxCorpusCompatibilitySummary
  readonly elapsedMs: number
  readonly error?: string
}

export interface WorkPaperXlsxCorpusMismatch {
  readonly path: string
  readonly fileName: string
  readonly sheetName: string
  readonly address: string
  readonly formula: string
  readonly expected: CachedFormulaValue
  readonly actual: CachedFormulaValue
}

export interface FormulaCellRecord {
  readonly sheetName: string
  readonly address: string
  readonly row: number
  readonly col: number
  readonly formula: string
  readonly cachedValue?: CachedFormulaValue
  readonly skipReason?: WorkPaperXlsxFormulaSkipReason
}

export interface PreparedWorkbook {
  readonly sheets: WorkPaperSheets
  readonly formulaCells: readonly FormulaCellRecord[]
  readonly compatibility: WorkPaperXlsxCorpusCompatibilitySummary
  readonly maxRows: number
  readonly maxColumns: number
}
