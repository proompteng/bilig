export type FormulaComparisonClassification =
  | 'bilig_matches_excel'
  | 'bilig_mismatches_excel'
  | 'cache_stale_bilig_matches_excel'
  | 'cache_stale_bilig_mismatches_excel'
  | 'cache_fresh_bilig_mismatches_excel'
  | 'missing_excel_oracle'
  | 'volatile_skipped'
  | 'parser_failure'
  | 'timeout_failure'

export type NormalizedFormulaValue =
  | { kind: 'blank' }
  | { kind: 'boolean'; value: boolean }
  | { kind: 'error'; value: string }
  | { kind: 'number'; value: number }
  | { kind: 'string'; value: string }

export interface FormulaComparisonInput {
  readonly actualBiligValue?: NormalizedFormulaValue
  readonly embeddedCacheValue?: NormalizedFormulaValue
  readonly excelOracleValue?: NormalizedFormulaValue
  readonly formula: string
  readonly parserFailed?: boolean
  readonly timedOut?: boolean
  readonly volatile?: boolean
}

export interface SanitizedFormulaSample {
  readonly actualBiligValue?: NormalizedFormulaValue
  readonly address: string
  readonly classification: FormulaComparisonClassification
  readonly embeddedCacheValue?: NormalizedFormulaValue
  readonly expectedExcelValue?: NormalizedFormulaValue
  readonly formula: string
  readonly functionFamilies: readonly string[]
  readonly reproNotes: string
  readonly sheet: string
  readonly workbookId: string
}

export interface FormulaCellComparison extends SanitizedFormulaSample {
  readonly cacheMatchesExcel?: boolean
  readonly biligMatchesExcel?: boolean
}

export interface WorkbookEvaluation {
  readonly elapsedMs: number
  readonly error?: string
  readonly formulaCells: number
  readonly id: string
  readonly status: 'ok' | 'parser_failure' | 'timeout_failure'
  readonly workbook: string
  readonly comparisons: readonly FormulaCellComparison[]
}

export interface OracleHarnessSummary {
  readonly biligVsFreshExcelMatchRate: number | null
  readonly cacheOnlyDiagnosticCells: number
  readonly comparableFormulaCells: number
  readonly embeddedCacheFreshnessRate: number | null
  readonly importParserFailures: number
  readonly realBiligMismatches: number
  readonly staleCacheFalsePositives: number
  readonly timeoutFailures: number
  readonly topMismatchFormulaFamilies: readonly { readonly family: string; readonly count: number }[]
  readonly totalFormulaCells: number
  readonly totalWorkbooksEvaluated: number
}

export interface OracleHarnessReport {
  readonly generatedAt: string
  readonly mode: 'cache-diagnostic' | 'excel-oracle'
  readonly notes: readonly string[]
  readonly schemaVersion: 1
  readonly summary: OracleHarnessSummary
  readonly workbooks: readonly WorkbookEvaluation[]
}

export const volatileFormulaPattern = /\b(CELL|INFO|INDIRECT|NOW|OFFSET|RAND|RANDBETWEEN|TODAY)\s*\(/iu

export const trueMismatchClassifications = new Set<FormulaComparisonClassification>([
  'bilig_mismatches_excel',
  'cache_stale_bilig_mismatches_excel',
  'cache_fresh_bilig_mismatches_excel',
])

export function classifyFormulaComparison(input: FormulaComparisonInput): FormulaComparisonClassification {
  if (input.timedOut) {
    return 'timeout_failure'
  }
  if (input.parserFailed) {
    return 'parser_failure'
  }
  if (input.volatile || volatileFormulaPattern.test(input.formula)) {
    return 'volatile_skipped'
  }
  if (input.excelOracleValue === undefined || input.actualBiligValue === undefined) {
    return 'missing_excel_oracle'
  }

  const biligMatchesExcel = normalizedValuesEqual(input.actualBiligValue, input.excelOracleValue)
  const cacheMatchesExcel =
    input.embeddedCacheValue === undefined ? undefined : normalizedValuesEqual(input.embeddedCacheValue, input.excelOracleValue)

  if (biligMatchesExcel) {
    return cacheMatchesExcel === false ? 'cache_stale_bilig_matches_excel' : 'bilig_matches_excel'
  }
  if (cacheMatchesExcel === false) {
    return 'cache_stale_bilig_mismatches_excel'
  }
  if (cacheMatchesExcel === true) {
    return 'cache_fresh_bilig_mismatches_excel'
  }
  return 'bilig_mismatches_excel'
}

export function buildReportSummary(report: Pick<OracleHarnessReport, 'workbooks'>): OracleHarnessSummary {
  const comparisons = report.workbooks.flatMap((workbook) => workbook.comparisons)
  const trueOracleComparisons = comparisons.filter((comparison) => comparison.biligMatchesExcel !== undefined)
  const cacheFreshnessComparisons = comparisons.filter((comparison) => comparison.cacheMatchesExcel !== undefined)
  const realMismatches = comparisons.filter((comparison) => trueMismatchClassifications.has(comparison.classification))
  const familyCounts = new Map<string, number>()
  for (const mismatch of realMismatches) {
    for (const family of mismatch.functionFamilies) {
      familyCounts.set(family, (familyCounts.get(family) ?? 0) + 1)
    }
  }

  return {
    totalWorkbooksEvaluated: report.workbooks.length,
    importParserFailures: report.workbooks.filter((workbook) => workbook.status === 'parser_failure').length,
    timeoutFailures: report.workbooks.filter((workbook) => workbook.status === 'timeout_failure').length,
    totalFormulaCells: report.workbooks.reduce((sum, workbook) => sum + workbook.formulaCells, 0),
    comparableFormulaCells: trueOracleComparisons.length,
    biligVsFreshExcelMatchRate:
      trueOracleComparisons.length === 0
        ? null
        : ratio(trueOracleComparisons.filter((comparison) => comparison.biligMatchesExcel).length, trueOracleComparisons.length),
    embeddedCacheFreshnessRate:
      cacheFreshnessComparisons.length === 0
        ? null
        : ratio(cacheFreshnessComparisons.filter((comparison) => comparison.cacheMatchesExcel).length, cacheFreshnessComparisons.length),
    staleCacheFalsePositives: comparisons.filter((comparison) => comparison.classification === 'cache_stale_bilig_matches_excel').length,
    realBiligMismatches: realMismatches.length,
    cacheOnlyDiagnosticCells: comparisons.filter((comparison) => comparison.classification === 'missing_excel_oracle').length,
    topMismatchFormulaFamilies: [...familyCounts.entries()]
      .toSorted((left, right) => right[1] - left[1] || left[0].localeCompare(right[0]))
      .slice(0, 10)
      .map(([family, count]) => ({ family, count })),
  }
}

export function normalizedValuesEqual(left: NormalizedFormulaValue, right: NormalizedFormulaValue): boolean {
  if (left.kind !== right.kind) {
    return false
  }
  switch (left.kind) {
    case 'blank':
      return true
    case 'boolean':
      return right.kind === 'boolean' && left.value === right.value
    case 'error':
      return right.kind === 'error' && left.value === right.value
    case 'string':
      return right.kind === 'string' && left.value === right.value
    case 'number':
      return right.kind === 'number' && numbersEqual(left.value, right.value)
  }
}

export function sanitizeFormula(formula: string): string {
  return formula
    .replace(/"(?:""|[^"])*"/gu, '"<text>"')
    .replace(/\[[^\]]+\]/gu, '[workbook]')
    .replace(/'[^']+'!/gu, "'<sheet>'!")
    .replace(/\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b/giu, '<email>')
}

export function sanitizeErrorMessage(message: string): string {
  return message
    .replace(/\/(?:Users|home|var|tmp|private)\/[^\s'"]+/gu, '<path>')
    .replace(/[A-Z]:\\[^\s'"]+/giu, '<path>')
    .replace(/\b[0-9a-f]{32,64}\b/giu, '<hash>')
}

export function formulaFamilies(formula: string): string[] {
  const families = new Set<string>()
  for (const match of formula.matchAll(/\b([A-Z][A-Z0-9.]*)\s*\(/giu)) {
    families.add(match[1]?.toUpperCase() ?? 'UNKNOWN')
  }
  return families.size === 0 ? ['arithmetic'] : [...families].toSorted((left, right) => left.localeCompare(right))
}

export function reproNotesFor(classification: FormulaComparisonClassification): string {
  switch (classification) {
    case 'bilig_matches_excel':
      return 'Bilig matched a fresh Excel recalculation.'
    case 'bilig_mismatches_excel':
    case 'cache_fresh_bilig_mismatches_excel':
    case 'cache_stale_bilig_mismatches_excel':
      return 'Fresh Excel expected value, Bilig actual value, and sanitized formula are present; this can be investigated as a correctness bug.'
    case 'cache_stale_bilig_matches_excel':
      return 'Embedded cache was stale, but Bilig matched fresh Excel. Treat any cache-only mismatch as a false positive.'
    case 'missing_excel_oracle':
      return 'No fresh Excel oracle was available. This cell is diagnostic only.'
    case 'parser_failure':
      return 'Workbook import or parser setup failed before an authoritative formula comparison could run.'
    case 'timeout_failure':
      return 'Workbook execution timed out before an authoritative formula comparison could run.'
    case 'volatile_skipped':
      return 'Volatile or environment-dependent formula skipped.'
  }
}

export function formatNormalizedValue(value: NormalizedFormulaValue | undefined): string {
  if (!value) {
    return '<missing>'
  }
  switch (value.kind) {
    case 'blank':
      return '<blank>'
    case 'boolean':
    case 'error':
    case 'number':
      return String(value.value)
    case 'string':
      return JSON.stringify(value.value)
  }
}

export function formatNullableRate(value: number | null): string {
  return value === null ? 'n/a' : `${String(Number((value * 100).toFixed(2)))}%`
}

function numbersEqual(left: number, right: number): boolean {
  if (Object.is(left, right)) {
    return true
  }
  const scale = Math.max(1, Math.abs(left), Math.abs(right))
  return Math.abs(left - right) <= Math.max(1e-7, scale * 1e-12)
}

function ratio(numerator: number, denominator: number): number {
  return denominator === 0 ? 0 : Number((numerator / denominator).toFixed(6))
}
