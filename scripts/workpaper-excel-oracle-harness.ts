#!/usr/bin/env bun

import { spawnSync } from 'node:child_process'
import { createHash } from 'node:crypto'
import { copyFileSync, existsSync, mkdirSync, readFileSync, readdirSync, statSync, unlinkSync, writeFileSync } from 'node:fs'
import { basename, dirname, extname, join, relative, resolve } from 'node:path'
import { performance } from 'node:perf_hooks'
import { fileURLToPath } from 'node:url'
import * as XLSX from 'xlsx'

import { attachRuntimeSnapshot } from '@bilig/core'
import { importXlsx } from '@bilig/excel-import'
import { WorkPaper, type WorkPaperConfig, type WorkPaperSheet, type WorkPaperSheets } from '@bilig/headless'
import { formatErrorCode, ValueTag, type CellValue } from '@bilig/protocol'

import {
  buildReportSummary,
  classifyFormulaComparison,
  formatNormalizedValue,
  formatNullableRate,
  formulaFamilies,
  normalizedValuesEqual,
  reproNotesFor,
  sanitizeErrorMessage,
  sanitizeFormula,
  trueMismatchClassifications,
  volatileFormulaPattern,
  type FormulaCellComparison,
  type FormulaComparisonClassification,
  type NormalizedFormulaValue,
  type OracleHarnessReport,
  type WorkbookEvaluation,
} from './workpaper-excel-oracle-harness-core.ts'

export { buildReportSummary, classifyFormulaComparison } from './workpaper-excel-oracle-harness-core.ts'
export type {
  FormulaCellComparison,
  FormulaComparisonClassification,
  FormulaComparisonInput,
  NormalizedFormulaValue,
  OracleHarnessReport,
  OracleHarnessSummary,
  SanitizedFormulaSample,
  WorkbookEvaluation,
} from './workpaper-excel-oracle-harness-core.ts'

interface FormulaCellRecord {
  readonly address: string
  readonly cachedValue?: NormalizedFormulaValue
  readonly col: number
  readonly formula: string
  readonly row: number
  readonly sheetIndex: number
  readonly sheetName: string
  readonly volatile: boolean
}

interface PreparedWorkbook {
  readonly formulaCells: readonly FormulaCellRecord[]
  readonly maxColumns: number
  readonly maxRows: number
  readonly sheets: WorkPaperSheets
}

interface HarnessOptions {
  readonly sampleLimit: number
  readonly timeoutMs: number
}

const defaultSampleLimit = 25
const defaultTimeoutMs = 30_000
const ignoredDirectoryNames = new Set(['.git', 'build', 'dist', 'node_modules'])
const workbookExtensions = new Set(['.xls', '.xlsm', '.xlsx'])
const xlsxErrorTextByCode = new Map<number, string>([
  [0, '#NULL!'],
  [7, '#DIV/0!'],
  [15, '#VALUE!'],
  [23, '#REF!'],
  [29, '#NAME?'],
  [36, '#NUM!'],
  [42, '#N/A'],
  [43, '#GETTING_DATA'],
])

export function runCacheDiagnostic(inputDir: string, outputDir: string, options: HarnessOptions): OracleHarnessReport {
  const files = collectWorkbookFiles(inputDir)
  const workbooks = files.map((filePath) => evaluateWorkbookAgainstOracle({ inputDir, filePath, oraclePath: undefined, options }))
  const report = buildReport('cache-diagnostic', workbooks, [
    'Embedded XLSX cached formula values are diagnostic only and are not an accuracy oracle.',
    'A cache-only mismatch must not be promoted to a Bilig correctness bug.',
  ])
  writeJson(join(outputDir, 'cache-diagnostic.json'), report)
  return report
}

export function runExcelOracleEvaluation(
  originalDir: string,
  recalculatedDir: string,
  outputDir: string,
  options: HarnessOptions,
): OracleHarnessReport {
  const files = collectWorkbookFiles(originalDir)
  const workbooks = files.map((filePath) => {
    const oraclePath = join(recalculatedDir, relative(originalDir, filePath))
    return evaluateWorkbookAgainstOracle({
      inputDir: originalDir,
      filePath,
      oraclePath: existsSync(oraclePath) ? oraclePath : undefined,
      options,
    })
  })
  const report = buildReport('excel-oracle', workbooks, [
    'Fresh Excel-recalculated formula results are the only authoritative accuracy oracle in this report.',
    'OpenPyXL-style parsers and SheetJS are used only for extracting formulas and cached values, not for recalculation.',
  ])
  writeJson(join(outputDir, 'excel-oracle-report.json'), report)
  writeGithubIssueDrafts(report, join(outputDir, 'github-issues'))
  return report
}

export function writeSummary(reportDir: string): { readonly markdown: string; readonly report: OracleHarnessReport | null } {
  const oracleReport = readOptionalReport(join(reportDir, 'excel-oracle-report.json'))
  const cacheReport = readOptionalReport(join(reportDir, 'cache-diagnostic.json'))
  const report = oracleReport ?? cacheReport
  const markdown = report ? formatSummaryMarkdown(report, cacheReport) : '# WorkPaper Excel Oracle Summary\n\nNo report JSON files found.\n'
  writeFile(join(reportDir, 'summary.md'), `${markdown.trimEnd()}\n`)
  writeJson(
    join(reportDir, 'summary.json'),
    report ? { generatedAt: new Date().toISOString(), summary: report.summary } : { summary: null },
  )
  return { markdown, report }
}

export function prepareExcelOracle(inputDir: string, outputDir: string): void {
  const files = collectWorkbookFiles(inputDir)
  const recalculatedDir = join(outputDir, 'recalculated')
  const excelAvailable = isExcelAutomationAvailable()
  const workbooks = files.map((filePath) => {
    const relativePath = relative(inputDir, filePath)
    const outputPath = join(recalculatedDir, relativePath)
    if (!excelAvailable) {
      return {
        id: workbookIdForFile(filePath),
        workbook: sanitizedWorkbookName(workbookIdForFile(filePath), filePath),
        status: 'missing_excel_oracle',
      }
    }
    try {
      recalculateWorkbookWithExcel(filePath, outputPath)
      return {
        id: workbookIdForFile(filePath),
        workbook: sanitizedWorkbookName(workbookIdForFile(filePath), filePath),
        status: 'recalculated',
      }
    } catch (error) {
      return {
        error: sanitizeErrorMessage(errorMessage(error)),
        id: workbookIdForFile(filePath),
        workbook: sanitizedWorkbookName(workbookIdForFile(filePath), filePath),
        status: 'oracle_failure',
      }
    }
  })
  writeJson(join(outputDir, 'prepare-oracle-report.json'), {
    schemaVersion: 1,
    generatedAt: new Date().toISOString(),
    excelAvailable,
    inputWorkbookCount: files.length,
    recalculatedDir: excelAvailable ? 'recalculated' : null,
    workbooks,
    notes: excelAvailable
      ? ['Original files were preserved; recalculated copies were written under the output folder.']
      : ['Microsoft Excel automation was not available; evaluate-oracle will classify workbooks as missing_excel_oracle.'],
  })
}

function evaluateWorkbookAgainstOracle(args: {
  readonly filePath: string
  readonly inputDir: string
  readonly options: HarnessOptions
  readonly oraclePath?: string
}): WorkbookEvaluation {
  const startedAt = performance.now()
  const id = workbookIdForFile(args.filePath)
  const workbookName = sanitizedWorkbookName(id, args.filePath)
  let prepared: PreparedWorkbook | undefined
  try {
    prepared = prepareWorkbook(args.filePath)
    const oracleCells = args.oraclePath ? readFormulaCacheMap(args.oraclePath) : new Map<string, NormalizedFormulaValue>()
    const workbook = WorkPaper.buildFromSheets(prepared.sheets, workPaperConfigFor(prepared, args.options))
    try {
      const comparisons = compareFormulaCells({ id, oracleCells, prepared, sampleLimit: args.options.sampleLimit, workbook })
      return {
        id,
        workbook: workbookName,
        status: 'ok',
        formulaCells: prepared.formulaCells.length,
        comparisons,
        elapsedMs: roundElapsed(performance.now() - startedAt),
      }
    } finally {
      workbook.dispose()
    }
  } catch (error) {
    const status = isTimeoutError(error) ? 'timeout_failure' : 'parser_failure'
    return {
      id,
      workbook: workbookName,
      status,
      formulaCells: prepared?.formulaCells.length ?? 0,
      comparisons: prepared
        ? prepared.formulaCells.slice(0, args.options.sampleLimit).map((cell) =>
            comparisonForFailure({
              cell,
              classification: status,
              id,
            }),
          )
        : [],
      elapsedMs: roundElapsed(performance.now() - startedAt),
      error: sanitizeErrorMessage(errorMessage(error)),
    }
  }
}

function compareFormulaCells(args: {
  readonly id: string
  readonly oracleCells: ReadonlyMap<string, NormalizedFormulaValue>
  readonly prepared: PreparedWorkbook
  readonly sampleLimit: number
  readonly workbook: WorkPaper
}): FormulaCellComparison[] {
  const comparisons: FormulaCellComparison[] = []
  for (const cell of args.prepared.formulaCells) {
    const sheetId = args.workbook.getSheetId(cell.sheetName)
    const actualBiligValue =
      sheetId === undefined
        ? ({ kind: 'error', value: 'missing-sheet' } satisfies NormalizedFormulaValue)
        : normalizeCellValue(args.workbook.getCellValue({ sheet: sheetId, row: cell.row, col: cell.col }))
    const excelOracleValue = args.oracleCells.get(formulaCellKey(cell.sheetName, cell.address))
    const classification = classifyFormulaComparison({
      actualBiligValue,
      embeddedCacheValue: cell.cachedValue,
      excelOracleValue,
      formula: cell.formula,
      volatile: cell.volatile,
    })
    const cacheMatchesExcel =
      cell.cachedValue === undefined || excelOracleValue === undefined
        ? undefined
        : normalizedValuesEqual(cell.cachedValue, excelOracleValue)
    const biligMatchesExcel = excelOracleValue === undefined ? undefined : normalizedValuesEqual(actualBiligValue, excelOracleValue)
    if (comparisons.length < args.sampleLimit || trueMismatchClassifications.has(classification)) {
      comparisons.push({
        actualBiligValue,
        address: cell.address,
        biligMatchesExcel,
        cacheMatchesExcel,
        classification,
        embeddedCacheValue: cell.cachedValue,
        expectedExcelValue: excelOracleValue,
        formula: sanitizeFormula(cell.formula),
        functionFamilies: formulaFamilies(cell.formula),
        reproNotes: reproNotesFor(classification),
        sheet: `sheet-${String(cell.sheetIndex + 1)}`,
        workbookId: args.id,
      })
    }
  }
  return comparisons
}

function comparisonForFailure(args: {
  readonly cell: FormulaCellRecord
  readonly classification: FormulaComparisonClassification
  readonly id: string
}): FormulaCellComparison {
  return {
    address: args.cell.address,
    classification: args.classification,
    embeddedCacheValue: args.cell.cachedValue,
    formula: sanitizeFormula(args.cell.formula),
    functionFamilies: formulaFamilies(args.cell.formula),
    reproNotes: reproNotesFor(args.classification),
    sheet: `sheet-${String(args.cell.sheetIndex + 1)}`,
    workbookId: args.id,
  }
}

function prepareWorkbook(filePath: string): PreparedWorkbook {
  const bytes = readFileSync(filePath)
  const workbook = XLSX.read(bytes, {
    type: 'buffer',
    cellFormula: true,
    cellNF: true,
    cellText: false,
    cellDates: false,
    bookFiles: true,
    bookVBA: true,
  })
  const sheets: Record<string, WorkPaperSheet> = {}
  const formulaCells: FormulaCellRecord[] = []
  let maxRows = 1
  let maxColumns = 1

  workbook.SheetNames.forEach((sheetName, sheetIndex) => {
    const sheet = workbook.Sheets[sheetName]
    const range = sheet?.['!ref'] ? XLSX.utils.decode_range(sheet['!ref']) : undefined
    const rowCount = range ? range.e.r + 1 : 1
    const columnCount = range ? range.e.c + 1 : 1
    maxRows = Math.max(maxRows, rowCount)
    maxColumns = Math.max(maxColumns, columnCount)
    const rows: NormalizedCellContent[][] = Array.from({ length: rowCount }, () =>
      Array.from<NormalizedCellContent>({ length: columnCount }).fill(null),
    )

    if (sheet && range) {
      for (let row = range.s.r; row <= range.e.r; row += 1) {
        for (let col = range.s.c; col <= range.e.c; col += 1) {
          const address = XLSX.utils.encode_cell({ r: row, c: col })
          const cell = sheet[address]
          if (!cell) {
            continue
          }
          const formula = formulaText(cell)
          if (formula) {
            rows[row][col] = formula.startsWith('=') ? formula : `=${formula}`
            formulaCells.push({
              address,
              cachedValue: cachedFormulaValue(cell),
              col,
              formula,
              row,
              sheetIndex,
              sheetName,
              volatile: volatileFormulaPattern.test(formula),
            })
          } else {
            rows[row][col] = literalCellContent(cell)
          }
        }
      }
    }
    sheets[sheetName] = rows
  })

  return {
    formulaCells,
    maxColumns,
    maxRows,
    sheets: formulaCells.length === 0 ? sheets : attachRuntimeSnapshot(sheets, importXlsx(bytes, basename(filePath)).snapshot),
  }
}

type NormalizedCellContent = WorkPaperSheet[number][number]

function readFormulaCacheMap(filePath: string): Map<string, NormalizedFormulaValue> {
  const workbook = XLSX.read(readFileSync(filePath), {
    type: 'buffer',
    cellFormula: true,
    cellText: false,
    cellDates: false,
  })
  const values = new Map<string, NormalizedFormulaValue>()
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName]
    if (!sheet?.['!ref']) {
      continue
    }
    const range = XLSX.utils.decode_range(sheet['!ref'])
    for (let row = range.s.r; row <= range.e.r; row += 1) {
      for (let col = range.s.c; col <= range.e.c; col += 1) {
        const address = XLSX.utils.encode_cell({ r: row, c: col })
        const cell = sheet[address]
        if (cell && formulaText(cell)) {
          const cached = cachedFormulaValue(cell)
          if (cached) {
            values.set(formulaCellKey(sheetName, address), cached)
          }
        }
      }
    }
  }
  return values
}

function formulaText(cell: XLSX.CellObject): string | null {
  if (typeof cell.f !== 'string') {
    return null
  }
  const trimmed = cell.f.trim()
  return trimmed.length > 0 ? trimmed : null
}

function literalCellContent(cell: XLSX.CellObject): NormalizedCellContent {
  if (typeof cell.v === 'number' || typeof cell.v === 'string' || typeof cell.v === 'boolean') {
    return cell.v
  }
  return null
}

function cachedFormulaValue(cell: XLSX.CellObject): NormalizedFormulaValue | undefined {
  if (!Object.hasOwn(cell, 'v') || cell.v === undefined) {
    return undefined
  }
  if (cell.t === 'e') {
    return { kind: 'error', value: cachedErrorText(cell) }
  }
  if (cell.v === null) {
    return { kind: 'blank' }
  }
  if (typeof cell.v === 'number') {
    return Number.isFinite(cell.v) ? { kind: 'number', value: cell.v } : undefined
  }
  if (typeof cell.v === 'string') {
    return { kind: cell.t === 'e' ? 'error' : 'string', value: cell.v }
  }
  if (typeof cell.v === 'boolean') {
    return { kind: 'boolean', value: cell.v }
  }
  return undefined
}

function cachedErrorText(cell: XLSX.CellObject): string {
  if (typeof cell.w === 'string') {
    return cell.w
  }
  if (typeof cell.v === 'number') {
    return xlsxErrorTextByCode.get(cell.v) ?? String(cell.v)
  }
  return String(cell.v)
}

function normalizeCellValue(value: CellValue): NormalizedFormulaValue {
  switch (value.tag) {
    case ValueTag.Empty:
      return { kind: 'blank' }
    case ValueTag.Boolean:
      return { kind: 'boolean', value: value.value }
    case ValueTag.Error:
      return { kind: 'error', value: formatErrorCode(value.code) }
    case ValueTag.Number:
      return { kind: 'number', value: value.value }
    case ValueTag.String:
      return { kind: 'string', value: value.value }
  }
}

function formulaCellKey(sheetName: string, address: string): string {
  return `${sheetName}\u0000${address}`
}

function buildReport(
  mode: OracleHarnessReport['mode'],
  workbooks: readonly WorkbookEvaluation[],
  notes: readonly string[],
): OracleHarnessReport {
  const reportWithoutSummary = {
    schemaVersion: 1 as const,
    mode,
    generatedAt: new Date().toISOString(),
    notes,
    workbooks,
  }
  return {
    ...reportWithoutSummary,
    summary: buildReportSummary(reportWithoutSummary),
  }
}

function workPaperConfigFor(prepared: PreparedWorkbook, options: HarnessOptions): WorkPaperConfig {
  return {
    evaluationTimeoutMs: options.timeoutMs,
    maxColumns: prepared.maxColumns,
    maxRows: prepared.maxRows,
    useColumnIndex: true,
  }
}

function collectWorkbookFiles(inputPath: string): string[] {
  const resolved = resolve(inputPath)
  const files: string[] = []
  collectWorkbookFilesFromPath(resolved, files)
  return files.toSorted((left, right) => left.localeCompare(right))
}

function collectWorkbookFilesFromPath(path: string, files: string[]): void {
  if (!existsSync(path)) {
    throw new Error(`Workbook input path does not exist: ${path}`)
  }
  const stat = statSync(path)
  if (stat.isFile()) {
    if (!basename(path).startsWith('~$') && workbookExtensions.has(extname(path).toLowerCase())) {
      files.push(path)
    }
    return
  }
  if (!stat.isDirectory()) {
    return
  }
  for (const entry of readdirSync(path).toSorted((left, right) => left.localeCompare(right))) {
    if (!ignoredDirectoryNames.has(entry)) {
      collectWorkbookFilesFromPath(join(path, entry), files)
    }
  }
}

function workbookIdForFile(filePath: string): string {
  const bytes = readFileSync(filePath)
  const hash = createHash('sha256').update(bytes).digest('hex')
  return `workbook-${hash.slice(0, 16)}`
}

function sanitizedWorkbookName(id: string, filePath: string): string {
  const extension = extname(filePath).toLowerCase() || '.xlsx'
  return `${id}${extension}`
}

function formatSummaryMarkdown(report: OracleHarnessReport, cacheReport: OracleHarnessReport | null): string {
  const summary = report.summary
  const samples = report.workbooks
    .flatMap((workbook) => workbook.comparisons)
    .filter((comparison) => trueMismatchClassifications.has(comparison.classification))
    .slice(0, defaultSampleLimit)
  const lines = [
    '# WorkPaper Excel Oracle Summary',
    '',
    `Mode: ${report.mode}`,
    '',
    'Embedded XLSX cached values are diagnostic only. Bilig correctness bugs require fresh Excel expected value, Bilig actual value, formula text, and repro notes.',
    '',
    `- Total workbooks evaluated: ${String(summary.totalWorkbooksEvaluated)}`,
    `- Import/parser failures: ${String(summary.importParserFailures)}`,
    `- Timeout failures: ${String(summary.timeoutFailures)}`,
    `- Total formula cells: ${String(summary.totalFormulaCells)}`,
    `- Comparable formula cells: ${String(summary.comparableFormulaCells)}`,
    `- Bilig vs fresh Excel match rate: ${formatNullableRate(summary.biligVsFreshExcelMatchRate)}`,
    `- Embedded-cache freshness rate: ${formatNullableRate(summary.embeddedCacheFreshnessRate)}`,
    `- Stale-cache false positives: ${String(summary.staleCacheFalsePositives)}`,
    `- Real Bilig mismatches: ${String(summary.realBiligMismatches)}`,
  ]
  if (cacheReport && report.mode === 'excel-oracle') {
    lines.push(`- Cache-only diagnostic cells from cache report: ${String(cacheReport.summary.cacheOnlyDiagnosticCells)}`)
  }
  lines.push('', '## Top Formula Families For True Mismatches', '')
  if (summary.topMismatchFormulaFamilies.length === 0) {
    lines.push('None.')
  } else {
    lines.push(...summary.topMismatchFormulaFamilies.map((entry) => `- ${entry.family}: ${String(entry.count)}`))
  }
  lines.push('', '## Sanitized True-Mismatch Samples', '')
  if (samples.length === 0) {
    lines.push('None.')
  } else {
    for (const sample of samples) {
      lines.push(
        `- ${sample.workbookId} ${sample.sheet}!${sample.address}: ${sample.formula}`,
        `  - expected Excel: ${formatNormalizedValue(sample.expectedExcelValue)}`,
        `  - actual Bilig: ${formatNormalizedValue(sample.actualBiligValue)}`,
        `  - classification: ${sample.classification}`,
        `  - repro notes: ${sample.reproNotes}`,
      )
    }
  }
  return lines.join('\n')
}

function writeGithubIssueDrafts(report: OracleHarnessReport, outputDir: string): void {
  const mismatches = report.workbooks
    .flatMap((workbook) => workbook.comparisons)
    .filter((comparison) => trueMismatchClassifications.has(comparison.classification))
  if (mismatches.length === 0) {
    return
  }
  mkdirSync(outputDir, { recursive: true })
  mismatches.slice(0, defaultSampleLimit).forEach((sample, index) => {
    const body = [
      `# Formula mismatch against fresh Excel oracle ${String(index + 1)}`,
      '',
      `Workbook id: ${sample.workbookId}`,
      `Cell: ${sample.sheet}!${sample.address}`,
      `Formula: \`${sample.formula}\``,
      `Expected fresh Excel value: \`${formatNormalizedValue(sample.expectedExcelValue)}\``,
      `Actual Bilig value: \`${formatNormalizedValue(sample.actualBiligValue)}\``,
      `Classification: \`${sample.classification}\``,
      '',
      'Repro notes:',
      `- ${sample.reproNotes}`,
      '- Use the original workbook privately; do not attach private files to the issue.',
      '- Recalculate the workbook in Microsoft Excel and compare the recalculated copy against Bilig.',
      '- Embedded XLSX cached values alone are not an accuracy oracle.',
    ].join('\n')
    writeFile(join(outputDir, `formula-mismatch-${String(index + 1).padStart(2, '0')}.md`), `${body}\n`)
  })
}

function readOptionalReport(path: string): OracleHarnessReport | null {
  if (!existsSync(path)) {
    return null
  }
  const parsed: unknown = JSON.parse(readFileSync(path, 'utf8'))
  if (!isOracleHarnessReport(parsed)) {
    throw new Error(`Report JSON does not match the Excel oracle harness schema: ${path}`)
  }
  return parsed
}

function isOracleHarnessReport(value: unknown): value is OracleHarnessReport {
  return (
    isRecord(value) &&
    value['schemaVersion'] === 1 &&
    (value['mode'] === 'cache-diagnostic' || value['mode'] === 'excel-oracle') &&
    typeof value['generatedAt'] === 'string' &&
    Array.isArray(value['notes']) &&
    isRecord(value['summary']) &&
    Array.isArray(value['workbooks'])
  )
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function writeJson(path: string, value: unknown): void {
  writeFile(path, `${JSON.stringify(value, null, 2)}\n`)
}

function writeFile(path: string, value: string): void {
  mkdirSync(dirname(path), { recursive: true })
  writeFileSync(path, value)
}

function isExcelAutomationAvailable(): boolean {
  if (process.env['BILIG_EXCEL_ORACLE_DISABLE'] === '1') {
    return false
  }
  if (process.platform === 'darwin') {
    return spawnSync('osascript', ['-e', 'id of application "Microsoft Excel"'], { encoding: 'utf8' }).status === 0
  }
  if (process.platform === 'win32') {
    return (
      spawnSync(
        'powershell.exe',
        ['-NoProfile', '-Command', 'try { $x = New-Object -ComObject Excel.Application; $x.Quit(); exit 0 } catch { exit 1 }'],
        { encoding: 'utf8' },
      ).status === 0
    )
  }
  return false
}

function recalculateWorkbookWithExcel(inputPath: string, outputPath: string): void {
  mkdirSync(dirname(outputPath), { recursive: true })
  copyFileSync(inputPath, outputPath)
  if (process.platform === 'darwin') {
    const script = [
      'tell application "Microsoft Excel"',
      'set display alerts to false',
      `open POSIX file ${JSON.stringify(outputPath)}`,
      'set activeWorkbook to active workbook',
      'calculate full rebuild',
      'save activeWorkbook',
      'close activeWorkbook saving no',
      'end tell',
    ].join('\n')
    const result = spawnSync('osascript', ['-e', script], { encoding: 'utf8', timeout: 120_000 })
    if (result.status !== 0) {
      removeIfExists(outputPath)
      throw new Error(result.stderr.trim() || result.stdout.trim() || 'Excel AppleScript recalculation failed')
    }
    removeIfExists(join(dirname(outputPath), `~$${basename(outputPath)}`))
    return
  }
  if (process.platform === 'win32') {
    const command = [
      '$excel = New-Object -ComObject Excel.Application',
      '$excel.DisplayAlerts = $false',
      `$workbook = $excel.Workbooks.Open(${JSON.stringify(outputPath)})`,
      '$excel.CalculateFullRebuild()',
      '$workbook.Save()',
      '$workbook.Close($false)',
      '$excel.Quit()',
    ].join('; ')
    const result = spawnSync('powershell.exe', ['-NoProfile', '-Command', command], { encoding: 'utf8', timeout: 120_000 })
    if (result.status !== 0) {
      removeIfExists(outputPath)
      throw new Error(result.stderr.trim() || result.stdout.trim() || 'Excel COM recalculation failed')
    }
    removeIfExists(join(dirname(outputPath), `~$${basename(outputPath)}`))
    return
  }
  removeIfExists(outputPath)
  throw new Error('Microsoft Excel automation is unavailable on this platform')
}

function removeIfExists(path: string): void {
  if (existsSync(path)) {
    unlinkSync(path)
  }
}

function roundElapsed(value: number): number {
  return Number(value.toFixed(2))
}

function isTimeoutError(error: unknown): boolean {
  return error instanceof Error && error.name === 'WorkPaperEvaluationTimeoutError'
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}

function usageText(): string {
  return [
    'Usage:',
    '  bun scripts/workpaper-excel-oracle-harness.ts prepare-oracle <input-dir> <output-dir>',
    '  bun scripts/workpaper-excel-oracle-harness.ts evaluate-cache <input-dir> <output-dir> [--timeout-ms n] [--sample-limit n]',
    '  bun scripts/workpaper-excel-oracle-harness.ts evaluate-oracle <original-dir> <recalculated-dir> <output-dir> [--timeout-ms n] [--sample-limit n]',
    '  bun scripts/workpaper-excel-oracle-harness.ts summarize <report-dir>',
  ].join('\n')
}

function parseOptions(argv: readonly string[]): HarnessOptions {
  let sampleLimit = defaultSampleLimit
  let timeoutMs = defaultTimeoutMs
  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index]
    switch (arg) {
      case '--sample-limit':
        sampleLimit = parsePositiveInteger(requiredValue(argv, index, arg), arg)
        index += 1
        break
      case '--timeout-ms':
        timeoutMs = parsePositiveInteger(requiredValue(argv, index, arg), arg)
        index += 1
        break
      default:
        if (arg.startsWith('-')) {
          throw new Error(`Unknown option: ${arg}\n\n${usageText()}`)
        }
    }
  }
  return { sampleLimit, timeoutMs }
}

function requiredValue(argv: readonly string[], index: number, option: string): string {
  const value = argv[index + 1]
  if (!value || value.startsWith('-')) {
    throw new Error(`Missing value for ${option}`)
  }
  return value
}

function parsePositiveInteger(value: string, option: string): number {
  const parsed = Number.parseInt(value, 10)
  if (!Number.isFinite(parsed) || parsed <= 0 || String(parsed) !== value) {
    throw new Error(`${option} expects a positive integer, got ${value}`)
  }
  return parsed
}

function runCli(argv: readonly string[]): void {
  const [command, first, second, third, ...rest] = argv
  if (!command || command === '--help' || command === '-h') {
    process.stdout.write(`${usageText()}\n`)
    return
  }
  if (command === 'prepare-oracle' && first && second) {
    prepareExcelOracle(resolve(first), resolve(second))
    return
  }
  if (command === 'evaluate-cache' && first && second) {
    runCacheDiagnostic(resolve(first), resolve(second), parseOptions(rest))
    return
  }
  if (command === 'evaluate-oracle' && first && second && third) {
    runExcelOracleEvaluation(resolve(first), resolve(second), resolve(third), parseOptions(rest))
    return
  }
  if (command === 'summarize' && first) {
    writeSummary(resolve(first))
    return
  }
  throw new Error(`Invalid arguments.\n\n${usageText()}`)
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  try {
    runCli(process.argv.slice(2))
  } catch (error) {
    process.stderr.write(`${errorMessage(error)}\n`)
    process.exitCode = 1
  }
}
