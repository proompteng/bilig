#!/usr/bin/env bun

import { existsSync, readFileSync, readdirSync, statSync, writeFileSync } from 'node:fs'
import { spawnSync } from 'node:child_process'
import { basename, extname, join, resolve } from 'node:path'
import { performance } from 'node:perf_hooks'
import { fileURLToPath } from 'node:url'
import * as XLSX from 'xlsx'

import { WorkPaper, type WorkPaperConfig, type WorkPaperSheet, type WorkPaperSheets } from '@bilig/headless'
import { formatErrorCode, ValueTag, type CellValue } from '@bilig/protocol'

type CellContent = WorkPaperSheet[number][number]

export type CachedFormulaValue =
  | { kind: 'blank' }
  | { kind: 'boolean'; value: boolean }
  | { kind: 'error'; value: string }
  | { kind: 'number'; value: number }
  | { kind: 'string'; value: string }

export type WorkPaperXlsxFormulaSkipReason =
  | 'missing-cached-result'
  | 'unsupported-cached-result-type'
  | 'volatile-or-environment-dependent-formula'

export interface WorkPaperXlsxCorpusOptions {
  readonly childProcessTimeoutMs?: number
  readonly evaluationTimeoutMs?: number
  readonly mismatchSampleLimit?: number
}

export interface WorkPaperXlsxCorpusResult {
  readonly summary: WorkPaperXlsxCorpusSummary
  readonly files: readonly WorkPaperXlsxCorpusFileResult[]
  readonly mismatches: readonly WorkPaperXlsxCorpusMismatch[]
  readonly skippedByReason: Readonly<Record<WorkPaperXlsxFormulaSkipReason, number>>
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

interface FormulaCellRecord {
  readonly sheetName: string
  readonly address: string
  readonly row: number
  readonly col: number
  readonly formula: string
  readonly cachedValue?: CachedFormulaValue
  readonly skipReason?: WorkPaperXlsxFormulaSkipReason
}

interface PreparedWorkbook {
  readonly sheets: WorkPaperSheets
  readonly formulaCells: readonly FormulaCellRecord[]
  readonly maxRows: number
  readonly maxColumns: number
}

interface CliOptions extends WorkPaperXlsxCorpusOptions {
  readonly isolateFiles: boolean
  readonly paths: readonly string[]
  readonly jsonOut?: string
  readonly maxMismatches: number
  readonly minMatchRate: number
}

const defaultEvaluationTimeoutMs = 30_000
const childProcessTimeoutPaddingMs = 1_000
const defaultMismatchSampleLimit = 25
const ignoredDirectoryNames = new Set(['.git', 'build', 'dist', 'node_modules'])
const skipReasons: readonly WorkPaperXlsxFormulaSkipReason[] = [
  'missing-cached-result',
  'unsupported-cached-result-type',
  'volatile-or-environment-dependent-formula',
]
const xlsxExtensions = new Set(['.xls', '.xlsm', '.xlsx'])
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
const volatileOrEnvironmentFunctionPattern = /\b(CELL|INFO|NOW|RAND|RANDBETWEEN|TODAY)\s*\(/i

export function runWorkPaperXlsxCorpus(paths: readonly string[], options: WorkPaperXlsxCorpusOptions = {}): WorkPaperXlsxCorpusResult {
  const startedAt = performance.now()
  const files = collectXlsxFiles(paths)
  return runWorkPaperXlsxCorpusFiles(files, startedAt, options, (filePath, skippedByReason, mismatches, mismatchSampleLimit) =>
    checkWorkbookFile(filePath, options, skippedByReason, mismatches, mismatchSampleLimit),
  )
}

export function runWorkPaperXlsxCorpusInChildProcesses(
  paths: readonly string[],
  options: WorkPaperXlsxCorpusOptions = {},
): WorkPaperXlsxCorpusResult {
  const startedAt = performance.now()
  const files = collectXlsxFiles(paths)
  return runWorkPaperXlsxCorpusFiles(files, startedAt, options, (filePath, skippedByReason, mismatches, mismatchSampleLimit) =>
    checkWorkbookFileInChildProcess(filePath, options, skippedByReason, mismatches, mismatchSampleLimit),
  )
}

function runWorkPaperXlsxCorpusFiles(
  files: readonly string[],
  startedAt: number,
  options: WorkPaperXlsxCorpusOptions,
  checkFile: (
    filePath: string,
    skippedByReason: Record<WorkPaperXlsxFormulaSkipReason, number>,
    mismatches: WorkPaperXlsxCorpusMismatch[],
    mismatchSampleLimit: number,
  ) => WorkPaperXlsxCorpusFileResult,
): WorkPaperXlsxCorpusResult {
  const mismatchSampleLimit = options.mismatchSampleLimit ?? defaultMismatchSampleLimit
  const skippedByReason: Record<WorkPaperXlsxFormulaSkipReason, number> = {
    'missing-cached-result': 0,
    'unsupported-cached-result-type': 0,
    'volatile-or-environment-dependent-formula': 0,
  }

  const fileResults: WorkPaperXlsxCorpusFileResult[] = []
  const mismatches: WorkPaperXlsxCorpusMismatch[] = []
  for (const filePath of files) {
    const result = checkFile(filePath, skippedByReason, mismatches, mismatchSampleLimit)
    fileResults.push(result)
  }

  const summaryBase = fileResults.reduce(
    (summary, result) => ({
      failedErrors: summary.failedErrors + (result.status === 'error' ? 1 : 0),
      failedTimeouts: summary.failedTimeouts + (result.status === 'timeout' ? 1 : 0),
      formulaCells: summary.formulaCells + result.formulaCells,
      comparableFormulaCells: summary.comparableFormulaCells + result.comparableFormulaCells,
      matchingFormulaCells: summary.matchingFormulaCells + result.matchingFormulaCells,
      mismatchedFormulaCells: summary.mismatchedFormulaCells + result.mismatchedFormulaCells,
      ok: summary.ok + (result.status === 'ok' ? 1 : 0),
      skippedFormulaCells: summary.skippedFormulaCells + result.skippedFormulaCells,
    }),
    {
      failedErrors: 0,
      failedTimeouts: 0,
      formulaCells: 0,
      comparableFormulaCells: 0,
      matchingFormulaCells: 0,
      mismatchedFormulaCells: 0,
      ok: 0,
      skippedFormulaCells: 0,
    },
  )

  return {
    summary: {
      totalFiles: files.length,
      filesProcessed: fileResults.length,
      ...summaryBase,
      matchRate: ratio(summaryBase.matchingFormulaCells, summaryBase.comparableFormulaCells),
      elapsedMs: roundElapsed(performance.now() - startedAt),
    },
    files: fileResults,
    mismatches,
    skippedByReason,
  }
}

function checkWorkbookFileInChildProcess(
  filePath: string,
  options: WorkPaperXlsxCorpusOptions,
  skippedByReason: Record<WorkPaperXlsxFormulaSkipReason, number>,
  mismatches: WorkPaperXlsxCorpusMismatch[],
  mismatchSampleLimit: number,
): WorkPaperXlsxCorpusFileResult {
  const childTimeoutMs =
    options.childProcessTimeoutMs ?? (options.evaluationTimeoutMs ?? defaultEvaluationTimeoutMs) + childProcessTimeoutPaddingMs
  const child = spawnSync(
    'bun',
    [
      fileURLToPath(import.meta.url),
      '--internal-check-file-json',
      filePath,
      '--timeout-ms',
      String(options.evaluationTimeoutMs ?? defaultEvaluationTimeoutMs),
      '--mismatch-sample-limit',
      String(mismatchSampleLimit),
    ],
    {
      cwd: process.cwd(),
      encoding: 'utf8',
      timeout: childTimeoutMs,
    },
  )
  const fileName = basename(filePath)

  if (child.error) {
    const status = child.error.message.includes('ETIMEDOUT') ? 'timeout' : 'error'
    return {
      path: filePath,
      fileName,
      status,
      ...emptyCounts(0),
      matchRate: 1,
      elapsedMs: childTimeoutMs,
      error: child.error.message,
    }
  }

  if (child.signal || child.status !== 0) {
    return {
      path: filePath,
      fileName,
      status: child.signal === 'SIGTERM' || child.signal === 'SIGKILL' ? 'timeout' : 'error',
      ...emptyCounts(0),
      matchRate: 1,
      elapsedMs: childTimeoutMs,
      error: child.stderr.trim() || child.stdout.trim() || `child process exited with ${String(child.status)}`,
    }
  }

  let result: WorkPaperXlsxCorpusResult
  try {
    result = parseChildCorpusResult(child.stdout, filePath)
  } catch (error) {
    return {
      path: filePath,
      fileName,
      status: 'error',
      ...emptyCounts(0),
      matchRate: 1,
      elapsedMs: childTimeoutMs,
      error: errorMessage(error),
    }
  }
  for (const reason of skipReasons) {
    skippedByReason[reason] += result.skippedByReason[reason]
  }
  for (const mismatch of result.mismatches) {
    addMismatch(mismatches, mismatchSampleLimit, mismatch)
  }
  return (
    result.files[0] ?? {
      path: filePath,
      fileName,
      status: 'error',
      ...emptyCounts(0),
      matchRate: 1,
      elapsedMs: 0,
      error: 'child process did not return a file result',
    }
  )
}

function parseChildCorpusResult(stdout: string, filePath: string): WorkPaperXlsxCorpusResult {
  try {
    const parsed: unknown = JSON.parse(stdout)
    if (!isWorkPaperXlsxCorpusResult(parsed)) {
      throw new Error('child output did not match the XLSX corpus result shape')
    }
    return parsed
  } catch (error) {
    throw new Error(`Failed to parse isolated XLSX corpus result for ${filePath}: ${errorMessage(error)}`, { cause: error })
  }
}

function isWorkPaperXlsxCorpusResult(value: unknown): value is WorkPaperXlsxCorpusResult {
  if (!isRecord(value)) {
    return false
  }
  return (
    isRecord(value['summary']) &&
    Array.isArray(value['files']) &&
    Array.isArray(value['mismatches']) &&
    isSkippedByReason(value['skippedByReason'])
  )
}

function isSkippedByReason(value: unknown): value is Readonly<Record<WorkPaperXlsxFormulaSkipReason, number>> {
  if (!isRecord(value)) {
    return false
  }
  return skipReasons.every((reason) => typeof value[reason] === 'number')
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function checkWorkbookFile(
  filePath: string,
  options: WorkPaperXlsxCorpusOptions,
  skippedByReason: Record<WorkPaperXlsxFormulaSkipReason, number>,
  mismatches: WorkPaperXlsxCorpusMismatch[],
  mismatchSampleLimit: number,
): WorkPaperXlsxCorpusFileResult {
  const startedAt = performance.now()
  const fileName = basename(filePath)
  let prepared: PreparedWorkbook | undefined
  try {
    prepared = prepareWorkbook(filePath, skippedByReason)
    const workbook = WorkPaper.buildFromSheets(prepared.sheets, workPaperConfigFor(prepared, options))
    try {
      const counts = compareWorkbookFormulaCells(workbook, filePath, fileName, prepared.formulaCells, mismatches, mismatchSampleLimit)
      return {
        path: filePath,
        fileName,
        status: counts.mismatchedFormulaCells === 0 ? 'ok' : 'mismatched',
        ...counts,
        matchRate: ratio(counts.matchingFormulaCells, counts.comparableFormulaCells),
        elapsedMs: roundElapsed(performance.now() - startedAt),
      }
    } finally {
      workbook.dispose()
    }
  } catch (error) {
    const status = isTimeoutError(error) ? 'timeout' : 'error'
    const counts = emptyCounts(prepared?.formulaCells.length ?? 0)
    return {
      path: filePath,
      fileName,
      status,
      ...counts,
      matchRate: ratio(counts.matchingFormulaCells, counts.comparableFormulaCells),
      elapsedMs: roundElapsed(performance.now() - startedAt),
      error: errorMessage(error),
    }
  }
}

function prepareWorkbook(filePath: string, skippedByReason: Record<WorkPaperXlsxFormulaSkipReason, number>): PreparedWorkbook {
  const workbook = XLSX.read(readFileSync(filePath), {
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

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName]
    const range = sheet?.['!ref'] ? XLSX.utils.decode_range(sheet['!ref']) : undefined
    const rowCount = range ? range.e.r + 1 : 1
    const columnCount = range ? range.e.c + 1 : 1
    maxRows = Math.max(maxRows, rowCount)
    maxColumns = Math.max(maxColumns, columnCount)
    const rows: CellContent[][] = Array.from({ length: rowCount }, () => Array.from<CellContent>({ length: columnCount }).fill(null))

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
            rows[row][col] = formulaInput(formula)
            const formulaRecord = formulaCellRecord(sheetName, address, row, col, formula, cell)
            formulaCells.push(formulaRecord)
            if (formulaRecord.skipReason) {
              skippedByReason[formulaRecord.skipReason] += 1
            }
            continue
          }
          rows[row][col] = literalCellContent(cell)
        }
      }
    }

    sheets[sheetName] = rows
  }

  return {
    sheets,
    formulaCells,
    maxRows,
    maxColumns,
  }
}

function formulaCellRecord(
  sheetName: string,
  address: string,
  row: number,
  col: number,
  formula: string,
  cell: XLSX.CellObject,
): FormulaCellRecord {
  if (volatileOrEnvironmentFunctionPattern.test(formula)) {
    return {
      sheetName,
      address,
      row,
      col,
      formula,
      skipReason: 'volatile-or-environment-dependent-formula',
    }
  }

  const cachedValue = cachedFormulaValue(cell)
  if (!cachedValue) {
    return {
      sheetName,
      address,
      row,
      col,
      formula,
      skipReason: cellHasCachedValue(cell) ? 'unsupported-cached-result-type' : 'missing-cached-result',
    }
  }

  return {
    sheetName,
    address,
    row,
    col,
    formula,
    cachedValue,
  }
}

function compareWorkbookFormulaCells(
  workbook: WorkPaper,
  filePath: string,
  fileName: string,
  formulaCells: readonly FormulaCellRecord[],
  mismatches: WorkPaperXlsxCorpusMismatch[],
  mismatchSampleLimit: number,
): {
  readonly formulaCells: number
  readonly comparableFormulaCells: number
  readonly matchingFormulaCells: number
  readonly mismatchedFormulaCells: number
  readonly skippedFormulaCells: number
} {
  let comparableFormulaCells = 0
  let matchingFormulaCells = 0
  let mismatchedFormulaCells = 0
  let skippedFormulaCells = 0

  for (const formulaCell of formulaCells) {
    if (formulaCell.skipReason || !formulaCell.cachedValue) {
      skippedFormulaCells += 1
      continue
    }

    comparableFormulaCells += 1
    const sheetId = workbook.getSheetId(formulaCell.sheetName)
    if (sheetId === undefined) {
      mismatchedFormulaCells += 1
      addMismatch(mismatches, mismatchSampleLimit, {
        path: filePath,
        fileName,
        sheetName: formulaCell.sheetName,
        address: formulaCell.address,
        formula: formulaCell.formula,
        expected: formulaCell.cachedValue,
        actual: { kind: 'error', value: 'missing-sheet' },
      })
      continue
    }

    const actual = normalizeCellValue(workbook.getCellValue({ sheet: sheetId, row: formulaCell.row, col: formulaCell.col }))
    if (cachedValuesEqual(formulaCell.cachedValue, actual)) {
      matchingFormulaCells += 1
      continue
    }

    mismatchedFormulaCells += 1
    addMismatch(mismatches, mismatchSampleLimit, {
      path: filePath,
      fileName,
      sheetName: formulaCell.sheetName,
      address: formulaCell.address,
      formula: formulaCell.formula,
      expected: formulaCell.cachedValue,
      actual,
    })
  }

  return {
    formulaCells: formulaCells.length,
    comparableFormulaCells,
    matchingFormulaCells,
    mismatchedFormulaCells,
    skippedFormulaCells,
  }
}

function workPaperConfigFor(prepared: PreparedWorkbook, options: WorkPaperXlsxCorpusOptions): WorkPaperConfig {
  return {
    evaluationTimeoutMs: options.evaluationTimeoutMs ?? defaultEvaluationTimeoutMs,
    maxColumns: prepared.maxColumns,
    maxRows: prepared.maxRows,
    useColumnIndex: true,
  }
}

function collectXlsxFiles(paths: readonly string[]): string[] {
  const files: string[] = []
  for (const inputPath of paths) {
    collectXlsxFilesFromPath(resolve(inputPath), files)
  }
  return files.toSorted((left, right) => left.localeCompare(right))
}

function collectXlsxFilesFromPath(path: string, files: string[]): void {
  if (!existsSync(path)) {
    throw new Error(`XLSX corpus path does not exist: ${path}`)
  }

  const stat = statSync(path)
  if (stat.isFile()) {
    if (xlsxExtensions.has(extname(path).toLowerCase())) {
      files.push(path)
    }
    return
  }

  if (!stat.isDirectory()) {
    return
  }

  for (const entry of readdirSync(path).toSorted((left, right) => left.localeCompare(right))) {
    if (ignoredDirectoryNames.has(entry)) {
      continue
    }
    collectXlsxFilesFromPath(join(path, entry), files)
  }
}

function formulaText(cell: XLSX.CellObject): string | null {
  if (typeof cell.f !== 'string') {
    return null
  }
  const trimmed = cell.f.trim()
  return trimmed.length > 0 ? trimmed : null
}

function formulaInput(formula: string): string {
  return formula.startsWith('=') ? formula : `=${formula}`
}

function literalCellContent(cell: XLSX.CellObject): CellContent {
  if (typeof cell.v === 'number' || typeof cell.v === 'string' || typeof cell.v === 'boolean') {
    return cell.v
  }
  return null
}

function cachedFormulaValue(cell: XLSX.CellObject): CachedFormulaValue | null {
  if (!cellHasCachedValue(cell)) {
    return null
  }
  if (cell.t === 'e') {
    return { kind: 'error', value: cachedErrorText(cell) }
  }
  if (cell.v === null) {
    return { kind: 'blank' }
  }
  if (typeof cell.v === 'number') {
    return Number.isFinite(cell.v) ? { kind: 'number', value: cell.v } : null
  }
  if (typeof cell.v === 'string') {
    return { kind: cell.t === 'e' ? 'error' : 'string', value: cell.v }
  }
  if (typeof cell.v === 'boolean') {
    return { kind: 'boolean', value: cell.v }
  }
  return null
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

function cellHasCachedValue(cell: XLSX.CellObject): boolean {
  return Object.hasOwn(cell, 'v') && cell.v !== undefined
}

function normalizeCellValue(value: CellValue): CachedFormulaValue {
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

function cachedValuesEqual(expected: CachedFormulaValue, actual: CachedFormulaValue): boolean {
  if (expected.kind !== actual.kind) {
    return false
  }
  switch (expected.kind) {
    case 'blank':
      return true
    case 'boolean':
    case 'error':
    case 'string':
      return expected.value === actual.value
    case 'number':
      return actual.kind === 'number' && numbersEqual(expected.value, actual.value)
  }
}

function numbersEqual(expected: number, actual: number): boolean {
  if (Object.is(expected, actual)) {
    return true
  }
  const absoluteDifference = Math.abs(expected - actual)
  const scale = Math.max(1, Math.abs(expected), Math.abs(actual))
  return absoluteDifference <= scale * 1e-9
}

function addMismatch(mismatches: WorkPaperXlsxCorpusMismatch[], mismatchSampleLimit: number, mismatch: WorkPaperXlsxCorpusMismatch): void {
  if (mismatches.length < mismatchSampleLimit) {
    mismatches.push(mismatch)
  }
}

function emptyCounts(formulaCells: number): {
  readonly formulaCells: number
  readonly comparableFormulaCells: number
  readonly matchingFormulaCells: number
  readonly mismatchedFormulaCells: number
  readonly skippedFormulaCells: number
} {
  return {
    formulaCells,
    comparableFormulaCells: 0,
    matchingFormulaCells: 0,
    mismatchedFormulaCells: 0,
    skippedFormulaCells: formulaCells,
  }
}

function ratio(numerator: number, denominator: number): number {
  if (denominator === 0) {
    return 1
  }
  return Number((numerator / denominator).toFixed(4))
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

function parseCliArgs(argv: readonly string[]): CliOptions {
  const paths: string[] = []
  let jsonOut: string | undefined
  let childProcessTimeoutMs: number | undefined
  let maxMismatches = 0
  let minMatchRate = 1
  let evaluationTimeoutMs: number | undefined
  let isolateFiles = true
  let mismatchSampleLimit: number | undefined

  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index]
    switch (arg) {
      case '--help':
      case '-h':
        throw new CliUsageError(usageText(), 0)
      case '--allow-mismatches':
        maxMismatches = Number.POSITIVE_INFINITY
        minMatchRate = 0
        break
      case '--child-timeout-ms':
        childProcessTimeoutMs = parseNonNegativeInteger(requiredArgValue(argv, index, arg), arg)
        index += 1
        break
      case '--json-out':
        jsonOut = requiredArgValue(argv, index, arg)
        index += 1
        break
      case '--max-mismatches':
        maxMismatches = parseNonNegativeInteger(requiredArgValue(argv, index, arg), arg)
        index += 1
        break
      case '--min-match-rate':
        minMatchRate = parseMatchRate(requiredArgValue(argv, index, arg), arg)
        index += 1
        break
      case '--mismatch-sample-limit':
        mismatchSampleLimit = parseNonNegativeInteger(requiredArgValue(argv, index, arg), arg)
        index += 1
        break
      case '--timeout-ms':
        evaluationTimeoutMs = parseNonNegativeInteger(requiredArgValue(argv, index, arg), arg)
        index += 1
        break
      case '--no-isolate':
        isolateFiles = false
        break
      default:
        if (arg.startsWith('-')) {
          throw new CliUsageError(`Unknown option: ${arg}\n\n${usageText()}`, 2)
        }
        paths.push(arg)
        break
    }
  }

  if (paths.length === 0) {
    throw new CliUsageError(`Missing XLSX file or directory path.\n\n${usageText()}`, 2)
  }

  return {
    childProcessTimeoutMs,
    isolateFiles,
    paths,
    jsonOut,
    maxMismatches,
    minMatchRate,
    evaluationTimeoutMs,
    mismatchSampleLimit,
  }
}

function requiredArgValue(argv: readonly string[], index: number, option: string): string {
  const value = argv[index + 1]
  if (value === undefined || value.startsWith('-')) {
    throw new CliUsageError(`Missing value for ${option}\n\n${usageText()}`, 2)
  }
  return value
}

function parseNonNegativeInteger(value: string, option: string): number {
  const parsed = Number.parseInt(value, 10)
  if (!Number.isFinite(parsed) || parsed < 0 || String(parsed) !== value) {
    throw new CliUsageError(`${option} expects a non-negative integer, got ${value}`, 2)
  }
  return parsed
}

function parseMatchRate(value: string, option: string): number {
  const parsed = Number.parseFloat(value)
  if (!Number.isFinite(parsed) || parsed < 0 || parsed > 1) {
    throw new CliUsageError(`${option} expects a number between 0 and 1, got ${value}`, 2)
  }
  return parsed
}

function usageText(): string {
  return [
    'Usage: bun scripts/check-workpaper-xlsx-corpus.ts [options] <xlsx-file-or-directory> [...]',
    '',
    'Options:',
    '  --timeout-ms <ms>              WorkPaper initial evaluation timeout per workbook. Default: 30000.',
    '  --child-timeout-ms <ms>        Child-process timeout per workbook. Default: timeout-ms + 1000.',
    '  --no-isolate                  Check files in the current process instead of one child process per workbook.',
    '  --max-mismatches <count>       Maximum comparable cached-result mismatches before failing. Default: 0.',
    '  --min-match-rate <ratio>       Minimum comparable cached-result match rate before failing. Default: 1.',
    '  --mismatch-sample-limit <n>    Number of mismatch samples to keep in JSON output. Default: 25.',
    '  --json-out <path>              Also write the JSON report to a file.',
    '  --allow-mismatches             Report mismatches without failing the process.',
  ].join('\n')
}

class CliUsageError extends Error {
  readonly exitCode: number

  constructor(message: string, exitCode: number) {
    super(message)
    this.name = 'CliUsageError'
    this.exitCode = exitCode
  }
}

function runCli(): void {
  try {
    if (process.argv.includes('--internal-check-file-json')) {
      runInternalFileCheckCli(process.argv.slice(2))
      return
    }
    const options = parseCliArgs(process.argv.slice(2))
    const result = options.isolateFiles
      ? runWorkPaperXlsxCorpusInChildProcesses(options.paths, options)
      : runWorkPaperXlsxCorpus(options.paths, options)
    const output = `${JSON.stringify(result, null, 2)}\n`
    if (options.jsonOut) {
      writeFileSync(resolve(options.jsonOut), output)
    }
    process.stdout.write(output)
    if (
      result.summary.failedErrors > 0 ||
      result.summary.failedTimeouts > 0 ||
      result.summary.mismatchedFormulaCells > options.maxMismatches ||
      result.summary.matchRate < options.minMatchRate
    ) {
      process.exitCode = 1
    }
  } catch (error) {
    if (error instanceof CliUsageError) {
      process.stderr.write(`${error.message}\n`)
      process.exitCode = error.exitCode
      return
    }
    throw error
  }
}

function runInternalFileCheckCli(argv: readonly string[]): void {
  const paths: string[] = []
  let evaluationTimeoutMs: number | undefined
  let mismatchSampleLimit: number | undefined

  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index]
    switch (arg) {
      case '--internal-check-file-json':
        paths.push(requiredArgValue(argv, index, arg))
        index += 1
        break
      case '--mismatch-sample-limit':
        mismatchSampleLimit = parseNonNegativeInteger(requiredArgValue(argv, index, arg), arg)
        index += 1
        break
      case '--timeout-ms':
        evaluationTimeoutMs = parseNonNegativeInteger(requiredArgValue(argv, index, arg), arg)
        index += 1
        break
      default:
        throw new CliUsageError(`Unknown internal option: ${arg}`, 2)
    }
  }

  if (paths.length !== 1) {
    throw new CliUsageError('Internal XLSX corpus check expects exactly one file path.', 2)
  }
  process.stdout.write(`${JSON.stringify(runWorkPaperXlsxCorpus(paths, { evaluationTimeoutMs, mismatchSampleLimit }))}\n`)
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  runCli()
}
