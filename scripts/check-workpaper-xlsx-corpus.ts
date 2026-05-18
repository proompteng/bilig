#!/usr/bin/env bun

import { existsSync, readFileSync, readdirSync, statSync, writeFileSync } from 'node:fs'
import { spawnSync } from 'node:child_process'
import { basename, extname, join, resolve } from 'node:path'
import { performance } from 'node:perf_hooks'
import { fileURLToPath } from 'node:url'
import * as XLSX from 'xlsx'

import { attachRuntimeSnapshot } from '@bilig/core'
import { importXlsx } from '@bilig/excel-import'
import { WorkPaper, type WorkPaperConfig, type WorkPaperSheet } from '@bilig/headless'
import { formatErrorCode, ValueTag, type CellValue, type WorkbookFormulaAuditEntrySnapshot, type WorkbookSnapshot } from '@bilig/protocol'
import type {
  CachedFormulaValue,
  CellContent,
  FormulaCellRecord,
  PreparedWorkbook,
  WorkPaperXlsxCorpusFileResult,
  WorkPaperXlsxCorpusCompatibilitySummary,
  WorkPaperXlsxCorpusMismatch,
  WorkPaperXlsxCorpusOptions,
  WorkPaperXlsxCorpusResult,
  WorkPaperXlsxFormulaSkipReason,
} from './check-workpaper-xlsx-corpus-types.ts'
export type {
  CachedFormulaValue,
  WorkPaperXlsxCorpusFileResult,
  WorkPaperXlsxCorpusMismatch,
  WorkPaperXlsxCorpusOptions,
  WorkPaperXlsxCorpusResult,
  WorkPaperXlsxCorpusSummary,
  WorkPaperXlsxCorpusCompatibilitySummary,
  WorkPaperXlsxFormulaSkipReason,
} from './check-workpaper-xlsx-corpus-types.ts'
import { formatByteSize } from './public-workbook-corpus-process.ts'
import {
  assertBroadCorpusSweepNotStopped,
  assertUnisolatedCliDebuggerPath,
  CliUsageError,
  parseWorkPaperXlsxCorpusCliArgs,
  parseWorkPaperXlsxCorpusInternalCliArgs,
  type WorkPaperXlsxCorpusCliOptions,
  type WorkPaperXlsxCorpusInternalCliOptions,
} from './workpaper-xlsx-corpus-cli.ts'
import { markVolatileDependentFormulaCells } from './workpaper-xlsx-volatile-dependencies.ts'

const defaultEvaluationTimeoutMs = 30_000
const childProcessTimeoutPaddingMs = 1_000
const defaultMaxFileBytes = 50 * 1024 * 1024
const defaultMismatchSampleLimit = 25
const ignoredDirectoryNames = new Set(['.git', 'build', 'dist', 'node_modules'])
const skipReasons: readonly WorkPaperXlsxFormulaSkipReason[] = [
  'missing-cached-result',
  'stale-cached-result',
  'stale-cached-name-error',
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
const volatileOrEnvironmentFunctionPattern = /\b(CELL|FILTERXML|IMAGE|INFO|NOW|RAND|RANDBETWEEN|STOCKHISTORY|TODAY|WEBSERVICE)\s*\(/i

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
  const skippedByReason = emptySkippedByReason()

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
      compatibility: addCompatibilitySummaries(summary.compatibility, result.compatibility),
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
      compatibility: emptyCompatibilitySummary(),
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
  const oversizeResult = oversizedWorkbookResult(filePath, options)
  if (oversizeResult) {
    return oversizeResult
  }
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
      '--max-file-bytes',
      String(maxFileBytesFor(options)),
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
      compatibility: emptyCompatibilitySummary(),
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
      compatibility: emptyCompatibilitySummary(),
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
      compatibility: emptyCompatibilitySummary(),
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
      compatibility: emptyCompatibilitySummary(),
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
  const oversizeResult = oversizedWorkbookResult(filePath, options, startedAt)
  if (oversizeResult) {
    return oversizeResult
  }
  let prepared: PreparedWorkbook | undefined
  try {
    prepared = prepareWorkbook(filePath, skippedByReason)
    const workbook = WorkPaper.buildFromSheets(prepared.sheets, workPaperConfigFor(prepared, options))
    try {
      workbook.rebuildAndRecalculate()
      const counts = compareWorkbookFormulaCells(
        workbook,
        filePath,
        fileName,
        prepared.formulaCells,
        skippedByReason,
        mismatches,
        mismatchSampleLimit,
      )
      return {
        path: filePath,
        fileName,
        status: counts.mismatchedFormulaCells === 0 ? 'ok' : 'mismatched',
        ...counts,
        matchRate: ratio(counts.matchingFormulaCells, counts.comparableFormulaCells),
        compatibility: prepared.compatibility,
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
      compatibility: prepared?.compatibility ?? emptyCompatibilitySummary(),
      elapsedMs: roundElapsed(performance.now() - startedAt),
      error: errorMessage(error),
    }
  }
}

function emptySkippedByReason(): Record<WorkPaperXlsxFormulaSkipReason, number> {
  return {
    'missing-cached-result': 0,
    'stale-cached-result': 0,
    'stale-cached-name-error': 0,
    'unsupported-cached-result-type': 0,
    'volatile-or-environment-dependent-formula': 0,
  }
}

function prepareWorkbook(filePath: string, skippedByReason: Record<WorkPaperXlsxFormulaSkipReason, number>): PreparedWorkbook {
  const workbookBytes = readFileSync(filePath)
  const workbook = XLSX.read(workbookBytes, {
    type: 'buffer',
    cellFormula: true,
    cellNF: true,
    cellText: false,
    cellDates: false,
    bookFiles: true,
    bookVBA: true,
  })
  const sheets: Record<string, WorkPaperSheet> = {}
  let formulaCells: FormulaCellRecord[] = []
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

  const importedSnapshot =
    formulaCells.length > 0 || workbookHasFormulaMarkup(workbook) ? importXlsx(workbookBytes, basename(filePath)).snapshot : null
  if (importedSnapshot) {
    formulaCells = addFormulaAuditFallbackCells(sheets, formulaCells, importedSnapshot, skippedByReason)
  }
  formulaCells = markVolatileDependentFormulaCells(formulaCells, skippedByReason)
  for (const formulaCell of formulaCells) {
    maxRows = Math.max(maxRows, formulaCell.row + 1)
    maxColumns = Math.max(maxColumns, formulaCell.col + 1)
  }

  return {
    sheets: formulaCells.length === 0 || !importedSnapshot ? sheets : attachRuntimeSnapshot(sheets, importedSnapshot),
    formulaCells,
    compatibility: importedSnapshot ? compatibilitySummaryForSnapshot(importedSnapshot) : emptyCompatibilitySummary(),
    maxRows,
    maxColumns,
  }
}

function workbookHasFormulaMarkup(workbook: XLSX.WorkBook): boolean {
  if (!('files' in workbook)) {
    return false
  }
  const files = workbook['files']
  if (!isRecord(files)) {
    return false
  }
  for (const [path, file] of Object.entries(files)) {
    if (!path.startsWith('xl/worksheets/') || !path.endsWith('.xml') || !isRecord(file)) {
      continue
    }
    const content = file['content']
    const xml = typeof content === 'string' ? content : content instanceof Uint8Array ? new TextDecoder().decode(content) : null
    if (xml && /<f\b/u.test(xml)) {
      return true
    }
  }
  return false
}

function addFormulaAuditFallbackCells(
  sheets: Record<string, WorkPaperSheet>,
  formulaCells: FormulaCellRecord[],
  importedSnapshot: WorkbookSnapshot,
  skippedByReason: Record<WorkPaperXlsxFormulaSkipReason, number>,
): FormulaCellRecord[] {
  const existingFormulaCells = new Set(formulaCells.map((record) => formulaCellRecordKey(record)))
  for (const auditEntry of importedSnapshot.workbook.metadata?.formulaAudit?.formulas ?? []) {
    const record = formulaAuditFallbackCellRecord(auditEntry)
    if (!record || existingFormulaCells.has(formulaCellRecordKey(record))) {
      continue
    }
    ensureSheetCell(sheets, record.sheetName, record.row, record.col, formulaInput(record.formula))
    existingFormulaCells.add(formulaCellRecordKey(record))
    formulaCells.push(record)
    if (record.skipReason) {
      skippedByReason[record.skipReason] += 1
    }
  }
  return formulaCells
}

function formulaAuditFallbackCellRecord(auditEntry: WorkbookFormulaAuditEntrySnapshot): FormulaCellRecord | null {
  if (
    auditEntry.context !== 'worksheet-cell' ||
    typeof auditEntry.sheetName !== 'string' ||
    typeof auditEntry.address !== 'string' ||
    auditEntry.formula.trim().length === 0
  ) {
    return null
  }

  let decodedAddress: XLSX.CellAddress
  try {
    decodedAddress = XLSX.utils.decode_cell(auditEntry.address)
  } catch {
    return null
  }

  const baseRecord = {
    sheetName: auditEntry.sheetName,
    address: XLSX.utils.encode_cell(decodedAddress),
    row: decodedAddress.r,
    col: decodedAddress.c,
    formula: auditEntry.formula,
  }

  if (volatileOrEnvironmentFunctionPattern.test(auditEntry.formula)) {
    return {
      ...baseRecord,
      skipReason: 'volatile-or-environment-dependent-formula',
    }
  }
  if (auditEntry.cacheStatus === 'missing' || auditEntry.cachedValueRaw === '') {
    return {
      ...baseRecord,
      skipReason: 'missing-cached-result',
    }
  }
  if (auditEntry.cacheStatus !== 'trustedCached') {
    return {
      ...baseRecord,
      skipReason: 'stale-cached-result',
    }
  }

  const cachedValue = cachedFormulaValueFromAudit(auditEntry)
  return cachedValue
    ? {
        ...baseRecord,
        cachedValue,
      }
    : {
        ...baseRecord,
        skipReason: 'unsupported-cached-result-type',
      }
}

function ensureSheetCell(sheets: Record<string, WorkPaperSheet>, sheetName: string, row: number, col: number, value: CellContent): void {
  let sheet = sheets[sheetName]
  if (!sheet) {
    sheet = []
    sheets[sheetName] = sheet
  }
  while (sheet.length <= row) {
    sheet.push([])
  }
  const targetRow = sheet[row]
  while (targetRow.length <= col) {
    targetRow.push(null)
  }
  targetRow[col] = value
}

function formulaCellRecordKey(record: Pick<FormulaCellRecord, 'sheetName' | 'address'>): string {
  return `${record.sheetName}!${record.address}`
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
  skippedByReason: Record<WorkPaperXlsxFormulaSkipReason, number>,
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

    const sheetId = workbook.getSheetId(formulaCell.sheetName)
    if (sheetId === undefined) {
      comparableFormulaCells += 1
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
    if (isStaleCachedNameError(formulaCell.cachedValue, actual)) {
      skippedFormulaCells += 1
      skippedByReason['stale-cached-name-error'] += 1
      continue
    }

    comparableFormulaCells += 1
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

function oversizedWorkbookResult(
  filePath: string,
  options: WorkPaperXlsxCorpusOptions,
  startedAt = performance.now(),
): WorkPaperXlsxCorpusFileResult | null {
  const maxFileBytes = maxFileBytesFor(options)
  const fileSizeBytes = statSync(filePath).size
  if (fileSizeBytes <= maxFileBytes) {
    return null
  }
  return {
    path: filePath,
    fileName: basename(filePath),
    status: 'error',
    ...emptyCounts(0),
    matchRate: 1,
    compatibility: emptyCompatibilitySummary(),
    elapsedMs: roundElapsed(performance.now() - startedAt),
    error: `XLSX file exceeds max file size: ${formatByteSize(fileSizeBytes)} > ${formatByteSize(maxFileBytes)}`,
  }
}

function maxFileBytesFor(options: WorkPaperXlsxCorpusOptions): number {
  const maxFileBytes = options.maxFileBytes ?? defaultMaxFileBytes
  if (!Number.isFinite(maxFileBytes) || maxFileBytes <= 0) {
    throw new Error(`maxFileBytes must be a positive finite integer, got ${String(maxFileBytes)}`)
  }
  return Math.trunc(maxFileBytes)
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

function cachedFormulaValueFromAudit(auditEntry: WorkbookFormulaAuditEntrySnapshot): CachedFormulaValue | null {
  const value = auditEntry.cachedValue
  if (value === undefined) {
    return null
  }
  if (value === null) {
    return { kind: 'blank' }
  }
  if (typeof value === 'number') {
    return Number.isFinite(value) ? { kind: 'number', value } : null
  }
  if (typeof value === 'boolean') {
    return { kind: 'boolean', value }
  }
  if (typeof value === 'string') {
    return { kind: auditEntry.cellValueType === 'e' ? 'error' : 'string', value }
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

function isStaleCachedNameError(expected: CachedFormulaValue, actual: CachedFormulaValue): boolean {
  return expected.kind === 'error' && expected.value === '#NAME?' && actual.kind !== 'error'
}

function numbersEqual(expected: number, actual: number): boolean {
  if (Object.is(expected, actual)) {
    return true
  }
  const absoluteDifference = Math.abs(expected - actual)
  const scale = Math.max(1, Math.abs(expected), Math.abs(actual))
  return absoluteDifference <= Math.max(1e-7, scale * 1e-12)
}

function addMismatch(mismatches: WorkPaperXlsxCorpusMismatch[], mismatchSampleLimit: number, mismatch: WorkPaperXlsxCorpusMismatch): void {
  if (mismatches.length < mismatchSampleLimit) {
    mismatches.push(mismatch)
  }
}

function emptyCompatibilitySummary(): WorkPaperXlsxCorpusCompatibilitySummary {
  return {
    formulaContextDiagnosticCount: 0,
    staleCacheRiskFormulaCount: 0,
    externalCacheOnlyPivotCount: 0,
    unsupportedRefreshFeatureCount: 0,
  }
}

function addCompatibilitySummaries(
  left: WorkPaperXlsxCorpusCompatibilitySummary,
  right: WorkPaperXlsxCorpusCompatibilitySummary,
): WorkPaperXlsxCorpusCompatibilitySummary {
  return {
    formulaContextDiagnosticCount: left.formulaContextDiagnosticCount + right.formulaContextDiagnosticCount,
    staleCacheRiskFormulaCount: left.staleCacheRiskFormulaCount + right.staleCacheRiskFormulaCount,
    externalCacheOnlyPivotCount: left.externalCacheOnlyPivotCount + right.externalCacheOnlyPivotCount,
    unsupportedRefreshFeatureCount: left.unsupportedRefreshFeatureCount + right.unsupportedRefreshFeatureCount,
  }
}

function compatibilitySummaryForSnapshot(snapshot: WorkbookSnapshot): WorkPaperXlsxCorpusCompatibilitySummary {
  const metadata = snapshot.workbook.metadata
  const formulaAudit = metadata?.formulaAudit
  const externalConnections = metadata?.externalConnections
  const connections = externalConnections?.connections ?? []
  const externalLinks = externalConnections?.externalLinks ?? []
  const externalRefreshFeatureCount =
    connections.filter((connection) => connection.refreshOnLoad === true).length +
    externalLinks.length +
    (metadata?.externalWorkbookReferences?.length ?? 0)
  const externalPivotCount =
    (metadata?.pivots?.filter((pivot) => pivot.sourceKind === 'external-cache-only' || pivot.cacheOnly === true).length ?? 0) +
    (metadata?.unsupportedPivots?.filter((pivot) => pivot.kind === 'external-cache' || pivot.sourceType === 'external').length ?? 0)
  return {
    formulaContextDiagnosticCount: formulaAudit?.diagnostics?.length ?? 0,
    staleCacheRiskFormulaCount: formulaAudit?.formulas.filter((formula) => formula.cacheStatus === 'staleRisk').length ?? 0,
    externalCacheOnlyPivotCount: externalPivotCount,
    unsupportedRefreshFeatureCount: externalRefreshFeatureCount,
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

function runCli(): void {
  try {
    if (process.argv.includes('--internal-check-file-json')) {
      runInternalFileCheckCli(process.argv.slice(2))
      return
    }
    const options: WorkPaperXlsxCorpusCliOptions = parseWorkPaperXlsxCorpusCliArgs(process.argv.slice(2))
    if (!options.isolateFiles) {
      assertUnisolatedCliDebuggerPath(options.paths)
    }
    assertBroadCorpusSweepNotStopped(options.paths, options.stopMarkerPath)
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
  const options: WorkPaperXlsxCorpusInternalCliOptions = parseWorkPaperXlsxCorpusInternalCliArgs(argv)
  process.stdout.write(`${JSON.stringify(runWorkPaperXlsxCorpus(options.paths, options))}\n`)
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  runCli()
}
