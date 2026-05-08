#!/usr/bin/env bun

import { execFileSync } from 'node:child_process'
import { existsSync, mkdtempSync, mkdirSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join, resolve } from 'node:path'
import { pathToFileURL } from 'node:url'
import { performance } from 'node:perf_hooks'

import { WorkPaper } from '@bilig/headless'
import { ValueTag } from '@bilig/protocol'
import * as XLSX from 'xlsx'
import { summarizeNumbers, type NumericSummary } from '../packages/benchmarks/src/stats.js'
import { buildOverlappingAggregateSheet, buildStructuralColumnSheet } from '../packages/benchmarks/src/workpaper-benchmark-fixtures.js'
import {
  arrayField,
  asObject,
  booleanField,
  literalField,
  numberField,
  objectField,
  readJsonObject,
  stringArrayField,
  stringField,
} from './json-scorecard-helpers.ts'
import { formatJsonForRepo } from './scorecard-format.ts'

export type MicrosoftExcelLiveStructuralOperation =
  | 'insert-rows'
  | 'delete-rows'
  | 'move-rows'
  | 'insert-columns'
  | 'delete-columns'
  | 'move-columns'

export interface MicrosoftExcelLiveStructuralCase {
  readonly id: string
  readonly operation: MicrosoftExcelLiveStructuralOperation
  readonly axis: 'rows' | 'columns'
  readonly rowCount: number
  readonly sampleCount: number
  readonly workpaperElapsedMs: NumericSummary
  readonly microsoftExcelElapsedMs: NumericSummary
  readonly workpaperToMicrosoftExcelMeanRatio: number
  readonly workpaperToMicrosoftExcelP95Ratio: number
  readonly tenXMeanAndP95: boolean
  readonly verification: {
    readonly workpaper: Record<string, boolean | number | string | null>
    readonly microsoftExcel: Record<string, boolean | number | string | null>
    readonly equivalent: boolean
  }
  readonly passed: boolean
}

export interface MicrosoftExcelLiveStructuralScorecard {
  readonly schemaVersion: 1
  readonly suite: 'microsoft-excel-live-structural-performance'
  readonly generatedAt: string
  readonly host: {
    readonly arch: string
    readonly platform: string
  }
  readonly source: {
    readonly artifactGenerator: 'scripts/gen-microsoft-excel-live-structural-scorecard.ts'
    readonly implementationPackage: 'packages/headless'
    readonly evidenceKind: 'live-local-microsoft-excel-automation'
    readonly appleScriptTransport: 'osascript'
  }
  readonly benchmark: {
    readonly rowCount: number
    readonly sampleCount: number
    readonly screenUpdating: false
  }
  readonly microsoftExcel: {
    readonly appPath: '/Applications/Microsoft Excel.app'
    readonly version: string
  }
  readonly summary: {
    readonly allRequiredCasesPassed: boolean
    readonly requiredCaseCount: number
    readonly tenXMeanAndP95CaseCount: number
    readonly workpaperWins: number
    readonly coveredOperations: MicrosoftExcelLiveStructuralOperation[]
    readonly googleSheetsEvidence: 'not-covered-by-this-artifact'
  }
  readonly cases: MicrosoftExcelLiveStructuralCase[]
}

interface StructuralCaseSpec {
  readonly id: string
  readonly operation: MicrosoftExcelLiveStructuralOperation
  readonly axis: 'rows' | 'columns'
}

interface ExcelSampleResult {
  readonly elapsedMs: number
  readonly excelVersion: string
  readonly verification: Record<string, boolean | number | string | null>
}

interface StructuralCaseRun {
  readonly caseResult: MicrosoftExcelLiveStructuralCase
  readonly excelVersion: string
}

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const outputPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'microsoft-excel-live-structural-scorecard.json')
const excelAppPath = '/Applications/Microsoft Excel.app' as const
const worksheetName = 'Bench'
const rowCount = 500
const sampleCount = 5
const TEN_X_RATIO = 0.1
const caseSpecs = [
  { id: 'excel-live-structural-insert-rows', operation: 'insert-rows', axis: 'rows' },
  { id: 'excel-live-structural-delete-rows', operation: 'delete-rows', axis: 'rows' },
  { id: 'excel-live-structural-move-rows', operation: 'move-rows', axis: 'rows' },
  { id: 'excel-live-structural-insert-columns', operation: 'insert-columns', axis: 'columns' },
  { id: 'excel-live-structural-delete-columns', operation: 'delete-columns', axis: 'columns' },
  { id: 'excel-live-structural-move-columns', operation: 'move-columns', axis: 'columns' },
] as const satisfies readonly StructuralCaseSpec[]

function main(): void {
  const isCheckMode = process.argv.includes('--check')
  if (isCheckMode) {
    if (!existsSync(outputPath)) {
      throw new Error(
        `Microsoft Excel live structural scorecard is missing. Run: bun scripts/gen-microsoft-excel-live-structural-scorecard.ts`,
      )
    }
    const scorecard = parseMicrosoftExcelLiveStructuralScorecard(readJsonObject(outputPath))
    validateMicrosoftExcelLiveStructuralScorecard(scorecard)
    logResult('check', scorecard)
    return
  }

  const scorecard = buildMicrosoftExcelLiveStructuralScorecard(new Date().toISOString())
  validateMicrosoftExcelLiveStructuralScorecard(scorecard)
  mkdirSync(dirname(outputPath), { recursive: true })
  writeFileSync(outputPath, formatJsonForRepo(`${JSON.stringify(scorecard, null, 2)}\n`))
  logResult('write', scorecard)
}

export function buildMicrosoftExcelLiveStructuralScorecard(generatedAt: string): MicrosoftExcelLiveStructuralScorecard {
  if (!existsSync(excelAppPath)) {
    throw new Error(`Microsoft Excel app is not installed at ${excelAppPath}`)
  }

  const caseRuns = caseSpecs.map(runStructuralCase)
  const cases = caseRuns.map((entry) => entry.caseResult)
  const excelVersions = new Set(caseRuns.map((entry) => entry.excelVersion))
  if (excelVersions.size !== 1) {
    throw new Error(`Microsoft Excel version changed during structural benchmark: ${[...excelVersions].join(', ')}`)
  }
  const excelVersion = [...excelVersions][0] ?? ''
  const tenXMeanAndP95CaseCount = cases.filter((entry) => entry.tenXMeanAndP95).length
  const workpaperWins = cases.filter((entry) => entry.workpaperElapsedMs.mean <= entry.microsoftExcelElapsedMs.mean).length

  return {
    schemaVersion: 1,
    suite: 'microsoft-excel-live-structural-performance',
    generatedAt,
    host: {
      arch: process.arch,
      platform: process.platform,
    },
    source: {
      artifactGenerator: 'scripts/gen-microsoft-excel-live-structural-scorecard.ts',
      implementationPackage: 'packages/headless',
      evidenceKind: 'live-local-microsoft-excel-automation',
      appleScriptTransport: 'osascript',
    },
    benchmark: {
      rowCount,
      sampleCount,
      screenUpdating: false,
    },
    microsoftExcel: {
      appPath: excelAppPath,
      version: excelVersion,
    },
    summary: {
      allRequiredCasesPassed: cases.every((entry) => entry.passed),
      requiredCaseCount: cases.length,
      tenXMeanAndP95CaseCount,
      workpaperWins,
      coveredOperations: caseSpecs.map((entry) => entry.operation),
      googleSheetsEvidence: 'not-covered-by-this-artifact',
    },
    cases,
  }
}

export function parseMicrosoftExcelLiveStructuralScorecard(value: Record<string, unknown>): MicrosoftExcelLiveStructuralScorecard {
  const host = objectField(value, 'host')
  const source = objectField(value, 'source')
  const benchmark = objectField(value, 'benchmark')
  const microsoftExcel = objectField(value, 'microsoftExcel')
  const summary = objectField(value, 'summary')
  return {
    schemaVersion: literalField(value, 'schemaVersion', 1),
    suite: literalField(value, 'suite', 'microsoft-excel-live-structural-performance'),
    generatedAt: stringField(value, 'generatedAt'),
    host: {
      arch: stringField(host, 'arch'),
      platform: stringField(host, 'platform'),
    },
    source: {
      artifactGenerator: literalField(source, 'artifactGenerator', 'scripts/gen-microsoft-excel-live-structural-scorecard.ts'),
      implementationPackage: literalField(source, 'implementationPackage', 'packages/headless'),
      evidenceKind: literalField(source, 'evidenceKind', 'live-local-microsoft-excel-automation'),
      appleScriptTransport: literalField(source, 'appleScriptTransport', 'osascript'),
    },
    benchmark: {
      rowCount: numberField(benchmark, 'rowCount'),
      sampleCount: numberField(benchmark, 'sampleCount'),
      screenUpdating: literalField(benchmark, 'screenUpdating', false),
    },
    microsoftExcel: {
      appPath: literalField(microsoftExcel, 'appPath', excelAppPath),
      version: stringField(microsoftExcel, 'version'),
    },
    summary: {
      allRequiredCasesPassed: booleanField(summary, 'allRequiredCasesPassed'),
      requiredCaseCount: numberField(summary, 'requiredCaseCount'),
      tenXMeanAndP95CaseCount: numberField(summary, 'tenXMeanAndP95CaseCount'),
      workpaperWins: numberField(summary, 'workpaperWins'),
      coveredOperations: stringArrayField(summary, 'coveredOperations').map(parseOperation),
      googleSheetsEvidence: literalField(summary, 'googleSheetsEvidence', 'not-covered-by-this-artifact'),
    },
    cases: arrayField(value, 'cases').map(parseStructuralCase),
  }
}

export function validateMicrosoftExcelLiveStructuralScorecard(scorecard: MicrosoftExcelLiveStructuralScorecard): void {
  const expectedIds = caseSpecs.map((entry) => entry.id)
  const expectedOperations = caseSpecs.map((entry) => entry.operation)
  if (scorecard.microsoftExcel.version.trim().length === 0) {
    throw new Error('Microsoft Excel live structural scorecard must record an Excel version')
  }
  if (scorecard.benchmark.rowCount !== rowCount || scorecard.benchmark.sampleCount !== sampleCount) {
    throw new Error('Microsoft Excel live structural scorecard benchmark settings are stale')
  }
  if (
    scorecard.summary.requiredCaseCount !== expectedIds.length ||
    JSON.stringify(scorecard.cases.map((entry) => entry.id)) !== JSON.stringify(expectedIds)
  ) {
    throw new Error('Microsoft Excel live structural scorecard required cases are stale')
  }
  if (JSON.stringify(scorecard.summary.coveredOperations) !== JSON.stringify(expectedOperations)) {
    throw new Error('Microsoft Excel live structural scorecard covered operations are stale')
  }
  if (
    scorecard.summary.workpaperWins !==
    scorecard.cases.filter((entry) => entry.workpaperElapsedMs.mean <= entry.microsoftExcelElapsedMs.mean).length
  ) {
    throw new Error('Microsoft Excel live structural scorecard workpaper win count is inconsistent')
  }
  if (scorecard.summary.tenXMeanAndP95CaseCount !== scorecard.cases.filter((entry) => entry.tenXMeanAndP95).length) {
    throw new Error('Microsoft Excel live structural scorecard 10x count is inconsistent')
  }
  const failingCases = scorecard.cases.filter((entry) => !entry.passed)
  if (!scorecard.summary.allRequiredCasesPassed || failingCases.length > 0) {
    throw new Error(
      `Microsoft Excel live structural scorecard has failing required cases: ${failingCases
        .map(
          (entry) =>
            `${entry.id} WorkPaper=${JSON.stringify(entry.verification.workpaper)} Excel=${JSON.stringify(entry.verification.microsoftExcel)}`,
        )
        .join(', ')}`,
    )
  }
  for (const entry of scorecard.cases) {
    validateNumericSummary(entry.workpaperElapsedMs, `${entry.id} WorkPaper elapsedMs`)
    validateNumericSummary(entry.microsoftExcelElapsedMs, `${entry.id} Microsoft Excel elapsedMs`)
    if (!entry.verification.equivalent) {
      throw new Error(`Microsoft Excel live structural scorecard verification mismatch: ${entry.id}`)
    }
  }
}

function runStructuralCase(caseSpec: StructuralCaseSpec): StructuralCaseRun {
  const workpaperSamples: number[] = []
  const excelSamples: number[] = []
  const workpaperVerifications: Array<Record<string, boolean | number | string | null>> = []
  const excelVerifications: Array<Record<string, boolean | number | string | null>> = []
  const excelVersions = new Set<string>()

  for (let index = 0; index < sampleCount; index += 1) {
    const workpaperSample = runWorkPaperSample(caseSpec.operation)
    const excelSample = runExcelSample(caseSpec.operation)
    workpaperSamples.push(workpaperSample.elapsedMs)
    excelSamples.push(excelSample.elapsedMs)
    workpaperVerifications.push(workpaperSample.verification)
    excelVerifications.push(excelSample.verification)
    excelVersions.add(excelSample.excelVersion)
  }

  const workpaperElapsedMs = summarizeNumbers(workpaperSamples)
  const microsoftExcelElapsedMs = summarizeNumbers(excelSamples)
  const workpaperToMicrosoftExcelMeanRatio = workpaperElapsedMs.mean / microsoftExcelElapsedMs.mean
  const workpaperToMicrosoftExcelP95Ratio = workpaperElapsedMs.p95 / microsoftExcelElapsedMs.p95
  const verification = {
    workpaper: workpaperVerifications[0] ?? {},
    microsoftExcel: excelVerifications[0] ?? {},
    equivalent: verificationsEquivalent(workpaperVerifications, excelVerifications),
  }

  if (excelVersions.size !== 1) {
    throw new Error(`Microsoft Excel version changed during structural benchmark: ${[...excelVersions].join(', ')}`)
  }

  return {
    excelVersion: [...excelVersions][0] ?? '',
    caseResult: {
      id: caseSpec.id,
      operation: caseSpec.operation,
      axis: caseSpec.axis,
      rowCount,
      sampleCount,
      workpaperElapsedMs,
      microsoftExcelElapsedMs,
      workpaperToMicrosoftExcelMeanRatio,
      workpaperToMicrosoftExcelP95Ratio,
      tenXMeanAndP95: workpaperToMicrosoftExcelMeanRatio <= TEN_X_RATIO && workpaperToMicrosoftExcelP95Ratio <= TEN_X_RATIO,
      verification,
      passed:
        verification.equivalent &&
        workpaperElapsedMs.samples.length === sampleCount &&
        microsoftExcelElapsedMs.samples.length === sampleCount,
    },
  }
}

function runWorkPaperSample(operation: MicrosoftExcelLiveStructuralOperation): {
  elapsedMs: number
  verification: Record<string, boolean | number | string | null>
} {
  const workbook = WorkPaper.buildFromSheets({
    [worksheetName]: operation.includes('columns') ? buildStructuralColumnSheet(rowCount) : buildOverlappingAggregateSheet(rowCount),
  })
  const sheetId = workbook.getSheetId(worksheetName)
  if (sheetId === undefined) {
    workbook.dispose()
    throw new Error(`Missing WorkPaper structural worksheet: ${worksheetName}`)
  }

  const startedAt = performance.now()
  switch (operation) {
    case 'insert-rows':
      workbook.addRows(sheetId, Math.floor(rowCount / 2), 1)
      break
    case 'delete-rows':
      workbook.removeRows(sheetId, Math.floor(rowCount / 2), 1)
      break
    case 'move-rows':
      workbook.moveRows(sheetId, Math.floor(rowCount / 2), 1, 0)
      break
    case 'insert-columns':
      workbook.addColumns(sheetId, 1, 1)
      break
    case 'delete-columns':
      workbook.removeColumns(sheetId, 1, 1)
      break
    case 'move-columns':
      workbook.moveColumns(sheetId, 1, 1, 0)
      break
  }
  const elapsedMs = performance.now() - startedAt
  const verification = workpaperVerification(workbook, sheetId, operation)
  workbook.dispose()
  return { elapsedMs, verification }
}

function runExcelSample(operation: MicrosoftExcelLiveStructuralOperation): ExcelSampleResult {
  const tempDir = mkdtempSync(join(tmpdir(), 'bilig-excel-live-structural-'))
  const workbookPath = join(tempDir, 'structural.xlsx')
  const scriptPath = join(tempDir, 'run-structural.scpt')
  try {
    writeFileSync(workbookPath, createExcelWorkbookBytes(operation))
    writeFileSync(scriptPath, createStructuralAppleScript(operation))
    return parseExcelSampleOutput(execFileSync('osascript', [scriptPath, workbookPath], { encoding: 'utf8' }).trim())
  } finally {
    rmSync(tempDir, { recursive: true, force: true })
  }
}

function createExcelWorkbookBytes(operation: MicrosoftExcelLiveStructuralOperation): Uint8Array {
  const worksheet = XLSX.utils.aoa_to_sheet(
    operation.includes('columns') ? buildStructuralColumnSheet(rowCount) : buildOverlappingAggregateSheet(rowCount),
  )
  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(workbook, worksheet, worksheetName)
  return XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' })
}

function createStructuralAppleScript(operation: MicrosoftExcelLiveStructuralOperation): string {
  return `use framework "Foundation"

on run argv
  set workbookPath to POSIX file (item 1 of argv)
  set output to ""
  tell application "Microsoft Excel"
    set display alerts to false
    set screen updating to false
    open workbook workbook file name workbookPath
    calculate full rebuild
    set startedAt to current application's NSDate's timeIntervalSinceReferenceDate()
${excelOperationStatement(operation)}
    calculate full rebuild
    set elapsedMs to ((current application's NSDate's timeIntervalSinceReferenceDate()) - startedAt) * 1000
    set output to "version=" & (version as string)
    set output to output & linefeed & "elapsedMs=" & (elapsedMs as string)
${excelVerificationStatements(operation)}
    close active workbook saving no
    set screen updating to true
  end tell
  return output
end run
`
}

function excelOperationStatement(operation: MicrosoftExcelLiveStructuralOperation): string {
  switch (operation) {
    case 'insert-rows':
      return `    insert into range row ${String(Math.floor(rowCount / 2) + 1)} of worksheet "${worksheetName}" of active workbook`
    case 'delete-rows':
      return `    delete row ${String(Math.floor(rowCount / 2) + 1)} of worksheet "${worksheetName}" of active workbook`
    case 'move-rows':
      return `    cut row ${String(Math.floor(rowCount / 2) + 1)} of worksheet "${worksheetName}" of active workbook
    insert into range row 1 of worksheet "${worksheetName}" of active workbook`
    case 'insert-columns':
      return `    insert into range column 2 of worksheet "${worksheetName}" of active workbook`
    case 'delete-columns':
      return `    delete column 2 of worksheet "${worksheetName}" of active workbook`
    case 'move-columns':
      return `    cut column 2 of worksheet "${worksheetName}" of active workbook
    insert into range column 1 of worksheet "${worksheetName}" of active workbook`
  }
}

function excelVerificationStatements(operation: MicrosoftExcelLiveStructuralOperation): string {
  const valueCell = excelVerificationValueCell(operation)
  return `    set output to output & linefeed & "height=" & ((count of rows of used range of worksheet "${worksheetName}" of active workbook) as string)
    set output to output & linefeed & "width=" & ((count of columns of used range of worksheet "${worksheetName}" of active workbook) as string)
    set output to output & linefeed & "value=" & ((value of range "${valueCell}" of worksheet "${worksheetName}" of active workbook) as string)`
}

function excelVerificationValueCell(operation: MicrosoftExcelLiveStructuralOperation): string {
  switch (operation) {
    case 'insert-rows':
      return `A${String(rowCount + 1)}`
    case 'delete-rows':
      return `A${String(rowCount - 1)}`
    case 'move-rows':
      return 'A1'
    case 'insert-columns':
      return `A${String(rowCount)}`
    case 'delete-columns':
      return `A${String(rowCount)}`
    case 'move-columns':
      return `A${String(rowCount)}`
  }
}

function parseExcelSampleOutput(rawOutput: string): ExcelSampleResult {
  const values = new Map<string, string>()
  for (const line of rawOutput.split(/\r?\n/u)) {
    const separatorIndex = line.indexOf('=')
    if (separatorIndex > 0) {
      values.set(line.slice(0, separatorIndex), line.slice(separatorIndex + 1))
    }
  }
  return {
    excelVersion: requiredString(values, 'version'),
    elapsedMs: Number(requiredString(values, 'elapsedMs')),
    verification: {
      height: Number(requiredString(values, 'height')),
      width: Number(requiredString(values, 'width')),
      value: Number(requiredString(values, 'value')),
    },
  }
}

function workpaperVerification(
  workbook: WorkPaper,
  sheetId: number,
  operation: MicrosoftExcelLiveStructuralOperation,
): Record<string, boolean | number | string | null> {
  const dimensions = workbook.getSheetDimensions(sheetId)
  const cell = workpaperVerificationCell(operation)
  return {
    height: dimensions.height,
    width: dimensions.width,
    value: normalizeProtocolValue(workbook.getCellValue({ sheet: sheetId, row: cell.row, col: cell.col })),
  }
}

function workpaperVerificationCell(operation: MicrosoftExcelLiveStructuralOperation): { row: number; col: number } {
  switch (operation) {
    case 'insert-rows':
      return { row: rowCount, col: 0 }
    case 'delete-rows':
      return { row: rowCount - 2, col: 0 }
    case 'move-rows':
      return { row: 0, col: 0 }
    case 'insert-columns':
      return { row: rowCount - 1, col: 0 }
    case 'delete-columns':
      return { row: rowCount - 1, col: 0 }
    case 'move-columns':
      return { row: rowCount - 1, col: 0 }
  }
}

function normalizeProtocolValue(value: unknown): boolean | number | string | null {
  if (!isProtocolValueLike(value)) {
    throw new Error('Unexpected non-protocol WorkPaper value in structural benchmark')
  }
  switch (value.tag) {
    case ValueTag.Empty:
      return null
    case ValueTag.Number:
    case ValueTag.Boolean:
    case ValueTag.String:
      return value.value ?? null
    case ValueTag.Error:
      return null
    default:
      return null
  }
}

function parseStructuralCase(value: unknown): MicrosoftExcelLiveStructuralCase {
  const record = asObject(value, 'Microsoft Excel live structural case')
  const verification = objectField(record, 'verification')
  return {
    id: stringField(record, 'id'),
    operation: parseOperation(stringField(record, 'operation')),
    axis: parseAxis(stringField(record, 'axis')),
    rowCount: numberField(record, 'rowCount'),
    sampleCount: numberField(record, 'sampleCount'),
    workpaperElapsedMs: parseNumericSummary(objectField(record, 'workpaperElapsedMs')),
    microsoftExcelElapsedMs: parseNumericSummary(objectField(record, 'microsoftExcelElapsedMs')),
    workpaperToMicrosoftExcelMeanRatio: numberField(record, 'workpaperToMicrosoftExcelMeanRatio'),
    workpaperToMicrosoftExcelP95Ratio: numberField(record, 'workpaperToMicrosoftExcelP95Ratio'),
    tenXMeanAndP95: booleanField(record, 'tenXMeanAndP95'),
    verification: {
      workpaper: parseVerificationRecord(objectField(verification, 'workpaper')),
      microsoftExcel: parseVerificationRecord(objectField(verification, 'microsoftExcel')),
      equivalent: booleanField(verification, 'equivalent'),
    },
    passed: booleanField(record, 'passed'),
  }
}

function parseNumericSummary(value: Record<string, unknown>): NumericSummary {
  const confidence95 = objectField(value, 'confidence95')
  return {
    samples: arrayField(value, 'samples').map((entry) => {
      if (typeof entry !== 'number' || !Number.isFinite(entry)) {
        throw new Error('Expected numeric summary samples to be finite numbers')
      }
      return entry
    }),
    min: numberField(value, 'min'),
    median: numberField(value, 'median'),
    p95: numberField(value, 'p95'),
    max: numberField(value, 'max'),
    mean: numberField(value, 'mean'),
    standardDeviation: numberField(value, 'standardDeviation'),
    relativeStandardDeviation: numberField(value, 'relativeStandardDeviation'),
    standardError: numberField(value, 'standardError'),
    confidence95: {
      low: numberField(confidence95, 'low'),
      high: numberField(confidence95, 'high'),
    },
  }
}

function parseVerificationRecord(value: Record<string, unknown>): Record<string, boolean | number | string | null> {
  const result: Record<string, boolean | number | string | null> = {}
  for (const [key, entry] of Object.entries(value)) {
    if (entry === null || typeof entry === 'boolean' || typeof entry === 'number' || typeof entry === 'string') {
      result[key] = entry
      continue
    }
    throw new Error(`Unsupported verification value for ${key}`)
  }
  return result
}

function parseOperation(value: string): MicrosoftExcelLiveStructuralOperation {
  if (
    value === 'insert-rows' ||
    value === 'delete-rows' ||
    value === 'move-rows' ||
    value === 'insert-columns' ||
    value === 'delete-columns' ||
    value === 'move-columns'
  ) {
    return value
  }
  throw new Error(`Unexpected Microsoft Excel live structural operation: ${value}`)
}

function parseAxis(value: string): 'rows' | 'columns' {
  if (value === 'rows' || value === 'columns') {
    return value
  }
  throw new Error(`Unexpected Microsoft Excel live structural axis: ${value}`)
}

function verificationsEquivalent(
  workpaperVerifications: readonly Record<string, boolean | number | string | null>[],
  excelVerifications: readonly Record<string, boolean | number | string | null>[],
): boolean {
  if (workpaperVerifications.length !== excelVerifications.length) {
    return false
  }
  return workpaperVerifications.every((workpaperSnapshot, index) => {
    const excelVerification = excelVerifications[index]
    return excelVerification !== undefined && JSON.stringify(workpaperSnapshot) === JSON.stringify(excelVerification)
  })
}

function requiredString(values: ReadonlyMap<string, string>, key: string): string {
  const value = values.get(key)
  if (value === undefined || value.length === 0) {
    throw new Error(`Missing Microsoft Excel structural output field: ${key}`)
  }
  return value
}

function validateNumericSummary(summary: NumericSummary, label: string): void {
  if (summary.samples.length !== sampleCount || summary.samples.some((entry) => !Number.isFinite(entry) || entry < 0)) {
    throw new Error(`${label} has invalid samples`)
  }
}

function isProtocolValueLike(value: unknown): value is { tag: ValueTag; value?: boolean | number | string } {
  if (value === null || typeof value !== 'object') {
    return false
  }
  const tag = Reflect.get(value, 'tag')
  return tag === ValueTag.Empty || tag === ValueTag.Number || tag === ValueTag.Boolean || tag === ValueTag.String || tag === ValueTag.Error
}

function logResult(mode: 'check' | 'write', scorecard: MicrosoftExcelLiveStructuralScorecard): void {
  console.log(
    JSON.stringify(
      {
        mode,
        outputPath,
        excelVersion: scorecard.microsoftExcel.version,
        allRequiredCasesPassed: scorecard.summary.allRequiredCasesPassed,
        workpaperWins: scorecard.summary.workpaperWins,
        tenXMeanAndP95CaseCount: scorecard.summary.tenXMeanAndP95CaseCount,
      },
      null,
      2,
    ),
  )
}

if (process.argv[1] && pathToFileURL(resolve(process.argv[1])).href === import.meta.url) {
  main()
}
