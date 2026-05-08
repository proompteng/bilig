#!/usr/bin/env bun

import { existsSync, mkdirSync, writeFileSync } from 'node:fs'
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
  numberArrayField,
  numberField,
  objectField,
  readJsonObject,
  stringArrayField,
  stringField,
} from './json-scorecard-helpers.ts'
import { formatJsonForRepo } from './scorecard-format.ts'

export type GoogleSheetsLiveStructuralOperation =
  | 'insert-rows'
  | 'delete-rows'
  | 'move-rows'
  | 'insert-columns'
  | 'delete-columns'
  | 'move-columns'

type VerificationValue = boolean | number | string | null
type StructuralSheet = ReadonlyArray<ReadonlyArray<boolean | number | string | null>>

export interface GoogleSheetsLiveStructuralCapture {
  readonly schemaVersion: 1
  readonly generatedAt: string
  readonly capture: {
    readonly transport: 'google-drive-connector'
    readonly sourceWorkbook: 'xlsx-native-google-sheets-conversion'
    readonly valueRenderOption: 'UNFORMATTED_VALUE'
    readonly measuredGoogleSheetsOperation: 'structural-edit-and-read-verification-values'
    readonly sampleCount: number
    readonly samplingOrder: 'engine-isolated-workpaper-then-google-sheets'
  }
  readonly googleSheets: {
    readonly spreadsheets: GoogleSheetsLiveStructuralSpreadsheet[]
  }
  readonly cases: GoogleSheetsLiveStructuralCaptureCase[]
}

export interface GoogleSheetsLiveStructuralSpreadsheet {
  readonly caseId: string
  readonly sampleIndex: number
  readonly spreadsheetId: string
  readonly spreadsheetUrl: string
  readonly title: string
}

export interface GoogleSheetsLiveStructuralCaptureCase {
  readonly id: string
  readonly operation: GoogleSheetsLiveStructuralOperation
  readonly axis: 'rows' | 'columns'
  readonly rowCount: number
  readonly googleSheetsElapsedMsSamples: number[]
  readonly googleSheetsVerificationSamples: Array<Record<string, VerificationValue>>
}

export interface GoogleSheetsLiveStructuralCase {
  readonly id: string
  readonly operation: GoogleSheetsLiveStructuralOperation
  readonly axis: 'rows' | 'columns'
  readonly rowCount: number
  readonly sampleCount: number
  readonly workpaperElapsedMs: NumericSummary
  readonly googleSheetsElapsedMs: NumericSummary
  readonly workpaperToGoogleSheetsMeanRatio: number
  readonly workpaperToGoogleSheetsP95Ratio: number
  readonly tenXMeanAndP95: boolean
  readonly verification: {
    readonly workpaper: Record<string, VerificationValue>
    readonly googleSheets: Record<string, VerificationValue>
    readonly equivalent: boolean
  }
  readonly passed: boolean
}

export interface GoogleSheetsLiveStructuralScorecard {
  readonly schemaVersion: 1
  readonly suite: 'google-sheets-live-structural-performance'
  readonly generatedAt: string
  readonly host: {
    readonly arch: string
    readonly platform: string
  }
  readonly source: {
    readonly artifactGenerator: 'scripts/gen-google-sheets-live-structural-scorecard.ts'
    readonly implementationPackage: 'packages/headless'
    readonly evidenceKind: 'live-google-sheets-native-conversion-via-google-drive-connector'
    readonly captureTransport: 'google-drive-connector'
  }
  readonly benchmark: {
    readonly rowCount: number
    readonly sampleCount: number
    readonly valueRenderOption: 'UNFORMATTED_VALUE'
    readonly measuredGoogleSheetsOperation: 'structural-edit-and-read-verification-values'
    readonly measuredWorkpaperOperation: 'structural-edit'
    readonly samplingOrder: 'engine-isolated-workpaper-then-google-sheets'
  }
  readonly googleSheets: {
    readonly spreadsheets: GoogleSheetsLiveStructuralSpreadsheet[]
  }
  readonly summary: {
    readonly allRequiredCasesPassed: boolean
    readonly requiredCaseCount: number
    readonly tenXMeanAndP95CaseCount: number
    readonly workpaperWins: number
    readonly coveredOperations: GoogleSheetsLiveStructuralOperation[]
    readonly microsoftExcelEvidence: 'not-covered-by-this-artifact'
  }
  readonly cases: GoogleSheetsLiveStructuralCase[]
}

interface StructuralCaseSpec {
  readonly id: string
  readonly operation: GoogleSheetsLiveStructuralOperation
  readonly axis: 'rows' | 'columns'
}

interface WorkPaperSampleResult {
  readonly elapsedMs: number
  readonly verification: Record<string, VerificationValue>
}

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const outputPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'google-sheets-live-structural-scorecard.json')
const worksheetName = 'Bench'
const rowCount = 500
const sampleCount = 3
const TEN_X_RATIO = 0.1
const caseSpecs = [
  { id: 'google-sheets-live-structural-insert-rows', operation: 'insert-rows', axis: 'rows' },
  { id: 'google-sheets-live-structural-delete-rows', operation: 'delete-rows', axis: 'rows' },
  { id: 'google-sheets-live-structural-move-rows', operation: 'move-rows', axis: 'rows' },
  { id: 'google-sheets-live-structural-insert-columns', operation: 'insert-columns', axis: 'columns' },
  { id: 'google-sheets-live-structural-delete-columns', operation: 'delete-columns', axis: 'columns' },
  { id: 'google-sheets-live-structural-move-columns', operation: 'move-columns', axis: 'columns' },
] as const satisfies readonly StructuralCaseSpec[]

async function main(): Promise<void> {
  const isCheckMode = process.argv.includes('--check')
  const emitXlsxIndex = process.argv.indexOf('--emit-xlsx')
  const captureIndex = process.argv.indexOf('--capture')

  if (emitXlsxIndex >= 0) {
    const targetDirectory = process.argv[emitXlsxIndex + 1]
    if (!targetDirectory) {
      throw new Error('Missing directory after --emit-xlsx')
    }
    emitGoogleSheetsStructuralXlsx(resolve(targetDirectory))
    return
  }

  if (isCheckMode) {
    if (!existsSync(outputPath)) {
      throw new Error(
        `Google Sheets live structural scorecard is missing. Run: bun scripts/gen-google-sheets-live-structural-scorecard.ts --capture <capture.json>`,
      )
    }
    const scorecard = parseGoogleSheetsLiveStructuralScorecard(readJsonObject(outputPath))
    validateGoogleSheetsLiveStructuralScorecard(scorecard)
    logResult('check', scorecard)
    return
  }

  if (captureIndex < 0) {
    throw new Error(
      'Missing --capture <capture.json>. First emit XLSX files with --emit-xlsx, import them as native Google Sheets, apply the structural edit, then capture edit+read timings and UNFORMATTED_VALUE verification cells through the Google Drive connector.',
    )
  }
  const capturePath = process.argv[captureIndex + 1]
  if (!capturePath) {
    throw new Error('Missing path after --capture')
  }
  const scorecard = buildGoogleSheetsLiveStructuralScorecard(parseGoogleSheetsLiveStructuralCapture(readJsonObject(capturePath)))
  validateGoogleSheetsLiveStructuralScorecard(scorecard)
  mkdirSync(dirname(outputPath), { recursive: true })
  writeFileSync(outputPath, formatJsonForRepo(`${JSON.stringify(scorecard, null, 2)}\n`))
  logResult('write', scorecard)
}

export function emitGoogleSheetsStructuralXlsx(targetDirectory: string): void {
  mkdirSync(targetDirectory, { recursive: true })
  const outputs = caseSpecs.map((caseSpec) => {
    const outputFile = join(targetDirectory, `${caseSpec.operation}.xlsx`)
    writeFileSync(outputFile, createGoogleSheetsWorkbookBytes(caseSpec.operation))
    return {
      caseId: caseSpec.id,
      operation: caseSpec.operation,
      axis: caseSpec.axis,
      rowCount,
      outputFile,
      sheetName: worksheetName,
      batchUpdateRequest: googleSheetsOperationInstruction(caseSpec.operation),
      verificationRanges: googleSheetsVerificationRanges(caseSpec.operation),
    }
  })
  console.log(
    JSON.stringify(
      {
        mode: 'emit-xlsx',
        targetDirectory,
        uploadMode: 'native_google_sheets',
        valueRenderOption: 'UNFORMATTED_VALUE',
        measuredGoogleSheetsOperation: 'structural-edit-and-read-verification-values',
        outputs,
      },
      null,
      2,
    ),
  )
}

export function buildGoogleSheetsLiveStructuralScorecard(capture: GoogleSheetsLiveStructuralCapture): GoogleSheetsLiveStructuralScorecard {
  if (capture.capture.sampleCount !== sampleCount) {
    throw new Error(`Google Sheets live structural capture sample settings must be sample=${String(sampleCount)}`)
  }

  const cases = caseSpecs.map((caseSpec) => runWorkPaperCaseAgainstCapture(capture, caseSpec))
  const tenXMeanAndP95CaseCount = cases.filter((entry) => entry.tenXMeanAndP95).length
  const workpaperWins = cases.filter((entry) => entry.workpaperElapsedMs.mean <= entry.googleSheetsElapsedMs.mean).length

  return {
    schemaVersion: 1,
    suite: 'google-sheets-live-structural-performance',
    generatedAt: capture.generatedAt,
    host: {
      arch: process.arch,
      platform: process.platform,
    },
    source: {
      artifactGenerator: 'scripts/gen-google-sheets-live-structural-scorecard.ts',
      implementationPackage: 'packages/headless',
      evidenceKind: 'live-google-sheets-native-conversion-via-google-drive-connector',
      captureTransport: capture.capture.transport,
    },
    benchmark: {
      rowCount,
      sampleCount,
      valueRenderOption: capture.capture.valueRenderOption,
      measuredGoogleSheetsOperation: capture.capture.measuredGoogleSheetsOperation,
      measuredWorkpaperOperation: 'structural-edit',
      samplingOrder: capture.capture.samplingOrder,
    },
    googleSheets: {
      spreadsheets: caseSpecs.flatMap((caseSpec) =>
        Array.from({ length: sampleCount }, (_unused, sampleIndex) => requiredSpreadsheet(capture, caseSpec.id, sampleIndex)),
      ),
    },
    summary: {
      allRequiredCasesPassed: cases.every((entry) => entry.passed),
      requiredCaseCount: cases.length,
      tenXMeanAndP95CaseCount,
      workpaperWins,
      coveredOperations: caseSpecs.map((entry) => entry.operation),
      microsoftExcelEvidence: 'not-covered-by-this-artifact',
    },
    cases,
  }
}

export function parseGoogleSheetsLiveStructuralCapture(value: Record<string, unknown>): GoogleSheetsLiveStructuralCapture {
  const capture = objectField(value, 'capture')
  const googleSheets = objectField(value, 'googleSheets')
  return {
    schemaVersion: literalField(value, 'schemaVersion', 1),
    generatedAt: stringField(value, 'generatedAt'),
    capture: {
      transport: literalField(capture, 'transport', 'google-drive-connector'),
      sourceWorkbook: literalField(capture, 'sourceWorkbook', 'xlsx-native-google-sheets-conversion'),
      valueRenderOption: literalField(capture, 'valueRenderOption', 'UNFORMATTED_VALUE'),
      measuredGoogleSheetsOperation: literalField(capture, 'measuredGoogleSheetsOperation', 'structural-edit-and-read-verification-values'),
      sampleCount: numberField(capture, 'sampleCount'),
      samplingOrder: literalField(capture, 'samplingOrder', 'engine-isolated-workpaper-then-google-sheets'),
    },
    googleSheets: {
      spreadsheets: arrayField(googleSheets, 'spreadsheets').map(parseSpreadsheet),
    },
    cases: arrayField(value, 'cases').map(parseCaptureCase),
  }
}

export function parseGoogleSheetsLiveStructuralScorecard(value: Record<string, unknown>): GoogleSheetsLiveStructuralScorecard {
  const host = objectField(value, 'host')
  const source = objectField(value, 'source')
  const benchmark = objectField(value, 'benchmark')
  const googleSheets = objectField(value, 'googleSheets')
  const summary = objectField(value, 'summary')
  return {
    schemaVersion: literalField(value, 'schemaVersion', 1),
    suite: literalField(value, 'suite', 'google-sheets-live-structural-performance'),
    generatedAt: stringField(value, 'generatedAt'),
    host: {
      arch: stringField(host, 'arch'),
      platform: stringField(host, 'platform'),
    },
    source: {
      artifactGenerator: literalField(source, 'artifactGenerator', 'scripts/gen-google-sheets-live-structural-scorecard.ts'),
      implementationPackage: literalField(source, 'implementationPackage', 'packages/headless'),
      evidenceKind: literalField(source, 'evidenceKind', 'live-google-sheets-native-conversion-via-google-drive-connector'),
      captureTransport: literalField(source, 'captureTransport', 'google-drive-connector'),
    },
    benchmark: {
      rowCount: numberField(benchmark, 'rowCount'),
      sampleCount: numberField(benchmark, 'sampleCount'),
      valueRenderOption: literalField(benchmark, 'valueRenderOption', 'UNFORMATTED_VALUE'),
      measuredGoogleSheetsOperation: literalField(
        benchmark,
        'measuredGoogleSheetsOperation',
        'structural-edit-and-read-verification-values',
      ),
      measuredWorkpaperOperation: literalField(benchmark, 'measuredWorkpaperOperation', 'structural-edit'),
      samplingOrder: literalField(benchmark, 'samplingOrder', 'engine-isolated-workpaper-then-google-sheets'),
    },
    googleSheets: {
      spreadsheets: arrayField(googleSheets, 'spreadsheets').map(parseSpreadsheet),
    },
    summary: {
      allRequiredCasesPassed: booleanField(summary, 'allRequiredCasesPassed'),
      requiredCaseCount: numberField(summary, 'requiredCaseCount'),
      tenXMeanAndP95CaseCount: numberField(summary, 'tenXMeanAndP95CaseCount'),
      workpaperWins: numberField(summary, 'workpaperWins'),
      coveredOperations: stringArrayField(summary, 'coveredOperations').map(parseOperation),
      microsoftExcelEvidence: literalField(summary, 'microsoftExcelEvidence', 'not-covered-by-this-artifact'),
    },
    cases: arrayField(value, 'cases').map(parseStructuralCase),
  }
}

export function validateGoogleSheetsLiveStructuralScorecard(scorecard: GoogleSheetsLiveStructuralScorecard): void {
  const expectedIds = caseSpecs.map((entry) => entry.id)
  const expectedOperations = caseSpecs.map((entry) => entry.operation)
  const spreadsheetSampleKeys = scorecard.googleSheets.spreadsheets.map((entry) => `${entry.caseId}:${String(entry.sampleIndex)}`)
  const expectedSpreadsheetSampleKeys = caseSpecs.flatMap((entry) =>
    Array.from({ length: sampleCount }, (_unused, sampleIndex) => `${entry.id}:${String(sampleIndex)}`),
  )
  if (scorecard.benchmark.rowCount !== rowCount || scorecard.benchmark.sampleCount !== sampleCount) {
    throw new Error('Google Sheets live structural scorecard benchmark settings are stale')
  }
  if (
    scorecard.summary.requiredCaseCount !== expectedIds.length ||
    JSON.stringify(scorecard.cases.map((entry) => entry.id)) !== JSON.stringify(expectedIds)
  ) {
    throw new Error('Google Sheets live structural scorecard required cases are stale')
  }
  if (JSON.stringify(spreadsheetSampleKeys) !== JSON.stringify(expectedSpreadsheetSampleKeys)) {
    throw new Error('Google Sheets live structural scorecard spreadsheet evidence is stale')
  }
  for (const spreadsheet of scorecard.googleSheets.spreadsheets) {
    if (
      spreadsheet.spreadsheetId.trim().length === 0 ||
      spreadsheet.spreadsheetUrl.trim().length === 0 ||
      spreadsheet.title.trim().length === 0
    ) {
      throw new Error(`Google Sheets live structural spreadsheet evidence is incomplete for ${spreadsheet.caseId}`)
    }
  }
  if (JSON.stringify(scorecard.summary.coveredOperations) !== JSON.stringify(expectedOperations)) {
    throw new Error('Google Sheets live structural scorecard covered operations are stale')
  }
  if (
    scorecard.summary.workpaperWins !==
    scorecard.cases.filter((entry) => entry.workpaperElapsedMs.mean <= entry.googleSheetsElapsedMs.mean).length
  ) {
    throw new Error('Google Sheets live structural scorecard workpaper win count is inconsistent')
  }
  if (scorecard.summary.tenXMeanAndP95CaseCount !== scorecard.cases.filter((entry) => entry.tenXMeanAndP95).length) {
    throw new Error('Google Sheets live structural scorecard 10x count is inconsistent')
  }
  const failingCases = scorecard.cases.filter((entry) => !entry.passed)
  if (!scorecard.summary.allRequiredCasesPassed || failingCases.length > 0) {
    throw new Error(
      `Google Sheets live structural scorecard has failing required cases: ${failingCases
        .map(
          (entry) =>
            `${entry.id} WorkPaper=${JSON.stringify(entry.verification.workpaper)} GoogleSheets=${JSON.stringify(
              entry.verification.googleSheets,
            )}`,
        )
        .join(', ')}`,
    )
  }
  for (const entry of scorecard.cases) {
    validateNumericSummary(entry.workpaperElapsedMs, `${entry.id} WorkPaper elapsedMs`)
    validateNumericSummary(entry.googleSheetsElapsedMs, `${entry.id} Google Sheets elapsedMs`)
    const expectedTenXMeanAndP95 =
      entry.workpaperToGoogleSheetsMeanRatio <= TEN_X_RATIO && entry.workpaperToGoogleSheetsP95Ratio <= TEN_X_RATIO
    if (entry.tenXMeanAndP95 !== expectedTenXMeanAndP95) {
      throw new Error(`Google Sheets live structural 10x flag is stale: ${entry.id}`)
    }
    if (!entry.verification.equivalent) {
      throw new Error(`Google Sheets live structural scorecard verification mismatch: ${entry.id}`)
    }
    const currentVerification = runWorkPaperSample(entry.operation).verification
    if (JSON.stringify(entry.verification.workpaper) !== JSON.stringify(currentVerification)) {
      throw new Error(`Google Sheets live structural WorkPaper verification is stale: ${entry.id}`)
    }
  }
}

function runWorkPaperCaseAgainstCapture(
  capture: GoogleSheetsLiveStructuralCapture,
  caseSpec: StructuralCaseSpec,
): GoogleSheetsLiveStructuralCase {
  const captureCase = requiredCaptureCase(capture, caseSpec)
  const workpaperSamples: number[] = []
  const workpaperVerifications: Array<Record<string, VerificationValue>> = []
  for (let index = 0; index < sampleCount; index += 1) {
    const sample = runWorkPaperSample(caseSpec.operation)
    workpaperSamples.push(sample.elapsedMs)
    workpaperVerifications.push(sample.verification)
  }

  const workpaperElapsedMs = summarizeNumbers(workpaperSamples)
  const googleSheetsElapsedMs = summarizeNumbers(captureCase.googleSheetsElapsedMsSamples)
  const workpaperToGoogleSheetsMeanRatio = workpaperElapsedMs.mean / googleSheetsElapsedMs.mean
  const workpaperToGoogleSheetsP95Ratio = workpaperElapsedMs.p95 / googleSheetsElapsedMs.p95
  const verification = {
    workpaper: workpaperVerifications[0] ?? {},
    googleSheets: captureCase.googleSheetsVerificationSamples[0] ?? {},
    equivalent: verificationsEquivalent(workpaperVerifications, captureCase.googleSheetsVerificationSamples),
  }

  return {
    id: caseSpec.id,
    operation: caseSpec.operation,
    axis: caseSpec.axis,
    rowCount,
    sampleCount,
    workpaperElapsedMs,
    googleSheetsElapsedMs,
    workpaperToGoogleSheetsMeanRatio,
    workpaperToGoogleSheetsP95Ratio,
    tenXMeanAndP95: workpaperToGoogleSheetsMeanRatio <= TEN_X_RATIO && workpaperToGoogleSheetsP95Ratio <= TEN_X_RATIO,
    verification,
    passed:
      verification.equivalent && workpaperElapsedMs.samples.length === sampleCount && googleSheetsElapsedMs.samples.length === sampleCount,
  }
}

function runWorkPaperSample(operation: GoogleSheetsLiveStructuralOperation): WorkPaperSampleResult {
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

function createGoogleSheetsWorkbookBytes(operation: GoogleSheetsLiveStructuralOperation): Uint8Array {
  const worksheet = aoaToFormulaWorksheet(sheetForOperation(operation))
  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(workbook, worksheet, worksheetName)
  return XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' })
}

function sheetForOperation(operation: GoogleSheetsLiveStructuralOperation): StructuralSheet {
  return operation.includes('columns') ? buildStructuralColumnSheet(rowCount) : buildOverlappingAggregateSheet(rowCount)
}

function workpaperVerification(
  workbook: WorkPaper,
  sheetId: number,
  operation: GoogleSheetsLiveStructuralOperation,
): Record<string, VerificationValue> {
  const cell = verificationCell(operation)
  return {
    targetCell: cell.a1,
    value: normalizeProtocolValue(workbook.getCellValue({ sheet: sheetId, row: cell.row, col: cell.col })),
  }
}

function verificationCell(operation: GoogleSheetsLiveStructuralOperation): { row: number; col: number; a1: string } {
  switch (operation) {
    case 'insert-rows':
      return { row: rowCount, col: 0, a1: `A${String(rowCount + 1)}` }
    case 'delete-rows':
      return { row: rowCount - 2, col: 0, a1: `A${String(rowCount - 1)}` }
    case 'move-rows':
      return { row: 0, col: 0, a1: 'A1' }
    case 'insert-columns':
      return { row: rowCount - 1, col: 0, a1: `A${String(rowCount)}` }
    case 'delete-columns':
      return { row: rowCount - 1, col: 0, a1: `A${String(rowCount)}` }
    case 'move-columns':
      return { row: rowCount - 1, col: 0, a1: `A${String(rowCount)}` }
  }
}

function googleSheetsOperationInstruction(operation: GoogleSheetsLiveStructuralOperation): Record<string, VerificationValue> {
  switch (operation) {
    case 'insert-rows':
      return { request: 'insertDimension', dimension: 'ROWS', startIndex: Math.floor(rowCount / 2), endIndex: Math.floor(rowCount / 2) + 1 }
    case 'delete-rows':
      return { request: 'deleteDimension', dimension: 'ROWS', startIndex: Math.floor(rowCount / 2), endIndex: Math.floor(rowCount / 2) + 1 }
    case 'move-rows':
      return { request: 'moveDimension', dimension: 'ROWS', startIndex: Math.floor(rowCount / 2), endIndex: Math.floor(rowCount / 2) + 1 }
    case 'insert-columns':
      return { request: 'insertDimension', dimension: 'COLUMNS', startIndex: 1, endIndex: 2 }
    case 'delete-columns':
      return { request: 'deleteDimension', dimension: 'COLUMNS', startIndex: 1, endIndex: 2 }
    case 'move-columns':
      return { request: 'moveDimension', dimension: 'COLUMNS', startIndex: 1, endIndex: 2 }
  }
}

function googleSheetsVerificationRanges(operation: GoogleSheetsLiveStructuralOperation): string[] {
  return [verificationCell(operation).a1]
}

function aoaToFormulaWorksheet(sheet: StructuralSheet): XLSX.WorkSheet {
  const worksheet = XLSX.utils.aoa_to_sheet(
    sheet.map((row) => row.map((cell) => (typeof cell === 'string' && cell.startsWith('=') ? null : cell))),
  )
  for (let rowIndex = 0; rowIndex < sheet.length; rowIndex += 1) {
    const row = sheet[rowIndex] ?? []
    for (let colIndex = 0; colIndex < row.length; colIndex += 1) {
      const cell = row[colIndex]
      if (typeof cell === 'string' && cell.startsWith('=')) {
        worksheet[XLSX.utils.encode_cell({ r: rowIndex, c: colIndex })] = { t: 'n', f: cell.slice(1) }
      }
    }
  }
  return worksheet
}

function parseStructuralCase(value: unknown): GoogleSheetsLiveStructuralCase {
  const record = asObject(value, 'Google Sheets live structural case')
  const verification = objectField(record, 'verification')
  return {
    id: stringField(record, 'id'),
    operation: parseOperation(stringField(record, 'operation')),
    axis: parseAxis(stringField(record, 'axis')),
    rowCount: numberField(record, 'rowCount'),
    sampleCount: numberField(record, 'sampleCount'),
    workpaperElapsedMs: parseNumericSummary(objectField(record, 'workpaperElapsedMs')),
    googleSheetsElapsedMs: parseNumericSummary(objectField(record, 'googleSheetsElapsedMs')),
    workpaperToGoogleSheetsMeanRatio: numberField(record, 'workpaperToGoogleSheetsMeanRatio'),
    workpaperToGoogleSheetsP95Ratio: numberField(record, 'workpaperToGoogleSheetsP95Ratio'),
    tenXMeanAndP95: booleanField(record, 'tenXMeanAndP95'),
    verification: {
      workpaper: parseVerificationRecord(objectField(verification, 'workpaper')),
      googleSheets: parseVerificationRecord(objectField(verification, 'googleSheets')),
      equivalent: booleanField(verification, 'equivalent'),
    },
    passed: booleanField(record, 'passed'),
  }
}

function parseCaptureCase(value: unknown): GoogleSheetsLiveStructuralCaptureCase {
  const record = asObject(value, 'Google Sheets live structural capture case')
  return {
    id: stringField(record, 'id'),
    operation: parseOperation(stringField(record, 'operation')),
    axis: parseAxis(stringField(record, 'axis')),
    rowCount: numberField(record, 'rowCount'),
    googleSheetsElapsedMsSamples: numberArrayField(record, 'googleSheetsElapsedMsSamples'),
    googleSheetsVerificationSamples: arrayField(record, 'googleSheetsVerificationSamples').map((entry) =>
      parseVerificationRecord(asObject(entry, 'Google Sheets live structural verification sample')),
    ),
  }
}

function parseSpreadsheet(value: unknown): GoogleSheetsLiveStructuralSpreadsheet {
  const record = asObject(value, 'Google Sheets live structural spreadsheet')
  return {
    caseId: stringField(record, 'caseId'),
    sampleIndex: numberField(record, 'sampleIndex'),
    spreadsheetId: stringField(record, 'spreadsheetId'),
    spreadsheetUrl: stringField(record, 'spreadsheetUrl'),
    title: stringField(record, 'title'),
  }
}

function parseNumericSummary(value: Record<string, unknown>): NumericSummary {
  const confidence95 = objectField(value, 'confidence95')
  return {
    samples: numberArrayField(value, 'samples'),
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

function parseVerificationRecord(value: Record<string, unknown>): Record<string, VerificationValue> {
  const result: Record<string, VerificationValue> = {}
  for (const [key, entry] of Object.entries(value)) {
    if (entry === null || typeof entry === 'boolean' || typeof entry === 'number' || typeof entry === 'string') {
      result[key] = entry
      continue
    }
    throw new Error(`Unsupported Google Sheets structural verification value for ${key}`)
  }
  return result
}

function parseOperation(value: string): GoogleSheetsLiveStructuralOperation {
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
  throw new Error(`Unexpected Google Sheets live structural operation: ${value}`)
}

function parseAxis(value: string): 'rows' | 'columns' {
  if (value === 'rows' || value === 'columns') {
    return value
  }
  throw new Error(`Unexpected Google Sheets live structural axis: ${value}`)
}

function requiredCaptureCase(
  capture: GoogleSheetsLiveStructuralCapture,
  caseSpec: StructuralCaseSpec,
): GoogleSheetsLiveStructuralCaptureCase {
  const captureCase = capture.cases.find((entry) => entry.id === caseSpec.id)
  if (!captureCase) {
    throw new Error(`Google Sheets live structural capture is missing required case: ${caseSpec.id}`)
  }
  if (
    captureCase.operation !== caseSpec.operation ||
    captureCase.axis !== caseSpec.axis ||
    captureCase.rowCount !== rowCount ||
    captureCase.googleSheetsElapsedMsSamples.length !== sampleCount ||
    captureCase.googleSheetsVerificationSamples.length !== sampleCount ||
    captureCase.googleSheetsElapsedMsSamples.some((entry) => !Number.isFinite(entry) || entry < 0)
  ) {
    throw new Error(`Google Sheets live structural capture case settings are stale: ${caseSpec.id}`)
  }
  return captureCase
}

function requiredSpreadsheet(
  capture: GoogleSheetsLiveStructuralCapture,
  caseId: string,
  sampleIndex: number,
): GoogleSheetsLiveStructuralSpreadsheet {
  const spreadsheet = capture.googleSheets.spreadsheets.find((entry) => entry.caseId === caseId && entry.sampleIndex === sampleIndex)
  if (!spreadsheet) {
    throw new Error(`Google Sheets live structural capture is missing spreadsheet evidence for ${caseId} sample ${String(sampleIndex)}`)
  }
  return spreadsheet
}

function verificationsEquivalent(
  workpaperVerifications: readonly Record<string, VerificationValue>[],
  googleSheetsVerifications: readonly Record<string, VerificationValue>[],
): boolean {
  if (workpaperVerifications.length !== googleSheetsVerifications.length) {
    return false
  }
  const expected = workpaperVerifications[0]
  if (expected === undefined) {
    return false
  }
  return (
    workpaperVerifications.every((entry) => JSON.stringify(entry) === JSON.stringify(expected)) &&
    googleSheetsVerifications.every((entry) => JSON.stringify(entry) === JSON.stringify(expected))
  )
}

function normalizeProtocolValue(value: unknown): VerificationValue {
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

function isProtocolValueLike(value: unknown): value is { tag: ValueTag; value?: boolean | number | string } {
  if (value === null || typeof value !== 'object') {
    return false
  }
  const tag = Reflect.get(value, 'tag')
  return tag === ValueTag.Empty || tag === ValueTag.Number || tag === ValueTag.Boolean || tag === ValueTag.String || tag === ValueTag.Error
}

function validateNumericSummary(summary: NumericSummary, label: string): void {
  if (summary.samples.length !== sampleCount || summary.samples.some((entry) => !Number.isFinite(entry) || entry < 0)) {
    throw new Error(`${label} has invalid samples`)
  }
}

function logResult(mode: 'check' | 'write', scorecard: GoogleSheetsLiveStructuralScorecard): void {
  console.log(
    JSON.stringify(
      {
        mode,
        outputPath,
        allRequiredCasesPassed: scorecard.summary.allRequiredCasesPassed,
        workpaperWins: scorecard.summary.workpaperWins,
        tenXMeanAndP95CaseCount: scorecard.summary.tenXMeanAndP95CaseCount,
        spreadsheetIds: scorecard.googleSheets.spreadsheets.map((entry) => entry.spreadsheetId),
      },
      null,
      2,
    ),
  )
}

if (process.argv[1] && pathToFileURL(resolve(process.argv[1])).href === import.meta.url) {
  await main()
}
