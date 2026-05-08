#!/usr/bin/env bun

import { existsSync, mkdirSync, writeFileSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { pathToFileURL } from 'node:url'
import { performance } from 'node:perf_hooks'

import { WorkPaper } from '@bilig/headless'
import { ValueTag } from '@bilig/protocol'
import * as XLSX from 'xlsx'
import { summarizeNumbers, type NumericSummary } from '../packages/benchmarks/src/stats.js'
import {
  buildConditionalAggregationSheet,
  buildFormulaFanoutRow,
  buildParserCacheTemplateSheet,
  buildValueFormulaRows,
} from '../packages/benchmarks/src/workpaper-benchmark-fixtures.js'
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

export type GoogleSheetsLiveRecalculationWorkload =
  | 'dirty-fanout-edit'
  | 'suspended-batch-single-column-edit'
  | 'conditional-aggregation-criteria-edit'
  | 'full-rebuild-recalculate'

type VerificationValue = boolean | number | string | null
type RecalculationSheet = ReadonlyArray<ReadonlyArray<boolean | number | string | null>>

export interface GoogleSheetsLiveRecalculationCapture {
  readonly schemaVersion: 1
  readonly generatedAt: string
  readonly capture: {
    readonly transport: 'google-drive-connector'
    readonly sourceWorkbook: 'xlsx-native-google-sheets-conversion'
    readonly valueRenderOption: 'UNFORMATTED_VALUE'
    readonly measuredGoogleSheetsOperation: 'edit-and-read-recalculated-values'
    readonly sampleCount: number
    readonly warmupCount: 0
    readonly samplingOrder: 'engine-isolated-workpaper-then-google-sheets'
  }
  readonly googleSheets: {
    readonly spreadsheets: GoogleSheetsLiveRecalculationSpreadsheet[]
  }
  readonly cases: GoogleSheetsLiveRecalculationCaptureCase[]
}

export interface GoogleSheetsLiveRecalculationSpreadsheet {
  readonly caseId: string
  readonly sampleIndex: number
  readonly spreadsheetId: string
  readonly spreadsheetUrl: string
  readonly title: string
}

export interface GoogleSheetsLiveRecalculationCaptureCase {
  readonly id: string
  readonly workload: GoogleSheetsLiveRecalculationWorkload
  readonly fixture: RecalculationFixture
  readonly googleSheetsElapsedMsSamples: number[]
  readonly googleSheetsVerificationSamples: Array<Record<string, VerificationValue>>
}

export interface RecalculationFixture {
  readonly rowCount: number
  readonly formulaCount: number
  readonly materializedCells: number
}

export interface GoogleSheetsLiveRecalculationCase {
  readonly id: string
  readonly workload: GoogleSheetsLiveRecalculationWorkload
  readonly fixture: RecalculationFixture
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

export interface GoogleSheetsLiveRecalculationScorecard {
  readonly schemaVersion: 1
  readonly suite: 'google-sheets-live-recalculation-performance'
  readonly generatedAt: string
  readonly host: {
    readonly arch: string
    readonly platform: string
  }
  readonly source: {
    readonly artifactGenerator: 'scripts/gen-google-sheets-live-recalculation-scorecard.ts'
    readonly implementationPackage: 'packages/headless'
    readonly evidenceKind: 'live-google-sheets-native-conversion-via-google-drive-connector'
    readonly captureTransport: 'google-drive-connector'
  }
  readonly benchmark: {
    readonly sampleCount: number
    readonly warmupCount: 0
    readonly valueRenderOption: 'UNFORMATTED_VALUE'
    readonly measuredGoogleSheetsOperation: 'edit-and-read-recalculated-values'
    readonly measuredWorkpaperOperation: 'mutate-and-recalculate'
    readonly samplingOrder: 'engine-isolated-workpaper-then-google-sheets'
  }
  readonly googleSheets: {
    readonly spreadsheets: GoogleSheetsLiveRecalculationSpreadsheet[]
  }
  readonly summary: {
    readonly allRequiredCasesPassed: boolean
    readonly requiredCaseCount: number
    readonly tenXMeanAndP95CaseCount: number
    readonly workpaperWins: number
    readonly coveredWorkloads: GoogleSheetsLiveRecalculationWorkload[]
    readonly microsoftExcelEvidence: 'not-covered-by-this-artifact'
  }
  readonly cases: GoogleSheetsLiveRecalculationCase[]
}

interface RecalculationCaseSpec {
  readonly id: string
  readonly workload: GoogleSheetsLiveRecalculationWorkload
  readonly fixture: RecalculationFixture
}

interface WorkPaperSampleResult {
  readonly elapsedMs: number
  readonly verification: Record<string, VerificationValue>
}

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const outputPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'google-sheets-live-recalculation-scorecard.json')
const worksheetName = 'Bench'
const sampleCount = 3
const warmupCount = 0
const TEN_X_RATIO = 0.1
const fanoutCount = 1_000
const singleColumnRowCount = 1_000
const aggregationRowCount = 1_000
const aggregationFormulaCopies = 8
const rebuildRowCount = 1_000
const caseSpecs = [
  {
    id: 'google-sheets-live-recalculation-dirty-fanout-edit',
    workload: 'dirty-fanout-edit',
    fixture: {
      rowCount: 1,
      formulaCount: fanoutCount,
      materializedCells: fanoutCount + 1,
    },
  },
  {
    id: 'google-sheets-live-recalculation-suspended-batch-single-column-edit',
    workload: 'suspended-batch-single-column-edit',
    fixture: {
      rowCount: singleColumnRowCount,
      formulaCount: singleColumnRowCount,
      materializedCells: singleColumnRowCount * 2,
    },
  },
  {
    id: 'google-sheets-live-recalculation-conditional-aggregation-criteria-edit',
    workload: 'conditional-aggregation-criteria-edit',
    fixture: {
      rowCount: aggregationRowCount,
      formulaCount: aggregationFormulaCopies * 2,
      materializedCells: aggregationRowCount * 2 + aggregationFormulaCopies * 2 + 4,
    },
  },
  {
    id: 'google-sheets-live-recalculation-full-rebuild-recalculate',
    workload: 'full-rebuild-recalculate',
    fixture: {
      rowCount: rebuildRowCount,
      formulaCount: rebuildRowCount * 4,
      materializedCells: rebuildRowCount * 6,
    },
  },
] as const satisfies readonly RecalculationCaseSpec[]

async function main(): Promise<void> {
  const isCheckMode = process.argv.includes('--check')
  const emitXlsxIndex = process.argv.indexOf('--emit-xlsx')
  const captureIndex = process.argv.indexOf('--capture')

  if (emitXlsxIndex >= 0) {
    const targetDirectory = process.argv[emitXlsxIndex + 1]
    if (!targetDirectory) {
      throw new Error('Missing directory after --emit-xlsx')
    }
    emitGoogleSheetsRecalculationXlsx(resolve(targetDirectory))
    return
  }

  if (isCheckMode) {
    if (!existsSync(outputPath)) {
      throw new Error(
        `Google Sheets live recalculation scorecard is missing. Run: bun scripts/gen-google-sheets-live-recalculation-scorecard.ts --capture <capture.json>`,
      )
    }
    const scorecard = parseGoogleSheetsLiveRecalculationScorecard(readJsonObject(outputPath))
    validateGoogleSheetsLiveRecalculationScorecard(scorecard)
    logResult('check', scorecard)
    return
  }

  if (captureIndex < 0) {
    throw new Error(
      'Missing --capture <capture.json>. First emit XLSX files with --emit-xlsx, import them as native Google Sheets, apply the workload edit, then capture edit+read timings and UNFORMATTED_VALUE verification cells through the Google Drive connector.',
    )
  }
  const capturePath = process.argv[captureIndex + 1]
  if (!capturePath) {
    throw new Error('Missing path after --capture')
  }
  const scorecard = buildGoogleSheetsLiveRecalculationScorecard(parseGoogleSheetsLiveRecalculationCapture(readJsonObject(capturePath)))
  validateGoogleSheetsLiveRecalculationScorecard(scorecard)
  mkdirSync(dirname(outputPath), { recursive: true })
  writeFileSync(outputPath, formatJsonForRepo(`${JSON.stringify(scorecard, null, 2)}\n`))
  logResult('write', scorecard)
}

export function emitGoogleSheetsRecalculationXlsx(targetDirectory: string): void {
  mkdirSync(targetDirectory, { recursive: true })
  const outputs = caseSpecs.map((caseSpec) => {
    const outputFile = join(targetDirectory, `${caseSpec.workload}.xlsx`)
    writeFileSync(outputFile, createGoogleSheetsWorkbookBytes(caseSpec.workload))
    return {
      caseId: caseSpec.id,
      workload: caseSpec.workload,
      fixture: caseSpec.fixture,
      outputFile,
      sheetName: worksheetName,
      edit: googleSheetsEditInstruction(caseSpec.workload),
      verificationRanges: googleSheetsVerificationRanges(caseSpec.workload),
    }
  })
  console.log(
    JSON.stringify(
      {
        mode: 'emit-xlsx',
        targetDirectory,
        uploadMode: 'native_google_sheets',
        valueRenderOption: 'UNFORMATTED_VALUE',
        measuredGoogleSheetsOperation: 'edit-and-read-recalculated-values',
        outputs,
      },
      null,
      2,
    ),
  )
}

export function buildGoogleSheetsLiveRecalculationScorecard(
  capture: GoogleSheetsLiveRecalculationCapture,
): GoogleSheetsLiveRecalculationScorecard {
  if (capture.capture.sampleCount !== sampleCount || capture.capture.warmupCount !== warmupCount) {
    throw new Error(`Google Sheets live recalculation capture sample settings must be sample=${String(sampleCount)} warmup=0`)
  }

  const cases = caseSpecs.map((caseSpec) => runWorkPaperCaseAgainstCapture(capture, caseSpec))
  const tenXMeanAndP95CaseCount = cases.filter((entry) => entry.tenXMeanAndP95).length
  const workpaperWins = cases.filter((entry) => entry.workpaperElapsedMs.mean <= entry.googleSheetsElapsedMs.mean).length

  return {
    schemaVersion: 1,
    suite: 'google-sheets-live-recalculation-performance',
    generatedAt: capture.generatedAt,
    host: {
      arch: process.arch,
      platform: process.platform,
    },
    source: {
      artifactGenerator: 'scripts/gen-google-sheets-live-recalculation-scorecard.ts',
      implementationPackage: 'packages/headless',
      evidenceKind: 'live-google-sheets-native-conversion-via-google-drive-connector',
      captureTransport: capture.capture.transport,
    },
    benchmark: {
      sampleCount,
      warmupCount,
      valueRenderOption: capture.capture.valueRenderOption,
      measuredGoogleSheetsOperation: capture.capture.measuredGoogleSheetsOperation,
      measuredWorkpaperOperation: 'mutate-and-recalculate',
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
      coveredWorkloads: caseSpecs.map((entry) => entry.workload),
      microsoftExcelEvidence: 'not-covered-by-this-artifact',
    },
    cases,
  }
}

export function parseGoogleSheetsLiveRecalculationCapture(value: Record<string, unknown>): GoogleSheetsLiveRecalculationCapture {
  const capture = objectField(value, 'capture')
  const googleSheets = objectField(value, 'googleSheets')
  return {
    schemaVersion: literalField(value, 'schemaVersion', 1),
    generatedAt: stringField(value, 'generatedAt'),
    capture: {
      transport: literalField(capture, 'transport', 'google-drive-connector'),
      sourceWorkbook: literalField(capture, 'sourceWorkbook', 'xlsx-native-google-sheets-conversion'),
      valueRenderOption: literalField(capture, 'valueRenderOption', 'UNFORMATTED_VALUE'),
      measuredGoogleSheetsOperation: literalField(capture, 'measuredGoogleSheetsOperation', 'edit-and-read-recalculated-values'),
      sampleCount: numberField(capture, 'sampleCount'),
      warmupCount: literalField(capture, 'warmupCount', 0),
      samplingOrder: literalField(capture, 'samplingOrder', 'engine-isolated-workpaper-then-google-sheets'),
    },
    googleSheets: {
      spreadsheets: arrayField(googleSheets, 'spreadsheets').map(parseSpreadsheet),
    },
    cases: arrayField(value, 'cases').map(parseCaptureCase),
  }
}

export function parseGoogleSheetsLiveRecalculationScorecard(value: Record<string, unknown>): GoogleSheetsLiveRecalculationScorecard {
  const host = objectField(value, 'host')
  const source = objectField(value, 'source')
  const benchmark = objectField(value, 'benchmark')
  const googleSheets = objectField(value, 'googleSheets')
  const summary = objectField(value, 'summary')
  return {
    schemaVersion: literalField(value, 'schemaVersion', 1),
    suite: literalField(value, 'suite', 'google-sheets-live-recalculation-performance'),
    generatedAt: stringField(value, 'generatedAt'),
    host: {
      arch: stringField(host, 'arch'),
      platform: stringField(host, 'platform'),
    },
    source: {
      artifactGenerator: literalField(source, 'artifactGenerator', 'scripts/gen-google-sheets-live-recalculation-scorecard.ts'),
      implementationPackage: literalField(source, 'implementationPackage', 'packages/headless'),
      evidenceKind: literalField(source, 'evidenceKind', 'live-google-sheets-native-conversion-via-google-drive-connector'),
      captureTransport: literalField(source, 'captureTransport', 'google-drive-connector'),
    },
    benchmark: {
      sampleCount: numberField(benchmark, 'sampleCount'),
      warmupCount: literalField(benchmark, 'warmupCount', 0),
      valueRenderOption: literalField(benchmark, 'valueRenderOption', 'UNFORMATTED_VALUE'),
      measuredGoogleSheetsOperation: literalField(benchmark, 'measuredGoogleSheetsOperation', 'edit-and-read-recalculated-values'),
      measuredWorkpaperOperation: literalField(benchmark, 'measuredWorkpaperOperation', 'mutate-and-recalculate'),
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
      coveredWorkloads: stringArrayField(summary, 'coveredWorkloads').map(parseWorkload),
      microsoftExcelEvidence: literalField(summary, 'microsoftExcelEvidence', 'not-covered-by-this-artifact'),
    },
    cases: arrayField(value, 'cases').map(parseRecalculationCase),
  }
}

export function validateGoogleSheetsLiveRecalculationScorecard(scorecard: GoogleSheetsLiveRecalculationScorecard): void {
  const expectedIds = caseSpecs.map((entry) => entry.id)
  const expectedWorkloads = caseSpecs.map((entry) => entry.workload)
  const spreadsheetSampleKeys = scorecard.googleSheets.spreadsheets.map((entry) => `${entry.caseId}:${String(entry.sampleIndex)}`)
  const expectedSpreadsheetSampleKeys = caseSpecs.flatMap((entry) =>
    Array.from({ length: sampleCount }, (_unused, sampleIndex) => `${entry.id}:${String(sampleIndex)}`),
  )
  if (scorecard.benchmark.sampleCount !== sampleCount || scorecard.benchmark.warmupCount !== warmupCount) {
    throw new Error('Google Sheets live recalculation scorecard benchmark settings are stale')
  }
  if (
    scorecard.summary.requiredCaseCount !== expectedIds.length ||
    JSON.stringify(scorecard.cases.map((entry) => entry.id)) !== JSON.stringify(expectedIds)
  ) {
    throw new Error('Google Sheets live recalculation scorecard required cases are stale')
  }
  if (JSON.stringify(spreadsheetSampleKeys) !== JSON.stringify(expectedSpreadsheetSampleKeys)) {
    throw new Error('Google Sheets live recalculation scorecard spreadsheet evidence is stale')
  }
  for (const spreadsheet of scorecard.googleSheets.spreadsheets) {
    if (
      spreadsheet.spreadsheetId.trim().length === 0 ||
      spreadsheet.spreadsheetUrl.trim().length === 0 ||
      spreadsheet.title.trim().length === 0
    ) {
      throw new Error(`Google Sheets live recalculation spreadsheet evidence is incomplete for ${spreadsheet.caseId}`)
    }
  }
  if (JSON.stringify(scorecard.summary.coveredWorkloads) !== JSON.stringify(expectedWorkloads)) {
    throw new Error('Google Sheets live recalculation scorecard covered workloads are stale')
  }
  if (
    scorecard.summary.workpaperWins !==
    scorecard.cases.filter((entry) => entry.workpaperElapsedMs.mean <= entry.googleSheetsElapsedMs.mean).length
  ) {
    throw new Error('Google Sheets live recalculation scorecard workpaper win count is inconsistent')
  }
  if (scorecard.summary.tenXMeanAndP95CaseCount !== scorecard.cases.filter((entry) => entry.tenXMeanAndP95).length) {
    throw new Error('Google Sheets live recalculation scorecard 10x count is inconsistent')
  }
  const failingCases = scorecard.cases.filter((entry) => !entry.passed)
  if (!scorecard.summary.allRequiredCasesPassed || failingCases.length > 0) {
    throw new Error(
      `Google Sheets live recalculation scorecard has failing required cases: ${failingCases
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
      throw new Error(`Google Sheets live recalculation 10x flag is stale: ${entry.id}`)
    }
    if (!entry.verification.equivalent) {
      throw new Error(`Google Sheets live recalculation scorecard verification mismatch: ${entry.id}`)
    }
    const currentVerification = runWorkPaperSample(entry.workload).verification
    if (JSON.stringify(entry.verification.workpaper) !== JSON.stringify(currentVerification)) {
      throw new Error(`Google Sheets live recalculation WorkPaper verification is stale: ${entry.id}`)
    }
  }
}

function runWorkPaperCaseAgainstCapture(
  capture: GoogleSheetsLiveRecalculationCapture,
  caseSpec: RecalculationCaseSpec,
): GoogleSheetsLiveRecalculationCase {
  const captureCase = requiredCaptureCase(capture, caseSpec)
  const workpaperSamples: number[] = []
  const workpaperVerifications: Array<Record<string, VerificationValue>> = []
  for (let index = 0; index < sampleCount; index += 1) {
    const sample = runWorkPaperSample(caseSpec.workload)
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
    workload: caseSpec.workload,
    fixture: caseSpec.fixture,
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

function runWorkPaperSample(workload: GoogleSheetsLiveRecalculationWorkload): WorkPaperSampleResult {
  const workbook = WorkPaper.buildFromSheets({ [worksheetName]: sheetForWorkload(workload) })
  const sheetId = workbook.getSheetId(worksheetName)
  if (sheetId === undefined) {
    workbook.dispose()
    throw new Error(`Missing WorkPaper recalculation worksheet: ${worksheetName}`)
  }

  const startedAt = performance.now()
  switch (workload) {
    case 'dirty-fanout-edit':
      workbook.setCellContents({ sheet: sheetId, row: 0, col: 0 }, 7)
      break
    case 'suspended-batch-single-column-edit':
      workbook.suspendEvaluation()
      for (let row = 0; row < singleColumnRowCount; row += 1) {
        workbook.setCellContents({ sheet: sheetId, row, col: 0 }, row * 7)
      }
      workbook.resumeEvaluation()
      break
    case 'conditional-aggregation-criteria-edit':
      workbook.setCellContents({ sheet: sheetId, row: 0, col: 3 }, 'B')
      break
    case 'full-rebuild-recalculate':
      workbook.rebuildAndRecalculate()
      break
  }
  const elapsedMs = performance.now() - startedAt
  const verification = workpaperVerification(workbook, sheetId, workload)
  workbook.dispose()
  return { elapsedMs, verification }
}

function createGoogleSheetsWorkbookBytes(workload: GoogleSheetsLiveRecalculationWorkload): Uint8Array {
  const worksheet = aoaToFormulaWorksheet(sheetForWorkload(workload))
  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(workbook, worksheet, worksheetName)
  return XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' })
}

function sheetForWorkload(workload: GoogleSheetsLiveRecalculationWorkload): RecalculationSheet {
  switch (workload) {
    case 'dirty-fanout-edit':
      return [buildFormulaFanoutRow(fanoutCount)]
    case 'suspended-batch-single-column-edit':
      return buildValueFormulaRows(singleColumnRowCount)
    case 'conditional-aggregation-criteria-edit':
      return buildConditionalAggregationSheet(aggregationRowCount, aggregationFormulaCopies)
    case 'full-rebuild-recalculate':
      return buildParserCacheTemplateSheet(rebuildRowCount)
  }
}

function workpaperVerification(
  workbook: WorkPaper,
  sheetId: number,
  workload: GoogleSheetsLiveRecalculationWorkload,
): Record<string, VerificationValue> {
  const dimensions = workbook.getSheetDimensions(sheetId)
  switch (workload) {
    case 'dirty-fanout-edit':
      return {
        terminalValue: normalizeProtocolValue(workbook.getCellValue({ sheet: sheetId, row: 0, col: fanoutCount })),
        width: dimensions.width,
      }
    case 'suspended-batch-single-column-edit':
      return {
        sampleFormulaValue: normalizeProtocolValue(workbook.getCellValue({ sheet: sheetId, row: singleColumnRowCount - 1, col: 1 })),
        height: dimensions.height,
      }
    case 'conditional-aggregation-criteria-edit':
      return {
        sampleSumValue: normalizeProtocolValue(workbook.getCellValue({ sheet: sheetId, row: 0, col: 4 })),
        sampleCountValue: normalizeProtocolValue(workbook.getCellValue({ sheet: sheetId, row: 0, col: 4 + aggregationFormulaCopies })),
      }
    case 'full-rebuild-recalculate':
      return {
        terminalAggregateValue: normalizeProtocolValue(workbook.getCellValue({ sheet: sheetId, row: rebuildRowCount - 1, col: 4 })),
        terminalChainValue: normalizeProtocolValue(workbook.getCellValue({ sheet: sheetId, row: rebuildRowCount - 1, col: 5 })),
      }
  }
}

function googleSheetsEditInstruction(workload: GoogleSheetsLiveRecalculationWorkload): Record<string, VerificationValue> {
  switch (workload) {
    case 'dirty-fanout-edit':
      return { range: 'A1', value: 7 }
    case 'suspended-batch-single-column-edit':
      return { range: `A1:A${String(singleColumnRowCount)}`, valuePattern: 'rowIndexZeroBased * 7' }
    case 'conditional-aggregation-criteria-edit':
      return { range: 'D1', value: 'B' }
    case 'full-rebuild-recalculate':
      return { range: 'A1', value: 1 }
  }
}

function googleSheetsVerificationRanges(workload: GoogleSheetsLiveRecalculationWorkload): string[] {
  switch (workload) {
    case 'dirty-fanout-edit':
      return [`${columnName(fanoutCount)}1`]
    case 'suspended-batch-single-column-edit':
      return [`B${String(singleColumnRowCount)}`]
    case 'conditional-aggregation-criteria-edit':
      return ['E1', `${columnName(4 + aggregationFormulaCopies)}1`]
    case 'full-rebuild-recalculate':
      return [`E${String(rebuildRowCount)}`, `F${String(rebuildRowCount)}`]
  }
}

function aoaToFormulaWorksheet(sheet: RecalculationSheet): XLSX.WorkSheet {
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

function parseRecalculationCase(value: unknown): GoogleSheetsLiveRecalculationCase {
  const record = asObject(value, 'Google Sheets live recalculation case')
  const verification = objectField(record, 'verification')
  return {
    id: stringField(record, 'id'),
    workload: parseWorkload(stringField(record, 'workload')),
    fixture: parseFixture(objectField(record, 'fixture')),
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

function parseCaptureCase(value: unknown): GoogleSheetsLiveRecalculationCaptureCase {
  const record = asObject(value, 'Google Sheets live recalculation capture case')
  return {
    id: stringField(record, 'id'),
    workload: parseWorkload(stringField(record, 'workload')),
    fixture: parseFixture(objectField(record, 'fixture')),
    googleSheetsElapsedMsSamples: numberArrayField(record, 'googleSheetsElapsedMsSamples'),
    googleSheetsVerificationSamples: arrayField(record, 'googleSheetsVerificationSamples').map((entry) =>
      parseVerificationRecord(asObject(entry, 'Google Sheets live recalculation verification sample')),
    ),
  }
}

function parseSpreadsheet(value: unknown): GoogleSheetsLiveRecalculationSpreadsheet {
  const record = asObject(value, 'Google Sheets live recalculation spreadsheet')
  return {
    caseId: stringField(record, 'caseId'),
    sampleIndex: numberField(record, 'sampleIndex'),
    spreadsheetId: stringField(record, 'spreadsheetId'),
    spreadsheetUrl: stringField(record, 'spreadsheetUrl'),
    title: stringField(record, 'title'),
  }
}

function parseFixture(value: Record<string, unknown>): RecalculationFixture {
  return {
    rowCount: numberField(value, 'rowCount'),
    formulaCount: numberField(value, 'formulaCount'),
    materializedCells: numberField(value, 'materializedCells'),
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
    throw new Error(`Unsupported Google Sheets recalculation verification value for ${key}`)
  }
  return result
}

function parseWorkload(value: string): GoogleSheetsLiveRecalculationWorkload {
  if (
    value === 'dirty-fanout-edit' ||
    value === 'suspended-batch-single-column-edit' ||
    value === 'conditional-aggregation-criteria-edit' ||
    value === 'full-rebuild-recalculate'
  ) {
    return value
  }
  throw new Error(`Unexpected Google Sheets live recalculation workload: ${value}`)
}

function requiredCaptureCase(
  capture: GoogleSheetsLiveRecalculationCapture,
  caseSpec: RecalculationCaseSpec,
): GoogleSheetsLiveRecalculationCaptureCase {
  const captureCase = capture.cases.find((entry) => entry.id === caseSpec.id)
  if (!captureCase) {
    throw new Error(`Google Sheets live recalculation capture is missing required case: ${caseSpec.id}`)
  }
  if (
    captureCase.workload !== caseSpec.workload ||
    JSON.stringify(captureCase.fixture) !== JSON.stringify(caseSpec.fixture) ||
    captureCase.googleSheetsElapsedMsSamples.length !== sampleCount ||
    captureCase.googleSheetsVerificationSamples.length !== sampleCount ||
    captureCase.googleSheetsElapsedMsSamples.some((entry) => !Number.isFinite(entry) || entry < 0)
  ) {
    throw new Error(`Google Sheets live recalculation capture case settings are stale: ${caseSpec.id}`)
  }
  return captureCase
}

function requiredSpreadsheet(
  capture: GoogleSheetsLiveRecalculationCapture,
  caseId: string,
  sampleIndex: number,
): GoogleSheetsLiveRecalculationSpreadsheet {
  const spreadsheet = capture.googleSheets.spreadsheets.find((entry) => entry.caseId === caseId && entry.sampleIndex === sampleIndex)
  if (!spreadsheet) {
    throw new Error(`Google Sheets live recalculation capture is missing spreadsheet evidence for ${caseId} sample ${String(sampleIndex)}`)
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
    throw new Error('Unexpected non-protocol WorkPaper value in recalculation benchmark')
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

function columnName(columnIndex: number): string {
  let index = columnIndex + 1
  let name = ''
  while (index > 0) {
    const remainder = (index - 1) % 26
    name = String.fromCharCode(65 + remainder) + name
    index = Math.floor((index - 1) / 26)
  }
  return name
}

function logResult(mode: 'check' | 'write', scorecard: GoogleSheetsLiveRecalculationScorecard): void {
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
