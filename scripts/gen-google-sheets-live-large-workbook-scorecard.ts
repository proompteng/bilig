#!/usr/bin/env bun

import { existsSync, mkdtempSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join, resolve } from 'node:path'
import { pathToFileURL } from 'node:url'
import { performance } from 'node:perf_hooks'

import { SpreadsheetEngine } from '@bilig/core'
import { exportXlsx } from '@bilig/excel-import'
import { ValueTag, type WorkbookSnapshot } from '@bilig/protocol'
import { summarizeNumbers, type NumericSummary } from '../packages/benchmarks/src/stats.js'
import { buildWorkbookBenchmarkCorpus, type WorkbookBenchmarkCorpusId } from '../packages/benchmarks/src/workbook-corpus.js'
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

export type GoogleSheetsLiveLargeWorkbookWorkload = 'native-import-read-dense-mixed-100k' | 'native-import-read-dense-mixed-250k'

type VerificationValue = boolean | number | string | null

export interface GoogleSheetsLiveLargeWorkbookVerification {
  readonly sheetCount: number
  readonly height: number
  readonly width: number
  readonly usedRangeCells: number
  readonly terminalAddress: string
  readonly terminalValue: VerificationValue
}

export interface GoogleSheetsLiveLargeWorkbookCapture {
  readonly schemaVersion: 1
  readonly generatedAt: string
  readonly capture: {
    readonly transport: 'google-drive-connector'
    readonly sourceWorkbook: 'xlsx-native-google-sheets-conversion'
    readonly valueRenderOption: 'UNFORMATTED_VALUE'
    readonly measuredGoogleSheetsOperation: 'native-xlsx-import-and-read-terminal-cell'
    readonly sampleCount: number
    readonly samplingOrder: 'engine-isolated-bilig-then-google-sheets'
  }
  readonly googleSheets: {
    readonly spreadsheets: GoogleSheetsLiveLargeWorkbookSpreadsheet[]
  }
  readonly cases: GoogleSheetsLiveLargeWorkbookCaptureCase[]
}

export interface GoogleSheetsLiveLargeWorkbookSpreadsheet {
  readonly caseId: string
  readonly sampleIndex: number
  readonly spreadsheetId: string
  readonly spreadsheetUrl: string
  readonly title: string
}

export interface GoogleSheetsLiveLargeWorkbookCaptureCase {
  readonly id: string
  readonly workload: GoogleSheetsLiveLargeWorkbookWorkload
  readonly corpusCaseId: WorkbookBenchmarkCorpusId
  readonly sheetName: string
  readonly terminalAddress: string
  readonly googleSheetsElapsedMsSamples: number[]
  readonly googleSheetsVerificationSamples: GoogleSheetsLiveLargeWorkbookVerification[]
}

export interface GoogleSheetsLiveLargeWorkbookCase {
  readonly id: string
  readonly workload: GoogleSheetsLiveLargeWorkbookWorkload
  readonly corpusCaseId: WorkbookBenchmarkCorpusId
  readonly materializedCells: number
  readonly sampleCount: number
  readonly biligElapsedMs: NumericSummary
  readonly googleSheetsElapsedMs: NumericSummary
  readonly biligToGoogleSheetsMeanRatio: number
  readonly biligToGoogleSheetsP95Ratio: number
  readonly tenXMeanAndP95: boolean
  readonly verification: {
    readonly bilig: GoogleSheetsLiveLargeWorkbookVerification
    readonly googleSheets: GoogleSheetsLiveLargeWorkbookVerification
    readonly equivalent: boolean
  }
  readonly passed: boolean
}

export interface GoogleSheetsLiveLargeWorkbookScorecard {
  readonly schemaVersion: 1
  readonly suite: 'google-sheets-live-large-workbook-performance'
  readonly generatedAt: string
  readonly host: {
    readonly arch: string
    readonly platform: string
  }
  readonly source: {
    readonly artifactGenerator: 'scripts/gen-google-sheets-live-large-workbook-scorecard.ts'
    readonly implementationPackage: 'packages/core'
    readonly xlsxExportPackage: 'packages/excel-import'
    readonly corpusPackage: 'packages/benchmarks'
    readonly evidenceKind: 'live-google-sheets-native-conversion-via-google-drive-connector'
    readonly captureTransport: 'google-drive-connector'
  }
  readonly benchmark: {
    readonly sampleCount: number
    readonly valueRenderOption: 'UNFORMATTED_VALUE'
    readonly measuredGoogleSheetsOperation: 'native-xlsx-import-and-read-terminal-cell'
    readonly measuredBiligOperation: 'import-snapshot'
    readonly samplingOrder: 'engine-isolated-bilig-then-google-sheets'
  }
  readonly googleSheets: {
    readonly spreadsheets: GoogleSheetsLiveLargeWorkbookSpreadsheet[]
  }
  readonly summary: {
    readonly allRequiredCasesPassed: boolean
    readonly requiredCaseCount: number
    readonly tenXMeanAndP95CaseCount: number
    readonly biligWins: number
    readonly coveredCorpusCaseIds: WorkbookBenchmarkCorpusId[]
    readonly coveredMaterializedCells: number[]
    readonly microsoftExcelEvidence: 'not-covered-by-this-artifact'
  }
  readonly cases: GoogleSheetsLiveLargeWorkbookCase[]
}

interface LargeWorkbookCaseSpec {
  readonly id: string
  readonly workload: GoogleSheetsLiveLargeWorkbookWorkload
  readonly corpusCaseId: WorkbookBenchmarkCorpusId
  readonly sheetName: string
  readonly terminalAddress: string
}

interface BiligSampleResult {
  readonly elapsedMs: number
  readonly verification: GoogleSheetsLiveLargeWorkbookVerification
}

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const outputPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'google-sheets-live-large-workbook-scorecard.json')
const sampleCount = 3
const TEN_X_RATIO = 0.1
const caseSpecs = [
  {
    id: 'google-sheets-live-large-workbook-import-read-dense-mixed-100k',
    workload: 'native-import-read-dense-mixed-100k',
    corpusCaseId: 'dense-mixed-100k',
    sheetName: 'Grid',
    terminalAddress: 'C25000',
  },
  {
    id: 'google-sheets-live-large-workbook-import-read-dense-mixed-250k',
    workload: 'native-import-read-dense-mixed-250k',
    corpusCaseId: 'dense-mixed-250k',
    sheetName: 'Grid',
    terminalAddress: 'C62500',
  },
] as const satisfies readonly LargeWorkbookCaseSpec[]

async function main(): Promise<void> {
  const isCheckMode = process.argv.includes('--check')
  const emitXlsxIndex = process.argv.indexOf('--emit-xlsx')
  const captureIndex = process.argv.indexOf('--capture')

  if (emitXlsxIndex >= 0) {
    const targetDirectory = process.argv[emitXlsxIndex + 1]
    if (!targetDirectory) {
      throw new Error('Missing directory after --emit-xlsx')
    }
    emitGoogleSheetsLargeWorkbookXlsx(resolve(targetDirectory))
    return
  }

  if (isCheckMode) {
    if (!existsSync(outputPath)) {
      throw new Error(
        `Google Sheets live large-workbook scorecard is missing. Run: bun scripts/gen-google-sheets-live-large-workbook-scorecard.ts --capture <capture.json>`,
      )
    }
    const scorecard = parseGoogleSheetsLiveLargeWorkbookScorecard(readJsonObject(outputPath))
    validateGoogleSheetsLiveLargeWorkbookScorecard(scorecard)
    logResult('check', scorecard)
    return
  }

  if (captureIndex < 0) {
    throw new Error(
      'Missing --capture <capture.json>. First emit XLSX files with --emit-xlsx, import them as native Google Sheets, then capture import/read timings and UNFORMATTED_VALUE terminal cells through the Google Drive connector.',
    )
  }
  const capturePath = process.argv[captureIndex + 1]
  if (!capturePath) {
    throw new Error('Missing path after --capture')
  }
  const scorecard = await buildGoogleSheetsLiveLargeWorkbookScorecard(
    parseGoogleSheetsLiveLargeWorkbookCapture(readJsonObject(capturePath)),
  )
  validateGoogleSheetsLiveLargeWorkbookScorecard(scorecard)
  mkdirSync(dirname(outputPath), { recursive: true })
  writeFileSync(outputPath, formatJsonForRepo(`${JSON.stringify(scorecard, null, 2)}\n`))
  logResult('write', scorecard)
}

export function emitGoogleSheetsLargeWorkbookXlsx(targetDirectory: string): void {
  mkdirSync(targetDirectory, { recursive: true })
  const outputs = caseSpecs.map((caseSpec) => {
    const corpus = buildWorkbookBenchmarkCorpus(caseSpec.corpusCaseId)
    const outputFile = join(targetDirectory, `${caseSpec.corpusCaseId}.xlsx`)
    writeFileSync(outputFile, exportXlsx(corpus.snapshot))
    return {
      caseId: caseSpec.id,
      corpusCaseId: caseSpec.corpusCaseId,
      materializedCells: corpus.materializedCellCount,
      outputFile,
      sheetName: caseSpec.sheetName,
      terminalAddress: caseSpec.terminalAddress,
    }
  })
  console.log(
    JSON.stringify(
      {
        mode: 'emit-xlsx',
        targetDirectory,
        uploadMode: 'native_google_sheets',
        valueRenderOption: 'UNFORMATTED_VALUE',
        measuredGoogleSheetsOperation: 'native-xlsx-import-and-read-terminal-cell',
        outputs,
      },
      null,
      2,
    ),
  )
}

export async function buildGoogleSheetsLiveLargeWorkbookScorecard(
  capture: GoogleSheetsLiveLargeWorkbookCapture,
): Promise<GoogleSheetsLiveLargeWorkbookScorecard> {
  if (capture.capture.sampleCount !== sampleCount) {
    throw new Error(`Google Sheets live large-workbook capture sample count must be ${String(sampleCount)}`)
  }

  const cases = await runBiligCasesAgainstCapture(capture, [...caseSpecs])
  const tenXMeanAndP95CaseCount = cases.filter((entry) => entry.tenXMeanAndP95).length
  const biligWins = cases.filter((entry) => entry.biligElapsedMs.mean <= entry.googleSheetsElapsedMs.mean).length

  return {
    schemaVersion: 1,
    suite: 'google-sheets-live-large-workbook-performance',
    generatedAt: capture.generatedAt,
    host: {
      arch: process.arch,
      platform: process.platform,
    },
    source: {
      artifactGenerator: 'scripts/gen-google-sheets-live-large-workbook-scorecard.ts',
      implementationPackage: 'packages/core',
      xlsxExportPackage: 'packages/excel-import',
      corpusPackage: 'packages/benchmarks',
      evidenceKind: 'live-google-sheets-native-conversion-via-google-drive-connector',
      captureTransport: capture.capture.transport,
    },
    benchmark: {
      sampleCount,
      valueRenderOption: capture.capture.valueRenderOption,
      measuredGoogleSheetsOperation: capture.capture.measuredGoogleSheetsOperation,
      measuredBiligOperation: 'import-snapshot',
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
      biligWins,
      coveredCorpusCaseIds: caseSpecs.map((entry) => entry.corpusCaseId),
      coveredMaterializedCells: cases.map((entry) => entry.materializedCells),
      microsoftExcelEvidence: 'not-covered-by-this-artifact',
    },
    cases,
  }
}

export function parseGoogleSheetsLiveLargeWorkbookCapture(value: Record<string, unknown>): GoogleSheetsLiveLargeWorkbookCapture {
  const capture = objectField(value, 'capture')
  const googleSheets = objectField(value, 'googleSheets')
  return {
    schemaVersion: literalField(value, 'schemaVersion', 1),
    generatedAt: stringField(value, 'generatedAt'),
    capture: {
      transport: literalField(capture, 'transport', 'google-drive-connector'),
      sourceWorkbook: literalField(capture, 'sourceWorkbook', 'xlsx-native-google-sheets-conversion'),
      valueRenderOption: literalField(capture, 'valueRenderOption', 'UNFORMATTED_VALUE'),
      measuredGoogleSheetsOperation: literalField(capture, 'measuredGoogleSheetsOperation', 'native-xlsx-import-and-read-terminal-cell'),
      sampleCount: numberField(capture, 'sampleCount'),
      samplingOrder: literalField(capture, 'samplingOrder', 'engine-isolated-bilig-then-google-sheets'),
    },
    googleSheets: {
      spreadsheets: arrayField(googleSheets, 'spreadsheets').map(parseSpreadsheet),
    },
    cases: arrayField(value, 'cases').map(parseCaptureCase),
  }
}

export function parseGoogleSheetsLiveLargeWorkbookScorecard(value: Record<string, unknown>): GoogleSheetsLiveLargeWorkbookScorecard {
  const host = objectField(value, 'host')
  const source = objectField(value, 'source')
  const benchmark = objectField(value, 'benchmark')
  const googleSheets = objectField(value, 'googleSheets')
  const summary = objectField(value, 'summary')
  return {
    schemaVersion: literalField(value, 'schemaVersion', 1),
    suite: literalField(value, 'suite', 'google-sheets-live-large-workbook-performance'),
    generatedAt: stringField(value, 'generatedAt'),
    host: {
      arch: stringField(host, 'arch'),
      platform: stringField(host, 'platform'),
    },
    source: {
      artifactGenerator: literalField(source, 'artifactGenerator', 'scripts/gen-google-sheets-live-large-workbook-scorecard.ts'),
      implementationPackage: literalField(source, 'implementationPackage', 'packages/core'),
      xlsxExportPackage: literalField(source, 'xlsxExportPackage', 'packages/excel-import'),
      corpusPackage: literalField(source, 'corpusPackage', 'packages/benchmarks'),
      evidenceKind: literalField(source, 'evidenceKind', 'live-google-sheets-native-conversion-via-google-drive-connector'),
      captureTransport: literalField(source, 'captureTransport', 'google-drive-connector'),
    },
    benchmark: {
      sampleCount: numberField(benchmark, 'sampleCount'),
      valueRenderOption: literalField(benchmark, 'valueRenderOption', 'UNFORMATTED_VALUE'),
      measuredGoogleSheetsOperation: literalField(benchmark, 'measuredGoogleSheetsOperation', 'native-xlsx-import-and-read-terminal-cell'),
      measuredBiligOperation: literalField(benchmark, 'measuredBiligOperation', 'import-snapshot'),
      samplingOrder: literalField(benchmark, 'samplingOrder', 'engine-isolated-bilig-then-google-sheets'),
    },
    googleSheets: {
      spreadsheets: arrayField(googleSheets, 'spreadsheets').map(parseSpreadsheet),
    },
    summary: {
      allRequiredCasesPassed: booleanField(summary, 'allRequiredCasesPassed'),
      requiredCaseCount: numberField(summary, 'requiredCaseCount'),
      tenXMeanAndP95CaseCount: numberField(summary, 'tenXMeanAndP95CaseCount'),
      biligWins: numberField(summary, 'biligWins'),
      coveredCorpusCaseIds: stringArrayField(summary, 'coveredCorpusCaseIds').map(parseWorkbookBenchmarkCorpusId),
      coveredMaterializedCells: arrayField(summary, 'coveredMaterializedCells').map((entry) => {
        if (typeof entry !== 'number' || !Number.isFinite(entry)) {
          throw new Error('Google Sheets live large-workbook coveredMaterializedCells must be finite numbers')
        }
        return entry
      }),
      microsoftExcelEvidence: literalField(summary, 'microsoftExcelEvidence', 'not-covered-by-this-artifact'),
    },
    cases: arrayField(value, 'cases').map(parseLargeWorkbookCase),
  }
}

export function validateGoogleSheetsLiveLargeWorkbookScorecard(scorecard: GoogleSheetsLiveLargeWorkbookScorecard): void {
  const expectedIds = caseSpecs.map((entry) => entry.id)
  const expectedCorpusCaseIds = caseSpecs.map((entry) => entry.corpusCaseId)
  const expectedMaterializedCells = caseSpecs.map((entry) => buildWorkbookBenchmarkCorpus(entry.corpusCaseId).materializedCellCount)
  const spreadsheetSampleKeys = scorecard.googleSheets.spreadsheets.map((entry) => `${entry.caseId}:${String(entry.sampleIndex)}`)
  const expectedSpreadsheetSampleKeys = caseSpecs.flatMap((entry) =>
    Array.from({ length: sampleCount }, (_unused, sampleIndex) => `${entry.id}:${String(sampleIndex)}`),
  )
  if (scorecard.benchmark.sampleCount !== sampleCount) {
    throw new Error('Google Sheets live large-workbook scorecard benchmark settings are stale')
  }
  if (
    scorecard.summary.requiredCaseCount !== expectedIds.length ||
    JSON.stringify(scorecard.cases.map((entry) => entry.id)) !== JSON.stringify(expectedIds)
  ) {
    throw new Error('Google Sheets live large-workbook scorecard required cases are stale')
  }
  if (JSON.stringify(spreadsheetSampleKeys) !== JSON.stringify(expectedSpreadsheetSampleKeys)) {
    throw new Error('Google Sheets live large-workbook scorecard spreadsheet evidence is stale')
  }
  for (const spreadsheet of scorecard.googleSheets.spreadsheets) {
    if (
      spreadsheet.spreadsheetId.trim().length === 0 ||
      spreadsheet.spreadsheetUrl.trim().length === 0 ||
      spreadsheet.title.trim().length === 0
    ) {
      throw new Error(`Google Sheets live large-workbook spreadsheet evidence is incomplete for ${spreadsheet.caseId}`)
    }
  }
  if (JSON.stringify(scorecard.summary.coveredCorpusCaseIds) !== JSON.stringify(expectedCorpusCaseIds)) {
    throw new Error('Google Sheets live large-workbook scorecard covered corpus cases are stale')
  }
  if (JSON.stringify(scorecard.summary.coveredMaterializedCells) !== JSON.stringify(expectedMaterializedCells)) {
    throw new Error('Google Sheets live large-workbook scorecard materialized-cell coverage is stale')
  }
  if (
    scorecard.summary.biligWins !== scorecard.cases.filter((entry) => entry.biligElapsedMs.mean <= entry.googleSheetsElapsedMs.mean).length
  ) {
    throw new Error('Google Sheets live large-workbook scorecard Bilig win count is inconsistent')
  }
  if (scorecard.summary.tenXMeanAndP95CaseCount !== scorecard.cases.filter((entry) => entry.tenXMeanAndP95).length) {
    throw new Error('Google Sheets live large-workbook scorecard 10x count is inconsistent')
  }
  const failingCases = scorecard.cases.filter((entry) => !entry.passed)
  if (!scorecard.summary.allRequiredCasesPassed || failingCases.length > 0) {
    throw new Error(
      `Google Sheets live large-workbook scorecard has failing required cases: ${failingCases
        .map(
          (entry) =>
            `${entry.id} Bilig=${JSON.stringify(entry.verification.bilig)} GoogleSheets=${JSON.stringify(entry.verification.googleSheets)}`,
        )
        .join(', ')}`,
    )
  }
  for (const entry of scorecard.cases) {
    validateNumericSummary(entry.biligElapsedMs, `${entry.id} Bilig elapsedMs`)
    validateNumericSummary(entry.googleSheetsElapsedMs, `${entry.id} Google Sheets elapsedMs`)
    const expectedTenXMeanAndP95 = entry.biligToGoogleSheetsMeanRatio <= TEN_X_RATIO && entry.biligToGoogleSheetsP95Ratio <= TEN_X_RATIO
    if (entry.tenXMeanAndP95 !== expectedTenXMeanAndP95) {
      throw new Error(`Google Sheets live large-workbook 10x flag is stale: ${entry.id}`)
    }
    if (!entry.verification.equivalent) {
      throw new Error(`Google Sheets live large-workbook scorecard verification mismatch: ${entry.id}`)
    }
  }
}

async function runBiligCasesAgainstCapture(
  capture: GoogleSheetsLiveLargeWorkbookCapture,
  caseSpecsToRun: readonly LargeWorkbookCaseSpec[],
): Promise<GoogleSheetsLiveLargeWorkbookCase[]> {
  if (caseSpecsToRun.length === 0) {
    return []
  }
  const [firstCaseSpec, ...remainingCaseSpecs] = caseSpecsToRun
  if (!firstCaseSpec) {
    return []
  }
  return [await runBiligCaseAgainstCapture(capture, firstCaseSpec), ...(await runBiligCasesAgainstCapture(capture, remainingCaseSpecs))]
}

async function runBiligCaseAgainstCapture(
  capture: GoogleSheetsLiveLargeWorkbookCapture,
  caseSpec: LargeWorkbookCaseSpec,
): Promise<GoogleSheetsLiveLargeWorkbookCase> {
  const corpus = buildWorkbookBenchmarkCorpus(caseSpec.corpusCaseId)
  const captureCase = requiredCaptureCase(capture, caseSpec)
  const biligSamples: number[] = []
  const biligVerifications: GoogleSheetsLiveLargeWorkbookVerification[] = []

  const collectBiligSamples = async (index: number): Promise<void> => {
    if (index >= sampleCount) return
    const sample = await runBiligSample(caseSpec, corpus.snapshot)
    biligSamples.push(sample.elapsedMs)
    biligVerifications.push(sample.verification)
    await collectBiligSamples(index + 1)
  }
  await collectBiligSamples(0)

  const biligElapsedMs = summarizeNumbers(biligSamples)
  const googleSheetsElapsedMs = summarizeNumbers(captureCase.googleSheetsElapsedMsSamples)
  const biligToGoogleSheetsMeanRatio = biligElapsedMs.mean / googleSheetsElapsedMs.mean
  const biligToGoogleSheetsP95Ratio = biligElapsedMs.p95 / googleSheetsElapsedMs.p95
  const verification = {
    bilig: biligVerifications[0] ?? failMissingVerification('Bilig'),
    googleSheets: captureCase.googleSheetsVerificationSamples[0] ?? failMissingVerification('Google Sheets'),
    equivalent: verificationsEquivalent(biligVerifications, captureCase.googleSheetsVerificationSamples),
  }

  return {
    id: caseSpec.id,
    workload: caseSpec.workload,
    corpusCaseId: caseSpec.corpusCaseId,
    materializedCells: corpus.materializedCellCount,
    sampleCount,
    biligElapsedMs,
    googleSheetsElapsedMs,
    biligToGoogleSheetsMeanRatio,
    biligToGoogleSheetsP95Ratio,
    tenXMeanAndP95: biligToGoogleSheetsMeanRatio <= TEN_X_RATIO && biligToGoogleSheetsP95Ratio <= TEN_X_RATIO,
    verification,
    passed:
      verification.equivalent && biligElapsedMs.samples.length === sampleCount && googleSheetsElapsedMs.samples.length === sampleCount,
  }
}

async function runBiligSample(caseSpec: LargeWorkbookCaseSpec, snapshot: WorkbookSnapshot): Promise<BiligSampleResult> {
  const engine = new SpreadsheetEngine({ workbookName: snapshot.workbook.name })
  await engine.ready()
  const startedAt = performance.now()
  engine.importSnapshot(snapshot)
  const elapsedMs = performance.now() - startedAt
  return {
    elapsedMs,
    verification: biligVerification(engine, caseSpec, snapshot),
  }
}

function biligVerification(
  engine: SpreadsheetEngine,
  caseSpec: LargeWorkbookCaseSpec,
  snapshot: WorkbookSnapshot,
): GoogleSheetsLiveLargeWorkbookVerification {
  const sheet = requiredSnapshotSheet(snapshot, caseSpec.sheetName)
  const bounds = sheetBounds(sheet)
  return {
    sheetCount: snapshot.sheets.length,
    height: bounds.height,
    width: bounds.width,
    usedRangeCells: bounds.height * bounds.width,
    terminalAddress: caseSpec.terminalAddress,
    terminalValue: normalizeProtocolValue(engine.getCellValue(caseSpec.sheetName, caseSpec.terminalAddress)),
  }
}

function parseLargeWorkbookCase(value: unknown): GoogleSheetsLiveLargeWorkbookCase {
  const record = asObject(value, 'Google Sheets live large-workbook case')
  const verification = objectField(record, 'verification')
  return {
    id: stringField(record, 'id'),
    workload: parseWorkload(stringField(record, 'workload')),
    corpusCaseId: parseWorkbookBenchmarkCorpusId(stringField(record, 'corpusCaseId')),
    materializedCells: numberField(record, 'materializedCells'),
    sampleCount: numberField(record, 'sampleCount'),
    biligElapsedMs: parseNumericSummary(objectField(record, 'biligElapsedMs')),
    googleSheetsElapsedMs: parseNumericSummary(objectField(record, 'googleSheetsElapsedMs')),
    biligToGoogleSheetsMeanRatio: numberField(record, 'biligToGoogleSheetsMeanRatio'),
    biligToGoogleSheetsP95Ratio: numberField(record, 'biligToGoogleSheetsP95Ratio'),
    tenXMeanAndP95: booleanField(record, 'tenXMeanAndP95'),
    verification: {
      bilig: parseVerification(objectField(verification, 'bilig')),
      googleSheets: parseVerification(objectField(verification, 'googleSheets')),
      equivalent: booleanField(verification, 'equivalent'),
    },
    passed: booleanField(record, 'passed'),
  }
}

function parseCaptureCase(value: unknown): GoogleSheetsLiveLargeWorkbookCaptureCase {
  const record = asObject(value, 'Google Sheets live large-workbook capture case')
  return {
    id: stringField(record, 'id'),
    workload: parseWorkload(stringField(record, 'workload')),
    corpusCaseId: parseWorkbookBenchmarkCorpusId(stringField(record, 'corpusCaseId')),
    sheetName: stringField(record, 'sheetName'),
    terminalAddress: stringField(record, 'terminalAddress'),
    googleSheetsElapsedMsSamples: numberArrayField(record, 'googleSheetsElapsedMsSamples'),
    googleSheetsVerificationSamples: arrayField(record, 'googleSheetsVerificationSamples').map((entry) =>
      parseVerification(asObject(entry, 'Google Sheets live large-workbook verification sample')),
    ),
  }
}

function parseSpreadsheet(value: unknown): GoogleSheetsLiveLargeWorkbookSpreadsheet {
  const record = asObject(value, 'Google Sheets live large-workbook spreadsheet')
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

function parseVerification(value: Record<string, unknown>): GoogleSheetsLiveLargeWorkbookVerification {
  return {
    sheetCount: numberField(value, 'sheetCount'),
    height: numberField(value, 'height'),
    width: numberField(value, 'width'),
    usedRangeCells: numberField(value, 'usedRangeCells'),
    terminalAddress: stringField(value, 'terminalAddress'),
    terminalValue: parseVerificationValue(value['terminalValue']),
  }
}

function parseVerificationValue(value: unknown): VerificationValue {
  if (value === null || typeof value === 'boolean' || typeof value === 'number' || typeof value === 'string') {
    return value
  }
  throw new Error('Unsupported Google Sheets live large-workbook verification value')
}

function parseWorkload(value: string): GoogleSheetsLiveLargeWorkbookWorkload {
  if (value === 'native-import-read-dense-mixed-100k' || value === 'native-import-read-dense-mixed-250k') {
    return value
  }
  throw new Error(`Unexpected Google Sheets live large-workbook workload: ${value}`)
}

function parseWorkbookBenchmarkCorpusId(value: string): WorkbookBenchmarkCorpusId {
  if (
    value === 'dense-mixed-100k' ||
    value === 'dense-mixed-250k' ||
    value === 'wide-mixed-250k' ||
    value === 'wide-mixed-frozen-250k' ||
    value === 'wide-mixed-variable-250k' ||
    value === 'analysis-multisheet-100k' ||
    value === 'analysis-multisheet-250k'
  ) {
    return value
  }
  throw new Error(`Unexpected workbook benchmark corpus id: ${value}`)
}

function requiredCaptureCase(
  capture: GoogleSheetsLiveLargeWorkbookCapture,
  caseSpec: LargeWorkbookCaseSpec,
): GoogleSheetsLiveLargeWorkbookCaptureCase {
  const captureCase = capture.cases.find((entry) => entry.id === caseSpec.id)
  if (!captureCase) {
    throw new Error(`Google Sheets live large-workbook capture is missing required case: ${caseSpec.id}`)
  }
  if (
    captureCase.workload !== caseSpec.workload ||
    captureCase.corpusCaseId !== caseSpec.corpusCaseId ||
    captureCase.sheetName !== caseSpec.sheetName ||
    captureCase.terminalAddress !== caseSpec.terminalAddress
  ) {
    throw new Error(`Google Sheets live large-workbook capture case settings are stale: ${caseSpec.id}`)
  }
  if (
    captureCase.googleSheetsElapsedMsSamples.length !== sampleCount ||
    captureCase.googleSheetsVerificationSamples.length !== sampleCount ||
    captureCase.googleSheetsElapsedMsSamples.some((entry) => !Number.isFinite(entry) || entry < 0)
  ) {
    throw new Error(`Google Sheets live large-workbook capture samples are incomplete for ${caseSpec.id}`)
  }
  return captureCase
}

function requiredSpreadsheet(
  capture: GoogleSheetsLiveLargeWorkbookCapture,
  caseId: string,
  sampleIndex: number,
): GoogleSheetsLiveLargeWorkbookSpreadsheet {
  const spreadsheet = capture.googleSheets.spreadsheets.find((entry) => entry.caseId === caseId && entry.sampleIndex === sampleIndex)
  if (!spreadsheet) {
    throw new Error(`Google Sheets live large-workbook capture is missing spreadsheet evidence for ${caseId} sample ${String(sampleIndex)}`)
  }
  return spreadsheet
}

function verificationsEquivalent(
  biligVerifications: readonly GoogleSheetsLiveLargeWorkbookVerification[],
  googleSheetsVerifications: readonly GoogleSheetsLiveLargeWorkbookVerification[],
): boolean {
  if (biligVerifications.length !== googleSheetsVerifications.length) {
    return false
  }
  const expected = biligVerifications[0]
  if (expected === undefined) {
    return false
  }
  return (
    biligVerifications.every((entry) => JSON.stringify(entry) === JSON.stringify(expected)) &&
    googleSheetsVerifications.every((entry) => JSON.stringify(entry) === JSON.stringify(expected))
  )
}

function failMissingVerification(label: string): never {
  throw new Error(`${label} verification sample is missing`)
}

function requiredSnapshotSheet(snapshot: WorkbookSnapshot, sheetName: string): WorkbookSnapshot['sheets'][number] {
  const sheet = snapshot.sheets.find((entry) => entry.name === sheetName)
  if (!sheet) {
    throw new Error(`Missing large-workbook benchmark sheet: ${sheetName}`)
  }
  return sheet
}

function sheetBounds(sheet: WorkbookSnapshot['sheets'][number]): { height: number; width: number } {
  let maxRow = -1
  let maxCol = -1
  for (const cell of sheet.cells) {
    const address = decodeA1Address(cell.address)
    maxRow = Math.max(maxRow, address.row)
    maxCol = Math.max(maxCol, address.col)
  }
  return {
    height: maxRow + 1,
    width: maxCol + 1,
  }
}

function decodeA1Address(address: string): { row: number; col: number } {
  const match = /^([A-Z]+)(\d+)$/u.exec(address)
  if (!match) {
    throw new Error(`Invalid A1 address in large-workbook benchmark: ${address}`)
  }
  const columnName = match[1] ?? ''
  const rowText = match[2] ?? ''
  let col = 0
  for (const character of columnName) {
    col = col * 26 + (character.charCodeAt(0) - 64)
  }
  return {
    row: Number.parseInt(rowText, 10) - 1,
    col: col - 1,
  }
}

function normalizeProtocolValue(value: unknown): VerificationValue {
  if (!isProtocolValueLike(value)) {
    throw new Error('Unexpected non-protocol Bilig value in large-workbook benchmark')
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

function logResult(mode: 'check' | 'write', scorecard: GoogleSheetsLiveLargeWorkbookScorecard): void {
  console.log(
    JSON.stringify(
      {
        mode,
        outputPath,
        allRequiredCasesPassed: scorecard.summary.allRequiredCasesPassed,
        biligWins: scorecard.summary.biligWins,
        tenXMeanAndP95CaseCount: scorecard.summary.tenXMeanAndP95CaseCount,
        spreadsheetIds: scorecard.googleSheets.spreadsheets.map((entry) => entry.spreadsheetId),
      },
      null,
      2,
    ),
  )
}

function formatJsonForRepo(serializedJson: string): string {
  const tempDir = mkdtempSync(join(tmpdir(), 'google-sheets-live-large-workbook-scorecard-'))
  const tempFilePath = join(tempDir, 'scorecard.json')
  writeFileSync(tempFilePath, serializedJson)
  const oxfmtPath = join(rootDir, 'node_modules', '.bin', 'oxfmt')

  const formatResult = Bun.spawnSync([oxfmtPath, '--write', tempFilePath], {
    stdin: 'ignore',
    stdout: 'pipe',
    stderr: 'pipe',
  })
  if (formatResult.exitCode !== 0) {
    rmSync(tempDir, { recursive: true, force: true })
    throw new Error(
      `Unable to format generated Google Sheets live large-workbook scorecard: ${new TextDecoder().decode(formatResult.stderr).trim()}`,
    )
  }

  const formattedJson = readFileSync(tempFilePath, 'utf8')
  rmSync(tempDir, { recursive: true, force: true })
  return formattedJson
}

if (process.argv[1] && pathToFileURL(resolve(process.argv[1])).href === import.meta.url) {
  await main()
}
