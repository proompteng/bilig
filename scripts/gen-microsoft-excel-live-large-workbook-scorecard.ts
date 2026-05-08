#!/usr/bin/env bun

import { execFileSync } from 'node:child_process'
import { existsSync, mkdtempSync, mkdirSync, rmSync, writeFileSync } from 'node:fs'
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
  numberField,
  objectField,
  readJsonObject,
  stringArrayField,
  stringField,
} from './json-scorecard-helpers.ts'
import { formatJsonForRepo } from './scorecard-format.ts'

export type MicrosoftExcelLiveLargeWorkbookWorkload = 'open-calculate-dense-mixed-100k' | 'open-calculate-dense-mixed-250k'

export interface MicrosoftExcelLiveLargeWorkbookCase {
  readonly id: string
  readonly workload: MicrosoftExcelLiveLargeWorkbookWorkload
  readonly corpusCaseId: WorkbookBenchmarkCorpusId
  readonly materializedCells: number
  readonly sampleCount: number
  readonly biligElapsedMs: NumericSummary
  readonly microsoftExcelElapsedMs: NumericSummary
  readonly biligToMicrosoftExcelMeanRatio: number
  readonly biligToMicrosoftExcelP95Ratio: number
  readonly tenXMeanAndP95: boolean
  readonly verification: {
    readonly bilig: Record<string, boolean | number | string | null>
    readonly microsoftExcel: Record<string, boolean | number | string | null>
    readonly equivalent: boolean
  }
  readonly passed: boolean
}

export interface MicrosoftExcelLiveLargeWorkbookScorecard {
  readonly schemaVersion: 1
  readonly suite: 'microsoft-excel-live-large-workbook-performance'
  readonly generatedAt: string
  readonly host: {
    readonly arch: string
    readonly platform: string
  }
  readonly source: {
    readonly artifactGenerator: 'scripts/gen-microsoft-excel-live-large-workbook-scorecard.ts'
    readonly implementationPackage: 'packages/core'
    readonly xlsxExportPackage: 'packages/excel-import'
    readonly corpusPackage: 'packages/benchmarks'
    readonly evidenceKind: 'live-local-microsoft-excel-automation'
    readonly appleScriptTransport: 'osascript'
  }
  readonly benchmark: {
    readonly sampleCount: number
    readonly screenUpdating: false
    readonly calculationMode: 'manual-during-open-and-calculate'
    readonly measuredExcelOperation: 'open-workbook-and-calculate-full-rebuild'
    readonly measuredBiligOperation: 'import-snapshot'
    readonly samplingOrder: 'engine-isolated-bilig-then-excel'
  }
  readonly microsoftExcel: {
    readonly appPath: '/Applications/Microsoft Excel.app'
    readonly version: string
  }
  readonly summary: {
    readonly allRequiredCasesPassed: boolean
    readonly requiredCaseCount: number
    readonly tenXMeanAndP95CaseCount: number
    readonly biligWins: number
    readonly coveredCorpusCaseIds: WorkbookBenchmarkCorpusId[]
    readonly coveredMaterializedCells: number[]
    readonly googleSheetsEvidence: 'not-covered-by-this-artifact'
  }
  readonly cases: MicrosoftExcelLiveLargeWorkbookCase[]
}

interface LargeWorkbookCaseSpec {
  readonly id: string
  readonly workload: MicrosoftExcelLiveLargeWorkbookWorkload
  readonly corpusCaseId: WorkbookBenchmarkCorpusId
  readonly sheetName: string
  readonly terminalAddress: string
}

interface SampleResult {
  readonly elapsedMs: number
  readonly verification: Record<string, boolean | number | string | null>
}

interface ExcelSampleResult extends SampleResult {
  readonly excelVersion: string
}

interface LargeWorkbookCaseRun {
  readonly caseResult: MicrosoftExcelLiveLargeWorkbookCase
  readonly excelVersion: string
}

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const outputPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'microsoft-excel-live-large-workbook-scorecard.json')
const excelAppPath = '/Applications/Microsoft Excel.app' as const
const sampleCount = 3
const TEN_X_RATIO = 0.1
const caseSpecs = [
  {
    id: 'excel-live-large-workbook-open-calculate-dense-mixed-100k',
    workload: 'open-calculate-dense-mixed-100k',
    corpusCaseId: 'dense-mixed-100k',
    sheetName: 'Grid',
    terminalAddress: 'C25000',
  },
  {
    id: 'excel-live-large-workbook-open-calculate-dense-mixed-250k',
    workload: 'open-calculate-dense-mixed-250k',
    corpusCaseId: 'dense-mixed-250k',
    sheetName: 'Grid',
    terminalAddress: 'C62500',
  },
] as const satisfies readonly LargeWorkbookCaseSpec[]

async function main(): Promise<void> {
  const isCheckMode = process.argv.includes('--check')
  if (isCheckMode) {
    if (!existsSync(outputPath)) {
      throw new Error(
        `Microsoft Excel live large-workbook scorecard is missing. Run: bun scripts/gen-microsoft-excel-live-large-workbook-scorecard.ts`,
      )
    }
    const scorecard = parseMicrosoftExcelLiveLargeWorkbookScorecard(readJsonObject(outputPath))
    validateMicrosoftExcelLiveLargeWorkbookScorecard(scorecard)
    logResult('check', scorecard)
    return
  }

  const scorecard = await buildMicrosoftExcelLiveLargeWorkbookScorecard(new Date().toISOString())
  validateMicrosoftExcelLiveLargeWorkbookScorecard(scorecard)
  mkdirSync(dirname(outputPath), { recursive: true })
  writeFileSync(outputPath, formatJsonForRepo(`${JSON.stringify(scorecard, null, 2)}\n`))
  logResult('write', scorecard)
}

export async function buildMicrosoftExcelLiveLargeWorkbookScorecard(
  generatedAt: string,
): Promise<MicrosoftExcelLiveLargeWorkbookScorecard> {
  if (!existsSync(excelAppPath)) {
    throw new Error(`Microsoft Excel app is not installed at ${excelAppPath}`)
  }

  const caseRuns = await runLargeWorkbookCasesSequentially([...caseSpecs])
  const cases = caseRuns.map((entry) => entry.caseResult)
  const excelVersions = new Set(caseRuns.map((entry) => entry.excelVersion))
  if (excelVersions.size !== 1) {
    throw new Error(`Microsoft Excel version changed during large-workbook benchmark: ${[...excelVersions].join(', ')}`)
  }
  const excelVersion = [...excelVersions][0] ?? ''
  const tenXMeanAndP95CaseCount = cases.filter((entry) => entry.tenXMeanAndP95).length
  const biligWins = cases.filter((entry) => entry.biligElapsedMs.mean <= entry.microsoftExcelElapsedMs.mean).length

  return {
    schemaVersion: 1,
    suite: 'microsoft-excel-live-large-workbook-performance',
    generatedAt,
    host: {
      arch: process.arch,
      platform: process.platform,
    },
    source: {
      artifactGenerator: 'scripts/gen-microsoft-excel-live-large-workbook-scorecard.ts',
      implementationPackage: 'packages/core',
      xlsxExportPackage: 'packages/excel-import',
      corpusPackage: 'packages/benchmarks',
      evidenceKind: 'live-local-microsoft-excel-automation',
      appleScriptTransport: 'osascript',
    },
    benchmark: {
      sampleCount,
      screenUpdating: false,
      calculationMode: 'manual-during-open-and-calculate',
      measuredExcelOperation: 'open-workbook-and-calculate-full-rebuild',
      measuredBiligOperation: 'import-snapshot',
      samplingOrder: 'engine-isolated-bilig-then-excel',
    },
    microsoftExcel: {
      appPath: excelAppPath,
      version: excelVersion,
    },
    summary: {
      allRequiredCasesPassed: cases.every((entry) => entry.passed),
      requiredCaseCount: cases.length,
      tenXMeanAndP95CaseCount,
      biligWins,
      coveredCorpusCaseIds: caseSpecs.map((entry) => entry.corpusCaseId),
      coveredMaterializedCells: cases.map((entry) => entry.materializedCells),
      googleSheetsEvidence: 'not-covered-by-this-artifact',
    },
    cases,
  }
}

export function parseMicrosoftExcelLiveLargeWorkbookScorecard(value: Record<string, unknown>): MicrosoftExcelLiveLargeWorkbookScorecard {
  const host = objectField(value, 'host')
  const source = objectField(value, 'source')
  const benchmark = objectField(value, 'benchmark')
  const microsoftExcel = objectField(value, 'microsoftExcel')
  const summary = objectField(value, 'summary')
  return {
    schemaVersion: literalField(value, 'schemaVersion', 1),
    suite: literalField(value, 'suite', 'microsoft-excel-live-large-workbook-performance'),
    generatedAt: stringField(value, 'generatedAt'),
    host: {
      arch: stringField(host, 'arch'),
      platform: stringField(host, 'platform'),
    },
    source: {
      artifactGenerator: literalField(source, 'artifactGenerator', 'scripts/gen-microsoft-excel-live-large-workbook-scorecard.ts'),
      implementationPackage: literalField(source, 'implementationPackage', 'packages/core'),
      xlsxExportPackage: literalField(source, 'xlsxExportPackage', 'packages/excel-import'),
      corpusPackage: literalField(source, 'corpusPackage', 'packages/benchmarks'),
      evidenceKind: literalField(source, 'evidenceKind', 'live-local-microsoft-excel-automation'),
      appleScriptTransport: literalField(source, 'appleScriptTransport', 'osascript'),
    },
    benchmark: {
      sampleCount: numberField(benchmark, 'sampleCount'),
      screenUpdating: literalField(benchmark, 'screenUpdating', false),
      calculationMode: literalField(benchmark, 'calculationMode', 'manual-during-open-and-calculate'),
      measuredExcelOperation: literalField(benchmark, 'measuredExcelOperation', 'open-workbook-and-calculate-full-rebuild'),
      measuredBiligOperation: literalField(benchmark, 'measuredBiligOperation', 'import-snapshot'),
      samplingOrder: literalField(benchmark, 'samplingOrder', 'engine-isolated-bilig-then-excel'),
    },
    microsoftExcel: {
      appPath: literalField(microsoftExcel, 'appPath', excelAppPath),
      version: stringField(microsoftExcel, 'version'),
    },
    summary: {
      allRequiredCasesPassed: booleanField(summary, 'allRequiredCasesPassed'),
      requiredCaseCount: numberField(summary, 'requiredCaseCount'),
      tenXMeanAndP95CaseCount: numberField(summary, 'tenXMeanAndP95CaseCount'),
      biligWins: numberField(summary, 'biligWins'),
      coveredCorpusCaseIds: stringArrayField(summary, 'coveredCorpusCaseIds').map(parseWorkbookBenchmarkCorpusId),
      coveredMaterializedCells: arrayField(summary, 'coveredMaterializedCells').map((entry) => {
        if (typeof entry !== 'number' || !Number.isFinite(entry)) {
          throw new Error('Microsoft Excel live large-workbook coveredMaterializedCells must be finite numbers')
        }
        return entry
      }),
      googleSheetsEvidence: literalField(summary, 'googleSheetsEvidence', 'not-covered-by-this-artifact'),
    },
    cases: arrayField(value, 'cases').map(parseLargeWorkbookCase),
  }
}

export function validateMicrosoftExcelLiveLargeWorkbookScorecard(scorecard: MicrosoftExcelLiveLargeWorkbookScorecard): void {
  const expectedIds = caseSpecs.map((entry) => entry.id)
  const expectedCorpusCaseIds = caseSpecs.map((entry) => entry.corpusCaseId)
  const expectedMaterializedCells = caseSpecs.map((entry) => buildWorkbookBenchmarkCorpus(entry.corpusCaseId).materializedCellCount)
  if (scorecard.microsoftExcel.version.trim().length === 0) {
    throw new Error('Microsoft Excel live large-workbook scorecard must record an Excel version')
  }
  if (scorecard.benchmark.sampleCount !== sampleCount) {
    throw new Error('Microsoft Excel live large-workbook scorecard benchmark settings are stale')
  }
  if (
    scorecard.summary.requiredCaseCount !== expectedIds.length ||
    JSON.stringify(scorecard.cases.map((entry) => entry.id)) !== JSON.stringify(expectedIds)
  ) {
    throw new Error('Microsoft Excel live large-workbook scorecard required cases are stale')
  }
  if (JSON.stringify(scorecard.summary.coveredCorpusCaseIds) !== JSON.stringify(expectedCorpusCaseIds)) {
    throw new Error('Microsoft Excel live large-workbook scorecard covered corpus cases are stale')
  }
  if (JSON.stringify(scorecard.summary.coveredMaterializedCells) !== JSON.stringify(expectedMaterializedCells)) {
    throw new Error('Microsoft Excel live large-workbook scorecard materialized-cell coverage is stale')
  }
  if (
    scorecard.summary.biligWins !==
    scorecard.cases.filter((entry) => entry.biligElapsedMs.mean <= entry.microsoftExcelElapsedMs.mean).length
  ) {
    throw new Error('Microsoft Excel live large-workbook scorecard Bilig win count is inconsistent')
  }
  if (scorecard.summary.tenXMeanAndP95CaseCount !== scorecard.cases.filter((entry) => entry.tenXMeanAndP95).length) {
    throw new Error('Microsoft Excel live large-workbook scorecard 10x count is inconsistent')
  }
  const failingCases = scorecard.cases.filter((entry) => !entry.passed)
  if (!scorecard.summary.allRequiredCasesPassed || failingCases.length > 0) {
    throw new Error(
      `Microsoft Excel live large-workbook scorecard has failing required cases: ${failingCases
        .map(
          (entry) =>
            `${entry.id} Bilig=${JSON.stringify(entry.verification.bilig)} Excel=${JSON.stringify(entry.verification.microsoftExcel)}`,
        )
        .join(', ')}`,
    )
  }
  for (const entry of scorecard.cases) {
    validateNumericSummary(entry.biligElapsedMs, `${entry.id} Bilig elapsedMs`)
    validateNumericSummary(entry.microsoftExcelElapsedMs, `${entry.id} Microsoft Excel elapsedMs`)
    if (!entry.verification.equivalent) {
      throw new Error(`Microsoft Excel live large-workbook scorecard verification mismatch: ${entry.id}`)
    }
  }
}

async function runLargeWorkbookCasesSequentially(caseSpecsToRun: readonly LargeWorkbookCaseSpec[]): Promise<LargeWorkbookCaseRun[]> {
  if (caseSpecsToRun.length === 0) {
    return []
  }
  const [firstCaseSpec, ...remainingCaseSpecs] = caseSpecsToRun
  if (!firstCaseSpec) {
    return []
  }
  return [await runLargeWorkbookCase(firstCaseSpec), ...(await runLargeWorkbookCasesSequentially(remainingCaseSpecs))]
}

async function runLargeWorkbookCase(caseSpec: LargeWorkbookCaseSpec): Promise<LargeWorkbookCaseRun> {
  const corpus = buildWorkbookBenchmarkCorpus(caseSpec.corpusCaseId)
  const workbookBytes = exportXlsx(corpus.snapshot)
  const biligSamples: number[] = []
  const excelSamples: number[] = []
  const biligVerifications: Array<Record<string, boolean | number | string | null>> = []
  const excelVerifications: Array<Record<string, boolean | number | string | null>> = []
  const excelVersions = new Set<string>()

  const collectBiligSamples = async (index: number): Promise<void> => {
    if (index >= sampleCount) return
    const biligSample = await runBiligSample(caseSpec, corpus.snapshot)
    biligSamples.push(biligSample.elapsedMs)
    biligVerifications.push(biligSample.verification)
    await collectBiligSamples(index + 1)
  }
  const collectExcelSamples = (index: number): void => {
    if (index >= sampleCount) return
    const excelSample = runExcelSample(caseSpec, workbookBytes)
    excelSamples.push(excelSample.elapsedMs)
    excelVerifications.push(excelSample.verification)
    excelVersions.add(excelSample.excelVersion)
    collectExcelSamples(index + 1)
  }
  await collectBiligSamples(0)
  collectExcelSamples(0)

  const biligElapsedMs = summarizeNumbers(biligSamples)
  const microsoftExcelElapsedMs = summarizeNumbers(excelSamples)
  const biligToMicrosoftExcelMeanRatio = biligElapsedMs.mean / microsoftExcelElapsedMs.mean
  const biligToMicrosoftExcelP95Ratio = biligElapsedMs.p95 / microsoftExcelElapsedMs.p95
  const verification = {
    bilig: biligVerifications[0] ?? {},
    microsoftExcel: excelVerifications[0] ?? {},
    equivalent: verificationsEquivalent(biligVerifications, excelVerifications),
  }

  if (excelVersions.size !== 1) {
    throw new Error(`Microsoft Excel version changed during large-workbook benchmark: ${[...excelVersions].join(', ')}`)
  }

  return {
    excelVersion: [...excelVersions][0] ?? '',
    caseResult: {
      id: caseSpec.id,
      workload: caseSpec.workload,
      corpusCaseId: caseSpec.corpusCaseId,
      materializedCells: corpus.materializedCellCount,
      sampleCount,
      biligElapsedMs,
      microsoftExcelElapsedMs,
      biligToMicrosoftExcelMeanRatio,
      biligToMicrosoftExcelP95Ratio,
      tenXMeanAndP95: biligToMicrosoftExcelMeanRatio <= TEN_X_RATIO && biligToMicrosoftExcelP95Ratio <= TEN_X_RATIO,
      verification,
      passed:
        verification.equivalent && biligElapsedMs.samples.length === sampleCount && microsoftExcelElapsedMs.samples.length === sampleCount,
    },
  }
}

async function runBiligSample(caseSpec: LargeWorkbookCaseSpec, snapshot: WorkbookSnapshot): Promise<SampleResult> {
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

function runExcelSample(caseSpec: LargeWorkbookCaseSpec, workbookBytes: Uint8Array): ExcelSampleResult {
  const tempDir = mkdtempSync(join(tmpdir(), 'bilig-excel-live-large-workbook-'))
  const workbookPath = join(tempDir, 'large-workbook.xlsx')
  const scriptPath = join(tempDir, 'run-large-workbook.scpt')
  try {
    writeFileSync(workbookPath, workbookBytes)
    writeFileSync(scriptPath, createLargeWorkbookAppleScript(caseSpec))
    return parseExcelSampleOutput(execFileSync('osascript', [scriptPath, workbookPath], { encoding: 'utf8' }).trim())
  } finally {
    rmSync(tempDir, { recursive: true, force: true })
  }
}

function biligVerification(
  engine: SpreadsheetEngine,
  caseSpec: LargeWorkbookCaseSpec,
  snapshot: WorkbookSnapshot,
): Record<string, boolean | number | string | null> {
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

function createLargeWorkbookAppleScript(caseSpec: LargeWorkbookCaseSpec): string {
  return `use framework "Foundation"

on run argv
  set workbookPath to POSIX file (item 1 of argv)
  set output to ""
  tell application "Microsoft Excel"
    set display alerts to false
    set screen updating to false
    set previousCalculation to calculation
    try
      set calculation to calculation manual
      set startedAt to current application's NSDate's timeIntervalSinceReferenceDate()
      open workbook workbook file name workbookPath
      calculate full rebuild
      set elapsedMs to ((current application's NSDate's timeIntervalSinceReferenceDate()) - startedAt) * 1000
      set targetSheet to worksheet "${caseSpec.sheetName}" of active workbook
      set usedRangeRows to count of rows of used range of targetSheet
      set usedRangeColumns to count of columns of used range of targetSheet
      set output to "version=" & (version as string)
      set output to output & linefeed & "elapsedMs=" & (elapsedMs as string)
      set output to output & linefeed & "sheetCount=" & ((count of worksheets of active workbook) as string)
      set output to output & linefeed & "height=" & (usedRangeRows as string)
      set output to output & linefeed & "width=" & (usedRangeColumns as string)
      set output to output & linefeed & "usedRangeCells=" & ((usedRangeRows * usedRangeColumns) as string)
      set output to output & linefeed & "terminalAddress=${caseSpec.terminalAddress}"
      set output to output & linefeed & "terminalValue=" & ((value of range "${caseSpec.terminalAddress}" of targetSheet) as string)
      close active workbook saving no
      set calculation to previousCalculation
      set screen updating to true
    on error errMsg number errNum
      try
        close active workbook saving no
      end try
      set calculation to previousCalculation
      set screen updating to true
      error errMsg number errNum
    end try
  end tell
  return output
end run
`
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
      sheetCount: Number(requiredString(values, 'sheetCount')),
      height: Number(requiredString(values, 'height')),
      width: Number(requiredString(values, 'width')),
      usedRangeCells: Number(requiredString(values, 'usedRangeCells')),
      terminalAddress: requiredString(values, 'terminalAddress'),
      terminalValue: parseExcelVerificationValue(requiredString(values, 'terminalValue')),
    },
  }
}

function parseExcelVerificationValue(value: string): number | string | null {
  if (value.length === 0 || value === 'missing value') {
    return null
  }
  const numericValue = Number(value)
  return Number.isFinite(numericValue) ? numericValue : value
}

function parseLargeWorkbookCase(value: unknown): MicrosoftExcelLiveLargeWorkbookCase {
  const record = asObject(value, 'Microsoft Excel live large-workbook case')
  const verification = objectField(record, 'verification')
  return {
    id: stringField(record, 'id'),
    workload: parseWorkload(stringField(record, 'workload')),
    corpusCaseId: parseWorkbookBenchmarkCorpusId(stringField(record, 'corpusCaseId')),
    materializedCells: numberField(record, 'materializedCells'),
    sampleCount: numberField(record, 'sampleCount'),
    biligElapsedMs: parseNumericSummary(objectField(record, 'biligElapsedMs')),
    microsoftExcelElapsedMs: parseNumericSummary(objectField(record, 'microsoftExcelElapsedMs')),
    biligToMicrosoftExcelMeanRatio: numberField(record, 'biligToMicrosoftExcelMeanRatio'),
    biligToMicrosoftExcelP95Ratio: numberField(record, 'biligToMicrosoftExcelP95Ratio'),
    tenXMeanAndP95: booleanField(record, 'tenXMeanAndP95'),
    verification: {
      bilig: parseVerificationRecord(objectField(verification, 'bilig')),
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

function parseWorkload(value: string): MicrosoftExcelLiveLargeWorkbookWorkload {
  if (value === 'open-calculate-dense-mixed-100k' || value === 'open-calculate-dense-mixed-250k') {
    return value
  }
  throw new Error(`Unexpected Microsoft Excel live large-workbook workload: ${value}`)
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

function verificationsEquivalent(
  biligVerifications: readonly Record<string, boolean | number | string | null>[],
  excelVerifications: readonly Record<string, boolean | number | string | null>[],
): boolean {
  if (biligVerifications.length !== excelVerifications.length) {
    return false
  }
  return biligVerifications.every((biligSnapshot, index) => {
    const excelVerification = excelVerifications[index]
    return excelVerification !== undefined && JSON.stringify(biligSnapshot) === JSON.stringify(excelVerification)
  })
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

function normalizeProtocolValue(value: unknown): boolean | number | string | null {
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

function requiredString(values: ReadonlyMap<string, string>, key: string): string {
  const value = values.get(key)
  if (value === undefined || value.length === 0) {
    throw new Error(`Missing Microsoft Excel large-workbook output field: ${key}`)
  }
  return value
}

function validateNumericSummary(summary: NumericSummary, label: string): void {
  if (summary.samples.length !== sampleCount || summary.samples.some((entry) => !Number.isFinite(entry) || entry < 0)) {
    throw new Error(`${label} has invalid samples`)
  }
}

function logResult(mode: 'check' | 'write', scorecard: MicrosoftExcelLiveLargeWorkbookScorecard): void {
  console.log(
    JSON.stringify(
      {
        mode,
        outputPath,
        excelVersion: scorecard.microsoftExcel.version,
        allRequiredCasesPassed: scorecard.summary.allRequiredCasesPassed,
        biligWins: scorecard.summary.biligWins,
        tenXMeanAndP95CaseCount: scorecard.summary.tenXMeanAndP95CaseCount,
      },
      null,
      2,
    ),
  )
}

if (process.argv[1] && pathToFileURL(resolve(process.argv[1])).href === import.meta.url) {
  await main()
}
