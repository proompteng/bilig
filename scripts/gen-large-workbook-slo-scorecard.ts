#!/usr/bin/env bun

import { existsSync, mkdtempSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join, resolve } from 'node:path'

interface NumericSummary {
  readonly samples: number[]
  readonly min: number
  readonly median: number
  readonly p95: number
  readonly max: number
  readonly mean: number
}

interface BenchContractsReport {
  readonly baseBudgets: Record<string, number>
  readonly budgets: Record<string, number>
  readonly toleranceMultiplier: number
  readonly sampleCounts: Record<string, number>
  readonly results: Record<string, unknown>
}

export interface LargeWorkbookSloMeasurement {
  readonly id: string
  readonly category: 'large-workbook-scale' | 'ui-responsiveness' | 'collaboration'
  readonly label: string
  readonly materializedCells: number
  readonly corpusCaseId: string | null
  readonly metric: string
  readonly actualP95: number
  readonly budgetP95: number
  readonly gateBudgetP95: number
  readonly sampleCount: number
  readonly passed: boolean
  readonly gatePassed: boolean
}

export interface LargeWorkbookSloScorecard {
  readonly schemaVersion: 1
  readonly suite: 'large-workbook-slo'
  readonly generatedAt: string
  readonly source: {
    readonly benchmarkCommand: 'CI=1 pnpm bench:contracts'
    readonly benchmarkScript: 'scripts/bench-contracts.ts'
    readonly artifactGenerator: 'scripts/gen-large-workbook-slo-scorecard.ts'
  }
  readonly summary: {
    readonly coveredLargeWorkbookRows: number[]
    readonly allSloBudgetsPassed: boolean
    readonly allGateBudgetsPassed: boolean
    readonly headedBrowserFrameP95Evidence: 'not-captured'
    readonly externalGoogleSheetsEvidence: 'not-captured'
    readonly externalMicrosoftExcelEvidence: 'not-captured'
  }
  readonly measurements: LargeWorkbookSloMeasurement[]
}

interface MeasurementSpec {
  readonly id: string
  readonly category: LargeWorkbookSloMeasurement['category']
  readonly label: string
  readonly resultKey: string
  readonly metricKey: string
  readonly budgetKey: string
}

const measurementSpecs = [
  {
    id: 'load100k',
    category: 'large-workbook-scale',
    label: '100k dense-mixed core snapshot load',
    resultKey: 'load100k',
    metricKey: 'elapsedMs',
    budgetKey: 'load100kP95Ms',
  },
  {
    id: 'load250k',
    category: 'large-workbook-scale',
    label: '250k dense-mixed core snapshot load',
    resultKey: 'load250k',
    metricKey: 'elapsedMs',
    budgetKey: 'load250kP95Ms',
  },
  {
    id: 'workerWarmStart100k',
    category: 'large-workbook-scale',
    label: '100k dense-mixed browser worker warm start',
    resultKey: 'workerWarmStart100k',
    metricKey: 'elapsedMs',
    budgetKey: 'workerWarmStart100kP95Ms',
  },
  {
    id: 'workerWarmStart250k',
    category: 'large-workbook-scale',
    label: '250k dense-mixed browser worker warm start',
    resultKey: 'workerWarmStart250k',
    metricKey: 'elapsedMs',
    budgetKey: 'workerWarmStart250kP95Ms',
  },
  {
    id: 'workerVisibleEdit10k',
    category: 'ui-responsiveness',
    label: '10k browser worker visible edit patch',
    resultKey: 'workerVisibleEdit10k',
    metricKey: 'visiblePatchMs',
    budgetKey: 'workerVisibleEdit10kP95Ms',
  },
  {
    id: 'workerReconnectCatchUp100Pending',
    category: 'collaboration',
    label: '100 pending browser worker reconnect catch-up',
    resultKey: 'workerReconnectCatchUp100Pending',
    metricKey: 'catchUpMs',
    budgetKey: 'workerReconnectCatchUp100PendingP95Ms',
  },
] as const satisfies readonly MeasurementSpec[]

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const outputPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'large-workbook-slo-scorecard.json')
const isCheckMode = process.argv.includes('--check')

function main(): void {
  if (isCheckMode) {
    if (!existsSync(outputPath)) {
      throw new Error(`Large workbook SLO scorecard is missing. Run: bun scripts/gen-large-workbook-slo-scorecard.ts`)
    }
    const scorecard = parseLargeWorkbookSloScorecard(JSON.parse(readFileSync(outputPath, 'utf8')) as unknown)
    validateLargeWorkbookSloScorecard(scorecard)
    logResult('check', scorecard)
    return
  }

  const scorecard = buildLargeWorkbookSloScorecard(runBenchContractsReport())
  mkdirSync(dirname(outputPath), { recursive: true })
  writeFileSync(outputPath, formatJsonForRepo(`${JSON.stringify(scorecard, null, 2)}\n`))
  logResult('write', scorecard)
}

export function buildLargeWorkbookSloScorecard(reportInput: unknown, generatedAt = new Date().toISOString()): LargeWorkbookSloScorecard {
  const report = parseBenchContractsReport(reportInput)
  const measurements = measurementSpecs.map((spec) => buildMeasurement(report, spec))

  return {
    schemaVersion: 1,
    suite: 'large-workbook-slo',
    generatedAt,
    source: {
      benchmarkCommand: 'CI=1 pnpm bench:contracts',
      benchmarkScript: 'scripts/bench-contracts.ts',
      artifactGenerator: 'scripts/gen-large-workbook-slo-scorecard.ts',
    },
    summary: {
      coveredLargeWorkbookRows: [
        ...new Set(
          measurements
            .filter((measurement) => measurement.category === 'large-workbook-scale')
            .map((measurement) => measurement.materializedCells),
        ),
      ].toSorted((left, right) => left - right),
      allSloBudgetsPassed: measurements.every((measurement) => measurement.passed),
      allGateBudgetsPassed: measurements.every((measurement) => measurement.gatePassed),
      headedBrowserFrameP95Evidence: 'not-captured',
      externalGoogleSheetsEvidence: 'not-captured',
      externalMicrosoftExcelEvidence: 'not-captured',
    },
    measurements,
  }
}

function buildMeasurement(report: BenchContractsReport, spec: MeasurementSpec): LargeWorkbookSloMeasurement {
  if (!(spec.resultKey in report.results)) {
    throw new Error(`Missing large workbook SLO benchmark result: ${spec.resultKey}`)
  }
  const result = recordField(report.results, spec.resultKey, `large workbook SLO benchmark result: ${spec.resultKey}`)
  const summary = numericSummaryField(result, spec.metricKey, `large workbook SLO metric: ${spec.resultKey}.${spec.metricKey}`)
  const materializedCells = numberField(result, 'materializedCells', `large workbook SLO materialized cell count: ${spec.resultKey}`)
  const budgetP95 = numberField(report.baseBudgets, spec.budgetKey, `large workbook SLO base budget: ${spec.budgetKey}`)
  const gateBudgetP95 = numberField(report.budgets, spec.budgetKey, `large workbook SLO gate budget: ${spec.budgetKey}`)

  return {
    id: spec.id,
    category: spec.category,
    label: spec.label,
    materializedCells,
    corpusCaseId: optionalStringField(result, 'corpusCaseId'),
    metric: `${spec.metricKey}.p95`,
    actualP95: summary.p95,
    budgetP95,
    gateBudgetP95,
    sampleCount: numberField(report.sampleCounts, spec.resultKey, `large workbook SLO sample count: ${spec.resultKey}`),
    passed: summary.p95 <= budgetP95,
    gatePassed: summary.p95 <= gateBudgetP95,
  }
}

function runBenchContractsReport(): unknown {
  const result = Bun.spawnSync(['bun', join(rootDir, 'scripts', 'bench-contracts.ts')], {
    cwd: rootDir,
    env: {
      ...process.env,
      CI: '1',
    },
    stdin: 'ignore',
    stdout: 'pipe',
    stderr: 'pipe',
  })

  if (result.exitCode !== 0) {
    const stderr = new TextDecoder().decode(result.stderr).trim()
    throw new Error(`Unable to generate large workbook SLO scorecard from bench contracts: ${stderr}`)
  }

  return JSON.parse(new TextDecoder().decode(result.stdout)) as unknown
}

function parseBenchContractsReport(value: unknown): BenchContractsReport {
  const record = toRecord(value, 'bench contracts report')
  return {
    baseBudgets: numberRecordField(record, 'baseBudgets'),
    budgets: numberRecordField(record, 'budgets'),
    toleranceMultiplier: numberField(record, 'toleranceMultiplier', 'bench contracts tolerance multiplier'),
    sampleCounts: numberRecordField(record, 'sampleCounts'),
    results: recordField(record, 'results', 'bench contracts results'),
  }
}

function parseLargeWorkbookSloScorecard(value: unknown): LargeWorkbookSloScorecard {
  const record = toRecord(value, 'large workbook SLO scorecard')
  if (record['schemaVersion'] !== 1 || record['suite'] !== 'large-workbook-slo') {
    throw new Error('Unexpected large workbook SLO scorecard header')
  }
  const summary = recordField(record, 'summary', 'large workbook SLO scorecard summary')
  const measurements = arrayField(record, 'measurements', 'large workbook SLO scorecard measurements').map((entry, index) => {
    const measurement = toRecord(entry, `large workbook SLO scorecard measurement ${String(index)}`)
    return {
      id: stringField(measurement, 'id', `large workbook SLO scorecard measurement ${String(index)} id`),
      category: parseCategory(stringField(measurement, 'category', `large workbook SLO scorecard measurement ${String(index)} category`)),
      label: stringField(measurement, 'label', `large workbook SLO scorecard measurement ${String(index)} label`),
      materializedCells: numberField(
        measurement,
        'materializedCells',
        `large workbook SLO scorecard measurement ${String(index)} materializedCells`,
      ),
      corpusCaseId: optionalStringField(measurement, 'corpusCaseId'),
      metric: stringField(measurement, 'metric', `large workbook SLO scorecard measurement ${String(index)} metric`),
      actualP95: numberField(measurement, 'actualP95', `large workbook SLO scorecard measurement ${String(index)} actualP95`),
      budgetP95: numberField(measurement, 'budgetP95', `large workbook SLO scorecard measurement ${String(index)} budgetP95`),
      gateBudgetP95: numberField(measurement, 'gateBudgetP95', `large workbook SLO scorecard measurement ${String(index)} gateBudgetP95`),
      sampleCount: numberField(measurement, 'sampleCount', `large workbook SLO scorecard measurement ${String(index)} sampleCount`),
      passed: booleanField(measurement, 'passed', `large workbook SLO scorecard measurement ${String(index)} passed`),
      gatePassed: booleanField(measurement, 'gatePassed', `large workbook SLO scorecard measurement ${String(index)} gatePassed`),
    }
  })

  return {
    schemaVersion: 1,
    suite: 'large-workbook-slo',
    generatedAt: stringField(record, 'generatedAt', 'large workbook SLO scorecard generatedAt'),
    source: {
      benchmarkCommand: 'CI=1 pnpm bench:contracts',
      benchmarkScript: 'scripts/bench-contracts.ts',
      artifactGenerator: 'scripts/gen-large-workbook-slo-scorecard.ts',
    },
    summary: {
      coveredLargeWorkbookRows: numberArrayField(summary, 'coveredLargeWorkbookRows'),
      allSloBudgetsPassed: booleanField(summary, 'allSloBudgetsPassed', 'large workbook SLO allSloBudgetsPassed'),
      allGateBudgetsPassed: booleanField(summary, 'allGateBudgetsPassed', 'large workbook SLO allGateBudgetsPassed'),
      headedBrowserFrameP95Evidence: literalField(summary, 'headedBrowserFrameP95Evidence', 'not-captured'),
      externalGoogleSheetsEvidence: literalField(summary, 'externalGoogleSheetsEvidence', 'not-captured'),
      externalMicrosoftExcelEvidence: literalField(summary, 'externalMicrosoftExcelEvidence', 'not-captured'),
    },
    measurements,
  }
}

function validateLargeWorkbookSloScorecard(scorecard: LargeWorkbookSloScorecard): void {
  const actualIds = scorecard.measurements.map((measurement) => measurement.id)
  const expectedIds = measurementSpecs.map((spec) => spec.id)
  if (JSON.stringify(actualIds) !== JSON.stringify(expectedIds)) {
    throw new Error(
      `Large workbook SLO scorecard measurement coverage is stale. Expected ${expectedIds.join(', ')}, got ${actualIds.join(', ')}`,
    )
  }
  if (JSON.stringify(scorecard.summary.coveredLargeWorkbookRows) !== JSON.stringify([100_000, 250_000])) {
    throw new Error('Large workbook SLO scorecard must cover 100k and 250k materialized-cell sessions')
  }
  const failed = scorecard.measurements.find((measurement) => !measurement.passed || !measurement.gatePassed)
  if (failed) {
    throw new Error(`Large workbook SLO scorecard contains a failed measurement: ${failed.id}`)
  }
}

function logResult(mode: 'check' | 'write', scorecard: LargeWorkbookSloScorecard): void {
  console.log(
    JSON.stringify(
      {
        mode,
        outputPath,
        coveredLargeWorkbookRows: scorecard.summary.coveredLargeWorkbookRows,
        allSloBudgetsPassed: scorecard.summary.allSloBudgetsPassed,
        allGateBudgetsPassed: scorecard.summary.allGateBudgetsPassed,
      },
      null,
      2,
    ),
  )
}

function numericSummaryField(record: Record<string, unknown>, key: string, context: string): NumericSummary {
  const summary = recordField(record, key, context)
  return {
    samples: numberArrayField(summary, 'samples'),
    min: numberField(summary, 'min', `${context}.min`),
    median: numberField(summary, 'median', `${context}.median`),
    p95: numberField(summary, 'p95', `${context}.p95`),
    max: numberField(summary, 'max', `${context}.max`),
    mean: numberField(summary, 'mean', `${context}.mean`),
  }
}

function numberRecordField(record: Record<string, unknown>, key: string): Record<string, number> {
  const source = recordField(record, key, key)
  return Object.fromEntries(
    Object.entries(source).map(([entryKey, entryValue]) => [entryKey, expectNumber(entryValue, `${key}.${entryKey}`)]),
  )
}

function recordField(record: Record<string, unknown>, key: string, context: string): Record<string, unknown> {
  return toRecord(record[key], context)
}

function arrayField(record: Record<string, unknown>, key: string, context: string): unknown[] {
  const value = record[key]
  if (!Array.isArray(value)) {
    throw new Error(`Expected ${context} to be an array`)
  }
  return value
}

function numberArrayField(record: Record<string, unknown>, key: string): number[] {
  return arrayField(record, key, key).map((value, index) => expectNumber(value, `${key}.${String(index)}`))
}

function numberField(record: Record<string, unknown>, key: string, context: string): number {
  return expectNumber(record[key], context)
}

function stringField(record: Record<string, unknown>, key: string, context: string): string {
  const value = record[key]
  if (typeof value !== 'string') {
    throw new Error(`Expected ${context} to be a string`)
  }
  return value
}

function optionalStringField(record: Record<string, unknown>, key: string): string | null {
  const value = record[key]
  if (value === null || value === undefined) {
    return null
  }
  if (typeof value !== 'string') {
    throw new Error(`Expected ${key} to be a string or null`)
  }
  return value
}

function booleanField(record: Record<string, unknown>, key: string, context: string): boolean {
  const value = record[key]
  if (typeof value !== 'boolean') {
    throw new Error(`Expected ${context} to be a boolean`)
  }
  return value
}

function literalField<T extends string>(record: Record<string, unknown>, key: string, expected: T): T {
  if (record[key] !== expected) {
    throw new Error(`Expected ${key} to be ${expected}`)
  }
  return expected
}

function parseCategory(value: string): LargeWorkbookSloMeasurement['category'] {
  if (value === 'large-workbook-scale' || value === 'ui-responsiveness' || value === 'collaboration') {
    return value
  }
  throw new Error(`Unexpected large workbook SLO category: ${value}`)
}

function expectNumber(value: unknown, context: string): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new Error(`Expected ${context} to be a finite number`)
  }
  return value
}

function toRecord(value: unknown, context: string): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new Error(`Expected ${context} to be an object`)
  }
  const record: Record<string, unknown> = {}
  for (const key of Object.keys(value)) {
    record[key] = Reflect.get(value, key)
  }
  return record
}

function formatJsonForRepo(serializedJson: string): string {
  const tempDir = mkdtempSync(join(tmpdir(), 'large-workbook-slo-'))
  const tempFilePath = join(tempDir, 'scorecard.json')
  writeFileSync(tempFilePath, serializedJson)
  const oxfmtPath = join(rootDir, 'node_modules', '.bin', 'oxfmt')

  const formatResult = Bun.spawnSync([oxfmtPath, '--write', tempFilePath], {
    stdin: 'ignore',
    stdout: 'pipe',
    stderr: 'pipe',
  })
  if (formatResult.exitCode !== 0) {
    const stderr = new TextDecoder().decode(formatResult.stderr).trim()
    rmSync(tempDir, { recursive: true, force: true })
    throw new Error(`Unable to format generated large workbook SLO scorecard: ${stderr}`)
  }

  const formattedJson = readFileSync(tempFilePath, 'utf8')
  rmSync(tempDir, { recursive: true, force: true })
  return formattedJson
}

if (import.meta.url === `file://${process.argv[1]}`) {
  main()
}
