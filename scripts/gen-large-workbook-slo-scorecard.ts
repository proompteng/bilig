#!/usr/bin/env bun

import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'

import {
  externalLargeWorkbookComparisonArtifactRepoPath,
  externalLargeWorkbookComparisonCoveredFeatures,
  parseExternalLargeWorkbookComparisonArtifact,
  validateExternalLargeWorkbookComparisonArtifact,
} from './large-workbook-external-sheets-excel-comparison.ts'
import {
  externalUiResponsivenessComparisonArtifactRepoPath,
  externalUiResponsivenessComparisonCoveredFeatures,
  parseExternalUiResponsivenessComparisonArtifact,
  validateExternalUiResponsivenessComparisonArtifact,
} from './ui-responsiveness-external-sheets-excel-comparison.ts'
import { formatJsonForRepo } from './scorecard-format.ts'

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
  readonly headedBrowserTestSource?: string
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

export interface HeadedBrowserFrameP95Contract {
  readonly id: string
  readonly category: 'large-workbook-scale' | 'ui-responsiveness'
  readonly label: string
  readonly materializedCells: number
  readonly corpusCaseId: string
  readonly metric: 'frameMs.p95' | 'mutationToVisibleMs.p95'
  readonly budgetP95: number
  readonly minSampleCount: number
  readonly playwrightTestFile: 'e2e/tests/web-shell-scroll-performance.pw.ts'
  readonly playwrightArtifactFile: string
  readonly command: 'pnpm test:browser:full'
  readonly passed: boolean
  readonly findings: string[]
}

export interface LargeWorkbookExternalComparisonEvidence {
  readonly artifact: 'packages/benchmarks/baselines/large-workbook-external-sheets-excel-comparison.json'
  readonly sourceBasis: string
  readonly officialGoogleSheetsSourceCount: number
  readonly officialMicrosoftExcelSourceCount: number
  readonly requiredDimensionsPassed: boolean
  readonly coveredFeatures: string[]
  readonly limitations: string[]
  readonly findings: string[]
}

export interface UiResponsivenessExternalComparisonEvidence {
  readonly artifact: 'packages/benchmarks/baselines/ui-responsiveness-external-sheets-excel-comparison.json'
  readonly sourceBasis: string
  readonly officialGoogleSheetsSourceCount: number
  readonly officialMicrosoftExcelSourceCount: number
  readonly requiredDimensionsPassed: boolean
  readonly coveredFeatures: string[]
  readonly limitations: string[]
  readonly findings: string[]
}

export interface LargeWorkbookSloScorecard {
  readonly schemaVersion: 1
  readonly suite: 'large-workbook-slo'
  readonly generatedAt: string
  readonly source: {
    readonly benchmarkCommand: 'CI=1 pnpm bench:contracts'
    readonly benchmarkScript: 'scripts/bench-contracts.ts'
    readonly headedBrowserCommand: 'pnpm test:browser:full'
    readonly headedBrowserTestFile: 'e2e/tests/web-shell-scroll-performance.pw.ts'
    readonly artifactGenerator: 'scripts/gen-large-workbook-slo-scorecard.ts'
    readonly externalLargeWorkbookComparisonArtifact: 'packages/benchmarks/baselines/large-workbook-external-sheets-excel-comparison.json'
    readonly externalUiResponsivenessComparisonArtifact: 'packages/benchmarks/baselines/ui-responsiveness-external-sheets-excel-comparison.json'
  }
  readonly summary: {
    readonly coveredLargeWorkbookRows: number[]
    readonly allSloBudgetsPassed: boolean
    readonly allGateBudgetsPassed: boolean
    readonly headedBrowserFrameP95Evidence: 'playwright-contracts'
    readonly headedBrowserFrameP95ContractsPassed: boolean
    readonly externalGoogleSheetsEvidence: 'official-docs-comparison-artifact'
    readonly externalMicrosoftExcelEvidence: 'official-docs-comparison-artifact'
    readonly externalUiResponsivenessGoogleSheetsEvidence: 'official-docs-comparison-artifact'
    readonly externalUiResponsivenessMicrosoftExcelEvidence: 'official-docs-comparison-artifact'
  }
  readonly measurements: LargeWorkbookSloMeasurement[]
  readonly headedBrowserFrameP95Contracts: HeadedBrowserFrameP95Contract[]
  readonly externalSheetsExcelComparison: LargeWorkbookExternalComparisonEvidence
  readonly uiResponsivenessExternalSheetsExcelComparison: UiResponsivenessExternalComparisonEvidence
}

interface MeasurementSpec {
  readonly id: string
  readonly category: LargeWorkbookSloMeasurement['category']
  readonly label: string
  readonly resultKey: string
  readonly metricKey: string
  readonly budgetKey: string
}

interface HeadedBrowserFrameP95ContractSpec {
  readonly id: string
  readonly category: HeadedBrowserFrameP95Contract['category']
  readonly label: string
  readonly testTitle: string
  readonly materializedCells: number
  readonly corpusCaseId: string
  readonly workload: string
  readonly metric: HeadedBrowserFrameP95Contract['metric']
  readonly budgetP95: number
  readonly minSampleCount: number
  readonly playwrightArtifactFile: string
  readonly verification: 'smooth-browse' | 'bounded-visible-mutation'
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

const headedBrowserTestFile = 'e2e/tests/web-shell-scroll-performance.pw.ts'
const headedBrowserFrameP95ContractSpecs = [
  {
    id: 'headedDense100kDiagonalBrowse',
    category: 'large-workbook-scale',
    label: '100k dense headed diagonal browse frame pacing',
    testTitle: 'keeps dense 100k browse inside headed frame budgets',
    materializedCells: 100_000,
    corpusCaseId: 'dense-mixed-100k',
    workload: 'dense-100k-diagonal-main-body',
    metric: 'frameMs.p95',
    budgetP95: 20,
    minSampleCount: 120,
    playwrightArtifactFile: 'scroll-perf-dense-100k-diagonal.json',
    verification: 'smooth-browse',
  },
  {
    id: 'headedWide250kMainBodyBrowse',
    category: 'large-workbook-scale',
    label: '250k wide headed main-body browse frame pacing',
    testTitle: 'keeps horizontal browse inside one resident window smooth and free of data-canvas redraw churn',
    materializedCells: 250_000,
    corpusCaseId: 'wide-mixed-250k',
    workload: 'wide-250k-main-body',
    metric: 'frameMs.p95',
    budgetP95: 20,
    minSampleCount: 120,
    playwrightArtifactFile: 'scroll-perf-wide-250k-main-body.json',
    verification: 'smooth-browse',
  },
  {
    id: 'headedWide250kVisibleEditCommit',
    category: 'ui-responsiveness',
    label: '250k wide headed visible edit commit response',
    testTitle: 'keeps visible edit commits bounded to dirty V3 tiles',
    materializedCells: 250_000,
    corpusCaseId: 'wide-mixed-250k',
    workload: 'wide-250k-visible-edit-commit',
    metric: 'mutationToVisibleMs.p95',
    budgetP95: 50,
    minSampleCount: 1,
    playwrightArtifactFile: 'scroll-perf-wide-250k-visible-edit.json',
    verification: 'bounded-visible-mutation',
  },
] as const satisfies readonly HeadedBrowserFrameP95ContractSpec[]

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const outputPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'large-workbook-slo-scorecard.json')
const externalLargeWorkbookComparisonArtifactPath = join(rootDir, externalLargeWorkbookComparisonArtifactRepoPath)
const externalUiResponsivenessComparisonArtifactPath = join(rootDir, externalUiResponsivenessComparisonArtifactRepoPath)
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

export function buildLargeWorkbookSloScorecard(
  reportInput: unknown,
  generatedAt = new Date().toISOString(),
  externalComparisonInput: unknown = readExternalLargeWorkbookComparisonArtifact(),
  externalUiResponsivenessComparisonInput: unknown = readExternalUiResponsivenessComparisonArtifact(),
): LargeWorkbookSloScorecard {
  const report = parseBenchContractsReport(reportInput)
  const externalSheetsExcelComparison = buildExternalComparisonEvidence(externalComparisonInput)
  const uiResponsivenessExternalSheetsExcelComparison = buildUiResponsivenessExternalComparisonEvidence(
    externalUiResponsivenessComparisonInput,
  )
  const measurements = measurementSpecs.map((spec) => buildMeasurement(report, spec))
  const headedBrowserFrameP95Contracts = buildHeadedBrowserFrameP95Contracts(
    report.headedBrowserTestSource ?? readFileSync(join(rootDir, headedBrowserTestFile), 'utf8'),
  )

  return {
    schemaVersion: 1,
    suite: 'large-workbook-slo',
    generatedAt,
    source: {
      benchmarkCommand: 'CI=1 pnpm bench:contracts',
      benchmarkScript: 'scripts/bench-contracts.ts',
      headedBrowserCommand: 'pnpm test:browser:full',
      headedBrowserTestFile,
      artifactGenerator: 'scripts/gen-large-workbook-slo-scorecard.ts',
      externalLargeWorkbookComparisonArtifact: externalLargeWorkbookComparisonArtifactRepoPath,
      externalUiResponsivenessComparisonArtifact: externalUiResponsivenessComparisonArtifactRepoPath,
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
      headedBrowserFrameP95Evidence: 'playwright-contracts',
      headedBrowserFrameP95ContractsPassed: headedBrowserFrameP95Contracts.every((contract) => contract.passed),
      externalGoogleSheetsEvidence: 'official-docs-comparison-artifact',
      externalMicrosoftExcelEvidence: 'official-docs-comparison-artifact',
      externalUiResponsivenessGoogleSheetsEvidence: 'official-docs-comparison-artifact',
      externalUiResponsivenessMicrosoftExcelEvidence: 'official-docs-comparison-artifact',
    },
    measurements,
    headedBrowserFrameP95Contracts,
    externalSheetsExcelComparison,
    uiResponsivenessExternalSheetsExcelComparison,
  }
}

function readExternalLargeWorkbookComparisonArtifact(): unknown {
  return JSON.parse(readFileSync(externalLargeWorkbookComparisonArtifactPath, 'utf8')) as unknown
}

function readExternalUiResponsivenessComparisonArtifact(): unknown {
  return JSON.parse(readFileSync(externalUiResponsivenessComparisonArtifactPath, 'utf8')) as unknown
}

function buildExternalComparisonEvidence(value: unknown): LargeWorkbookExternalComparisonEvidence {
  const artifact = parseExternalLargeWorkbookComparisonArtifact(value)
  const findings = validateExternalLargeWorkbookComparisonArtifact(artifact)
  return {
    artifact: externalLargeWorkbookComparisonArtifactRepoPath,
    sourceBasis: artifact.sourceBasis,
    officialGoogleSheetsSourceCount: artifact.officialSources.filter((source) => source.vendor === 'google-sheets').length,
    officialMicrosoftExcelSourceCount: artifact.officialSources.filter((source) => source.vendor === 'microsoft-excel').length,
    requiredDimensionsPassed: artifact.summary.requiredDimensionsPassed && findings.length === 0,
    coveredFeatures: artifact.summary.coveredFeatures,
    limitations: artifact.summary.limitations,
    findings,
  }
}

function buildUiResponsivenessExternalComparisonEvidence(value: unknown): UiResponsivenessExternalComparisonEvidence {
  const artifact = parseExternalUiResponsivenessComparisonArtifact(value)
  const findings = validateExternalUiResponsivenessComparisonArtifact(artifact)
  return {
    artifact: externalUiResponsivenessComparisonArtifactRepoPath,
    sourceBasis: artifact.sourceBasis,
    officialGoogleSheetsSourceCount: artifact.officialSources.filter((source) => source.vendor === 'google-sheets').length,
    officialMicrosoftExcelSourceCount: artifact.officialSources.filter((source) => source.vendor === 'microsoft-excel').length,
    requiredDimensionsPassed: artifact.summary.requiredDimensionsPassed && findings.length === 0,
    coveredFeatures: artifact.summary.coveredFeatures,
    limitations: artifact.summary.limitations,
    findings,
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

function buildHeadedBrowserFrameP95Contracts(source: string): HeadedBrowserFrameP95Contract[] {
  return headedBrowserFrameP95ContractSpecs.map((spec) => {
    const findings: string[] = []
    const testBlock = extractPlaywrightTestBlock(source, spec.testTitle)
    if (!testBlock) {
      findings.push(`missing Playwright test: ${spec.testTitle}`)
    } else {
      findings.push(...validateHeadedBrowserFrameP95TestBlock(source, testBlock, spec))
    }
    return {
      id: spec.id,
      category: spec.category,
      label: spec.label,
      materializedCells: spec.materializedCells,
      corpusCaseId: spec.corpusCaseId,
      metric: spec.metric,
      budgetP95: spec.budgetP95,
      minSampleCount: spec.minSampleCount,
      playwrightTestFile: headedBrowserTestFile,
      playwrightArtifactFile: spec.playwrightArtifactFile,
      command: 'pnpm test:browser:full',
      passed: findings.length === 0,
      findings,
    }
  })
}

function extractPlaywrightTestBlock(source: string, testTitle: string): string | null {
  const marker = `test('${testTitle}'`
  const start = source.indexOf(marker)
  if (start < 0) {
    return null
  }
  const endCandidates = ['\n  test(', '\n  remoteSyncTest(']
    .map((nextMarker) => source.indexOf(nextMarker, start + marker.length))
    .filter((index) => index >= 0)
  const end = endCandidates.length > 0 ? Math.min(...endCandidates) : source.length
  return source.slice(start, end)
}

function validateHeadedBrowserFrameP95TestBlock(source: string, testBlock: string, spec: HeadedBrowserFrameP95ContractSpec): string[] {
  const findings: string[] = []
  requireSnippet(testBlock, `benchmarkCorpus=${spec.corpusCaseId}`, 'loads the expected benchmark corpus', findings)
  requireSnippet(
    testBlock,
    `expect(benchmarkState.fixture?.id).toBe('${spec.corpusCaseId}')`,
    'asserts installed corpus identity',
    findings,
  )
  requireSnippet(testBlock, `expect(report.fixture?.id).toBe('${spec.corpusCaseId}')`, 'asserts sampled report corpus identity', findings)
  requireSnippet(testBlock, `warmStartWorkbookScrollPerf(page, '${spec.workload}'`, 'warms the renderer before sampling', findings)
  requireSnippet(testBlock, 'stopWorkbookScrollPerf(page)', 'stops and captures the headed browser perf report', findings)
  requireSnippet(testBlock, `testInfo.outputPath('${spec.playwrightArtifactFile}')`, 'writes the headed browser perf artifact', findings)

  if (spec.verification === 'smooth-browse') {
    requireSnippet(testBlock, 'expectSmoothBrowse(report', 'asserts smooth headed browsing', findings)
    requireSnippet(source, 'expect(frameSummary.p95).toBeLessThan(20)', 'keeps frame p95 under 20ms', findings)
    requireSnippet(source, 'expect(report.samples.frameMs.length).toBeGreaterThan(120)', 'captures at least 120 frame samples', findings)
  } else {
    requireSnippet(testBlock, 'expectBoundedVisibleMutation(report', 'asserts bounded visible mutation response', findings)
    requireSnippet(testBlock, 'mutationToVisibleP95Max: 50', 'keeps mutation-to-visible p95 under 50ms', findings)
    requireSnippet(
      source,
      'expect(report.samples.mutationToVisibleMs.length).toBeGreaterThan(0)',
      'captures mutation-to-visible samples',
      findings,
    )
  }

  return findings
}

function requireSnippet(source: string, snippet: string, label: string, findings: string[]): void {
  if (!source.includes(snippet)) {
    findings.push(`missing ${label}`)
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
    headedBrowserTestSource: optionalStringField(record, 'headedBrowserTestSource') ?? undefined,
  }
}

export function parseLargeWorkbookSloScorecard(value: unknown): LargeWorkbookSloScorecard {
  const record = toRecord(value, 'large workbook SLO scorecard')
  if (record['schemaVersion'] !== 1 || record['suite'] !== 'large-workbook-slo') {
    throw new Error('Unexpected large workbook SLO scorecard header')
  }
  const source = recordField(record, 'source', 'large workbook SLO scorecard source')
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
      benchmarkCommand: literalField(source, 'benchmarkCommand', 'CI=1 pnpm bench:contracts'),
      benchmarkScript: literalField(source, 'benchmarkScript', 'scripts/bench-contracts.ts'),
      headedBrowserCommand: literalField(source, 'headedBrowserCommand', 'pnpm test:browser:full'),
      headedBrowserTestFile: literalField(source, 'headedBrowserTestFile', headedBrowserTestFile),
      artifactGenerator: literalField(source, 'artifactGenerator', 'scripts/gen-large-workbook-slo-scorecard.ts'),
      externalLargeWorkbookComparisonArtifact: literalField(
        source,
        'externalLargeWorkbookComparisonArtifact',
        externalLargeWorkbookComparisonArtifactRepoPath,
      ),
      externalUiResponsivenessComparisonArtifact: literalField(
        source,
        'externalUiResponsivenessComparisonArtifact',
        externalUiResponsivenessComparisonArtifactRepoPath,
      ),
    },
    summary: {
      coveredLargeWorkbookRows: numberArrayField(summary, 'coveredLargeWorkbookRows'),
      allSloBudgetsPassed: booleanField(summary, 'allSloBudgetsPassed', 'large workbook SLO allSloBudgetsPassed'),
      allGateBudgetsPassed: booleanField(summary, 'allGateBudgetsPassed', 'large workbook SLO allGateBudgetsPassed'),
      headedBrowserFrameP95Evidence: literalField(summary, 'headedBrowserFrameP95Evidence', 'playwright-contracts'),
      headedBrowserFrameP95ContractsPassed: booleanField(
        summary,
        'headedBrowserFrameP95ContractsPassed',
        'large workbook headed browser frame contracts passed',
      ),
      externalGoogleSheetsEvidence: literalField(summary, 'externalGoogleSheetsEvidence', 'official-docs-comparison-artifact'),
      externalMicrosoftExcelEvidence: literalField(summary, 'externalMicrosoftExcelEvidence', 'official-docs-comparison-artifact'),
      externalUiResponsivenessGoogleSheetsEvidence: literalField(
        summary,
        'externalUiResponsivenessGoogleSheetsEvidence',
        'official-docs-comparison-artifact',
      ),
      externalUiResponsivenessMicrosoftExcelEvidence: literalField(
        summary,
        'externalUiResponsivenessMicrosoftExcelEvidence',
        'official-docs-comparison-artifact',
      ),
    },
    measurements,
    headedBrowserFrameP95Contracts: arrayField(
      record,
      'headedBrowserFrameP95Contracts',
      'large workbook headed browser frame contracts',
    ).map(parseHeadedBrowserFrameP95Contract),
    externalSheetsExcelComparison: parseExternalComparisonEvidence(
      recordField(record, 'externalSheetsExcelComparison', 'large workbook external Sheets/Excel comparison'),
    ),
    uiResponsivenessExternalSheetsExcelComparison: parseUiResponsivenessExternalComparisonEvidence(
      recordField(record, 'uiResponsivenessExternalSheetsExcelComparison', 'UI responsiveness external Sheets/Excel comparison'),
    ),
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
  const actualHeadedBrowserIds = scorecard.headedBrowserFrameP95Contracts.map((contract) => contract.id)
  const expectedHeadedBrowserIds = headedBrowserFrameP95ContractSpecs.map((spec) => spec.id)
  if (JSON.stringify(actualHeadedBrowserIds) !== JSON.stringify(expectedHeadedBrowserIds)) {
    throw new Error(
      `Large workbook SLO scorecard headed browser coverage is stale. Expected ${expectedHeadedBrowserIds.join(', ')}, got ${actualHeadedBrowserIds.join(', ')}`,
    )
  }
  const failed = scorecard.measurements.find((measurement) => !measurement.passed || !measurement.gatePassed)
  if (failed) {
    throw new Error(`Large workbook SLO scorecard contains a failed measurement: ${failed.id}`)
  }
  const failedContract = scorecard.headedBrowserFrameP95Contracts.find((contract) => !contract.passed)
  if (failedContract) {
    throw new Error(`Large workbook SLO scorecard contains a failed headed browser contract: ${failedContract.id}`)
  }
  if (!scorecard.externalSheetsExcelComparison.requiredDimensionsPassed || scorecard.externalSheetsExcelComparison.findings.length > 0) {
    throw new Error(
      `Large workbook SLO scorecard contains stale external Sheets/Excel comparison evidence: ${scorecard.externalSheetsExcelComparison.findings.join(
        ', ',
      )}`,
    )
  }
  if (!arrayEquals(scorecard.externalSheetsExcelComparison.coveredFeatures, externalLargeWorkbookComparisonCoveredFeatures)) {
    throw new Error('Large workbook SLO scorecard external comparison feature coverage is stale')
  }
  if (
    !scorecard.uiResponsivenessExternalSheetsExcelComparison.requiredDimensionsPassed ||
    scorecard.uiResponsivenessExternalSheetsExcelComparison.findings.length > 0
  ) {
    throw new Error(
      `Large workbook SLO scorecard contains stale UI responsiveness external comparison evidence: ${scorecard.uiResponsivenessExternalSheetsExcelComparison.findings.join(
        ', ',
      )}`,
    )
  }
  if (
    !arrayEquals(scorecard.uiResponsivenessExternalSheetsExcelComparison.coveredFeatures, externalUiResponsivenessComparisonCoveredFeatures)
  ) {
    throw new Error('Large workbook SLO scorecard UI responsiveness external comparison feature coverage is stale')
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
        headedBrowserFrameP95ContractsPassed: scorecard.summary.headedBrowserFrameP95ContractsPassed,
        externalGoogleSheetsEvidence: scorecard.summary.externalGoogleSheetsEvidence,
        externalMicrosoftExcelEvidence: scorecard.summary.externalMicrosoftExcelEvidence,
        externalUiResponsivenessGoogleSheetsEvidence: scorecard.summary.externalUiResponsivenessGoogleSheetsEvidence,
        externalUiResponsivenessMicrosoftExcelEvidence: scorecard.summary.externalUiResponsivenessMicrosoftExcelEvidence,
      },
      null,
      2,
    ),
  )
}

function parseExternalComparisonEvidence(record: Record<string, unknown>): LargeWorkbookExternalComparisonEvidence {
  return {
    artifact: literalField(record, 'artifact', externalLargeWorkbookComparisonArtifactRepoPath),
    sourceBasis: stringField(record, 'sourceBasis', 'large workbook external comparison sourceBasis'),
    officialGoogleSheetsSourceCount: numberField(
      record,
      'officialGoogleSheetsSourceCount',
      'large workbook external comparison officialGoogleSheetsSourceCount',
    ),
    officialMicrosoftExcelSourceCount: numberField(
      record,
      'officialMicrosoftExcelSourceCount',
      'large workbook external comparison officialMicrosoftExcelSourceCount',
    ),
    requiredDimensionsPassed: booleanField(
      record,
      'requiredDimensionsPassed',
      'large workbook external comparison requiredDimensionsPassed',
    ),
    coveredFeatures: stringArrayField(record, 'coveredFeatures', 'large workbook external comparison coveredFeatures'),
    limitations: stringArrayField(record, 'limitations', 'large workbook external comparison limitations'),
    findings: stringArrayField(record, 'findings', 'large workbook external comparison findings'),
  }
}

function parseHeadedBrowserFrameP95Contract(entry: unknown, index: number): HeadedBrowserFrameP95Contract {
  const contract = toRecord(entry, `large workbook headed browser frame contract ${String(index)}`)
  return {
    id: stringField(contract, 'id', `large workbook headed browser frame contract ${String(index)} id`),
    category: parseHeadedBrowserFrameP95Category(
      stringField(contract, 'category', `large workbook headed browser frame contract ${String(index)} category`),
    ),
    label: stringField(contract, 'label', `large workbook headed browser frame contract ${String(index)} label`),
    materializedCells: numberField(
      contract,
      'materializedCells',
      `large workbook headed browser frame contract ${String(index)} materializedCells`,
    ),
    corpusCaseId: stringField(contract, 'corpusCaseId', `large workbook headed browser frame contract ${String(index)} corpusCaseId`),
    metric: parseHeadedBrowserFrameP95Metric(
      stringField(contract, 'metric', `large workbook headed browser frame contract ${String(index)} metric`),
    ),
    budgetP95: numberField(contract, 'budgetP95', `large workbook headed browser frame contract ${String(index)} budgetP95`),
    minSampleCount: numberField(contract, 'minSampleCount', `large workbook headed browser frame contract ${String(index)} minSampleCount`),
    playwrightTestFile: literalField(contract, 'playwrightTestFile', headedBrowserTestFile),
    playwrightArtifactFile: stringField(
      contract,
      'playwrightArtifactFile',
      `large workbook headed browser frame contract ${String(index)} playwrightArtifactFile`,
    ),
    command: literalField(contract, 'command', 'pnpm test:browser:full'),
    passed: booleanField(contract, 'passed', `large workbook headed browser frame contract ${String(index)} passed`),
    findings: stringArrayField(contract, 'findings', `large workbook headed browser frame contract ${String(index)} findings`),
  }
}

function parseUiResponsivenessExternalComparisonEvidence(record: Record<string, unknown>): UiResponsivenessExternalComparisonEvidence {
  return {
    artifact: literalField(record, 'artifact', externalUiResponsivenessComparisonArtifactRepoPath),
    sourceBasis: stringField(record, 'sourceBasis', 'UI responsiveness external comparison sourceBasis'),
    officialGoogleSheetsSourceCount: numberField(
      record,
      'officialGoogleSheetsSourceCount',
      'UI responsiveness external comparison officialGoogleSheetsSourceCount',
    ),
    officialMicrosoftExcelSourceCount: numberField(
      record,
      'officialMicrosoftExcelSourceCount',
      'UI responsiveness external comparison officialMicrosoftExcelSourceCount',
    ),
    requiredDimensionsPassed: booleanField(
      record,
      'requiredDimensionsPassed',
      'UI responsiveness external comparison requiredDimensionsPassed',
    ),
    coveredFeatures: stringArrayField(record, 'coveredFeatures', 'UI responsiveness external comparison coveredFeatures'),
    limitations: stringArrayField(record, 'limitations', 'UI responsiveness external comparison limitations'),
    findings: stringArrayField(record, 'findings', 'UI responsiveness external comparison findings'),
  }
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

function parseHeadedBrowserFrameP95Category(value: string): HeadedBrowserFrameP95Contract['category'] {
  if (value === 'large-workbook-scale' || value === 'ui-responsiveness') {
    return value
  }
  throw new Error(`Unexpected large workbook headed browser category: ${value}`)
}

function parseHeadedBrowserFrameP95Metric(value: string): HeadedBrowserFrameP95Contract['metric'] {
  if (value === 'frameMs.p95' || value === 'mutationToVisibleMs.p95') {
    return value
  }
  throw new Error(`Unexpected large workbook headed browser metric: ${value}`)
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

function stringArrayField(record: Record<string, unknown>, key: string, context: string): string[] {
  return arrayField(record, key, context).map((value, index) => {
    if (typeof value !== 'string') {
      throw new Error(`Expected ${context}.${String(index)} to be a string`)
    }
    return value
  })
}

function arrayEquals(left: readonly string[], right: readonly string[]): boolean {
  return left.length === right.length && left.every((value, index) => value === right[index])
}

if (import.meta.url === `file://${process.argv[1]}`) {
  main()
}
