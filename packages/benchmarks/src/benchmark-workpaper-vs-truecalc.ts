import { performance } from 'node:perf_hooks'
import { WorkPaper, type WorkPaperSheet } from '@bilig/headless'
import { createEngine, type EvalResult } from '@truecalc/core'
import {
  DEFAULT_COMPETITIVE_SAMPLE_COUNT,
  DEFAULT_COMPETITIVE_WARMUP_COUNT,
  type ComparativeBenchmarkSuiteOptions,
  type ComparativeMeasuredEngineResult,
  type ComparativeMemorySummary,
} from './benchmark-workpaper-vs-hyperformula.js'
import { normalizeWorkPaperValue, type BenchmarkSample } from './benchmark-workpaper-vs-hyperformula-expanded-support.js'
import { measureMemory, sampleMemory, type MemoryMeasurement } from './metrics.js'
import { summarizeNumbers } from './stats.js'

export type WorkPaperTrueCalcScalarWorkload =
  | 'scalar-arithmetic'
  | 'scalar-branching'
  | 'scalar-financial-pmt'
  | 'scalar-math-nested'
  | 'scalar-text-concat'
  | 'scalar-text-length'
  | 'scalar-minmax'

export type TrueCalcVariableValue = boolean | number | string | null

export interface WorkPaperTrueCalcScalarFixture {
  readonly edit: {
    readonly col: number
    readonly row: number
    readonly value: TrueCalcVariableValue
  }
  readonly formula: string
  readonly result: {
    readonly col: number
    readonly row: number
  }
  readonly sheet: WorkPaperSheet
  readonly variables: Readonly<Record<string, TrueCalcVariableValue>>
}

export interface WorkPaperTrueCalcScalarComparison {
  readonly confidenceIntervalOverlaps: boolean
  readonly fasterEngine: 'truecalc' | 'workpaper'
  readonly maxRelativeNoise: number
  readonly meanSpeedup: number
  readonly verificationEquivalent: true
  readonly workpaperToTrueCalcMeanRatio: number
  readonly workpaperToTrueCalcMedianRatio: number
  readonly workpaperToTrueCalcP95Ratio: number
}

export interface WorkPaperTrueCalcScalarBenchmarkResult {
  readonly workload: WorkPaperTrueCalcScalarWorkload
  readonly category: 'scalar-formula'
  readonly comparable: true
  readonly fixture: WorkPaperTrueCalcScalarFixture
  readonly comparison: WorkPaperTrueCalcScalarComparison
  readonly engines: {
    readonly truecalc: ComparativeMeasuredEngineResult
    readonly workpaper: ComparativeMeasuredEngineResult
  }
}

export interface WorkPaperTrueCalcScalarScorecard {
  readonly comparableWorkloadCount: number
  readonly coverageNote: string
  readonly coverageTier: 'scalar-formula'
  readonly directionalMeanRatioGeomean: number
  readonly directionalP95RatioGeomean: number
  readonly meanAndP95WinCount: number
  readonly meanWinCount: number
  readonly p95WinCount: number
  readonly truecalcMeanWinCount: number
  readonly truecalcP95WinCount: number
  readonly worstMeanRatioWorkload: WorkPaperTrueCalcScalarWorkload
  readonly worstP95RatioWorkload: WorkPaperTrueCalcScalarWorkload
  readonly worstWorkpaperToTrueCalcMeanRatio: number
  readonly worstWorkpaperToTrueCalcP95Ratio: number
}

export interface WorkPaperTrueCalcScalarBenchmarkReport {
  readonly suite: 'workpaper-vs-truecalc-scalar'
  readonly scorecard: WorkPaperTrueCalcScalarScorecard
  readonly results: readonly WorkPaperTrueCalcScalarBenchmarkResult[]
}

interface ResolvedBenchmarkSuiteOptions {
  readonly sampleCount: number
  readonly warmupCount: number
}

export const TRUECALC_SCALAR_WORKLOADS = [
  'scalar-arithmetic',
  'scalar-branching',
  'scalar-financial-pmt',
  'scalar-math-nested',
  'scalar-text-concat',
  'scalar-text-length',
  'scalar-minmax',
] as const satisfies readonly WorkPaperTrueCalcScalarWorkload[]

export function runWorkPaperVsTrueCalcScalarBenchmarkSuite(
  options: ComparativeBenchmarkSuiteOptions = {},
): WorkPaperTrueCalcScalarBenchmarkResult[] {
  const resolvedOptions = resolveSuiteOptions(options)
  return TRUECALC_SCALAR_WORKLOADS.map((workload) => runScalarScenario(workload, scalarFixture(workload), resolvedOptions))
}

export function buildWorkPaperVsTrueCalcScalarBenchmarkReport(
  results: readonly WorkPaperTrueCalcScalarBenchmarkResult[],
): WorkPaperTrueCalcScalarBenchmarkReport {
  if (results.length === 0) {
    throw new Error('Cannot build a WorkPaper vs TrueCalc scorecard without benchmark results')
  }
  const meanWinCount = results.filter((result) => result.comparison.workpaperToTrueCalcMeanRatio < 1).length
  const p95WinCount = results.filter((result) => result.comparison.workpaperToTrueCalcP95Ratio < 1).length
  const meanAndP95WinCount = results.filter(
    (result) => result.comparison.workpaperToTrueCalcMeanRatio < 1 && result.comparison.workpaperToTrueCalcP95Ratio < 1,
  ).length

  return {
    suite: 'workpaper-vs-truecalc-scalar',
    scorecard: {
      comparableWorkloadCount: results.length,
      coverageNote:
        'TrueCalc is covered through its public scalar formula API. Its simple API does not cover workbook dependency graphs, ranges, structural edits, or full-sheet recalculation.',
      coverageTier: 'scalar-formula',
      directionalMeanRatioGeomean: geometricMean(results.map((result) => result.comparison.workpaperToTrueCalcMeanRatio)),
      directionalP95RatioGeomean: geometricMean(results.map((result) => result.comparison.workpaperToTrueCalcP95Ratio)),
      meanAndP95WinCount,
      meanWinCount,
      p95WinCount,
      truecalcMeanWinCount: results.length - meanWinCount,
      truecalcP95WinCount: results.length - p95WinCount,
      worstMeanRatioWorkload: maxComparableRatioWorkload(results, 'workpaperToTrueCalcMeanRatio'),
      worstP95RatioWorkload: maxComparableRatioWorkload(results, 'workpaperToTrueCalcP95Ratio'),
      worstWorkpaperToTrueCalcMeanRatio: maxComparableRatio(results, 'workpaperToTrueCalcMeanRatio'),
      worstWorkpaperToTrueCalcP95Ratio: maxComparableRatio(results, 'workpaperToTrueCalcP95Ratio'),
    },
    results,
  }
}

function runScalarScenario(
  workload: WorkPaperTrueCalcScalarWorkload,
  fixture: WorkPaperTrueCalcScalarFixture,
  options: ResolvedBenchmarkSuiteOptions,
): WorkPaperTrueCalcScalarBenchmarkResult {
  const workpaper = benchmarkSupportedEngine(() => measureWorkPaperScalarEvaluationSample(fixture), options)
  const truecalc = benchmarkSupportedEngine(() => measureTrueCalcScalarEvaluationSample(fixture), options)
  const workPaperVerification = JSON.stringify(workpaper.verification)
  const trueCalcVerification = JSON.stringify(truecalc.verification)
  if (workPaperVerification !== trueCalcVerification) {
    throw new Error(`Verification mismatch for ${workload}: WorkPaper ${workPaperVerification} !== TrueCalc ${trueCalcVerification}`)
  }

  const fasterEngine = workpaper.elapsedMs.mean <= truecalc.elapsedMs.mean ? 'workpaper' : 'truecalc'
  const fasterMean = fasterEngine === 'workpaper' ? workpaper.elapsedMs.mean : truecalc.elapsedMs.mean
  const slowerMean = fasterEngine === 'workpaper' ? truecalc.elapsedMs.mean : workpaper.elapsedMs.mean

  return {
    workload,
    category: 'scalar-formula',
    comparable: true,
    fixture,
    comparison: {
      confidenceIntervalOverlaps:
        workpaper.elapsedMs.confidence95.low <= truecalc.elapsedMs.confidence95.high &&
        truecalc.elapsedMs.confidence95.low <= workpaper.elapsedMs.confidence95.high,
      fasterEngine,
      maxRelativeNoise: Math.max(workpaper.elapsedMs.relativeStandardDeviation, truecalc.elapsedMs.relativeStandardDeviation),
      meanSpeedup: slowerMean / fasterMean,
      verificationEquivalent: true,
      workpaperToTrueCalcMeanRatio: workpaper.elapsedMs.mean / truecalc.elapsedMs.mean,
      workpaperToTrueCalcMedianRatio: workpaper.elapsedMs.median / truecalc.elapsedMs.median,
      workpaperToTrueCalcP95Ratio: workpaper.elapsedMs.p95 / truecalc.elapsedMs.p95,
    },
    engines: {
      truecalc,
      workpaper,
    },
  }
}

function measureWorkPaperScalarEvaluationSample(fixture: WorkPaperTrueCalcScalarFixture): BenchmarkSample {
  const workbook = WorkPaper.buildFromSheets({ Sheet1: fixture.sheet })
  const sheetId = workbook.getSheetId('Sheet1')
  if (sheetId === undefined) {
    workbook.dispose()
    throw new Error('WorkPaper TrueCalc benchmark fixture did not create Sheet1')
  }
  const memoryBefore = sampleMemory()
  const started = performance.now()
  workbook.setCellContents({ sheet: sheetId, row: fixture.edit.row, col: fixture.edit.col }, fixture.edit.value)
  const value = workbook.getCellValue({ sheet: sheetId, row: fixture.result.row, col: fixture.result.col })
  const elapsedMs = performance.now() - started
  const memoryAfter = sampleMemory()
  workbook.dispose()
  return {
    elapsedMs,
    memory: measureMemory(memoryBefore, memoryAfter),
    verification: {
      value: normalizeBenchmarkValue(normalizeWorkPaperValue(value)),
    },
  }
}

function measureTrueCalcScalarEvaluationSample(fixture: WorkPaperTrueCalcScalarFixture): BenchmarkSample {
  const engine = createEngine('google-sheets')
  const memoryBefore = sampleMemory()
  const started = performance.now()
  const value = engine.evaluate(fixture.formula, fixture.variables)
  const elapsedMs = performance.now() - started
  const memoryAfter = sampleMemory()
  engine.free()
  return {
    elapsedMs,
    memory: measureMemory(memoryBefore, memoryAfter),
    verification: {
      value: normalizeBenchmarkValue(normalizeTrueCalcValue(value)),
    },
  }
}

function benchmarkSupportedEngine(
  runSample: () => BenchmarkSample,
  options: ResolvedBenchmarkSuiteOptions,
): ComparativeMeasuredEngineResult {
  for (let warmup = 0; warmup < options.warmupCount; warmup += 1) {
    runSample()
  }

  const samples: BenchmarkSample[] = []
  for (let sample = 0; sample < options.sampleCount; sample += 1) {
    samples.push(runSample())
  }

  const verificationStrings = new Set(samples.map((sample) => JSON.stringify(sample.verification)))
  if (verificationStrings.size !== 1) {
    throw new Error('Benchmark verification drifted across samples')
  }

  return {
    status: 'supported',
    elapsedMs: summarizeNumbers(samples.map((sample) => sample.elapsedMs)),
    memoryDeltaBytes: summarizeMemory(samples.map((sample) => sample.memory)),
    verification: samples[0]?.verification ?? {},
  }
}

function summarizeMemory(samples: readonly MemoryMeasurement[]): ComparativeMemorySummary {
  return {
    rssBytes: summarizeNumbers(samples.map((sample) => sample.delta.rssBytes)),
    heapUsedBytes: summarizeNumbers(samples.map((sample) => sample.delta.heapUsedBytes)),
    heapTotalBytes: summarizeNumbers(samples.map((sample) => sample.delta.heapTotalBytes)),
    externalBytes: summarizeNumbers(samples.map((sample) => sample.delta.externalBytes)),
    arrayBuffersBytes: summarizeNumbers(samples.map((sample) => sample.delta.arrayBuffersBytes)),
  }
}

function scalarFixture(workload: WorkPaperTrueCalcScalarWorkload): WorkPaperTrueCalcScalarFixture {
  switch (workload) {
    case 'scalar-arithmetic':
      return {
        edit: { row: 0, col: 0, value: 101 },
        formula: 'A1+B1*2',
        result: { row: 0, col: 3 },
        sheet: [[100, 20, null, '=A1+B1*2']],
        variables: { A1: 101, B1: 20 },
      }
    case 'scalar-branching':
      return {
        edit: { row: 0, col: 0, value: -1 },
        formula: 'IF(A1>0,"yes","no")',
        result: { row: 0, col: 3 },
        sheet: [[100, null, null, '=IF(A1>0,"yes","no")']],
        variables: { A1: -1 },
      }
    case 'scalar-financial-pmt':
      return {
        edit: { row: 0, col: 0, value: 101 },
        formula: 'PMT(A1/12,A2,A3)',
        result: { row: 0, col: 3 },
        sheet: [[100, null, null, '=PMT(A1/12,A2,A3)'], [12], [1000]],
        variables: { A1: 101, A2: 12, A3: 1000 },
      }
    case 'scalar-math-nested':
      return {
        edit: { row: 0, col: 0, value: 121 },
        formula: 'ROUND(SQRT(A1),2)',
        result: { row: 0, col: 3 },
        sheet: [[100, null, null, '=ROUND(SQRT(A1),2)']],
        variables: { A1: 121 },
      }
    case 'scalar-text-concat':
      return {
        edit: { row: 0, col: 0, value: 'baz' },
        formula: 'CONCATENATE(A1,"-",B1)',
        result: { row: 0, col: 3 },
        sheet: [['foo', 'bar', null, '=CONCATENATE(A1,"-",B1)']],
        variables: { A1: 'baz', B1: 'bar' },
      }
    case 'scalar-text-length':
      return {
        edit: { row: 0, col: 1, value: 'quux' },
        formula: 'LEN(A1)+LEN(B1)',
        result: { row: 0, col: 3 },
        sheet: [['foo', 'bar', null, '=LEN(A1)+LEN(B1)']],
        variables: { A1: 'foo', B1: 'quux' },
      }
    case 'scalar-minmax':
      return {
        edit: { row: 0, col: 2, value: 200 },
        formula: 'MIN(A1,B1,C1)+MAX(A1,B1,C1)',
        result: { row: 0, col: 3 },
        sheet: [[100, 20, 5, '=MIN(A1,B1,C1)+MAX(A1,B1,C1)']],
        variables: { A1: 100, B1: 20, C1: 200 },
      }
  }
}

function normalizeTrueCalcValue(value: EvalResult): boolean | number | string | null | { error: string } {
  switch (value.type) {
    case 'bool':
    case 'number':
    case 'text':
      return value.value
    case 'empty':
      return null
    case 'error':
      return { error: value.error }
  }
}

function normalizeBenchmarkValue(
  value: boolean | number | string | null | { error: unknown },
): boolean | number | string | null | { error: string } {
  if (typeof value === 'number') {
    return Number(value.toPrecision(12))
  }
  if (isErrorRecord(value)) {
    return { error: String(value.error) }
  }
  return value
}

function isErrorRecord(value: unknown): value is { error: unknown } {
  return typeof value === 'object' && value !== null && 'error' in value
}

function geometricMean(values: readonly number[]): number {
  if (values.length === 0) {
    return Number.NaN
  }
  const totalLog = values.reduce((sum, value) => {
    if (value <= 0) {
      throw new Error(`Cannot compute geomean for non-positive value: ${String(value)}`)
    }
    return sum + Math.log(value)
  }, 0)
  return Math.exp(totalLog / values.length)
}

function maxComparableRatio(
  results: readonly WorkPaperTrueCalcScalarBenchmarkResult[],
  ratioKey: 'workpaperToTrueCalcMeanRatio' | 'workpaperToTrueCalcP95Ratio',
): number {
  return Math.max(...results.map((result) => result.comparison[ratioKey]))
}

function maxComparableRatioWorkload(
  results: readonly WorkPaperTrueCalcScalarBenchmarkResult[],
  ratioKey: 'workpaperToTrueCalcMeanRatio' | 'workpaperToTrueCalcP95Ratio',
): WorkPaperTrueCalcScalarWorkload {
  return results.reduce((worst, result) => (result.comparison[ratioKey] > worst.comparison[ratioKey] ? result : worst)).workload
}

function resolveSuiteOptions(options: ComparativeBenchmarkSuiteOptions): ResolvedBenchmarkSuiteOptions {
  return {
    sampleCount: options.sampleCount ?? DEFAULT_COMPETITIVE_SAMPLE_COUNT,
    warmupCount: options.warmupCount ?? DEFAULT_COMPETITIVE_WARMUP_COUNT,
  }
}
