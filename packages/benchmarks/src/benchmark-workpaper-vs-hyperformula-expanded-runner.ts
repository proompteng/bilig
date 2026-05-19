import { ENGINE_COUNTER_KEYS, type EngineCounters } from '../../core/src/perf/engine-counters.js'
import type {
  ComparativeBenchmarkSuiteOptions,
  ComparativeMeasuredEngineResult,
  ComparativeMemorySummary,
  ComparativeUnsupportedEngineResult,
} from './benchmark-workpaper-vs-hyperformula.js'
import type { ExpandedComparativeBenchmarkWorkload } from './expanded-competitive-workloads.js'
import type { MemoryMeasurement } from './metrics.js'
import { summarizeNumbers, type NumericSummary } from './stats.js'
import type { BenchmarkSample } from './benchmark-workpaper-vs-hyperformula-expanded-support.js'

export interface ExpandedComparativeComparableResult {
  workload: ExpandedComparativeBenchmarkWorkload
  category: 'directly-comparable'
  comparable: true
  fixture: Record<string, unknown>
  comparison: {
    fasterEngine: 'workpaper' | 'hyperformula'
    meanSpeedup: number
    workpaperToHyperFormulaMeanRatio: number
    workpaperToHyperFormulaMedianRatio: number
    workpaperToHyperFormulaP95Ratio: number
    maxRelativeNoise: number
    confidenceIntervalOverlaps: boolean
    resultConfidence: 'decisive' | 'inconclusive'
    decisiveFasterEngine: 'workpaper' | 'hyperformula' | 'inconclusive'
    verificationEquivalent: true
  }
  engines: {
    workpaper: ComparativeMeasuredEngineResult
    hyperformula: ComparativeMeasuredEngineResult
  }
}

export interface ExpandedComparativeLeadershipResult {
  workload: ExpandedComparativeBenchmarkWorkload
  category: 'leadership'
  comparable: false
  fixture: Record<string, unknown>
  note: string
  engines: {
    workpaper: ComparativeMeasuredEngineResult
    hyperformula: ComparativeUnsupportedEngineResult
  }
}

type EngineCounterSummary = Record<keyof EngineCounters, number>

export type EngineCounterNumericSummary = Record<keyof EngineCounters, NumericSummary>

export type ExpandedComparativeBenchmarkResult = ExpandedComparativeComparableResult | ExpandedComparativeLeadershipResult

export function runComparableScenario(
  workload: ExpandedComparativeBenchmarkWorkload,
  fixture: Record<string, unknown>,
  options: Required<ComparativeBenchmarkSuiteOptions>,
  runWorkPaperSample: () => BenchmarkSample,
  runHyperFormulaSample: () => BenchmarkSample,
): ExpandedComparativeComparableResult {
  const workpaper = benchmarkSupportedEngine(runWorkPaperSample, options)
  const hyperformula = benchmarkSupportedEngine(runHyperFormulaSample, options)
  const workPaperVerification = JSON.stringify(workpaper.verification)
  const hyperFormulaVerification = JSON.stringify(hyperformula.verification)
  if (workPaperVerification !== hyperFormulaVerification) {
    throw new Error(
      `Verification mismatch for ${workload}: WorkPaper ${workPaperVerification} !== HyperFormula ${hyperFormulaVerification}`,
    )
  }

  const fasterEngine = workpaper.elapsedMs.mean <= hyperformula.elapsedMs.mean ? 'workpaper' : 'hyperformula'
  const fasterMean = fasterEngine === 'workpaper' ? workpaper.elapsedMs.mean : hyperformula.elapsedMs.mean
  const slowerMean = fasterEngine === 'workpaper' ? hyperformula.elapsedMs.mean : workpaper.elapsedMs.mean
  const confidenceIntervalOverlaps =
    workpaper.elapsedMs.confidence95.low <= hyperformula.elapsedMs.confidence95.high &&
    hyperformula.elapsedMs.confidence95.low <= workpaper.elapsedMs.confidence95.high
  const resultConfidence = confidenceIntervalOverlaps ? 'inconclusive' : 'decisive'

  return {
    workload,
    category: 'directly-comparable',
    comparable: true,
    fixture,
    comparison: {
      fasterEngine,
      meanSpeedup: slowerMean / fasterMean,
      workpaperToHyperFormulaMeanRatio: workpaper.elapsedMs.mean / hyperformula.elapsedMs.mean,
      workpaperToHyperFormulaMedianRatio: workpaper.elapsedMs.median / hyperformula.elapsedMs.median,
      workpaperToHyperFormulaP95Ratio: workpaper.elapsedMs.p95 / hyperformula.elapsedMs.p95,
      maxRelativeNoise: Math.max(workpaper.elapsedMs.relativeStandardDeviation, hyperformula.elapsedMs.relativeStandardDeviation),
      confidenceIntervalOverlaps,
      resultConfidence,
      decisiveFasterEngine: resultConfidence === 'decisive' ? fasterEngine : 'inconclusive',
      verificationEquivalent: true,
    },
    engines: {
      workpaper,
      hyperformula,
    },
  }
}

export function runLeadershipScenario(
  workload: ExpandedComparativeBenchmarkWorkload,
  fixture: Record<string, unknown>,
  options: Required<ComparativeBenchmarkSuiteOptions>,
  runWorkPaperSample: () => BenchmarkSample,
  hyperformula: ComparativeUnsupportedEngineResult,
): ExpandedComparativeLeadershipResult {
  return {
    workload,
    category: 'leadership',
    comparable: false,
    fixture,
    note: 'This workload demonstrates capability leadership and is not an apples-to-apples speed comparison.',
    engines: {
      workpaper: benchmarkSupportedEngine(runWorkPaperSample, options),
      hyperformula,
    },
  }
}

function benchmarkSupportedEngine(
  runSample: () => BenchmarkSample,
  options: Required<ComparativeBenchmarkSuiteOptions>,
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
  const engineCounters = summarizeEngineCounters(samples)

  return {
    status: 'supported',
    elapsedMs: summarizeNumbers(samples.map((sample) => sample.elapsedMs)),
    memoryDeltaBytes: summarizeMemory(samples.map((sample) => sample.memory)),
    ...(engineCounters ? { engineCounters } : {}),
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

function summarizeEngineCounters(samples: readonly BenchmarkSample[]): EngineCounterNumericSummary | undefined {
  const counterSamples = samples
    .map((sample) => sample.engineCounters)
    .filter((counters): counters is EngineCounterSummary => counters !== undefined)
  if (counterSamples.length === 0) {
    return undefined
  }
  const zeroSummary = summarizeNumbers([0])
  const summaries: EngineCounterNumericSummary = {
    cellsRemapped: zeroSummary,
    rangesMaterialized: zeroSummary,
    rangeMembersExpanded: zeroSummary,
    formulasParsed: zeroSummary,
    formulasBound: zeroSummary,
    columnSliceBuilds: zeroSummary,
    exactIndexBuilds: zeroSummary,
    approxIndexBuilds: zeroSummary,
    topoRebuilds: zeroSummary,
    changedCellPayloadsBuilt: zeroSummary,
    snapshotOpsReplayed: zeroSummary,
    wasmFullUploads: zeroSummary,
    directAggregateScanEvaluations: zeroSummary,
    directAggregateScanCells: zeroSummary,
    directAggregatePrefixEvaluations: zeroSummary,
    nativeDirectAggregatePrefixEvaluations: zeroSummary,
    directAggregateDeltaApplications: zeroSummary,
    directAggregateDeltaOnlyRecalcSkips: zeroSummary,
    directScalarDeltaApplications: zeroSummary,
    directScalarDeltaOnlyRecalcSkips: zeroSummary,
    kernelSyncOnlyRecalcSkips: zeroSummary,
    directFormulaKernelSyncOnlyRecalcSkips: zeroSummary,
    directFormulaInitialEvaluations: zeroSummary,
    nativeDirectScalarInitialEvaluations: zeroSummary,
    nativeDirectScalarRecalcEvaluations: zeroSummary,
    nativeDirectLookupInitialEvaluations: zeroSummary,
    nativeDirectCriteriaAggregateEvaluations: zeroSummary,
    directCriteriaMatchCacheHits: zeroSummary,
    directCriteriaAggregateCacheHits: zeroSummary,
    structuralTransactions: zeroSummary,
    structuralPlannedCells: zeroSummary,
    structuralSurvivorCellsRemapped: zeroSummary,
    structuralRemovedCells: zeroSummary,
    structuralUndoCapturedCells: zeroSummary,
    structuralUndoCapturedFormulas: zeroSummary,
    structuralUndoFormulaDependencyScans: zeroSummary,
    structuralFormulaImpactCandidates: zeroSummary,
    structuralFormulaRebindInputs: zeroSummary,
    structuralRangeRetargets: zeroSummary,
    sheetGridBlockScans: zeroSummary,
    axisMapSplices: zeroSummary,
    axisMapMoves: zeroSummary,
    regionQueryIndexBuilds: zeroSummary,
    columnOwnerBuilds: zeroSummary,
    lookupOwnerBuilds: zeroSummary,
    calcChainFullScans: zeroSummary,
    cycleFormulaScans: zeroSummary,
    topoRepairs: zeroSummary,
    topoRepairFailures: zeroSummary,
    topoRepairAffectedFormulas: zeroSummary,
  }
  for (const key of ENGINE_COUNTER_KEYS) {
    summaries[key] = summarizeNumbers(counterSamples.map((counters) => counters[key]))
  }
  return summaries
}
