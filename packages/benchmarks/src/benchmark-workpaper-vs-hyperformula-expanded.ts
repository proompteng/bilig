import type { ComparativeBenchmarkSuiteOptions } from './benchmark-workpaper-vs-hyperformula.js'
import { DEFAULT_COMPETITIVE_SAMPLE_COUNT, DEFAULT_COMPETITIVE_WARMUP_COUNT } from './benchmark-workpaper-vs-hyperformula.js'
import { runWorkPaperVsHyperFormulaExpandedBenchmarkSuite } from './benchmark-workpaper-vs-hyperformula-expanded-scenarios.js'
import type { ExpandedComparativeBenchmarkResult } from './benchmark-workpaper-vs-hyperformula-expanded-runner.js'
import { buildExpandedCompetitiveFamilyReport, type ExpandedCompetitiveFamilySummary } from './report-competitive-families.js'

export { EXPANDED_COMPARATIVE_WORKLOADS } from './expanded-competitive-workloads.js'
export type { ExpandedComparativeBenchmarkWorkload } from './expanded-competitive-workloads.js'
export { runWorkPaperVsHyperFormulaExpandedBenchmarkSuite } from './benchmark-workpaper-vs-hyperformula-expanded-scenarios.js'
export type {
  EngineCounterNumericSummary,
  ExpandedComparativeBenchmarkResult,
  ExpandedComparativeComparableResult,
  ExpandedComparativeLeadershipResult,
} from './benchmark-workpaper-vs-hyperformula-expanded-runner.js'

export interface ExpandedComparativeBenchmarkReport {
  suite: 'workpaper-vs-hyperformula'
  results: readonly ExpandedComparativeBenchmarkResult[]
  families: readonly ExpandedCompetitiveFamilySummary[]
  scorecard: ReturnType<typeof buildExpandedCompetitiveFamilyReport>['scorecard']
}

export function buildExpandedComparativeBenchmarkReport(
  results: readonly ExpandedComparativeBenchmarkResult[],
): ExpandedComparativeBenchmarkReport {
  const familyReport = buildExpandedCompetitiveFamilyReport(results)
  return {
    suite: familyReport.suite,
    results: [...results],
    families: familyReport.families,
    scorecard: familyReport.scorecard,
  }
}

if (import.meta.url === `file://${process.argv[1]}`) {
  const cliOptions = parseExpandedBenchmarkCliOptions(process.argv.slice(2))
  const benchmarkResults = runWorkPaperVsHyperFormulaExpandedBenchmarkSuite({
    sampleCount: cliOptions.sampleCount ?? DEFAULT_COMPETITIVE_SAMPLE_COUNT,
    warmupCount: cliOptions.warmupCount ?? DEFAULT_COMPETITIVE_WARMUP_COUNT,
  })
  console.log(JSON.stringify(buildExpandedComparativeBenchmarkReport(benchmarkResults), null, 2))
}

export function parseExpandedBenchmarkCliOptions(args: readonly string[]): ComparativeBenchmarkSuiteOptions {
  const options: ComparativeBenchmarkSuiteOptions = {}
  for (let index = 0; index < args.length; index += 1) {
    const arg = args[index]!
    if (arg === '--sample-count') {
      const raw = args[index + 1]
      if (raw === undefined) {
        throw new Error('Missing value for --sample-count')
      }
      options.sampleCount = parsePositiveDecimalInteger(raw, '--sample-count')
      index += 1
      continue
    }
    if (arg === '--warmup-count') {
      const raw = args[index + 1]
      if (raw === undefined) {
        throw new Error('Missing value for --warmup-count')
      }
      options.warmupCount = parseNonNegativeDecimalInteger(raw, '--warmup-count')
      index += 1
      continue
    }
    throw new Error(`Unknown expanded benchmark argument: ${arg}`)
  }
  return options
}

function parsePositiveDecimalInteger(value: string, option: string): number {
  const parsed = parseNonNegativeDecimalInteger(value, option)
  if (parsed < 1) {
    throw new Error(`${option} expects a positive integer, got ${value}`)
  }
  return parsed
}

function parseNonNegativeDecimalInteger(value: string, option: string): number {
  if (!/^(?:0|[1-9]\d*)$/u.test(value)) {
    throw new Error(`${option} expects a non-negative integer, got ${value}`)
  }
  const parsed = Number(value)
  if (!Number.isSafeInteger(parsed)) {
    throw new Error(`${option} expects a safe integer, got ${value}`)
  }
  return parsed
}
