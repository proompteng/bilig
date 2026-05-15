import type { ExtraHeadlessComparisonEngineSummary, HeadlessComparisonCoverageTier } from './headless-comparison-engine-summary.ts'
import { arrayField, asObject, literalField, numberField, objectField, stringField } from './json-scorecard-helpers.ts'

export interface ParsedWorkPaperTrueCalcScalarArtifact {
  readonly generatedAt: string
  readonly engines: {
    readonly truecalc: {
      readonly version: string
    }
  }
  readonly results: readonly ParsedWorkPaperTrueCalcScalarResult[]
  readonly scorecard: ParsedWorkPaperTrueCalcScalarScorecard
  readonly suite: 'workpaper-vs-truecalc-scalar'
}

export interface ParsedWorkPaperTrueCalcScalarResult {
  readonly category: 'scalar-formula'
  readonly comparable: true
  readonly comparison: {
    readonly workpaperToTrueCalcMeanRatio: number
    readonly workpaperToTrueCalcP95Ratio: number
  }
  readonly fixture: {
    readonly formula: string
  }
  readonly workload: string
}

export interface ParsedWorkPaperTrueCalcScalarScorecard {
  readonly comparableWorkloadCount: number
  readonly coverageNote: string
  readonly coverageTier: HeadlessComparisonCoverageTier
  readonly directionalMeanRatioGeomean: number
  readonly directionalP95RatioGeomean: number
  readonly meanAndP95WinCount: number
  readonly meanWinCount: number
  readonly p95WinCount: number
  readonly truecalcMeanWinCount: number
  readonly truecalcP95WinCount: number
  readonly worstMeanRatioWorkload: string
  readonly worstP95RatioWorkload: string
  readonly worstWorkpaperToTrueCalcMeanRatio: number
  readonly worstWorkpaperToTrueCalcP95Ratio: number
}

export function parseWorkPaperTrueCalcScalarArtifact(value: Record<string, unknown>): ParsedWorkPaperTrueCalcScalarArtifact {
  const engines = objectField(value, 'engines')
  const truecalc = objectField(engines, 'truecalc')
  return {
    generatedAt: stringField(value, 'generatedAt'),
    engines: {
      truecalc: {
        version: stringField(truecalc, 'version'),
      },
    },
    results: arrayField(value, 'results').map(parseWorkPaperTrueCalcScalarResult),
    scorecard: parseWorkPaperTrueCalcScalarScorecard(objectField(value, 'scorecard')),
    suite: literalField(value, 'suite', 'workpaper-vs-truecalc-scalar'),
  }
}

export function parseWorkPaperTrueCalcExtraComparisonEngineSummary(
  value: Record<string, unknown>,
  artifactPath: string,
): ExtraHeadlessComparisonEngineSummary {
  const artifact = parseWorkPaperTrueCalcScalarArtifact(value)
  return {
    artifactPath,
    comparableWorkloadCount: artifact.scorecard.comparableWorkloadCount,
    coverageNote: artifact.scorecard.coverageNote,
    coverageTier: artifact.scorecard.coverageTier,
    engineName: 'TrueCalc',
    generatedAt: artifact.generatedAt,
    meanAndP95WinCount: artifact.scorecard.meanAndP95WinCount,
    meanWinCount: artifact.scorecard.meanWinCount,
    p95WinCount: artifact.scorecard.p95WinCount,
    version: artifact.engines.truecalc.version,
    workloadFamilies: ['scalar-formula'],
  }
}

export function deriveWorkPaperTrueCalcScalarScorecard(
  results: readonly ParsedWorkPaperTrueCalcScalarResult[],
  coverageNote: string,
): ParsedWorkPaperTrueCalcScalarScorecard {
  if (results.length === 0) {
    throw new Error('Cannot derive a WorkPaper vs TrueCalc scorecard without results')
  }
  const meanWinCount = results.filter((result) => result.comparison.workpaperToTrueCalcMeanRatio < 1).length
  const p95WinCount = results.filter((result) => result.comparison.workpaperToTrueCalcP95Ratio < 1).length
  const meanAndP95WinCount = results.filter(
    (result) => result.comparison.workpaperToTrueCalcMeanRatio < 1 && result.comparison.workpaperToTrueCalcP95Ratio < 1,
  ).length
  return {
    comparableWorkloadCount: results.length,
    coverageNote,
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
  }
}

function parseWorkPaperTrueCalcScalarScorecard(value: Record<string, unknown>): ParsedWorkPaperTrueCalcScalarScorecard {
  return {
    comparableWorkloadCount: numberField(value, 'comparableWorkloadCount'),
    coverageNote: stringField(value, 'coverageNote'),
    coverageTier: literalField(value, 'coverageTier', 'scalar-formula'),
    directionalMeanRatioGeomean: numberField(value, 'directionalMeanRatioGeomean'),
    directionalP95RatioGeomean: numberField(value, 'directionalP95RatioGeomean'),
    meanAndP95WinCount: numberField(value, 'meanAndP95WinCount'),
    meanWinCount: numberField(value, 'meanWinCount'),
    p95WinCount: numberField(value, 'p95WinCount'),
    truecalcMeanWinCount: numberField(value, 'truecalcMeanWinCount'),
    truecalcP95WinCount: numberField(value, 'truecalcP95WinCount'),
    worstMeanRatioWorkload: stringField(value, 'worstMeanRatioWorkload'),
    worstP95RatioWorkload: stringField(value, 'worstP95RatioWorkload'),
    worstWorkpaperToTrueCalcMeanRatio: numberField(value, 'worstWorkpaperToTrueCalcMeanRatio'),
    worstWorkpaperToTrueCalcP95Ratio: numberField(value, 'worstWorkpaperToTrueCalcP95Ratio'),
  }
}

function parseWorkPaperTrueCalcScalarResult(value: unknown): ParsedWorkPaperTrueCalcScalarResult {
  const result = asObject(value, 'WorkPaper TrueCalc scalar result')
  const comparison = objectField(result, 'comparison')
  return {
    category: literalField(result, 'category', 'scalar-formula'),
    comparable: literalField(result, 'comparable', true),
    comparison: {
      workpaperToTrueCalcMeanRatio: numberField(comparison, 'workpaperToTrueCalcMeanRatio'),
      workpaperToTrueCalcP95Ratio: numberField(comparison, 'workpaperToTrueCalcP95Ratio'),
    },
    fixture: {
      formula: stringField(objectField(result, 'fixture'), 'formula'),
    },
    workload: stringField(result, 'workload'),
  }
}

function geometricMean(values: readonly number[]): number {
  const totalLog = values.reduce((sum, value) => {
    if (value <= 0) {
      throw new Error(`Cannot compute geomean for non-positive value: ${String(value)}`)
    }
    return sum + Math.log(value)
  }, 0)
  return Math.exp(totalLog / values.length)
}

function maxComparableRatio(
  results: readonly ParsedWorkPaperTrueCalcScalarResult[],
  ratioKey: 'workpaperToTrueCalcMeanRatio' | 'workpaperToTrueCalcP95Ratio',
): number {
  return Math.max(...results.map((result) => result.comparison[ratioKey]))
}

function maxComparableRatioWorkload(
  results: readonly ParsedWorkPaperTrueCalcScalarResult[],
  ratioKey: 'workpaperToTrueCalcMeanRatio' | 'workpaperToTrueCalcP95Ratio',
): string {
  return results.reduce((worst, result) => (result.comparison[ratioKey] > worst.comparison[ratioKey] ? result : worst)).workload
}
