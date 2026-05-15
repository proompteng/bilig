import type { ExtraHeadlessComparisonEngineSummary, HeadlessComparisonCoverageTier } from './headless-comparison-engine-summary.ts'
import { arrayField, asObject, literalField, numberField, objectField, stringArrayField, stringField } from './json-scorecard-helpers.ts'

export interface ParsedWorkPaperXlsxCalcArtifact {
  readonly generatedAt: string
  readonly engines: {
    readonly xlsxCalc: {
      readonly version: string
    }
  }
  readonly results: readonly ParsedWorkPaperXlsxCalcResult[]
  readonly scorecard: ParsedWorkPaperXlsxCalcScorecard
  readonly suite: 'workpaper-vs-xlsx-calc'
}

export interface ParsedWorkPaperXlsxCalcResult {
  readonly category: 'workbook-wide-limited'
  readonly comparable: true
  readonly comparison: {
    readonly workpaperToXlsxCalcMeanRatio: number
    readonly workpaperToXlsxCalcP95Ratio: number
  }
  readonly fixture: {
    readonly family: string
    readonly rowCount: number
  }
  readonly workload: string
}

export interface ParsedWorkPaperXlsxCalcScorecard {
  readonly comparableWorkloadCount: number
  readonly coverageNote: string
  readonly coverageTier: HeadlessComparisonCoverageTier
  readonly directionalMeanRatioGeomean: number
  readonly directionalP95RatioGeomean: number
  readonly meanAndP95WinCount: number
  readonly meanWinCount: number
  readonly p95WinCount: number
  readonly workloadFamilies: readonly string[]
  readonly worstMeanRatioWorkload: string
  readonly worstP95RatioWorkload: string
  readonly worstWorkpaperToXlsxCalcMeanRatio: number
  readonly worstWorkpaperToXlsxCalcP95Ratio: number
  readonly xlsxCalcMeanWinCount: number
  readonly xlsxCalcP95WinCount: number
}

export function parseWorkPaperXlsxCalcArtifact(value: Record<string, unknown>): ParsedWorkPaperXlsxCalcArtifact {
  const engines = objectField(value, 'engines')
  const xlsxCalc = objectField(engines, 'xlsxCalc')
  return {
    generatedAt: stringField(value, 'generatedAt'),
    engines: {
      xlsxCalc: {
        version: stringField(xlsxCalc, 'version'),
      },
    },
    results: arrayField(value, 'results').map(parseWorkPaperXlsxCalcResult),
    scorecard: parseWorkPaperXlsxCalcScorecard(objectField(value, 'scorecard')),
    suite: literalField(value, 'suite', 'workpaper-vs-xlsx-calc'),
  }
}

export function parseWorkPaperXlsxCalcExtraComparisonEngineSummary(
  value: Record<string, unknown>,
  artifactPath: string,
): ExtraHeadlessComparisonEngineSummary {
  const artifact = parseWorkPaperXlsxCalcArtifact(value)
  return {
    artifactPath,
    comparableWorkloadCount: artifact.scorecard.comparableWorkloadCount,
    coverageNote: artifact.scorecard.coverageNote,
    coverageTier: artifact.scorecard.coverageTier,
    engineName: 'xlsx-calc',
    generatedAt: artifact.generatedAt,
    meanAndP95WinCount: artifact.scorecard.meanAndP95WinCount,
    meanWinCount: artifact.scorecard.meanWinCount,
    p95WinCount: artifact.scorecard.p95WinCount,
    version: artifact.engines.xlsxCalc.version,
    workloadFamilies: artifact.scorecard.workloadFamilies,
  }
}

export function deriveWorkPaperXlsxCalcScorecard(
  results: readonly ParsedWorkPaperXlsxCalcResult[],
  coverageNote: string,
): ParsedWorkPaperXlsxCalcScorecard {
  if (results.length === 0) {
    throw new Error('Cannot derive a WorkPaper vs xlsx-calc scorecard without results')
  }
  const meanWinCount = results.filter((result) => result.comparison.workpaperToXlsxCalcMeanRatio < 1).length
  const p95WinCount = results.filter((result) => result.comparison.workpaperToXlsxCalcP95Ratio < 1).length
  const meanAndP95WinCount = results.filter(
    (result) => result.comparison.workpaperToXlsxCalcMeanRatio < 1 && result.comparison.workpaperToXlsxCalcP95Ratio < 1,
  ).length
  return {
    comparableWorkloadCount: results.length,
    coverageNote,
    coverageTier: 'workbook-wide',
    directionalMeanRatioGeomean: geometricMean(results.map((result) => result.comparison.workpaperToXlsxCalcMeanRatio)),
    directionalP95RatioGeomean: geometricMean(results.map((result) => result.comparison.workpaperToXlsxCalcP95Ratio)),
    meanAndP95WinCount,
    meanWinCount,
    p95WinCount,
    workloadFamilies: orderedUnique(results.map((result) => result.fixture.family)),
    worstMeanRatioWorkload: maxComparableRatioWorkload(results, 'workpaperToXlsxCalcMeanRatio'),
    worstP95RatioWorkload: maxComparableRatioWorkload(results, 'workpaperToXlsxCalcP95Ratio'),
    worstWorkpaperToXlsxCalcMeanRatio: maxComparableRatio(results, 'workpaperToXlsxCalcMeanRatio'),
    worstWorkpaperToXlsxCalcP95Ratio: maxComparableRatio(results, 'workpaperToXlsxCalcP95Ratio'),
    xlsxCalcMeanWinCount: results.length - meanWinCount,
    xlsxCalcP95WinCount: results.length - p95WinCount,
  }
}

function parseWorkPaperXlsxCalcScorecard(value: Record<string, unknown>): ParsedWorkPaperXlsxCalcScorecard {
  return {
    comparableWorkloadCount: numberField(value, 'comparableWorkloadCount'),
    coverageNote: stringField(value, 'coverageNote'),
    coverageTier: literalField(value, 'coverageTier', 'workbook-wide'),
    directionalMeanRatioGeomean: numberField(value, 'directionalMeanRatioGeomean'),
    directionalP95RatioGeomean: numberField(value, 'directionalP95RatioGeomean'),
    meanAndP95WinCount: numberField(value, 'meanAndP95WinCount'),
    meanWinCount: numberField(value, 'meanWinCount'),
    p95WinCount: numberField(value, 'p95WinCount'),
    workloadFamilies: stringArrayField(value, 'workloadFamilies'),
    worstMeanRatioWorkload: stringField(value, 'worstMeanRatioWorkload'),
    worstP95RatioWorkload: stringField(value, 'worstP95RatioWorkload'),
    worstWorkpaperToXlsxCalcMeanRatio: numberField(value, 'worstWorkpaperToXlsxCalcMeanRatio'),
    worstWorkpaperToXlsxCalcP95Ratio: numberField(value, 'worstWorkpaperToXlsxCalcP95Ratio'),
    xlsxCalcMeanWinCount: numberField(value, 'xlsxCalcMeanWinCount'),
    xlsxCalcP95WinCount: numberField(value, 'xlsxCalcP95WinCount'),
  }
}

function parseWorkPaperXlsxCalcResult(value: unknown): ParsedWorkPaperXlsxCalcResult {
  const result = asObject(value, 'WorkPaper xlsx-calc result')
  const comparison = objectField(result, 'comparison')
  const fixture = objectField(result, 'fixture')
  return {
    category: literalField(result, 'category', 'workbook-wide-limited'),
    comparable: literalField(result, 'comparable', true),
    comparison: {
      workpaperToXlsxCalcMeanRatio: numberField(comparison, 'workpaperToXlsxCalcMeanRatio'),
      workpaperToXlsxCalcP95Ratio: numberField(comparison, 'workpaperToXlsxCalcP95Ratio'),
    },
    fixture: {
      family: stringField(fixture, 'family'),
      rowCount: numberField(fixture, 'rowCount'),
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
  results: readonly ParsedWorkPaperXlsxCalcResult[],
  ratioKey: 'workpaperToXlsxCalcMeanRatio' | 'workpaperToXlsxCalcP95Ratio',
): number {
  return Math.max(...results.map((result) => result.comparison[ratioKey]))
}

function maxComparableRatioWorkload(
  results: readonly ParsedWorkPaperXlsxCalcResult[],
  ratioKey: 'workpaperToXlsxCalcMeanRatio' | 'workpaperToXlsxCalcP95Ratio',
): string {
  return results.reduce((worst, result) => (result.comparison[ratioKey] > worst.comparison[ratioKey] ? result : worst)).workload
}

function orderedUnique(values: readonly string[]): string[] {
  return [...new Set(values)]
}
