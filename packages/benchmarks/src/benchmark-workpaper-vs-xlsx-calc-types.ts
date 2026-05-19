import type { ComparativeMeasuredEngineResult } from './benchmark-workpaper-vs-hyperformula.js'

export type WorkPaperXlsxCalcWorkload =
  | 'xlsx-calc-large-sum-recalc'
  | 'xlsx-calc-2d-sum-recalc'
  | 'xlsx-calc-overlapping-sum-recalc'
  | 'xlsx-calc-exact-match-recalc'
  | 'xlsx-calc-approximate-match-recalc'
  | 'xlsx-calc-formula-chain-recalc'
  | 'xlsx-calc-scalar-fanout-recalc'
  | 'xlsx-calc-deep-chain-recalc'
  | 'xlsx-calc-cross-sheet-sum-recalc'
  | 'xlsx-calc-cross-sheet-chain-recalc'
  | 'xlsx-calc-cross-sheet-scalar-fanout-recalc'
  | 'xlsx-calc-index-match-exact-text-recalc'
  | 'xlsx-calc-vlookup-text-exact-recalc'
  | 'xlsx-calc-vlookup-approximate-duplicates-recalc'
  | 'xlsx-calc-hlookup-exact-numeric-recalc'
  | 'xlsx-calc-range-stats-recalc'

export type WorkPaperXlsxCalcWorkloadFamily =
  | 'aggregate'
  | 'aggregate-2d'
  | 'overlapping-aggregate'
  | 'lookup-exact'
  | 'lookup-approximate'
  | 'formula-chain'
  | 'formula-fanout'
  | 'cross-sheet'
  | 'range-stats'

export type XlsxCalcEditableValue = boolean | number | string | null

export interface WorkPaperXlsxCalcFixture {
  readonly edit: {
    readonly address: string
    readonly col: number
    readonly row: number
    readonly sheetName: string
    readonly value: XlsxCalcEditableValue
  }
  readonly family: WorkPaperXlsxCalcWorkloadFamily
  readonly formula: string
  readonly result: {
    readonly address: string
    readonly col: number
    readonly row: number
    readonly sheetName: string
  }
  readonly rowCount: number
}

export interface WorkPaperXlsxCalcComparison {
  readonly confidenceIntervalOverlaps: boolean
  readonly fasterEngine: 'workpaper' | 'xlsx-calc'
  readonly maxRelativeNoise: number
  readonly meanSpeedup: number
  readonly verificationEquivalent: true
  readonly workpaperToXlsxCalcMeanRatio: number
  readonly workpaperToXlsxCalcMedianRatio: number
  readonly workpaperToXlsxCalcP95Ratio: number
}

export interface WorkPaperXlsxCalcBenchmarkResult {
  readonly workload: WorkPaperXlsxCalcWorkload
  readonly category: 'workbook-wide-limited'
  readonly comparable: true
  readonly fixture: WorkPaperXlsxCalcFixture
  readonly comparison: WorkPaperXlsxCalcComparison
  readonly engines: {
    readonly workpaper: ComparativeMeasuredEngineResult
    readonly xlsxCalc: ComparativeMeasuredEngineResult
  }
}

export interface WorkPaperXlsxCalcScorecard {
  readonly comparableWorkloadCount: number
  readonly coverageNote: string
  readonly coverageTier: 'workbook-wide-limited'
  readonly directionalMeanRatioGeomean: number
  readonly directionalP95RatioGeomean: number
  readonly meanAndP95WinCount: number
  readonly meanWinCount: number
  readonly p95WinCount: number
  readonly workloadFamilies: readonly WorkPaperXlsxCalcWorkloadFamily[]
  readonly worstMeanRatioWorkload: WorkPaperXlsxCalcWorkload
  readonly worstP95RatioWorkload: WorkPaperXlsxCalcWorkload
  readonly worstWorkpaperToXlsxCalcMeanRatio: number
  readonly worstWorkpaperToXlsxCalcP95Ratio: number
  readonly xlsxCalcMeanWinCount: number
  readonly xlsxCalcP95WinCount: number
}

export interface WorkPaperXlsxCalcBenchmarkReport {
  readonly suite: 'workpaper-vs-xlsx-calc'
  readonly scorecard: WorkPaperXlsxCalcScorecard
  readonly results: readonly WorkPaperXlsxCalcBenchmarkResult[]
}
