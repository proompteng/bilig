import { describe, expect, it } from 'vitest'

import {
  WORKPAPER_XLSX_CALC_WORKLOADS,
  buildWorkPaperVsXlsxCalcBenchmarkReport,
  runWorkPaperVsXlsxCalcBenchmarkSuite,
  type WorkPaperXlsxCalcWorkload,
} from '../benchmark-workpaper-vs-xlsx-calc.js'

const expectedVerificationByWorkload = {
  'xlsx-calc-large-sum-recalc': { value: 12_510_000 },
  'xlsx-calc-exact-match-recalc': { value: 31_000 },
  'xlsx-calc-approximate-match-recalc': { value: 31_000 },
  'xlsx-calc-formula-chain-recalc': { value: 11_998 },
} as const satisfies Record<WorkPaperXlsxCalcWorkload, { readonly value: number }>

describe('WorkPaper vs xlsx-calc benchmark', () => {
  it('covers a limited workbook-wide lane with real worksheet recalculation workloads', () => {
    expect(WORKPAPER_XLSX_CALC_WORKLOADS).toEqual([
      'xlsx-calc-large-sum-recalc',
      'xlsx-calc-exact-match-recalc',
      'xlsx-calc-approximate-match-recalc',
      'xlsx-calc-formula-chain-recalc',
    ])

    const results = runWorkPaperVsXlsxCalcBenchmarkSuite({ sampleCount: 1, warmupCount: 0 })

    expect(results).toHaveLength(WORKPAPER_XLSX_CALC_WORKLOADS.length)
    for (const result of results) {
      const expectedVerification = expectedVerificationByWorkload[result.workload]
      expect(result.category).toBe('workbook-wide-limited')
      expect(result.comparable).toBe(true)
      expect(result.comparison.verificationEquivalent).toBe(true)
      expect(result.engines.workpaper.verification).toEqual(expectedVerification)
      expect(result.engines.xlsxCalc.verification).toEqual(expectedVerification)
      expect(result.fixture.rowCount).toBeGreaterThanOrEqual(1000)
      expect(Number.isFinite(result.comparison.workpaperToXlsxCalcMeanRatio)).toBe(true)
      expect(Number.isFinite(result.comparison.workpaperToXlsxCalcP95Ratio)).toBe(true)
    }
  })

  it('derives scorecard totals from measured xlsx-calc results', () => {
    const results = runWorkPaperVsXlsxCalcBenchmarkSuite({ sampleCount: 1, warmupCount: 0 })
    const report = buildWorkPaperVsXlsxCalcBenchmarkReport(results)

    expect(report.suite).toBe('workpaper-vs-xlsx-calc')
    expect(report.scorecard.coverageTier).toBe('workbook-wide')
    expect(report.scorecard.comparableWorkloadCount).toBe(WORKPAPER_XLSX_CALC_WORKLOADS.length)
    expect(report.scorecard.meanWinCount + report.scorecard.xlsxCalcMeanWinCount).toBe(WORKPAPER_XLSX_CALC_WORKLOADS.length)
    expect(report.scorecard.p95WinCount + report.scorecard.xlsxCalcP95WinCount).toBe(WORKPAPER_XLSX_CALC_WORKLOADS.length)
    expect(report.scorecard.meanAndP95WinCount).toBeLessThanOrEqual(WORKPAPER_XLSX_CALC_WORKLOADS.length)
    expect(report.scorecard.workloadFamilies).toEqual(['aggregate', 'lookup-exact', 'lookup-approximate', 'formula-chain'])
  })
})
