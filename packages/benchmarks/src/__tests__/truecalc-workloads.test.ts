import { describe, expect, it } from 'vitest'

import {
  TRUECALC_SCALAR_WORKLOADS,
  buildWorkPaperVsTrueCalcScalarBenchmarkReport,
  runWorkPaperVsTrueCalcScalarBenchmarkSuite,
  type WorkPaperTrueCalcScalarWorkload,
} from '../benchmark-workpaper-vs-truecalc.js'

const expectedVerificationByWorkload = {
  'scalar-arithmetic': { value: 141 },
  'scalar-branching': { value: 'no' },
  'scalar-financial-pmt': { value: -8416.66666668 },
  'scalar-math-nested': { value: 11 },
  'scalar-text-concat': { value: 'baz-bar' },
  'scalar-text-length': { value: 7 },
  'scalar-minmax': { value: 220 },
} as const satisfies Record<WorkPaperTrueCalcScalarWorkload, { readonly value: number | string }>

describe('WorkPaper vs TrueCalc scalar benchmark', () => {
  it('keeps the TrueCalc lane scoped to scalar formulas without range references', () => {
    expect(TRUECALC_SCALAR_WORKLOADS).toEqual([
      'scalar-arithmetic',
      'scalar-branching',
      'scalar-financial-pmt',
      'scalar-math-nested',
      'scalar-text-concat',
      'scalar-text-length',
      'scalar-minmax',
    ])

    const results = runWorkPaperVsTrueCalcScalarBenchmarkSuite({ sampleCount: 1, warmupCount: 0 })

    expect(results).toHaveLength(TRUECALC_SCALAR_WORKLOADS.length)
    for (const result of results) {
      const expectedVerification = expectedVerificationByWorkload[result.workload]
      expect(result.category).toBe('scalar-formula')
      expect(result.comparable).toBe(true)
      expect(result.fixture.formula).not.toContain(':')
      expect(result.comparison.verificationEquivalent).toBe(true)
      expect(result.engines.workpaper.verification).toEqual(expectedVerification)
      expect(result.engines.truecalc.verification).toEqual(expectedVerification)
      expect(Number.isFinite(result.comparison.workpaperToTrueCalcMeanRatio)).toBe(true)
      expect(Number.isFinite(result.comparison.workpaperToTrueCalcP95Ratio)).toBe(true)
    }
  })

  it('derives scorecard totals from measured scalar formula results', () => {
    const results = runWorkPaperVsTrueCalcScalarBenchmarkSuite({ sampleCount: 1, warmupCount: 0 })
    const report = buildWorkPaperVsTrueCalcScalarBenchmarkReport(results)

    expect(report.suite).toBe('workpaper-vs-truecalc-scalar')
    expect(report.scorecard.coverageTier).toBe('scalar-formula')
    expect(report.scorecard.comparableWorkloadCount).toBe(TRUECALC_SCALAR_WORKLOADS.length)
    expect(report.scorecard.meanWinCount + report.scorecard.truecalcMeanWinCount).toBe(TRUECALC_SCALAR_WORKLOADS.length)
    expect(report.scorecard.p95WinCount + report.scorecard.truecalcP95WinCount).toBe(TRUECALC_SCALAR_WORKLOADS.length)
    expect(report.scorecard.meanAndP95WinCount).toBeLessThanOrEqual(TRUECALC_SCALAR_WORKLOADS.length)
    expect(report.scorecard.worstMeanRatioWorkload).toEqual(expect.any(String))
    expect(report.scorecard.worstP95RatioWorkload).toEqual(expect.any(String))
  })
})
