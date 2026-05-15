import { describe, expect, it } from 'vitest'

import { buildHeadlessPerformanceLeadershipScorecard } from '../gen-headless-performance-leadership-scorecard.ts'
import type { CompetitiveArtifact } from '../bilig-dominance-scorecard-types.ts'

const competitiveArtifact: CompetitiveArtifact = {
  generatedAt: '2026-05-14T00:00:00.000Z',
  engines: {
    hyperformula: {
      commit: 'hf-commit',
      version: '3.2.0',
    },
  },
  families: [
    family('build', 2, 0.8, 'build-mixed-content', 0.9),
    family('dirty-execution', 1, 0.72, 'single-edit-recalc', 0.8),
    family('batch-edit', 1, 0.7, 'batch-edit-recalc', 0.78),
    family('structural-rows', 1, 0.65, 'structural-insert-rows', 0.75),
    family('structural-columns', 1, 0.62, 'structural-insert-columns', 0.74),
    family('range-read', 1, 0.3, 'range-read', 0.35),
    family('aggregate-2d', 1, 0.4, 'aggregate-2d-ranges', 0.45),
    family('conditional-aggregation', 1, 0.52, 'conditional-aggregation-reused-ranges', 0.6),
    family('lookup-exact', 1, 0.5, 'lookup-with-column-index', 0.55),
    family('lookup-approximate', 1, 0.91, 'lookup-approximate-duplicates', 1.04),
    {
      comparableCount: 1,
      family: 'config-toggle',
      hyperformulaWins: 0,
      scorecardEligible: false,
      workpaperWins: 1,
      worstMeanRatioWorkload: 'rebuild-config-toggle',
      worstP95RatioWorkload: 'rebuild-config-toggle',
      worstWorkpaperToHyperFormulaMeanRatio: 0.01,
      worstWorkpaperToHyperFormulaP95Ratio: 0.02,
    },
  ],
  results: [
    {
      comparable: true,
      workload: 'build-mixed-content',
      comparison: {
        workpaperToHyperFormulaMeanRatio: 0.8,
        workpaperToHyperFormulaP95Ratio: 0.9,
      },
    },
    {
      comparable: true,
      workload: 'build-dense-literals',
      comparison: {
        workpaperToHyperFormulaMeanRatio: 0.6,
        workpaperToHyperFormulaP95Ratio: 0.7,
      },
    },
    {
      comparable: true,
      workload: 'lookup-approximate-duplicates',
      comparison: {
        workpaperToHyperFormulaMeanRatio: 0.91,
        workpaperToHyperFormulaP95Ratio: 1.04,
      },
    },
    {
      comparable: false,
      workload: 'dynamic-array-sort',
    },
  ],
  scorecard: {
    comparableCount: 3,
    directionalMeanRatioGeomean: 0.72,
    directionalP95RatioGeomean: 0.85,
    hyperformulaWins: 0,
    worstMeanRatioWorkload: 'lookup-approximate-duplicates',
    worstP95RatioWorkload: 'lookup-approximate-duplicates',
    worstWorkpaperToHyperFormulaMeanRatio: 0.91,
    worstWorkpaperToHyperFormulaP95Ratio: 1.04,
    workpaperWins: 3,
  },
}

describe('headless performance leadership scorecard', () => {
  it('blocks broad headless leadership claims when competitor coverage is singular and p95 has a holdout', () => {
    const scorecard = buildHeadlessPerformanceLeadershipScorecard({
      competitiveArtifact,
      competitiveArtifactPath: 'packages/benchmarks/baselines/workpaper-vs-hyperformula.json',
    })

    expect(scorecard.goalStatus).toBe('active-not-achieved')
    expect(scorecard.claimPolicy.blanketHeadlessPerformanceLeadershipClaimAllowed).toBe(false)
    expect(scorecard.completionAudit.allCriteriaPassed).toBe(false)
    expect(scorecard.summary).toMatchObject({
      comparisonEngineCount: 1,
      comparisonEngines: ['HyperFormula'],
      comparableWorkloadCount: 3,
      meanWinCount: 3,
      p95WinCount: 2,
      meanAndP95WinCount: 2,
      workbookWideComparisonEngineCount: 1,
      workbookWideComparisonEngines: ['HyperFormula'],
      worstP95RatioWorkload: 'lookup-approximate-duplicates',
    })
    expect(scorecard.summary.p95Holdouts).toEqual([
      {
        comparisonTarget: 'HyperFormula',
        p95Ratio: 1.04,
        workload: 'lookup-approximate-duplicates',
      },
    ])
    expect(scorecard.completionAudit.unmetRequirements).toEqual([
      'competitor-coverage: only HyperFormula is workbook-wide; add at least one more direct workbook-wide headless spreadsheet engine before broad headless leadership claims',
      'per-workload-mean-and-p95-wins: 2/3 comparable workloads win both mean and p95; p95 holdouts: lookup-approximate-duplicates',
    ])
  })

  it('tracks scalar formula competitor evidence without treating it as workbook-wide coverage', () => {
    const scorecard = buildHeadlessPerformanceLeadershipScorecard({
      competitiveArtifact,
      competitiveArtifactPath: 'packages/benchmarks/baselines/workpaper-vs-hyperformula.json',
      extraComparisonEngines: [
        {
          artifactPath: 'packages/benchmarks/baselines/workpaper-vs-truecalc.json',
          comparableWorkloadCount: 7,
          coverageNote: 'TrueCalc simple API evaluates scalar formulas but not workbook/range dependency workloads.',
          coverageTier: 'scalar-formula',
          engineName: 'TrueCalc',
          generatedAt: '2026-05-14T00:00:00.000Z',
          meanAndP95WinCount: 7,
          meanWinCount: 7,
          p95WinCount: 7,
          version: '0.6.4',
          workloadFamilies: ['scalar-formula'],
        },
      ],
    })

    expect(scorecard.goalStatus).toBe('active-not-achieved')
    expect(scorecard.summary.comparisonEngines).toEqual(['HyperFormula', 'TrueCalc'])
    expect(scorecard.summary.workbookWideComparisonEngines).toEqual(['HyperFormula'])
    expect(scorecard.summary.limitedComparisonEngines).toEqual([
      {
        comparableWorkloadCount: 7,
        coverageTier: 'scalar-formula',
        engineName: 'TrueCalc',
      },
    ])
    expect(scorecard.completionAudit.unmetRequirements).toContain(
      'competitor-coverage: only HyperFormula is workbook-wide; add at least one more direct workbook-wide headless spreadsheet engine before broad headless leadership claims',
    )
  })

  it('tracks partial workbook-wide competitor evidence without completing the broad claim', () => {
    const scorecard = buildHeadlessPerformanceLeadershipScorecard({
      competitiveArtifact,
      competitiveArtifactPath: 'packages/benchmarks/baselines/workpaper-vs-hyperformula.json',
      extraComparisonEngines: [
        {
          artifactPath: 'packages/benchmarks/baselines/workpaper-vs-xlsx-calc.json',
          comparableWorkloadCount: 2,
          coverageNote: 'xlsx-calc covers a workbook-wide subset of equivalent aggregate and lookup recalculation workloads.',
          coverageTier: 'workbook-wide',
          engineName: 'xlsx-calc',
          generatedAt: '2026-05-14T00:00:00.000Z',
          meanAndP95WinCount: 2,
          meanWinCount: 2,
          p95WinCount: 2,
          version: '0.9.2',
          workloadFamilies: ['aggregate', 'lookup-exact'],
        },
      ],
    })

    expect(scorecard.goalStatus).toBe('active-not-achieved')
    expect(scorecard.summary.comparisonEngines).toEqual(['HyperFormula', 'xlsx-calc'])
    expect(scorecard.summary.workbookWideComparisonEngines).toEqual(['HyperFormula', 'xlsx-calc'])
    expect(scorecard.summary.limitedComparisonEngines).toEqual([])
    expect(scorecard.completionAudit.unmetRequirements).toContain(
      'per-workload-mean-and-p95-wins: 2/3 comparable workloads win both mean and p95; p95 holdouts: lookup-approximate-duplicates; xlsx-calc workbook-wide comparison is incomplete: covers 2/3 comparable workloads and has 2/2 mean+p95 wins',
    )
  })

  it('allows the claim only when every criterion is directly covered', () => {
    const achievedArtifact: CompetitiveArtifact = {
      ...competitiveArtifact,
      results: competitiveArtifact.results.map((result) =>
        result.workload === 'lookup-approximate-duplicates'
          ? {
              ...result,
              comparison: {
                workpaperToHyperFormulaMeanRatio: 0.7,
                workpaperToHyperFormulaP95Ratio: 0.8,
              },
            }
          : result,
      ),
      scorecard: {
        ...competitiveArtifact.scorecard,
        worstWorkpaperToHyperFormulaP95Ratio: 0.9,
      },
    }

    const scorecard = buildHeadlessPerformanceLeadershipScorecard({
      competitiveArtifact: achievedArtifact,
      competitiveArtifactPath: 'packages/benchmarks/baselines/workpaper-vs-hyperformula.json',
      extraComparisonEngines: [
        {
          artifactPath: 'packages/benchmarks/baselines/workpaper-vs-other-headless-engine.json',
          comparableWorkloadCount: 3,
          coverageTier: 'workbook-wide',
          engineName: 'OtherHeadlessEngine',
          generatedAt: '2026-05-14T00:00:00.000Z',
          meanAndP95WinCount: 3,
          meanWinCount: 3,
          p95WinCount: 3,
          version: '1.0.0',
          workloadFamilies: ['build', 'dirty-execution', 'lookup-approximate'],
        },
      ],
    })

    expect(scorecard.goalStatus).toBe('achieved')
    expect(scorecard.claimPolicy.blanketHeadlessPerformanceLeadershipClaimAllowed).toBe(true)
    expect(scorecard.completionAudit.unmetRequirements).toEqual([])
  })
})

function family(familyName: string, comparableCount: number, meanRatio: number, workload: string, p95Ratio: number) {
  return {
    comparableCount,
    family: familyName,
    hyperformulaWins: 0,
    scorecardEligible: true,
    workpaperWins: comparableCount,
    worstMeanRatioWorkload: workload,
    worstP95RatioWorkload: workload,
    worstWorkpaperToHyperFormulaMeanRatio: meanRatio,
    worstWorkpaperToHyperFormulaP95Ratio: p95Ratio,
  }
}
