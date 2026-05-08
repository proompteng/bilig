import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
import { ENGINE_COUNTER_KEYS } from '../../../core/src/perf/engine-counters.js'
import {
  EXPANDED_COMPARATIVE_WORKLOADS,
  buildExpandedComparativeBenchmarkReport,
  type ExpandedComparativeBenchmarkResult,
  type ExpandedComparativeBenchmarkWorkload,
} from '../benchmark-workpaper-vs-hyperformula-expanded.js'
import {
  measureHyperFormula2dAggregateSample,
  measureHyperFormulaApproximateLookupDescendingSample,
  measureHyperFormulaApproximateLookupDuplicateSample,
  measureHyperFormulaConditionalAggregationMixedCriteriaSample,
  measureHyperFormulaConditionalAggregationSharedCriteriaSample,
  measureHyperFormulaNamedExpressionChangeSample,
  measureHyperFormulaParserCacheUniqueFormulaSample,
  measureHyperFormulaSheetRenameDependencySample,
  measureWorkPaper2dAggregateSample,
  measureWorkPaperApproximateLookupDescendingSample,
  measureWorkPaperApproximateLookupDuplicateSample,
  measureWorkPaperConditionalAggregationMixedCriteriaSample,
  measureWorkPaperConditionalAggregationSharedCriteriaSample,
  measureWorkPaperDynamicArraySortSample,
  measureWorkPaperDynamicArrayUniqueSample,
  measureWorkPaperNamedExpressionChangeSample,
  measureWorkPaperParserCacheUniqueFormulaSample,
  measureWorkPaperSheetRenameDependencySample,
  measureWorkPaperStructuralDeleteColumnsSample,
  measureWorkPaperStructuralDeleteRowsSample,
  measureWorkPaperStructuralInsertColumnsSample,
  measureWorkPaperStructuralInsertRowsSample,
  measureWorkPaperStructuralMoveColumnsSample,
  measureWorkPaperStructuralMoveRowsSample,
} from '../benchmark-workpaper-vs-hyperformula-expanded-additional-workloads.js'
import {
  EXPANDED_COMPARATIVE_FAMILY_GROUPS,
  EXPANDED_COMPARATIVE_FAMILY_ORDER,
  EXPANDED_COMPARATIVE_WORKLOAD_FAMILY,
  buildExpandedCompetitiveFamilyReport,
  formatExpandedCompetitiveFamilyReport,
  type ExpandedCompetitiveFamily,
} from '../report-competitive-families.js'

const expectedExpandedWorkloads: ExpandedComparativeBenchmarkWorkload[] = [
  'build-from-sheets',
  'build-dense-literals',
  'build-mixed-content',
  'build-parser-cache-row-templates',
  'build-parser-cache-mixed-templates',
  'build-parser-cache-unique-formulas',
  'build-many-sheets',
  'rebuild-and-recalculate',
  'rebuild-config-toggle',
  'rebuild-config-toggle-large',
  'rebuild-runtime-from-snapshot',
  'sheet-rename-dependencies',
  'named-expression-change',
  'single-edit-recalc',
  'single-edit-chain',
  'single-edit-fanout',
  'partial-recompute-mixed-frontier',
  'single-formula-edit-recalc',
  'batch-edit-recalc',
  'batch-edit-single-column',
  'batch-edit-multi-column',
  'batch-edit-single-column-with-undo',
  'batch-suspended-single-column',
  'batch-suspended-multi-column',
  'structural-insert-rows',
  'structural-delete-rows',
  'structural-move-rows',
  'structural-insert-columns',
  'structural-delete-columns',
  'structural-move-columns',
  'range-read',
  'range-read-dense',
  'aggregate-2d-ranges',
  'aggregate-overlapping-ranges',
  'aggregate-overlapping-sliding-window',
  'conditional-aggregation-reused-ranges',
  'conditional-aggregation-criteria-cell-edit',
  'conditional-aggregation-shared-criteria',
  'conditional-aggregation-mixed-criteria',
  'lookup-no-column-index',
  'lookup-with-column-index',
  'lookup-with-column-index-after-column-write',
  'lookup-with-column-index-after-batch-write',
  'lookup-approximate-sorted',
  'lookup-approximate-descending',
  'lookup-approximate-duplicates',
  'lookup-approximate-sorted-after-column-write',
  'lookup-text-exact',
  'lookup-reverse-search',
  'dynamic-array-filter',
  'dynamic-array-sort',
  'dynamic-array-unique',
]

const benchmarkDir = dirname(fileURLToPath(import.meta.url))
const expandedBaselinePath = join(benchmarkDir, '..', '..', 'baselines', 'workpaper-vs-hyperformula.json')
const emptyFamilyMetrics = {
  meanSpeedupGeomean: null,
  directionalMeanRatioGeomean: null,
  directionalP95RatioGeomean: null,
  worstWorkpaperToHyperFormulaMeanRatio: null,
  worstMeanRatioWorkload: null,
  worstWorkpaperToHyperFormulaP95Ratio: null,
  worstP95RatioWorkload: null,
}
const emptyOverallScorecard = {
  lane: 'overall',
  comparableCount: 0,
  workpaperWins: 0,
  hyperformulaWins: 0,
  directionalMeanRatioGeomean: null,
  directionalP95RatioGeomean: null,
  worstWorkpaperToHyperFormulaMeanRatio: null,
  worstMeanRatioWorkload: null,
  worstWorkpaperToHyperFormulaP95Ratio: null,
  worstP95RatioWorkload: null,
}
const emptyPublicScorecard = { ...emptyOverallScorecard, lane: 'public' }
const emptyHoldoutScorecard = { ...emptyOverallScorecard, lane: 'holdout' }

function familyEligibility(family: ExpandedCompetitiveFamily): { scorecardEligible: boolean; exclusionReason: string | null } {
  switch (family) {
    case 'build':
    case 'rebuild':
    case 'runtime-restore':
    case 'config-toggle':
      return family === 'config-toggle'
        ? {
            scorecardEligible: false,
            exclusionReason: 'Control-only rebuild toggle; not evidence of broad competitive victory.',
          }
        : { scorecardEligible: true, exclusionReason: null }
    case 'sheet-lifecycle':
    case 'named-expression':
    case 'dirty-execution':
    case 'batch-edit':
    case 'structural-rows':
    case 'structural-columns':
    case 'range-read':
    case 'aggregate-2d':
    case 'overlapping-aggregate':
    case 'sliding-window-aggregate':
    case 'conditional-aggregation':
    case 'lookup-exact':
    case 'lookup-after-write':
    case 'lookup-approximate':
    case 'lookup-approximate-after-write':
    case 'lookup-text':
    case 'dynamic-array':
      return {
        scorecardEligible: family !== 'dynamic-array',
        exclusionReason:
          family === 'dynamic-array' ? 'Leadership-only support lane; not an apples-to-apples performance scorecard input.' : null,
      }
  }
}

function readExpandedBaselineWorkloads(path: string): string[] {
  const parsed: unknown = JSON.parse(readFileSync(path, 'utf8'))
  if (
    parsed === null ||
    typeof parsed !== 'object' ||
    !('results' in parsed) ||
    !Array.isArray(parsed.results) ||
    !parsed.results.every(
      (result) => result !== null && typeof result === 'object' && 'workload' in result && typeof result.workload === 'string',
    )
  ) {
    throw new Error(`Unexpected expanded baseline format: ${path}`)
  }
  return parsed.results.map((result) => result.workload)
}

function readExpandedBaselineReport(path: string): {
  families: unknown
  scorecard: unknown
  results: ExpandedComparativeBenchmarkResult[]
} {
  const parsed: unknown = JSON.parse(readFileSync(path, 'utf8'))
  const results = parsed !== null && typeof parsed === 'object' && 'results' in parsed ? parsed.results : undefined
  if (
    parsed === null ||
    typeof parsed !== 'object' ||
    !('families' in parsed) ||
    !('scorecard' in parsed) ||
    !Array.isArray(results) ||
    !results.every(isExpandedComparativeBenchmarkResult)
  ) {
    throw new Error(`Unexpected expanded baseline report format: ${path}`)
  }
  return {
    families: parsed.families,
    scorecard: parsed.scorecard,
    results,
  }
}

function isExpandedComparativeBenchmarkResult(value: unknown): value is ExpandedComparativeBenchmarkResult {
  if (value === null || typeof value !== 'object') {
    return false
  }
  return (
    'workload' in value &&
    typeof value.workload === 'string' &&
    'category' in value &&
    typeof value.category === 'string' &&
    'comparable' in value &&
    typeof value.comparable === 'boolean' &&
    'fixture' in value &&
    value.fixture !== null &&
    typeof value.fixture === 'object' &&
    'engines' in value &&
    value.engines !== null &&
    typeof value.engines === 'object'
  )
}

function normalizeCompetitiveReportValue(value: unknown): unknown {
  if (typeof value === 'number') {
    return Number(value.toPrecision(12))
  }
  if (Array.isArray(value)) {
    return value.map((item) => normalizeCompetitiveReportValue(item))
  }
  if (value !== null && typeof value === 'object') {
    return Object.fromEntries(Object.entries(value).map(([key, item]) => [key, normalizeCompetitiveReportValue(item)]))
  }
  return value
}

describe('expanded comparative benchmark workloads', () => {
  it('enumerates the full expanded workload inventory without duplicates', () => {
    expect(new Set(EXPANDED_COMPARATIVE_WORKLOADS).size).toBe(EXPANDED_COMPARATIVE_WORKLOADS.length)
    expect(EXPANDED_COMPARATIVE_WORKLOADS).toEqual(expectedExpandedWorkloads)
  })

  it('checked-in expanded baseline covers every expanded workload exactly once', () => {
    expect(readExpandedBaselineWorkloads(expandedBaselinePath)).toEqual(expectedExpandedWorkloads)
  })

  it('keeps checked-in expanded baseline summaries derived from raw workload results', () => {
    const baseline = readExpandedBaselineReport(expandedBaselinePath)
    const expectedReport = buildExpandedCompetitiveFamilyReport(baseline.results)

    expect(normalizeCompetitiveReportValue(baseline.families)).toEqual(normalizeCompetitiveReportValue(expectedReport.families))
    expect(normalizeCompetitiveReportValue(baseline.scorecard)).toEqual(normalizeCompetitiveReportValue(expectedReport.scorecard))
  })

  it('assigns every expanded workload to exactly one family', () => {
    const seen = new Map<ExpandedComparativeBenchmarkWorkload, ExpandedCompetitiveFamily>()

    for (const family of EXPANDED_COMPARATIVE_FAMILY_ORDER) {
      const workloads = EXPANDED_COMPARATIVE_FAMILY_GROUPS[family]
      for (const workload of workloads) {
        expect(EXPANDED_COMPARATIVE_WORKLOAD_FAMILY[workload]).toBe(family)
        expect(seen.has(workload)).toBe(false)
        seen.set(workload, family)
      }
    }

    expect(seen.size).toBe(expectedExpandedWorkloads.length)
    expect([...seen.keys()].toSorted()).toEqual([...expectedExpandedWorkloads].toSorted())
  })

  it('formats an expanded family report with a machine-readable top-level families array', () => {
    const parsed: unknown = JSON.parse(formatExpandedCompetitiveFamilyReport([]))

    expect(parsed).toEqual({
      suite: 'workpaper-vs-hyperformula',
      families: EXPANDED_COMPARATIVE_FAMILY_ORDER.map((family) => ({
        family,
        workloads: EXPANDED_COMPARATIVE_FAMILY_GROUPS[family],
        ...familyEligibility(family),
        resultCount: 0,
        comparableCount: 0,
        leadershipCount: 0,
        workpaperWins: 0,
        hyperformulaWins: 0,
        ...emptyFamilyMetrics,
      })),
      scorecard: {
        ...emptyOverallScorecard,
        eligibleFamilies: EXPANDED_COMPARATIVE_FAMILY_ORDER.filter((family) => familyEligibility(family).scorecardEligible),
        excludedFamilies: EXPANDED_COMPARATIVE_FAMILY_ORDER.filter((family) => !familyEligibility(family).scorecardEligible),
        scorecards: {
          overall: emptyOverallScorecard,
          public: emptyPublicScorecard,
          holdout: emptyHoldoutScorecard,
        },
      },
    })
  })

  it('builds an expanded benchmark report with attached family summaries', () => {
    expect(buildExpandedComparativeBenchmarkReport([])).toEqual({
      suite: 'workpaper-vs-hyperformula',
      results: [],
      families: EXPANDED_COMPARATIVE_FAMILY_ORDER.map((family) => ({
        family,
        workloads: EXPANDED_COMPARATIVE_FAMILY_GROUPS[family],
        ...familyEligibility(family),
        resultCount: 0,
        comparableCount: 0,
        leadershipCount: 0,
        workpaperWins: 0,
        hyperformulaWins: 0,
        ...emptyFamilyMetrics,
      })),
      scorecard: {
        ...emptyOverallScorecard,
        eligibleFamilies: EXPANDED_COMPARATIVE_FAMILY_ORDER.filter((family) => familyEligibility(family).scorecardEligible),
        excludedFamilies: EXPANDED_COMPARATIVE_FAMILY_ORDER.filter((family) => !familyEligibility(family).scorecardEligible),
        scorecards: {
          overall: emptyOverallScorecard,
          public: emptyPublicScorecard,
          holdout: emptyHoldoutScorecard,
        },
      },
    })
  })

  it('emits engine counters for additional WorkPaper structural workload helpers', () => {
    const samples = [
      measureWorkPaperStructuralInsertRowsSample(32),
      measureWorkPaperStructuralDeleteRowsSample(32),
      measureWorkPaperStructuralMoveRowsSample(32),
      measureWorkPaperStructuralInsertColumnsSample(32),
      measureWorkPaperStructuralDeleteColumnsSample(32),
      measureWorkPaperStructuralMoveColumnsSample(32),
    ]

    for (const sample of samples) {
      expect(sample.engineCounters).toBeDefined()
      expect(Object.keys(sample.engineCounters ?? {}).toSorted()).toEqual([...ENGINE_COUNTER_KEYS].toSorted())
    }
  })

  it('keeps new comparable workload helper verifications equivalent', () => {
    const samplePairs = [
      [measureWorkPaperParserCacheUniqueFormulaSample(32), measureHyperFormulaParserCacheUniqueFormulaSample(32)],
      [measureWorkPaperSheetRenameDependencySample(), measureHyperFormulaSheetRenameDependencySample()],
      [measureWorkPaperNamedExpressionChangeSample(), measureHyperFormulaNamedExpressionChangeSample()],
      [measureWorkPaper2dAggregateSample(32), measureHyperFormula2dAggregateSample(32)],
      [
        measureWorkPaperConditionalAggregationSharedCriteriaSample(32, 4),
        measureHyperFormulaConditionalAggregationSharedCriteriaSample(32, 4),
      ],
      [
        measureWorkPaperConditionalAggregationMixedCriteriaSample(32, 4),
        measureHyperFormulaConditionalAggregationMixedCriteriaSample(32, 4),
      ],
      [measureWorkPaperApproximateLookupDescendingSample(32), measureHyperFormulaApproximateLookupDescendingSample(32)],
      [measureWorkPaperApproximateLookupDuplicateSample(32), measureHyperFormulaApproximateLookupDuplicateSample(32)],
    ]

    for (const [workpaper, hyperformula] of samplePairs) {
      expect(workpaper.verification).toEqual(hyperformula.verification)
    }
  })

  it('emits engine counters for new WorkPaper leadership workload helpers', () => {
    const samples = [measureWorkPaperDynamicArraySortSample(32), measureWorkPaperDynamicArrayUniqueSample(32)]

    for (const sample of samples) {
      expect(sample.engineCounters).toBeDefined()
      expect(Object.keys(sample.engineCounters ?? {}).toSorted()).toEqual([...ENGINE_COUNTER_KEYS].toSorted())
    }
  })
})
