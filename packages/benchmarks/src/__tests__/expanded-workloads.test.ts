import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
import { ENGINE_COUNTER_KEYS } from '../../../core/src/perf/engine-counters.js'
import {
  EXPANDED_COMPARATIVE_WORKLOAD_SCORECARD_LANE,
  EXPANDED_COMPARATIVE_WORKLOADS,
  buildExpandedComparativeBenchmarkReport,
  parseExpandedBenchmarkCliOptions,
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
  measureHyperFormulaCrossSheetAggregateSample,
  measureHyperFormulaCrossSheetDashboardBuildSample,
  measureHyperFormulaCrossSheetDashboardRecalcSample,
  measureHyperFormulaCrossSheetScalarFanoutSample,
  measureHyperFormulaIndexMatchExactSample,
  measureHyperFormulaIndexReferenceSample,
  measureHyperFormulaAppendFormulaRowsSample,
  measureHyperFormulaBatchClearRectangularBlockSample,
  measureHyperFormulaFormulaGridRangeReadSample,
  measureHyperFormulaRectangularBatchEditSample,
  measureHyperFormulaSparseWideRangeReadSample,
  measureWorkPaperCrossSheetAggregateSample,
  measureWorkPaperCrossSheetDashboardBuildSample,
  measureWorkPaperCrossSheetDashboardRecalcSample,
  measureWorkPaperCrossSheetScalarFanoutSample,
  measureWorkPaperIndexMatchExactSample,
  measureWorkPaperIndexReferenceSample,
  measureWorkPaperAppendFormulaRowsSample,
  measureWorkPaperBatchClearRectangularBlockSample,
  measureWorkPaperFormulaGridRangeReadSample,
  measureWorkPaperRectangularBatchEditSample,
  measureWorkPaperSparseWideRangeReadSample,
} from '../benchmark-workpaper-vs-hyperformula-expanded-workbook-shape-workloads.js'
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
  'build-dense-literals-wide',
  'build-dense-literals-tall',
  'build-mixed-content',
  'build-mixed-content-small',
  'build-mixed-content-large',
  'build-parser-cache-row-templates',
  'build-parser-cache-mixed-templates',
  'build-parser-cache-unique-formulas',
  'build-many-sheets',
  'build-many-sheets-wide',
  'build-many-sheets-narrow',
  'build-cross-sheet-dashboard',
  'build-cross-sheet-dashboard-small',
  'build-cross-sheet-dashboard-large',
  'rebuild-and-recalculate',
  'rebuild-and-recalculate-large',
  'rebuild-config-toggle',
  'rebuild-config-toggle-large',
  'rebuild-runtime-from-snapshot',
  'rebuild-runtime-from-snapshot-large',
  'sheet-rename-dependencies',
  'named-expression-change',
  'cross-sheet-scalar-recalc',
  'cross-sheet-aggregate-recalc',
  'cross-sheet-dashboard-recalc',
  'single-edit-recalc',
  'single-edit-chain',
  'single-edit-chain-small',
  'single-edit-chain-large',
  'single-edit-fanout',
  'single-edit-fanout-small',
  'single-edit-fanout-large',
  'partial-recompute-mixed-frontier',
  'single-formula-edit-recalc',
  'single-formula-edit-recalc-large',
  'batch-edit-recalc',
  'batch-edit-single-column',
  'batch-edit-single-column-small',
  'batch-edit-single-column-large',
  'batch-edit-multi-column-small',
  'batch-edit-multi-column',
  'batch-edit-multi-column-large',
  'batch-edit-rectangular-block',
  'batch-edit-rectangular-block-wide',
  'batch-clear-rectangular-block',
  'batch-clear-rectangular-block-wide',
  'batch-edit-single-column-with-undo',
  'batch-suspended-single-column',
  'batch-suspended-multi-column',
  'structural-insert-rows',
  'structural-insert-rows-small',
  'structural-insert-rows-large',
  'structural-append-formula-rows',
  'structural-append-formula-rows-small',
  'structural-append-formula-rows-large',
  'structural-delete-rows',
  'structural-move-rows',
  'structural-insert-columns',
  'structural-insert-columns-small',
  'structural-insert-columns-large',
  'structural-delete-columns',
  'structural-delete-columns-large',
  'structural-move-columns',
  'structural-move-columns-large',
  'range-read',
  'range-read-dense',
  'range-read-wide',
  'range-read-sparse-wide',
  'range-read-formula-grid',
  'range-read-formula-grid-wide',
  'aggregate-2d-ranges',
  'aggregate-2d-ranges-small',
  'aggregate-2d-ranges-large',
  'aggregate-overlapping-ranges',
  'aggregate-overlapping-ranges-small',
  'aggregate-overlapping-sliding-window',
  'aggregate-overlapping-sliding-window-wide',
  'conditional-aggregation-reused-ranges',
  'conditional-aggregation-reused-ranges-large',
  'conditional-aggregation-criteria-cell-edit',
  'conditional-aggregation-shared-criteria',
  'conditional-aggregation-mixed-criteria',
  'lookup-no-column-index',
  'lookup-no-column-index-small',
  'lookup-with-column-index',
  'lookup-with-column-index-large',
  'lookup-index-match-exact',
  'lookup-index-match-exact-large',
  'lookup-index-reference',
  'lookup-index-reference-large',
  'lookup-with-column-index-after-column-write',
  'lookup-with-column-index-after-batch-write',
  'lookup-with-column-index-after-batch-write-large',
  'lookup-approximate-sorted',
  'lookup-approximate-sorted-large',
  'lookup-approximate-descending',
  'lookup-approximate-duplicates',
  'lookup-approximate-sorted-after-column-write',
  'lookup-text-exact',
  'lookup-text-exact-large',
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
    case 'cross-sheet':
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
    expect(EXPANDED_COMPARATIVE_WORKLOADS).toHaveLength(106)
    expect(Object.keys(EXPANDED_COMPARATIVE_WORKLOAD_SCORECARD_LANE)).toHaveLength(106)
    expect(EXPANDED_COMPARATIVE_WORKLOADS).toEqual(expectedExpandedWorkloads)
  })

  it('keeps the scorecard-eligible comparable benchmark inventory at 100 workloads', () => {
    const leadershipOnlyWorkloads = new Set<ExpandedComparativeBenchmarkWorkload>([
      'lookup-reverse-search',
      'dynamic-array-filter',
      'dynamic-array-sort',
      'dynamic-array-unique',
    ])
    const scorecardEligibleWorkloads = EXPANDED_COMPARATIVE_FAMILY_ORDER.flatMap((family) =>
      familyEligibility(family).scorecardEligible
        ? EXPANDED_COMPARATIVE_FAMILY_GROUPS[family].filter((workload) => !leadershipOnlyWorkloads.has(workload))
        : [],
    )

    expect(scorecardEligibleWorkloads).toHaveLength(100)
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

  it('parses expanded benchmark CLI sample controls strictly', () => {
    expect(parseExpandedBenchmarkCliOptions(['--sample-count', '5', '--warmup-count', '0'])).toEqual({
      sampleCount: 5,
      warmupCount: 0,
    })
  })

  it('rejects malformed expanded benchmark CLI sample controls', () => {
    expect(() => parseExpandedBenchmarkCliOptions(['--sample-count', '1.5'])).toThrow(
      '--sample-count expects a non-negative integer, got 1.5',
    )
    expect(() => parseExpandedBenchmarkCliOptions(['--sample-count', '0'])).toThrow('--sample-count expects a positive integer, got 0')
    expect(() => parseExpandedBenchmarkCliOptions(['--warmup-count', '3abc'])).toThrow(
      '--warmup-count expects a non-negative integer, got 3abc',
    )
  })

  it('rejects unknown expanded benchmark CLI arguments', () => {
    expect(() => parseExpandedBenchmarkCliOptions(['--samples', '5'])).toThrow('Unknown expanded benchmark argument: --samples')
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
      [measureWorkPaperCrossSheetScalarFanoutSample(32), measureHyperFormulaCrossSheetScalarFanoutSample(32)],
      [measureWorkPaperCrossSheetAggregateSample(32), measureHyperFormulaCrossSheetAggregateSample(32)],
      [measureWorkPaperCrossSheetDashboardBuildSample(3, 32), measureHyperFormulaCrossSheetDashboardBuildSample(3, 32)],
      [measureWorkPaperCrossSheetDashboardRecalcSample(3, 32), measureHyperFormulaCrossSheetDashboardRecalcSample(3, 32)],
      [measureWorkPaperRectangularBatchEditSample(8, 4), measureHyperFormulaRectangularBatchEditSample(8, 4)],
      [measureWorkPaperBatchClearRectangularBlockSample(8, 4), measureHyperFormulaBatchClearRectangularBlockSample(8, 4)],
      [measureWorkPaperAppendFormulaRowsSample(8, 4, 4), measureHyperFormulaAppendFormulaRowsSample(8, 4, 4)],
      [measureWorkPaperSparseWideRangeReadSample(8, 8), measureHyperFormulaSparseWideRangeReadSample(8, 8)],
      [measureWorkPaperFormulaGridRangeReadSample(8, 3, 4), measureHyperFormulaFormulaGridRangeReadSample(8, 3, 4)],
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
      [measureWorkPaperIndexMatchExactSample(32), measureHyperFormulaIndexMatchExactSample(32)],
      [measureWorkPaperIndexReferenceSample(32), measureHyperFormulaIndexReferenceSample(32)],
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
