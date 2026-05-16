import { ENGINE_COUNTER_KEYS, type EngineCounters } from '../../core/src/perf/engine-counters.js'
import type {
  ComparativeBenchmarkSuiteOptions,
  ComparativeMeasuredEngineResult,
  ComparativeMemorySummary,
  ComparativeUnsupportedEngineResult,
} from './benchmark-workpaper-vs-hyperformula.js'
import type { ExpandedComparativeBenchmarkWorkload } from './expanded-competitive-workloads.js'
import { buildExpandedCompetitiveFamilyReport, type ExpandedCompetitiveFamilySummary } from './report-competitive-families.js'
import { DEFAULT_COMPETITIVE_SAMPLE_COUNT, DEFAULT_COMPETITIVE_WARMUP_COUNT } from './benchmark-workpaper-vs-hyperformula.js'
import type { MemoryMeasurement } from './metrics.js'
import { summarizeNumbers, type NumericSummary } from './stats.js'
import {
  measureHyperFormulaApproximateLookupSample,
  measureHyperFormulaBatchMultiColumnEditSample,
  measureHyperFormulaBatchSingleColumnEditSample,
  measureHyperFormulaDenseBuildSample,
  measureHyperFormulaFormulaEditSample,
  measureHyperFormulaLegacyBatchEditSample,
  measureHyperFormulaLegacySingleEditSample,
  measureHyperFormulaLookupSample,
  measureHyperFormulaManySheetsBuildSample,
  measureHyperFormulaMixedBuildSample,
  measureHyperFormulaRangeReadSample,
  measureHyperFormulaSingleChainEditSample,
  measureHyperFormulaSingleFanoutEditSample,
  measureHyperFormulaTextLookupSample,
  measureWorkPaperApproximateLookupSample,
  measureWorkPaperBatchMultiColumnEditSample,
  measureWorkPaperBatchSingleColumnEditSample,
  measureWorkPaperDenseBuildSample,
  measureWorkPaperDynamicArraySample,
  measureWorkPaperFormulaEditSample,
  measureWorkPaperLegacyBatchEditSample,
  measureWorkPaperLegacySingleEditSample,
  measureWorkPaperLookupSample,
  measureWorkPaperManySheetsBuildSample,
  measureWorkPaperMixedBuildSample,
  measureWorkPaperRangeReadSample,
  measureWorkPaperSingleChainEditSample,
  measureWorkPaperSingleFanoutEditSample,
  measureWorkPaperTextLookupSample,
} from './benchmark-workpaper-vs-hyperformula-expanded-core-workloads.js'
import {
  measureHyperFormulaConditionalAggregationCriteriaEditSample,
  measureHyperFormula2dAggregateSample,
  measureHyperFormulaApproximateLookupDescendingSample,
  measureHyperFormulaApproximateLookupDuplicateSample,
  measureHyperFormulaConditionalAggregationMixedCriteriaSample,
  measureHyperFormulaConditionalAggregationSharedCriteriaSample,
  measureHyperFormulaBatchSingleColumnUndoSample,
  measureHyperFormulaIndexedLookupAfterBatchWriteSample,
  measureHyperFormulaNamedExpressionChangeSample,
  measureHyperFormulaParserCacheMixedTemplateSample,
  measureHyperFormulaParserCacheUniqueFormulaSample,
  measureHyperFormulaRebuildRuntimeFromSnapshotSample,
  measureHyperFormulaSlidingAggregateSample,
  measureHyperFormulaSheetRenameDependencySample,
  measureHyperFormulaStructuralDeleteColumnsSample,
  measureHyperFormulaStructuralDeleteRowsSample,
  measureHyperFormulaStructuralInsertColumnsSample,
  measureHyperFormulaStructuralMoveColumnsSample,
  measureHyperFormulaStructuralMoveRowsSample,
  measureHyperFormulaApproximateLookupAfterColumnWriteSample,
  measureHyperFormulaConditionalAggregationSample,
  measureHyperFormulaConfigToggleSample,
  measureHyperFormulaIndexedLookupAfterColumnWriteSample,
  measureHyperFormulaMixedFrontierSample,
  measureHyperFormulaOverlappingAggregateSample,
  measureHyperFormulaParserCacheTemplateSample,
  measureHyperFormulaRebuildAndRecalculateSample,
  measureHyperFormulaStructuralInsertRowsSample,
  measureHyperFormulaSuspendedBatchMultiColumnEditSample,
  measureHyperFormulaSuspendedBatchSingleColumnEditSample,
  measureWorkPaperConditionalAggregationCriteriaEditSample,
  measureWorkPaper2dAggregateSample,
  measureWorkPaperApproximateLookupDescendingSample,
  measureWorkPaperApproximateLookupDuplicateSample,
  measureWorkPaperConditionalAggregationMixedCriteriaSample,
  measureWorkPaperConditionalAggregationSharedCriteriaSample,
  measureWorkPaperBatchSingleColumnUndoSample,
  measureWorkPaperIndexedLookupAfterBatchWriteSample,
  measureWorkPaperNamedExpressionChangeSample,
  measureWorkPaperParserCacheMixedTemplateSample,
  measureWorkPaperParserCacheUniqueFormulaSample,
  measureWorkPaperRebuildRuntimeFromSnapshotSample,
  measureWorkPaperSlidingAggregateSample,
  measureWorkPaperSheetRenameDependencySample,
  measureWorkPaperStructuralDeleteColumnsSample,
  measureWorkPaperStructuralDeleteRowsSample,
  measureWorkPaperStructuralInsertColumnsSample,
  measureWorkPaperStructuralMoveColumnsSample,
  measureWorkPaperStructuralMoveRowsSample,
  measureWorkPaperApproximateLookupAfterColumnWriteSample,
  measureWorkPaperConditionalAggregationSample,
  measureWorkPaperConfigToggleSample,
  measureWorkPaperIndexedLookupAfterColumnWriteSample,
  measureWorkPaperMixedFrontierSample,
  measureWorkPaperOverlappingAggregateSample,
  measureWorkPaperParserCacheTemplateSample,
  measureWorkPaperRebuildAndRecalculateSample,
  measureWorkPaperStructuralInsertRowsSample,
  measureWorkPaperSuspendedBatchMultiColumnEditSample,
  measureWorkPaperSuspendedBatchSingleColumnEditSample,
} from './benchmark-workpaper-vs-hyperformula-expanded-additional-workloads.js'
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
} from './benchmark-workpaper-vs-hyperformula-expanded-workbook-shape-workloads.js'
import {
  measureWorkPaperDynamicArraySortSample,
  measureWorkPaperDynamicArrayUniqueSample,
  measureWorkPaperReverseSearchLookupSample,
} from './benchmark-workpaper-vs-hyperformula-expanded-leadership-workloads.js'
import type { BenchmarkSample } from './benchmark-workpaper-vs-hyperformula-expanded-support.js'

export { EXPANDED_COMPARATIVE_WORKLOADS } from './expanded-competitive-workloads.js'
export type { ExpandedComparativeBenchmarkWorkload } from './expanded-competitive-workloads.js'

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

export function runWorkPaperVsHyperFormulaExpandedBenchmarkSuite(
  options: ComparativeBenchmarkSuiteOptions = {},
): ExpandedComparativeBenchmarkResult[] {
  const runtimeOptions = resolveSuiteOptions(options)
  return [
    runComparableScenario(
      'build-from-sheets',
      { cols: 24, rows: 160, materializedCells: 160 * 24 },
      runtimeOptions,
      () => measureWorkPaperDenseBuildSample(160, 24),
      () => measureHyperFormulaDenseBuildSample(160, 24),
    ),
    runComparableScenario(
      'build-dense-literals',
      { cols: 24, rows: 160, materializedCells: 160 * 24 },
      runtimeOptions,
      () => measureWorkPaperDenseBuildSample(160, 24),
      () => measureHyperFormulaDenseBuildSample(160, 24),
    ),
    runComparableScenario(
      'build-mixed-content',
      { cols: 6, rows: 750 },
      runtimeOptions,
      () => measureWorkPaperMixedBuildSample(750),
      () => measureHyperFormulaMixedBuildSample(750),
    ),
    runComparableScenario(
      'build-parser-cache-row-templates',
      { cols: 6, rows: 1_500 },
      runtimeOptions,
      () => measureWorkPaperParserCacheTemplateSample(1_500),
      () => measureHyperFormulaParserCacheTemplateSample(1_500),
    ),
    runComparableScenario(
      'build-parser-cache-mixed-templates',
      { cols: 6, rows: 1_500, templateShapes: 3 },
      runtimeOptions,
      () => measureWorkPaperParserCacheMixedTemplateSample(1_500),
      () => measureHyperFormulaParserCacheMixedTemplateSample(1_500),
    ),
    runComparableScenario(
      'build-parser-cache-unique-formulas',
      { cols: 6, rows: 1_500, templateShapes: 1_500 },
      runtimeOptions,
      () => measureWorkPaperParserCacheUniqueFormulaSample(1_500),
      () => measureHyperFormulaParserCacheUniqueFormulaSample(1_500),
    ),
    runComparableScenario(
      'build-many-sheets',
      { sheetCount: 8, rowsPerSheet: 120, colsPerSheet: 12 },
      runtimeOptions,
      () => measureWorkPaperManySheetsBuildSample(8, 120, 12),
      () => measureHyperFormulaManySheetsBuildSample(8, 120, 12),
    ),
    runComparableScenario(
      'build-cross-sheet-dashboard',
      { sheetCount: 4, rowsPerSheet: 500, summaryFormulas: 12 },
      runtimeOptions,
      () => measureWorkPaperCrossSheetDashboardBuildSample(4, 500),
      () => measureHyperFormulaCrossSheetDashboardBuildSample(4, 500),
    ),
    runComparableScenario(
      'rebuild-and-recalculate',
      { cols: 6, rows: 1_500 },
      runtimeOptions,
      () => measureWorkPaperRebuildAndRecalculateSample(1_500),
      () => measureHyperFormulaRebuildAndRecalculateSample(1_500),
    ),
    runComparableScenario(
      'rebuild-config-toggle',
      { rowCount: 5_000, config: 'useColumnIndex:false->true' },
      runtimeOptions,
      () => measureWorkPaperConfigToggleSample(5_000),
      () => measureHyperFormulaConfigToggleSample(5_000),
    ),
    runComparableScenario(
      'rebuild-config-toggle-large',
      { rowCount: 20_000, config: 'useColumnIndex:false->true' },
      runtimeOptions,
      () => measureWorkPaperConfigToggleSample(20_000),
      () => measureHyperFormulaConfigToggleSample(20_000),
    ),
    runComparableScenario(
      'rebuild-runtime-from-snapshot',
      { rowCount: 1_500, source: 'serialized-sheets' },
      runtimeOptions,
      () => measureWorkPaperRebuildRuntimeFromSnapshotSample(1_500),
      () => measureHyperFormulaRebuildRuntimeFromSnapshotSample(1_500),
    ),
    runComparableScenario(
      'sheet-rename-dependencies',
      { sheets: 2, renamedSheet: 'Data->Source', dependentFormulas: 2 },
      runtimeOptions,
      () => measureWorkPaperSheetRenameDependencySample(),
      () => measureHyperFormulaSheetRenameDependencySample(),
    ),
    runComparableScenario(
      'named-expression-change',
      { namedExpression: 'Rate', formulas: 2, mutation: '=2->=3' },
      runtimeOptions,
      () => measureWorkPaperNamedExpressionChangeSample(),
      () => measureHyperFormulaNamedExpressionChangeSample(),
    ),
    runComparableScenario(
      'cross-sheet-scalar-recalc',
      { sheets: 2, rowCount: 1_500, dependents: 1_500, mutation: 'Data!A1' },
      runtimeOptions,
      () => measureWorkPaperCrossSheetScalarFanoutSample(1_500),
      () => measureHyperFormulaCrossSheetScalarFanoutSample(1_500),
    ),
    runComparableScenario(
      'cross-sheet-aggregate-recalc',
      { sheets: 2, rowCount: 1_500, functionName: 'SUM', mutation: 'Data!A1' },
      runtimeOptions,
      () => measureWorkPaperCrossSheetAggregateSample(1_500),
      () => measureHyperFormulaCrossSheetAggregateSample(1_500),
    ),
    runComparableScenario(
      'cross-sheet-dashboard-recalc',
      { sheetCount: 4, rowsPerSheet: 1_000, summaryFormulas: 12, mutation: 'Data1!B1' },
      runtimeOptions,
      () => measureWorkPaperCrossSheetDashboardRecalcSample(4, 1_000),
      () => measureHyperFormulaCrossSheetDashboardRecalcSample(4, 1_000),
    ),
    runComparableScenario(
      'single-edit-recalc',
      { downstreamCount: 2_000 },
      runtimeOptions,
      () => measureWorkPaperLegacySingleEditSample(2_000),
      () => measureHyperFormulaLegacySingleEditSample(2_000),
    ),
    runComparableScenario(
      'single-edit-chain',
      { downstreamCount: 2_000 },
      runtimeOptions,
      () => measureWorkPaperSingleChainEditSample(2_000),
      () => measureHyperFormulaSingleChainEditSample(2_000),
    ),
    runComparableScenario(
      'single-edit-fanout',
      { downstreamCount: 2_000 },
      runtimeOptions,
      () => measureWorkPaperSingleFanoutEditSample(2_000),
      () => measureHyperFormulaSingleFanoutEditSample(2_000),
    ),
    runComparableScenario(
      'partial-recompute-mixed-frontier',
      { rowCount: 1_500, graphShape: 'scalar+range+fanout' },
      runtimeOptions,
      () => measureWorkPaperMixedFrontierSample(1_500),
      () => measureHyperFormulaMixedFrontierSample(1_500),
    ),
    runComparableScenario(
      'single-formula-edit-recalc',
      { downstreamCount: 1_500 },
      runtimeOptions,
      () => measureWorkPaperFormulaEditSample(1_500),
      () => measureHyperFormulaFormulaEditSample(1_500),
    ),
    runComparableScenario(
      'batch-edit-recalc',
      { editCount: 500 },
      runtimeOptions,
      () => measureWorkPaperLegacyBatchEditSample(500),
      () => measureHyperFormulaLegacyBatchEditSample(500),
    ),
    runComparableScenario(
      'batch-edit-single-column',
      { editCount: 500 },
      runtimeOptions,
      () => measureWorkPaperBatchSingleColumnEditSample(500),
      () => measureHyperFormulaBatchSingleColumnEditSample(500),
    ),
    runComparableScenario(
      'batch-edit-multi-column',
      { rowCount: 250, editsPerRow: 2 },
      runtimeOptions,
      () => measureWorkPaperBatchMultiColumnEditSample(250),
      () => measureHyperFormulaBatchMultiColumnEditSample(250),
    ),
    runComparableScenario(
      'batch-edit-rectangular-block',
      { rowCount: 64, inputCols: 12, editCount: 64 * 12 },
      runtimeOptions,
      () => measureWorkPaperRectangularBatchEditSample(64, 12),
      () => measureHyperFormulaRectangularBatchEditSample(64, 12),
    ),
    runComparableScenario(
      'batch-clear-rectangular-block',
      { rowCount: 64, inputCols: 12, editCount: 64 * 12 },
      runtimeOptions,
      () => measureWorkPaperBatchClearRectangularBlockSample(64, 12),
      () => measureHyperFormulaBatchClearRectangularBlockSample(64, 12),
    ),
    runComparableScenario(
      'batch-edit-single-column-with-undo',
      { editCount: 500, includesUndo: true },
      runtimeOptions,
      () => measureWorkPaperBatchSingleColumnUndoSample(500),
      () => measureHyperFormulaBatchSingleColumnUndoSample(500),
    ),
    runComparableScenario(
      'batch-suspended-single-column',
      { editCount: 500, mode: 'suspend-resume' },
      runtimeOptions,
      () => measureWorkPaperSuspendedBatchSingleColumnEditSample(500),
      () => measureHyperFormulaSuspendedBatchSingleColumnEditSample(500),
    ),
    runComparableScenario(
      'batch-suspended-multi-column',
      { rowCount: 250, editsPerRow: 2, mode: 'suspend-resume' },
      runtimeOptions,
      () => measureWorkPaperSuspendedBatchMultiColumnEditSample(250),
      () => measureHyperFormulaSuspendedBatchMultiColumnEditSample(250),
    ),
    runComparableScenario(
      'structural-insert-rows',
      { rowCount: 1_500, insertIndex: 750 },
      runtimeOptions,
      () => measureWorkPaperStructuralInsertRowsSample(1_500),
      () => measureHyperFormulaStructuralInsertRowsSample(1_500),
    ),
    runComparableScenario(
      'structural-append-formula-rows',
      { rowCount: 750, appendCount: 250, inputCols: 6 },
      runtimeOptions,
      () => measureWorkPaperAppendFormulaRowsSample(750, 6, 250),
      () => measureHyperFormulaAppendFormulaRowsSample(750, 6, 250),
    ),
    runComparableScenario(
      'structural-delete-rows',
      { rowCount: 1_500, deleteIndex: 750 },
      runtimeOptions,
      () => measureWorkPaperStructuralDeleteRowsSample(1_500),
      () => measureHyperFormulaStructuralDeleteRowsSample(1_500),
    ),
    runComparableScenario(
      'structural-move-rows',
      { rowCount: 1_500, start: 750, target: 0 },
      runtimeOptions,
      () => measureWorkPaperStructuralMoveRowsSample(1_500),
      () => measureHyperFormulaStructuralMoveRowsSample(1_500),
    ),
    runComparableScenario(
      'structural-insert-columns',
      { rowCount: 1_500, insertIndex: 1 },
      runtimeOptions,
      () => measureWorkPaperStructuralInsertColumnsSample(1_500),
      () => measureHyperFormulaStructuralInsertColumnsSample(1_500),
    ),
    runComparableScenario(
      'structural-delete-columns',
      { rowCount: 1_500, deleteIndex: 1 },
      runtimeOptions,
      () => measureWorkPaperStructuralDeleteColumnsSample(1_500),
      () => measureHyperFormulaStructuralDeleteColumnsSample(1_500),
    ),
    runComparableScenario(
      'structural-move-columns',
      { rowCount: 1_500, start: 1, target: 0 },
      runtimeOptions,
      () => measureWorkPaperStructuralMoveColumnsSample(1_500),
      () => measureHyperFormulaStructuralMoveColumnsSample(1_500),
    ),
    runComparableScenario(
      'range-read',
      { cols: 24, rows: 240, materializedCells: 240 * 24 },
      runtimeOptions,
      () => measureWorkPaperRangeReadSample(240, 24),
      () => measureHyperFormulaRangeReadSample(240, 24),
    ),
    runComparableScenario(
      'range-read-dense',
      { cols: 24, rows: 240, materializedCells: 240 * 24 },
      runtimeOptions,
      () => measureWorkPaperRangeReadSample(240, 24),
      () => measureHyperFormulaRangeReadSample(240, 24),
    ),
    runComparableScenario(
      'range-read-sparse-wide',
      { cols: 96, rows: 128, populatedCellsPerRow: 3, requestedCells: 128 * 96 },
      runtimeOptions,
      () => measureWorkPaperSparseWideRangeReadSample(128, 96),
      () => measureHyperFormulaSparseWideRangeReadSample(128, 96),
    ),
    runComparableScenario(
      'range-read-formula-grid',
      { rows: 256, inputCols: 4, formulaCols: 8, requestedCells: 256 * 8 },
      runtimeOptions,
      () => measureWorkPaperFormulaGridRangeReadSample(256, 4, 8),
      () => measureHyperFormulaFormulaGridRangeReadSample(256, 4, 8),
    ),
    runComparableScenario(
      'aggregate-2d-ranges',
      { rowCount: 1_500, functionName: 'SUM', rangeShape: 'growing-2d' },
      runtimeOptions,
      () => measureWorkPaper2dAggregateSample(1_500),
      () => measureHyperFormula2dAggregateSample(1_500),
    ),
    runComparableScenario(
      'aggregate-overlapping-ranges',
      { rowCount: 1_500, functionName: 'SUM' },
      runtimeOptions,
      () => measureWorkPaperOverlappingAggregateSample(1_500),
      () => measureHyperFormulaOverlappingAggregateSample(1_500),
    ),
    runComparableScenario(
      'aggregate-overlapping-sliding-window',
      { rowCount: 1_500, functionName: 'SUM', window: 32 },
      runtimeOptions,
      () => measureWorkPaperSlidingAggregateSample(1_500, 32),
      () => measureHyperFormulaSlidingAggregateSample(1_500, 32),
    ),
    runComparableScenario(
      'conditional-aggregation-reused-ranges',
      { rowCount: 2_000, formulaCopies: 32 },
      runtimeOptions,
      () => measureWorkPaperConditionalAggregationSample(2_000, 32),
      () => measureHyperFormulaConditionalAggregationSample(2_000, 32),
    ),
    runComparableScenario(
      'conditional-aggregation-criteria-cell-edit',
      { rowCount: 2_000, formulaCopies: 32, mutate: 'criteria-cell' },
      runtimeOptions,
      () => measureWorkPaperConditionalAggregationCriteriaEditSample(2_000, 32),
      () => measureHyperFormulaConditionalAggregationCriteriaEditSample(2_000, 32),
    ),
    runComparableScenario(
      'conditional-aggregation-shared-criteria',
      { rowCount: 2_000, criteriaCount: 32, mutate: 'shared-criteria-cell' },
      runtimeOptions,
      () => measureWorkPaperConditionalAggregationSharedCriteriaSample(2_000, 32),
      () => measureHyperFormulaConditionalAggregationSharedCriteriaSample(2_000, 32),
    ),
    runComparableScenario(
      'conditional-aggregation-mixed-criteria',
      { rowCount: 2_000, formulaCopies: 24, functions: ['COUNTIFS', 'SUMIFS'], mutate: 'threshold-cell' },
      runtimeOptions,
      () => measureWorkPaperConditionalAggregationMixedCriteriaSample(2_000, 24),
      () => measureHyperFormulaConditionalAggregationMixedCriteriaSample(2_000, 24),
    ),
    runComparableScenario(
      'lookup-no-column-index',
      { rowCount: 5_000, useColumnIndex: false },
      runtimeOptions,
      () => measureWorkPaperLookupSample(5_000, false),
      () => measureHyperFormulaLookupSample(5_000, false),
    ),
    runComparableScenario(
      'lookup-with-column-index',
      { rowCount: 5_000, useColumnIndex: true },
      runtimeOptions,
      () => measureWorkPaperLookupSample(5_000, true),
      () => measureHyperFormulaLookupSample(5_000, true),
    ),
    runComparableScenario(
      'lookup-index-match-exact',
      { rowCount: 5_000, functionName: 'INDEX+MATCH', matchMode: 'exact' },
      runtimeOptions,
      () => measureWorkPaperIndexMatchExactSample(5_000),
      () => measureHyperFormulaIndexMatchExactSample(5_000),
    ),
    runComparableScenario(
      'lookup-index-reference',
      { rowCount: 5_000, functionName: 'INDEX', lookupShape: '2d-table-row-index' },
      runtimeOptions,
      () => measureWorkPaperIndexReferenceSample(5_000),
      () => measureHyperFormulaIndexReferenceSample(5_000),
    ),
    runComparableScenario(
      'lookup-with-column-index-after-column-write',
      { rowCount: 5_000, useColumnIndex: true, mutate: 'lookup-column' },
      runtimeOptions,
      () => measureWorkPaperIndexedLookupAfterColumnWriteSample(5_000),
      () => measureHyperFormulaIndexedLookupAfterColumnWriteSample(5_000),
    ),
    runComparableScenario(
      'lookup-with-column-index-after-batch-write',
      { rowCount: 5_000, useColumnIndex: true, mutate: 'lookup-column-batch', editCount: 256 },
      runtimeOptions,
      () => measureWorkPaperIndexedLookupAfterBatchWriteSample(5_000, 256),
      () => measureHyperFormulaIndexedLookupAfterBatchWriteSample(5_000, 256),
    ),
    runComparableScenario(
      'lookup-approximate-sorted',
      { rowCount: 5_000 },
      runtimeOptions,
      () => measureWorkPaperApproximateLookupSample(5_000),
      () => measureHyperFormulaApproximateLookupSample(5_000),
    ),
    runComparableScenario(
      'lookup-approximate-descending',
      { rowCount: 5_000, matchMode: -1, ordering: 'descending' },
      runtimeOptions,
      () => measureWorkPaperApproximateLookupDescendingSample(5_000),
      () => measureHyperFormulaApproximateLookupDescendingSample(5_000),
    ),
    runComparableScenario(
      'lookup-approximate-duplicates',
      { rowCount: 5_000, matchMode: 1, duplicateKeys: true },
      runtimeOptions,
      () => measureWorkPaperApproximateLookupDuplicateSample(5_000),
      () => measureHyperFormulaApproximateLookupDuplicateSample(5_000),
    ),
    runComparableScenario(
      'lookup-approximate-sorted-after-column-write',
      { rowCount: 5_000, mutate: 'sorted-column-tail' },
      runtimeOptions,
      () => measureWorkPaperApproximateLookupAfterColumnWriteSample(5_000),
      () => measureHyperFormulaApproximateLookupAfterColumnWriteSample(5_000),
    ),
    runComparableScenario(
      'lookup-text-exact',
      { rowCount: 5_000 },
      runtimeOptions,
      () => measureWorkPaperTextLookupSample(5_000),
      () => measureHyperFormulaTextLookupSample(5_000),
    ),
    runLeadershipScenario(
      'lookup-reverse-search',
      { rowCount: 5_000, functionName: 'XMATCH', searchMode: -1 },
      runtimeOptions,
      () => measureWorkPaperReverseSearchLookupSample(5_000),
      {
        status: 'unsupported',
        evidence: ['HyperFormula 3.2.0 returns #NAME? for =XMATCH(D1,A2:A5001,0,-1) in this fixture.'],
        reason: 'HyperFormula 3.2.0 does not provide an equivalent XMATCH reverse-search result for this workload.',
      },
    ),
    runLeadershipScenario(
      'dynamic-array-filter',
      { rowCount: 750, formula: '=FILTER(A2:A751,A2:A751>B1)' },
      runtimeOptions,
      () => measureWorkPaperDynamicArraySample(750),
      {
        status: 'unsupported',
        evidence: [
          '/Users/gregkonush/github.com/hyperformula/docs/guide/known-limitations.md',
          '/Users/gregkonush/github.com/hyperformula/src/HyperFormula.ts',
        ],
        reason: 'HyperFormula 3.2.0 documents dynamic arrays as unsupported.',
      },
    ),
    runLeadershipScenario(
      'dynamic-array-sort',
      { rowCount: 750, formula: '=SORT(A2:A751)' },
      runtimeOptions,
      () => measureWorkPaperDynamicArraySortSample(750),
      {
        status: 'unsupported',
        evidence: [
          '/Users/gregkonush/github.com/hyperformula/docs/guide/known-limitations.md',
          '/Users/gregkonush/github.com/hyperformula/src/HyperFormula.ts',
        ],
        reason: 'HyperFormula 3.2.0 documents dynamic arrays as unsupported.',
      },
    ),
    runLeadershipScenario(
      'dynamic-array-unique',
      { rowCount: 750, formula: '=UNIQUE(A2:A751)' },
      runtimeOptions,
      () => measureWorkPaperDynamicArrayUniqueSample(750),
      {
        status: 'unsupported',
        evidence: [
          '/Users/gregkonush/github.com/hyperformula/docs/guide/known-limitations.md',
          '/Users/gregkonush/github.com/hyperformula/src/HyperFormula.ts',
        ],
        reason: 'HyperFormula 3.2.0 documents dynamic arrays as unsupported.',
      },
    ),
  ]
}

if (import.meta.url === `file://${process.argv[1]}`) {
  const cliOptions = parseExpandedBenchmarkCliOptions(process.argv.slice(2))
  const benchmarkResults = runWorkPaperVsHyperFormulaExpandedBenchmarkSuite({
    sampleCount: cliOptions.sampleCount ?? DEFAULT_COMPETITIVE_SAMPLE_COUNT,
    warmupCount: cliOptions.warmupCount ?? DEFAULT_COMPETITIVE_WARMUP_COUNT,
  })
  console.log(JSON.stringify(buildExpandedComparativeBenchmarkReport(benchmarkResults), null, 2))
}

function runComparableScenario(
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
      confidenceIntervalOverlaps:
        workpaper.elapsedMs.confidence95.low <= hyperformula.elapsedMs.confidence95.high &&
        hyperformula.elapsedMs.confidence95.low <= workpaper.elapsedMs.confidence95.high,
      verificationEquivalent: true,
    },
    engines: {
      workpaper,
      hyperformula,
    },
  }
}

function runLeadershipScenario(
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
    directAggregateDeltaApplications: zeroSummary,
    directAggregateDeltaOnlyRecalcSkips: zeroSummary,
    directScalarDeltaApplications: zeroSummary,
    directScalarDeltaOnlyRecalcSkips: zeroSummary,
    kernelSyncOnlyRecalcSkips: zeroSummary,
    directFormulaKernelSyncOnlyRecalcSkips: zeroSummary,
    directFormulaInitialEvaluations: zeroSummary,
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

function resolveSuiteOptions(options: ComparativeBenchmarkSuiteOptions): Required<ComparativeBenchmarkSuiteOptions> {
  return {
    sampleCount: options.sampleCount ?? DEFAULT_COMPETITIVE_SAMPLE_COUNT,
    warmupCount: options.warmupCount ?? DEFAULT_COMPETITIVE_WARMUP_COUNT,
  }
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
