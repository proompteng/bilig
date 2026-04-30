export type ExpandedComparativeBenchmarkWorkload =
  | 'build-from-sheets'
  | 'build-dense-literals'
  | 'build-mixed-content'
  | 'build-parser-cache-row-templates'
  | 'build-parser-cache-mixed-templates'
  | 'build-parser-cache-unique-formulas'
  | 'build-many-sheets'
  | 'rebuild-and-recalculate'
  | 'rebuild-config-toggle'
  | 'rebuild-runtime-from-snapshot'
  | 'sheet-rename-dependencies'
  | 'named-expression-change'
  | 'single-edit-recalc'
  | 'single-edit-chain'
  | 'single-edit-fanout'
  | 'partial-recompute-mixed-frontier'
  | 'single-formula-edit-recalc'
  | 'batch-edit-recalc'
  | 'batch-edit-single-column'
  | 'batch-edit-multi-column'
  | 'batch-edit-single-column-with-undo'
  | 'batch-suspended-single-column'
  | 'batch-suspended-multi-column'
  | 'structural-insert-rows'
  | 'structural-delete-rows'
  | 'structural-move-rows'
  | 'structural-insert-columns'
  | 'structural-delete-columns'
  | 'structural-move-columns'
  | 'range-read'
  | 'range-read-dense'
  | 'aggregate-2d-ranges'
  | 'aggregate-overlapping-ranges'
  | 'aggregate-overlapping-sliding-window'
  | 'conditional-aggregation-reused-ranges'
  | 'conditional-aggregation-criteria-cell-edit'
  | 'conditional-aggregation-shared-criteria'
  | 'conditional-aggregation-mixed-criteria'
  | 'lookup-no-column-index'
  | 'lookup-with-column-index'
  | 'lookup-with-column-index-after-column-write'
  | 'lookup-with-column-index-after-batch-write'
  | 'lookup-approximate-sorted'
  | 'lookup-approximate-descending'
  | 'lookup-approximate-duplicates'
  | 'lookup-approximate-sorted-after-column-write'
  | 'lookup-text-exact'
  | 'lookup-reverse-search'
  | 'dynamic-array-filter'
  | 'dynamic-array-sort'
  | 'dynamic-array-unique'

export const EXPANDED_COMPARATIVE_WORKLOADS = [
  'build-from-sheets',
  'build-dense-literals',
  'build-mixed-content',
  'build-parser-cache-row-templates',
  'build-parser-cache-mixed-templates',
  'build-parser-cache-unique-formulas',
  'build-many-sheets',
  'rebuild-and-recalculate',
  'rebuild-config-toggle',
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
] as const satisfies readonly ExpandedComparativeBenchmarkWorkload[]

export type ExpandedComparativeScorecardLane = 'public' | 'holdout'

export const EXPANDED_COMPARATIVE_WORKLOAD_SCORECARD_LANE = {
  'build-from-sheets': 'public',
  'build-dense-literals': 'public',
  'build-mixed-content': 'public',
  'build-parser-cache-row-templates': 'public',
  'build-parser-cache-mixed-templates': 'public',
  'build-parser-cache-unique-formulas': 'holdout',
  'build-many-sheets': 'public',
  'rebuild-and-recalculate': 'public',
  'rebuild-config-toggle': 'public',
  'rebuild-runtime-from-snapshot': 'public',
  'sheet-rename-dependencies': 'holdout',
  'named-expression-change': 'holdout',
  'single-edit-recalc': 'public',
  'single-edit-chain': 'public',
  'single-edit-fanout': 'public',
  'partial-recompute-mixed-frontier': 'public',
  'single-formula-edit-recalc': 'public',
  'batch-edit-recalc': 'public',
  'batch-edit-single-column': 'public',
  'batch-edit-multi-column': 'public',
  'batch-edit-single-column-with-undo': 'public',
  'batch-suspended-single-column': 'public',
  'batch-suspended-multi-column': 'public',
  'structural-insert-rows': 'public',
  'structural-delete-rows': 'public',
  'structural-move-rows': 'public',
  'structural-insert-columns': 'public',
  'structural-delete-columns': 'public',
  'structural-move-columns': 'public',
  'range-read': 'public',
  'range-read-dense': 'public',
  'aggregate-2d-ranges': 'holdout',
  'aggregate-overlapping-ranges': 'public',
  'aggregate-overlapping-sliding-window': 'public',
  'conditional-aggregation-reused-ranges': 'public',
  'conditional-aggregation-criteria-cell-edit': 'public',
  'conditional-aggregation-shared-criteria': 'holdout',
  'conditional-aggregation-mixed-criteria': 'holdout',
  'lookup-no-column-index': 'public',
  'lookup-with-column-index': 'public',
  'lookup-with-column-index-after-column-write': 'public',
  'lookup-with-column-index-after-batch-write': 'public',
  'lookup-approximate-sorted': 'public',
  'lookup-approximate-descending': 'holdout',
  'lookup-approximate-duplicates': 'holdout',
  'lookup-approximate-sorted-after-column-write': 'public',
  'lookup-text-exact': 'public',
  'lookup-reverse-search': 'holdout',
  'dynamic-array-filter': 'public',
  'dynamic-array-sort': 'holdout',
  'dynamic-array-unique': 'holdout',
} as const satisfies Record<ExpandedComparativeBenchmarkWorkload, ExpandedComparativeScorecardLane>
