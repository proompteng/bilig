export type ExpandedComparativeBenchmarkWorkload =
  | 'build-dense-literals'
  | 'build-mixed-content'
  | 'build-parser-cache-row-templates'
  | 'build-parser-cache-mixed-templates'
  | 'build-many-sheets'
  | 'rebuild-and-recalculate'
  | 'rebuild-config-toggle'
  | 'rebuild-runtime-from-snapshot'
  | 'single-edit-chain'
  | 'single-edit-fanout'
  | 'partial-recompute-mixed-frontier'
  | 'single-formula-edit-recalc'
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
  | 'range-read-dense'
  | 'aggregate-overlapping-ranges'
  | 'aggregate-overlapping-sliding-window'
  | 'conditional-aggregation-reused-ranges'
  | 'conditional-aggregation-criteria-cell-edit'
  | 'lookup-no-column-index'
  | 'lookup-with-column-index'
  | 'lookup-with-column-index-after-column-write'
  | 'lookup-with-column-index-after-batch-write'
  | 'lookup-approximate-sorted'
  | 'lookup-approximate-sorted-after-column-write'
  | 'lookup-text-exact'
  | 'dynamic-array-filter'

export const EXPANDED_COMPARATIVE_WORKLOADS = [
  'build-dense-literals',
  'build-mixed-content',
  'build-parser-cache-row-templates',
  'build-parser-cache-mixed-templates',
  'build-many-sheets',
  'rebuild-and-recalculate',
  'rebuild-config-toggle',
  'rebuild-runtime-from-snapshot',
  'single-edit-chain',
  'single-edit-fanout',
  'partial-recompute-mixed-frontier',
  'single-formula-edit-recalc',
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
  'range-read-dense',
  'aggregate-overlapping-ranges',
  'aggregate-overlapping-sliding-window',
  'conditional-aggregation-reused-ranges',
  'conditional-aggregation-criteria-cell-edit',
  'lookup-no-column-index',
  'lookup-with-column-index',
  'lookup-with-column-index-after-column-write',
  'lookup-with-column-index-after-batch-write',
  'lookup-approximate-sorted',
  'lookup-approximate-sorted-after-column-write',
  'lookup-text-exact',
  'dynamic-array-filter',
] as const satisfies readonly ExpandedComparativeBenchmarkWorkload[]
