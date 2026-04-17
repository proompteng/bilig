# WorkPaper Ultra-Performance Engine Delivery Plan

Date: `2026-04-12`

Status: `execution-grade, revised against current broader benchmark reality on main`

Related documents:

- `/Users/gregkonush/github.com/bilig2/docs/workpaper-ultra-performance-engine-architecture-2026-04-12.md`
- `/Users/gregkonush/github.com/bilig2/docs/workpaper-hyperformula-prior-art-audit-2026-04-12.md`
- `/Users/gregkonush/github.com/bilig2/docs/workpaper-hyperformula-targeted-reread-2026-04-13.md`
- `/Users/gregkonush/github.com/bilig2/docs/workpaper-engine-leadership-program.md`

## Purpose

This document turns the target architecture into a hard migration program for the current
`bilig2` codebase.

The standard is unchanged:

- no workaround-first design
- no fake benchmark wins
- no permanent dual ownership
- no “clean it up later”
- no hot path that still routes through displaced legacy code

Current reality is harsher than the earlier plan state: on the current broader suite, WorkPaper is
behind HyperFormula by a wide workload count margin. The job now is to recover leadership on the
broader suite without regressing the remaining green lanes.

## Current Benchmark Scoreboard

Current decision-driving artifact:

- `/tmp/workpaper-vs-hf-current-sample2.json`

Current position:

- `WorkPaper` wins `12/35` directly comparable workloads
- `HyperFormula` wins `23/35`
- `WorkPaper` retains `1` leadership-only workload that HyperFormula does not support

The older `sample-count 5` artifact is now historical only. It should not be used to justify
current execution order.

### Current green workloads

These are the comparable lanes WorkPaper currently wins on the current broader artifact:

- `build-dense-literals`
- `build-many-sheets`
- `rebuild-and-recalculate`
- `rebuild-config-toggle`
- `single-edit-fanout`
- `range-read-dense`
- `aggregate-overlapping-ranges`
- `conditional-aggregation-reused-ranges`
- `conditional-aggregation-criteria-cell-edit`
- `lookup-no-column-index`
- `lookup-with-column-index`
- `lookup-text-exact`

### Current red workloads and immediate owners

| Workload | WorkPaper mean | HyperFormula mean | Primary owner |
| --- | ---: | ---: | --- |
| `structural-insert-columns` | `28.460542 ms` | `0.472729 ms` | `StructuralTransformService` |
| `structural-insert-rows` | `87.243625 ms` | `4.599000 ms` | `StructuralTransformService` |
| `structural-delete-rows` | `90.903167 ms` | `5.498417 ms` | `StructuralTransformService` |
| `structural-move-rows` | `97.645687 ms` | `9.541813 ms` | `StructuralTransformService` |
| `lookup-approximate-sorted-after-column-write` | `0.532750 ms` | `0.052083 ms` | `SortedColumnSearchService` |
| `structural-delete-columns` | `44.231937 ms` | `8.920687 ms` | `StructuralTransformService` |
| `build-parser-cache-row-templates` | `156.698667 ms` | `34.312834 ms` | `FormulaTemplateNormalizationService` |
| `partial-recompute-mixed-frontier` | `13.216562 ms` | `3.930520 ms` | `DirtyFrontier` and `RangeAggregateCacheService` |
| `batch-edit-single-column-with-undo` | `3.913437 ms` | `1.428584 ms` | `SuspendedBulkMutationLane` |
| `aggregate-overlapping-sliding-window` | `0.248979 ms` | `0.091479 ms` | `RangeAggregateCacheService` |
| `structural-move-columns` | `18.526750 ms` | `7.899146 ms` | `StructuralTransformService` |
| `lookup-with-column-index-after-column-write` | `0.150729 ms` | `0.072020 ms` | `ExactColumnIndexService` |
| `build-parser-cache-mixed-templates` | `142.403709 ms` | `73.411562 ms` | `FormulaTemplateNormalizationService` |
| `rebuild-runtime-from-snapshot` | `70.406521 ms` | `39.959250 ms` | `RebuildExecutionPolicy` |
| `build-mixed-content` | `21.588395 ms` | `12.914250 ms` | `FormulaTemplateNormalizationService` |
| `lookup-with-column-index-after-batch-write` | `1.086792 ms` | `0.661750 ms` | `ExactColumnIndexService` |
| `batch-suspended-single-column` | `0.982688 ms` | `0.622895 ms` | `SuspendedBulkMutationLane` |

## Delivery Rules

1. JavaScript remains the semantic source of truth.
2. A new path may coexist with an old path only during an active cutover.
3. Every phase ends with explicit deletion or hard displacement of the path it replaces.
4. No benchmark win counts if the displaced legacy path is still the common hot path.
5. No subsystem may have permanent mixed ownership between evaluator code and engine services.
6. Structural edits must not route through large generic mutation loops after the structural phase
   is complete.
7. Rebuild must not default to persistence reconstruction once rebuild policy is complete.
8. WASM is allowed only for closed kernels with JS oracle parity and typed-memory inputs.
9. No phase is complete if it introduces semantic drift, flaky invalidation, or cleanup debt.

## Current Phase Status

This is the honest current repo status, not the aspirational one.

| Phase | Status | Reality |
| --- | --- | --- |
| `Phase 0: Measurement and invariants` | `done` | expanded benchmark exists, but the broader suite now says we are behind and must drive priority |
| `Phase 1: Direct change payload substrate` | `mostly delivered` | direct changed-cell payloads exist and headless no longer depends on ordinary snapshot diffs |
| `Phase 2: Compiled plan arena and formula slots` | `partial` | shared plans and descriptors exist, but formula-local ownership is not fully gone |
| `Phase 3: Rebuild execution policy` | `partial` | `rebuild-config-toggle` is green, but `rebuild-runtime-from-snapshot` is still red |
| `Phase 4: Formula template normalization` | `partial` | the current broader suite says both parser-template lanes are red again, so family reuse is incomplete |
| `Phase 5: Suspended bulk mutation lane` | `partial` | suspended literal fast queue exists, but all batch lanes are still red |
| `Phase 6: Range and criterion caches` | `partial` | criteria reuse is green and overlapping prefix ranges are green, but sliding-window aggregate reuse is still red |
| `Phase 7: Structural transform service` | `partial` | structural impact and dependency ownership are much better, but rows and columns are still the worst family |
| `Phase 8: Post-write lookup maintenance cutover` | `partial` | steady-state exact lookup is green, but after-write exact and sorted lanes are still red |
| `Phase 9: RuntimeColumnStore authority` | `partial` | runtime-owned column and range state grew meaningfully, but it is not authoritative for every remaining hot path |
| `Phase 10: WASM criteria and search kernels` | `not delivered` | existing wins came from ownership cuts, not the final WASM layer yet |
| `Phase 11: Delete displaced paths and lock gates` | `not delivered` | legacy ownership still exists in the remaining red families |

## Implemented And Retained

These changes are in the tree and have already survived reruns:

- `CriterionRangeCacheService`
  - made `conditional-aggregation-reused-ranges` and `conditional-aggregation-criteria-cell-edit`
    green
- `RangeAggregateCacheService`
  - made `aggregate-overlapping-ranges` green
- `FormulaTemplateNormalizationService`
  - proved the right architectural direction, but the current broader suite says both parser-cache
    template lanes are still red on `main`
- live runtime `useColumnIndex` policy
  - made `rebuild-config-toggle` decisively green
- mutation-owned exact and approximate lookup dirtying
  - made steady-state exact lookup green, but approximate sorted is still red on the current
    broader suite
- suspended literal fast queue
  - narrowed the batch gap without finishing it
- descriptor-owned structural impact tracking
  - massively improved structural lanes without finishing them

## Rejected Approaches

These were tried, benchmarked, and removed because they were not the right design:

- direct-aggregate structural astless rewrite
- cross-sheet mixed-sheet initialization batching
- prepared lookup refresh reuse shortcut

Do not reintroduce these without new evidence.

## HyperFormula Reread Implications

The targeted reread still defines the remaining order:

1. structural transforms remain the biggest architecture-owned red family
2. exact after-write and approximate after-write must remain separate cuts
3. parser-template normalization must finish prepared binding family reuse
4. sliding-window aggregate reuse needs smaller-range extension semantics

This is why the remaining order below does not start with “more WASM”.

## Remaining Program Order

The remaining execution order is now:

1. `StructuralTransformService`
2. `FormulaTemplateNormalizationService`
3. `RebuildExecutionPolicy`
4. `SortedColumnSearchService`
5. `ExactColumnIndexService`
6. `RangeAggregateCacheService`
7. `SuspendedBulkMutationLane`
8. `RuntimeColumnStore` authority expansion
9. WASM kernel cutover on the cleaned ownership boundaries
10. deletion of displaced paths

This order is fixed by the current stable reds:

- structural rows and columns are now the largest engine-owned misses by far
- row-template and mixed-template builds still over-materialize equivalent formula state
- snapshot rebuild cannot become truly cheap until build normalization is cheaper
- approximate-after-write is still a larger miss than exact-after-batch
- sliding-window reuse matters, but only after range and structural ownership are stable enough

## Performance Correctness Invariants

These are required to ship performance work safely:

1. Formula-result writes must update the same column and range invalidation machinery as literal
   writes.
2. Direct formulas must never read workbook state that is still sitting in a queued WASM batch.
3. Versioned explicit empty cells must survive snapshot, replica, and undo or redo roundtrips.
4. Literal-to-formula and formula-to-literal transitions must refresh dependent range topology
   narrowly and correctly.
5. Any performance phase that breaks snapshot parity, replica parity, or undo or redo correctness
   is incomplete by definition.

## Remaining Phase Details

### `StructuralTransformService`

Goal:

- make row and column transforms true engine transforms instead of partial transform plus broad
  rebuild

Still required:

- in-place range retargeting
- transform-owned dependency maintenance
- structural undo or redo on the same model

Proof gate:

- `structural-insert-rows`
- `structural-delete-rows`
- `structural-move-rows`
- `structural-insert-columns`
- `structural-delete-columns`
- `structural-move-columns`

### `FormulaTemplateNormalizationService`

Goal:

- stop rematerializing equivalent prepared binding state for repeated row-shifted and mixed
  repeated formulas

Still required:

- prepared binding family reuse
- mixed-content build cleanup
- tighter coupling to snapshot rebuild

Proof gate:

- `build-parser-cache-row-templates`
- `build-parser-cache-mixed-templates`
- `build-mixed-content`

### `RebuildExecutionPolicy`

Goal:

- keep the green config-toggle result and finish snapshot rebuild

Still required:

- cheaper `rebuildRuntimeFromSnapshot`
- build-normalization reuse during rebuild

Proof gate:

- keep `rebuild-config-toggle` green
- flip `rebuild-runtime-from-snapshot`

### `SortedColumnSearchService`

Goal:

- narrow approximate lookup after writes to descriptor maintenance plus one search

Proof gate:

- `lookup-approximate-sorted-after-column-write`

### `ExactColumnIndexService`

Goal:

- finish narrowing exact lookup after ordinary and batched writes

Proof gate:

- `lookup-with-column-index-after-column-write`
- `lookup-with-column-index-after-batch-write`

### `SuspendedBulkMutationLane`

Goal:

- one queue, one history record, one emission pass per batch

Proof gate:

- `batch-edit-single-column`
- `batch-edit-multi-column`
- `batch-suspended-single-column`
- `batch-suspended-multi-column`

### `RangeAggregateCacheService`

Goal:

- extend current prefix reuse into true sliding-window reuse

Proof gate:

- keep `aggregate-overlapping-ranges` green
- flip `aggregate-overlapping-sliding-window`

## Benchmarks And Test Gates

### Required tests

- exact lookup correctness after column writes
- approximate lookup correctness after column writes
- criteria-function correctness with reused range caches
- structural row and column transform correctness
- rebuild-config and snapshot-rebuild equivalence
- JS and WASM parity for every new kernel

### Required benchmark commands

Fast engineering loop:

- `pnpm exec tsx packages/benchmarks/src/benchmark-workpaper-vs-hyperformula-expanded.ts --sample-count 2 --warmup-count 0`

Stability gate after a candidate tranche:

- `pnpm exec tsx packages/benchmarks/src/benchmark-workpaper-vs-hyperformula-expanded.ts --sample-count 5 --warmup-count 0`

### Required benchmark rule

No phase is done if it flips only its target lane while turning a previously green lane red.

## Stop Conditions

The program must stop and correct course if any phase produces one of these outcomes:

- benchmark win with semantic drift
- benchmark win that still routes common hot work through the legacy path
- persistent mixed ownership between evaluator code and an engine service
- WASM kernel whose marshaling cost erases the gain
- structural phase that still leaves generic mutation loops on the hot path
- cache phase that still leaves repeated range rescans for supported shapes

These are not acceptable tradeoffs.

## Current Files And Target Ownership

The most relevant remaining files are still:

- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/runtime-state.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/live.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/range-registry.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-binding-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-template-normalization-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/operation-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/structure-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/exact-column-index-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/sorted-column-search-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/range-aggregate-cache-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/headless/src/initial-sheet-load.ts`
- `/Users/gregkonush/github.com/bilig2/packages/headless/src/work-paper-runtime.ts`

## Done Definition

This delivery program is complete only when all of the following are true:

1. `WorkPaper` wins all directly comparable workloads in the expanded benchmark suite
2. `conditional-aggregation-reused-ranges` and `conditional-aggregation-criteria-cell-edit` stay
   green through range-owned criteria caches
3. `aggregate-overlapping-ranges` and `aggregate-overlapping-sliding-window` are both green
4. `structural-insert-rows`, `structural-delete-rows`, and `structural-move-rows` are
   transformer-owned and green
5. `structural-insert-columns`, `structural-delete-columns`, and `structural-move-columns` are
   transformer-owned and green
6. `lookup-with-column-index-after-column-write` and
   `lookup-with-column-index-after-batch-write` are exact-index-owned and green
7. `lookup-approximate-sorted-after-column-write` is sorted-service-owned and green
8. `build-parser-cache-row-templates`, `build-parser-cache-mixed-templates`, and
   `build-mixed-content` are template-normalization-owned and green
9. `rebuild-config-toggle` and `rebuild-runtime-from-snapshot` are rebuild-policy-owned and green
10. structural undo and redo share the same transform model as forward structural edits
11. displaced legacy paths are deleted

Current state is behind on the broader suite. Final state is `35/35` comparable wins with one
architecture and no fallback-first leftovers.
