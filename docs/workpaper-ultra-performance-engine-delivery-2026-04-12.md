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

Current reality is much better than the earlier plan state: on the current
expanded suite, WorkPaper leads `44/46` scorecard-eligible comparable workloads
and `8/8` holdout workloads. The job now is to close the two remaining
confidence-overlap mean reds and harden the worst p95 tail without regressing
the current green lanes.

## Current Benchmark Scoreboard

Current decision-driving artifact:

- `packages/benchmarks/baselines/workpaper-vs-hyperformula.json`
- generated at `2026-04-29T14:47:16.831Z`

Current position:

- Total workloads: `51`
- Scorecard-eligible comparable workloads: `46`
- `WorkPaper` wins `44/46` overall
- `HyperFormula` wins `2/46` overall
- Public lane: WorkPaper `36/38`, HyperFormula `2/38`
- Holdout lane: WorkPaper `8/8`, HyperFormula `0/8`

The old `/tmp/workpaper-vs-hf-current-sample2.json` scorecard and the 12/35
position are no longer current. They remain useful only as dated evidence for
why the ownership phases existed.

### Current green workloads

These are the decision-critical rows that are green or should be protected from
regression in the current expanded artifact:

- all `8/8` holdout rows, including `build-parser-cache-unique-formulas`,
  `sheet-rename-dependencies`, `named-expression-change`, and
  `lookup-approximate-duplicates`
- aggregate rows that were previously noisy, including
  `aggregate-overlapping-sliding-window`
- steady-state exact and approximate lookup rows
- build/parser rows except the current small `build-mixed-content` mean loss
- structural rows and columns except the current small `structural-delete-rows`
  mean loss
- public and holdout scorecard reporting itself

### Current active rows and immediate owners

| Workload | Mean Ratio | Median Ratio | P95 Ratio | Confidence Overlap | Primary owner |
| --- | ---: | ---: | ---: | --- | --- |
| `build-mixed-content` | `1.0362639565590437` | `1.0069852963334736` | `1.156165042556` | yes | cold mixed build and formula initialization |
| `structural-delete-rows` | `1.0234049542127845` | `0.8750303474565914` | `1.267650293785557` | yes | structural row-delete metadata and headless result collection |
| `lookup-text-exact` p95 | mean green | mean green | `2.27208263805424` | n/a | text lookup normalization, index reuse, invalidation, allocation |

Implementation order:

1. `build-mixed-content`: reduce production cold-build allocation and duplicated
   initialization.
2. `structural-delete-rows`: narrow structural row-delete metadata and result
   collection.
3. `lookup-text-exact`: harden p95 tail latency without changing benchmark
   sampling or scoring.

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
| `Phase 0: Measurement and invariants` | `done` | expanded benchmark and public/holdout scorecards exist; current artifact is `44/46` overall and `8/8` holdout |
| `Phase 1: Direct change payload substrate` | `mostly delivered` | direct changed-cell payloads exist and headless no longer depends on ordinary snapshot diffs |
| `Phase 2: Compiled plan arena and formula slots` | `mostly delivered` | shared plans and descriptors are enough for the current parser/build holdout wins; mixed build still needs allocation trimming |
| `Phase 3: Rebuild execution policy` | `delivered for current scorecard` | rebuild rows are not current red lanes |
| `Phase 4: Formula template normalization` | `mostly delivered` | parser-template and unique-formula rows are green; `build-mixed-content` remains the build-family target |
| `Phase 5: Suspended bulk mutation lane` | `delivered for current scorecard` | batch lanes are not current scorecard blockers |
| `Phase 6: Range and criterion caches` | `delivered for current scorecard` | criteria reuse, overlapping ranges, 2D aggregates, and sliding-window aggregate are green in the current artifact |
| `Phase 7: Structural transform service` | `partial` | most structural rows and columns are green; `structural-delete-rows` remains a small confidence-overlap mean red |
| `Phase 8: Post-write lookup maintenance cutover` | `mostly delivered` | lookup approximate duplicates and after-write rows are green; `lookup-text-exact` p95 remains a tail-risk target |
| `Phase 9: RuntimeColumnStore authority` | `partial` | runtime-owned column and range state grew meaningfully, but it is not authoritative for every remaining hot path |
| `Phase 10: WASM criteria and search kernels` | `not delivered` | existing wins came from ownership cuts, not the final WASM layer yet |
| `Phase 11: Delete displaced paths and lock gates` | `partial` | the remaining lock gate is preserving `44/46` overall and `8/8` holdout while removing the two small mean reds |

## Implemented And Retained

These changes are in the tree and have already survived reruns:

- `CriterionRangeCacheService`
  - made `conditional-aggregation-reused-ranges` and `conditional-aggregation-criteria-cell-edit`
    green
- `RangeAggregateCacheService`
  - made `aggregate-overlapping-ranges` green
- `FormulaTemplateNormalizationService`
  - parser-cache and unique-formula rows are green; mixed-content cold build
    remains the active allocation target
- live runtime `useColumnIndex` policy
  - made `rebuild-config-toggle` decisively green
- mutation-owned exact and approximate lookup dirtying
  - made steady-state and after-write lookup rows green enough for the current
    scorecard; `lookup-text-exact` p95 remains the target
- suspended literal fast queue
  - removed batch lanes from the current blocker list
- descriptor-owned structural impact tracking
  - massively improved structural lanes; `structural-delete-rows` remains the
    only current structural mean red

## Rejected Approaches

These were tried, benchmarked, and removed because they were not the right design:

- direct-aggregate structural astless rewrite
- cross-sheet mixed-sheet initialization batching
- prepared lookup refresh reuse shortcut

Do not reintroduce these without new evidence.

## HyperFormula Reread Implications

The targeted reread still explains the architecture, but the current benchmark
artifact changes the remaining order:

1. `build-mixed-content` cold-build allocation and duplicated initialization.
2. `structural-delete-rows` row-delete metadata/result collection.
3. `lookup-text-exact` p95 tail latency.
4. Preservation of current green rows, especially holdout rows.

This is why the remaining order below does not start with “more WASM”.

## Remaining Program Order

The remaining execution order is now:

1. `build-mixed-content` cold-build hardening in initial sheet load, formula
   source registration, formula initialization, and binding allocation.
2. `structural-delete-rows` hardening in structural row metadata, dependency
   retargeting, undo, and headless changed-result collection.
3. `lookup-text-exact` p95 hardening in text-key normalization, index reuse,
   invalidation, and allocation.
4. Preservation checks for all current green holdout and public rows.
5. WASM kernel cutover only after the remaining ownership paths are clean enough
   to feed typed memory without JS object graph materialization.
6. deletion of displaced paths once the expanded scorecard is stable.

This order is fixed by the current checked artifact:

- `build-mixed-content` is the worst current mean ratio.
- `structural-delete-rows` is the only current structural mean red.
- `lookup-text-exact` is the worst p95 ratio.
- all holdout rows are green and must be protected.

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

- `structural-delete-rows`
- preservation reruns for `structural-insert-rows`, `structural-move-rows`,
  `structural-insert-columns`, `structural-delete-columns`, and
  `structural-move-columns`

### `FormulaTemplateNormalizationService`

Goal:

- stop rematerializing equivalent prepared binding state for repeated row-shifted and mixed
  repeated formulas

Still required:

- mixed-content build cleanup
- allocation and duplicated-initialization reductions that preserve parser and
  binding semantics
- preservation of `build-parser-cache-unique-formulas`

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
