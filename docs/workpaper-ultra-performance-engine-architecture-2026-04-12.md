# WorkPaper Ultra-Performance Engine Architecture

Date: `2026-04-12`

Status: `executing architecture, revised against current expanded-suite reality on main`

Related documents:

- `/Users/gregkonush/github.com/bilig2/docs/workpaper-ultra-performance-engine-delivery-2026-04-12.md`
- `/Users/gregkonush/github.com/bilig2/docs/workpaper-hyperformula-prior-art-audit-2026-04-12.md`
- `/Users/gregkonush/github.com/bilig2/docs/workpaper-hyperformula-targeted-reread-2026-04-13.md`
- `/Users/gregkonush/github.com/bilig2/docs/workpaper-engine-leadership-program.md`
- `/Users/gregkonush/github.com/bilig2/docs/workpaper-performance-acceleration-plan.md`

## Purpose

This document defines the engine architecture that can beat HyperFormula across the full expanded
competitive suite in this repo without fallback-first sludge, temporary hacks, or benchmark-only
special cases.

Several major ownership cuts are already in the tree, and the current expanded
suite shows WorkPaper leading `44/46` scorecard-eligible comparable workloads
with `8/8` holdout wins. This document now serves two jobs:

1. record the architecture cuts that actually survived implementation
2. define the remaining ownership work required to close the final two
   confidence-overlap mean reds and the `lookup-text-exact` p95 tail

The standard remains strict:

- no evaluator-owned primary lookup state
- no per-formula criteria rescans over reused ranges
- no structural row or column edits implemented as broad generic mutation loops
- no rebuild path that goes through public-sheet serialization when snapshot rebuild is valid
- no hidden legacy hot path under a benchmark-only fast path
- no WASM kernel fed by already-materialized JS object graphs

## Current Benchmark Reality

Current decision-driving artifact on `main`:

- `packages/benchmarks/baselines/workpaper-vs-hyperformula.json`
- Generated at `2026-04-29T14:47:16.831Z`

Current position on the expanded artifact:

- Total workloads: `51`
- Scorecard-eligible comparable workloads: `46`
- WorkPaper wins `44/46` overall
- HyperFormula wins `2/46` overall
- Public lane: WorkPaper `36/38`, HyperFormula `2/38`
- Holdout lane: WorkPaper `8/8`, HyperFormula `0/8`

The old `/tmp/workpaper-vs-hf-current-sample2.json` evidence and the 12/35
scorecard are no longer current. They remain useful only as proof that the
ownership cuts were necessary.

The remaining decision-driving rows are now much narrower:

| Workload | Mean Ratio | Median Ratio | P95 Ratio | Confidence Overlap | Primary owner that still must change | Real issue |
| --- | ---: | ---: | ---: | --- | --- | --- |
| `build-mixed-content` | `1.0362639565590437` | `1.0069852963334736` | `1.156165042556` | yes | formula initialization and cold mixed-sheet build | duplicate initialization and allocation still decide a close row |
| `structural-delete-rows` | `1.0234049542127845` | `0.8750303474565914` | `1.267650293785557` | yes | structural row-delete metadata and headless result collection | mean red with green median points to tail overhead, not a missing feature |
| `lookup-text-exact` | mean green | mean green | `2.27208263805424` | n/a for mean winner | text lookup owner and exact index invalidation | p95 allocation/cache churn is still too high |

Rows that this architecture document previously called active red lanes but that
are now green in the current artifact or no longer drive implementation order:

- `structural-insert-columns`
- `structural-insert-rows`
- `structural-move-rows`
- `structural-delete-columns`
- `structural-move-columns`
- `lookup-approximate-sorted-after-column-write`
- `lookup-with-column-index-after-column-write`
- `lookup-with-column-index-after-batch-write`
- `build-parser-cache-row-templates`
- `build-parser-cache-mixed-templates`
- `build-parser-cache-unique-formulas`
- `aggregate-overlapping-sliding-window`
- `conditional-aggregation-shared-criteria`
- `conditional-aggregation-mixed-criteria`
- `sheet-rename-dependencies`
- `named-expression-change`
- `lookup-approximate-duplicates`

The architecture only matters insofar as it removes the remaining production
costs without turning current greens red.

## Architecture Already In Tree

These are not aspirational. They are already implemented, benchmarked, and retained:

### `CriterionRangeCacheService`

Range-owned criteria cache reuse exists and already flipped the worst criteria lane.

Delivered effect:

- `conditional-aggregation-reused-ranges` is green
- `conditional-aggregation-criteria-cell-edit` is green

The architecture cut was correct: repeated criteria evaluation moved out of formula-local rescans
and into engine-owned range cache state.

### `RangeAggregateCacheService`

Engine-owned aggregate reuse now exists for direct aggregate families.

Delivered effect:

- `aggregate-overlapping-ranges` is still green
- `aggregate-overlapping-sliding-window` is green in the latest checked
  artifact after earlier noisy red runs

### `FormulaTemplateNormalizationService`

Template-family detection and astless translated-family reuse are in the tree.

Delivered effect:

- parser-template and unique-formula build rows are no longer current red lanes
  in the expanded artifact
- `build-parser-cache-unique-formulas` is green in the holdout lane

Not yet complete:

- `build-mixed-content` remains the active build-family target, with a small
  confidence-overlap mean loss
- remaining work is cold mixed-sheet initialization and binding allocation, not
  parser-cache benchmark control

That means template normalization is real, and the next useful work is reducing
production allocation in mixed build paths without weakening unique-formula
verification.

### Live `useColumnIndex` Runtime Policy

`useColumnIndex` is now a live runtime policy instead of a rebuild-only construction choice.

Delivered effect:

- `rebuild-config-toggle` is green by a huge margin

This confirmed the architectural point: config-toggle did not need a generic rebuild hot path.

### Mutation-Owned Lookup Dirtying

Exact and approximate lookup formulas now wake through mutation-owned service ownership rather than
through broad literal-cell reverse edges.

Delivered effect:

- `lookup-with-column-index` is green
- `lookup-no-column-index` is green
- `lookup-approximate-duplicates` is green
- `lookup-with-column-index-after-batch-write` is green

Not yet complete:

- `lookup-text-exact` has the worst p95 ratio and needs tail-latency hardening
  around text normalization, exact index reuse, invalidation, and allocation

### Descriptor-Owned Structural Impact Tracking

Structural impact collection now understands direct aggregate, direct criteria, and direct lookup
dependencies. Structural edits no longer blindly force the same formula teardown they used to.

Delivered effect:

- structural ownership got materially cleaner than the original full-rebind design
- most structural rows and columns are no longer current red lanes

Not yet complete:

- `structural-delete-rows` remains a small confidence-overlap mean loss, with a
  green median and red p95. The active issue is row-delete tail overhead and
  headless result collection, not the old catastrophic full-transform failure.

### Suspended-Eval Literal Fast Queue

Headless suspended literal edits now bypass the normal per-call mutation wrapper and go directly
into the deferred queue.

Delivered effect:

- batch lanes got closer
- batch lanes are not current scorecard blockers in the latest artifact

Not yet complete:

- preserve the current batch wins while fixing the active rows above

## HyperFormula Prior-Art Takeaways

The local reread in `/Users/gregkonush/github.com/hyperformula` still drives the correct remaining
order.

### Structural transforms are first-class engine operations

HyperFormula treats row and column add, remove, and move as dedicated graph, address, range, and
search transforms, not as large generic mutation loops. Its address-mapping layer is also a
first-class strategy subsystem.

Implication for WorkPaper:

- `StructuralTransformService` must own in-place range retargeting, address-mapping updates,
  transform-scoped dependency maintenance, and structural undo or redo on the same model

### Exact and approximate lookup are different systems

HyperFormula’s exact indexed lookup and approximate sorted lookup are not one shared abstraction.

Implication for WorkPaper:

- `ExactColumnIndexService` and `SortedColumnSearchService` must continue to evolve independently
- approximate-after-write should get narrower, not “more indexed”

### Parser caching is relative-template-aware

HyperFormula’s parser cache is centered on normalized token streams relative to base address.

Implication for WorkPaper:

- `FormulaTemplateNormalizationService` must finish the prepared binding family problem, not merely
  deduplicate compiled plans after expensive binding has already happened

### Range reuse is explicit and incremental

HyperFormula’s `RangeVertex` and aggregation code reuse smaller cached ranges instead of rescanning
overlapping windows from scratch.

Implication for WorkPaper:

- `RangeAggregateCacheService` needs smaller-range extension semantics for sliding windows

## Rejected Approaches

These were tried and are intentionally not part of the architecture anymore:

- direct-aggregate structural astless rewrite
  - regressed the suite and was removed
- cross-sheet mixed-sheet initialization batching
  - regressed mixed-sheet initialization and was removed
- prepared lookup refresh reuse shortcut
  - regressed the suite and was removed

These should stay dead unless new evidence proves a different design.

## Non-Negotiable Rules

1. JavaScript remains the semantic source of truth for formula meaning and correctness.
2. Exact lookup and approximate sorted lookup must not collapse into one shared primary service.
3. Criteria and aggregate reuse must be range-owned, not formula-owned.
4. Structural edits must not route through large generic cell-mutation loops on the hot path.
5. Rebuild modes must be explicit and selected by policy.
6. Headless and UI layers consume engine-emitted changes instead of diffing workbook state.
7. WASM accelerates closed deterministic kernels only after JS parity is already proven.
8. No benchmark win counts if the old path is still the common hot path.
9. No phase is complete if it introduces semantic drift, flaky invalidation, or cleanup debt.

## Performance Correctness Invariants

These are now hard architecture rules, not “test cleanup” details:

1. Formula-result writes must invalidate column versions, lookup freshness, and range caches the
   same way literal writes do.
2. Direct formulas must never evaluate against workbook state that is still sitting in a pending
   WASM batch.
3. Versioned explicit empty cells are real workbook state and must not be pruned as if they were
   dependency-only placeholders.
4. When a cell flips between literal and formula, dependent range topology must refresh narrowly
   and correctly.
5. Snapshot import, replica replay, undo or redo, and live local mutation must converge to the same
   visible workbook state.

## Runtime Layers

The runtime is still the same seven-layer design, but the ownership boundaries matter more than the
names:

1. `WorkbookPersistenceModel`
   - serializable workbook representation
2. `RuntimeColumnStore`
   - typed hot-path value storage and formula-result invalidation authority
3. `RangeEntityStore`
   - canonical range handles, prefix links, cache roots, and range-member topology
4. `CompiledFormulaPlanArena`
   - shared compiled plans and template families
5. `EngineServices`
   - exact lookup, sorted lookup, aggregate reuse, criteria reuse, structural transforms, rebuild
     policy, dirty frontier, and change emission
6. `ExecutionTier`
   - JS semantic tier with selective WASM acceleration
7. `Headless` and `UI Adapters`
   - consumers of already-materialized engine changes

```mermaid
flowchart LR
  PM["WorkbookPersistenceModel"] --> RB["RuntimeBuilder"]
  RB --> CS["RuntimeColumnStore"]
  RB --> RS["RangeEntityStore"]
  RB --> FP["CompiledFormulaPlanArena"]
  RB --> DG["DependencyGraph"]
  RB --> EXI["ExactColumnIndexService"]
  RB --> SSI["SortedColumnSearchService"]
  RB --> RAC["RangeAggregateCacheService"]
  RB --> CRC["CriterionRangeCacheService"]
  RB --> ST["StructuralTransformService"]
  RB --> RP["RebuildExecutionPolicy"]
  DG --> DFS["DirtyFrontierScheduler"]
  EXI --> EX["ExecutionTier"]
  SSI --> EX
  RAC --> EX
  CRC --> EX
  DFS --> EX
  EX --> CSE["ChangeSetEmitter"]
```

## Workload Ownership Matrix

| Workload | Primary subsystem | What must be true when done |
| --- | --- | --- |
| `build-mixed-content` | `FormulaTemplateNormalizationService` | mixed builds reuse prepared binding families instead of rematerializing equivalent dependency state |
| `build-parser-cache-row-templates` | `FormulaTemplateNormalizationService` | repeated row templates normalize to one compiled and prepared family |
| `rebuild-config-toggle` | `RebuildExecutionPolicy` | live config policy avoids generic rebuild and keeps this lane green |
| `rebuild-runtime-from-snapshot` | `RebuildExecutionPolicy` | snapshot rebuild avoids broad runtime reconstruction and reuses normalized plan families |
| `batch-edit-single-column` | `SuspendedBulkMutationLane` | single-column batches pay one queue setup, one history record, and one emission pass |
| `batch-edit-multi-column` | `SuspendedBulkMutationLane` | multi-column batches stay on the same narrow lane |
| `batch-suspended-single-column` | `SuspendedBulkMutationLane` | suspended single-column edits are a true deferred batch, not repeated local wrappers |
| `batch-suspended-multi-column` | `SuspendedBulkMutationLane` | suspended multi-column edits stay on that same narrow lane |
| `structural-insert-rows` | `StructuralTransformService` | row insert retargets ranges and dependency state in place |
| `structural-delete-rows` | `StructuralTransformService` | row delete reuses the same transform-owned model and undo records |
| `structural-move-rows` | `StructuralTransformService` | row move is one transform sequence, not broad rebuild bookkeeping |
| `aggregate-overlapping-ranges` | `RangeAggregateCacheService` | overlapping prefix ranges reuse parent cache state |
| `aggregate-overlapping-sliding-window` | `RangeAggregateCacheService` | sliding windows extend smaller cached windows instead of rescanning |
| `conditional-aggregation-reused-ranges` | `CriterionRangeCacheService` | repeated criteria formulas share one range-owned cache root |
| `lookup-with-column-index-after-column-write` | `ExactColumnIndexService` | post-write exact lookup is raw bucket maintenance plus a narrow dirty frontier |
| `lookup-with-column-index-after-batch-write` | `ExactColumnIndexService` | batched writes do not rebuild broad exact-index state |
| `lookup-approximate-sorted-after-column-write` | `SortedColumnSearchService` | post-write approximate lookup is narrow descriptor maintenance plus one search |

## Remaining Execution Order

This is the remaining order that actually matches the current checked artifact:

1. `build-mixed-content`
   - cold mixed-sheet initialization
   - formula source registration and binding allocation
   - preservation of parser-cache and unique-formula holdout wins
2. `structural-delete-rows`
   - row-delete metadata narrowing
   - transform-owned dependency/index maintenance
   - structural undo or redo on the same model
   - headless changed-result collection
3. `lookup-text-exact`
   - text-key normalization
   - exact index reuse and invalidation
   - p95 allocation tail reduction
3. `RebuildExecutionPolicy`
   - snapshot rebuild on top of cheaper normalized build machinery
4. `SortedColumnSearchService`
   - approximate-after-write narrowing
5. `ExactColumnIndexService`
   - after-write and batched-write cleanup
6. `SuspendedBulkMutationLane`
   - one queue, one history record, one emission pass
7. `RangeAggregateCacheService`
   - sliding-window smaller-range extension semantics
8. `RuntimeColumnStore` authority expansion
9. WASM kernels on the now-clean ownership boundaries
10. deletion of displaced paths

## Disallowed Fallbacks

These remain explicitly forbidden as primary architecture:

- evaluator-owned primary lookup descriptors
- evaluator-owned primary criteria or aggregate caches
- per-formula criteria rescans over reused identical ranges
- structural row or column edits implemented as large generic mutation loops
- rebuild paths that serialize sheets through the public surface when snapshot reuse is valid
- headless before-and-after workbook diff reconstruction on ordinary edits
- WASM kernels invoked only after JS already materialized object-heavy vectors

## Why This Can Beat HyperFormula Everywhere

WorkPaper is already ahead by stable workload count because the right architecture cuts were real:

- range-owned criteria reuse
- range-owned direct aggregate reuse
- template-family normalization
- live rebuild policy for `useColumnIndex`
- mutation-owned exact and approximate lookup dirtying
- deferred suspended literal queue

The remaining gaps are still architecture gaps, not “just optimize harder” gaps:

- structural transforms still rebuild too much range and dependency state
- row-template build still binds too much equivalent structure
- snapshot rebuild still reconstructs too much runtime
- approximate-after-write still wakes too much work
- sliding-window aggregate reuse is not yet incremental enough

If those are closed without regressions, the final all-green result does not require a different
engine, only completion of the current one.

## Acceptance Criteria

This architecture is not done until all of the following are true:

1. `WorkPaper` wins all directly comparable workloads in the expanded benchmark suite
2. `conditional-aggregation-reused-ranges` and `conditional-aggregation-criteria-cell-edit` stay
   green through range-owned criteria caches
3. `aggregate-overlapping-ranges` and `aggregate-overlapping-sliding-window` are both served by
   `RangeAggregateCacheService`
4. `structural-insert-rows`, `structural-delete-rows`, and `structural-move-rows` are all served
   by `StructuralTransformService` and are green
5. `structural-insert-columns`, `structural-delete-columns`, and
   `structural-move-columns` are all served by `StructuralTransformService` and are green
6. `lookup-with-column-index-after-column-write` and
   `lookup-with-column-index-after-batch-write` are exact-index-owned and green
7. `lookup-approximate-sorted-after-column-write` is sorted-service-owned and green
8. `build-parser-cache-row-templates`, `build-parser-cache-mixed-templates`, and
   `build-mixed-content` are template-normalization-owned and green
9. `rebuild-config-toggle` and `rebuild-runtime-from-snapshot` are rebuild-policy-owned and green
10. headless applies engine-emitted `WorkPaperChange[]` without workbook diff reconstruction
11. benchmark wins survive reruns on a clean committed tree

That is the bar: `35/35` comparable wins, one architecture, no benchmark-only cheats, and no
fallback-first leftovers.
