# WorkPaper Ultra-Performance Engine Delivery Plan

Date: `2026-04-12`

Status: `execution-grade, revised against current benchmark reality`

Related documents:

- `/Users/gregkonush/github.com/bilig2/docs/workpaper-ultra-performance-engine-architecture-2026-04-12.md`
- `/Users/gregkonush/github.com/bilig2/docs/workpaper-hyperformula-prior-art-audit-2026-04-12.md`
- `/Users/gregkonush/github.com/bilig2/docs/workpaper-hyperformula-closeout-plan-2026-04-12.md`
- `/Users/gregkonush/github.com/bilig2/docs/workpaper-engine-leadership-program.md`

## Purpose

This document turns the target architecture into a hard migration program for the current
`bilig2` codebase.

The standard is strict:

- no workarounds
- no fallback-first architecture
- no permanent dual ownership
- no “clean it up later”
- no benchmark win that depends on hidden legacy paths

The program must end with one primary engine architecture that beats HyperFormula across the full
expanded competitive suite in this repo.

## Current Benchmark Scoreboard

Latest local expanded artifact: `/tmp/workpaper-expanded-current.json`

Current position:

- `WorkPaper` wins `13/24` directly comparable workloads
- `HyperFormula` wins `11/24`
- overall geometric mean is still `1.193x` slower for `WorkPaper`

Current red workloads and their immediate owners:

| Workload | WorkPaper mean | HyperFormula mean | Primary owner |
| --- | ---: | ---: | --- |
| `build-mixed-content` | `21.340667 ms` | `16.372333 ms` | `FormulaTemplateNormalizationService` |
| `build-parser-cache-row-templates` | `108.853292 ms` | `50.341541 ms` | `FormulaTemplateNormalizationService` |
| `rebuild-config-toggle` | `33.938125 ms` | `14.454667 ms` | `RebuildExecutionPolicy` |
| `partial-recompute-mixed-frontier` | `5.098167 ms` | `4.714833 ms` | `DirtyFrontierScheduler` |
| `batch-edit-multi-column` | `1.241833 ms` | `1.133042 ms` | `SuspendedBulkMutationLane` |
| `batch-suspended-multi-column` | `1.061250 ms` | `0.799417 ms` | `SuspendedBulkMutationLane` |
| `structural-insert-rows` | `70.896209 ms` | `6.222416 ms` | `StructuralTransformService` |
| `aggregate-overlapping-ranges` | `4.290542 ms` | `3.015083 ms` | `RangeAggregateCacheService` |
| `conditional-aggregation-reused-ranges` | `17.697083 ms` | `1.188792 ms` | `CriterionRangeCacheService` |
| `lookup-with-column-index-after-column-write` | `1.899959 ms` | `0.098500 ms` | `ExactColumnIndexService` |
| `lookup-approximate-sorted-after-column-write` | `0.526209 ms` | `0.116708 ms` | `SortedColumnSearchService` |

This program is ordered to flip those lanes first. Any phase that does not map to a red workload
or a blocker for a red workload is not a top priority.

## Delivery Rules

1. JavaScript remains the semantic source of truth.
2. A new path may coexist with an old path only while a specific migration phase is actively
   cutting over.
3. Every phase ends with explicit deletion or hard displacement of the path it replaces.
4. No benchmark win counts if the displaced legacy path is still the common hot path.
5. No subsystem may have permanent mixed ownership between evaluator code and engine services.
6. Structural edits must not route through large ordinary cell-mutation loops after the structural
   transform phase lands.
7. Criteria-function reuse must not live on formulas after the criterion-cache phase lands.
8. Rebuild must not default to persistence reconstruction once rebuild policy is in place.
9. WASM is allowed only for closed kernels with JS oracle parity and typed-memory inputs.
10. No phase is complete if it introduces semantic drift, flaky invalidation, or new cleanup debt.

## Current Phase Status

This is the honest current repo status, not the aspirational one.

| Phase | Status | Reality |
| --- | --- | --- |
| `Phase 0: Measurement and invariants` | `done` | expanded benchmark exists and current red lanes are known |
| `Phase 1: Direct change payload substrate` | `mostly delivered` | direct changed-cell payloads exist and headless no longer depends on ordinary snapshot diffs |
| `Phase 2: Compiled plan arena and formula slots` | `partial` | plan separation exists, but formula-local ownership is not fully gone |
| `Phase 3: Rebuild execution policy` | `partial` | in-place `rebuildAndRecalculate` and snapshot rebuild path exist, but config-toggle is still red |
| `Phase 4: Formula template normalization` | `not delivered` | repeated row-template formulas are still too expensive |
| `Phase 5: Suspended bulk mutation lane` | `partial` | deferred local suspended batches improved, but multi-column suspended path is still red |
| `Phase 6: Range and criterion caches` | `not delivered` | criteria reuse is still missing as a first-class engine subsystem |
| `Phase 7: Structural transform service` | `not delivered` | structural inserts still behave too much like generic mutation |
| `Phase 8: Post-write lookup maintenance cutover` | `not delivered` | after-write exact and sorted lookup lanes are still badly red |
| `Phase 9: RuntimeColumnStore authority` | `partial` | runtime column store boundary exists, but it is not authoritative for all hot lanes |
| `Phase 10: WASM criteria and search kernels` | `not delivered` | WASM exists, but not yet for the remaining red criteria/search ownership problem |
| `Phase 11: Delete displaced paths and lock gates` | `not delivered` | legacy ownership still exists in multiple hot paths |

## Program Outcome

At the end of this program:

- workbook build precomputes all interactive lookup, range, and criteria state
- rebuild uses explicit mode selection instead of one generic expensive path
- formulas execute through shared compiled plans and template-normalized bindings
- criteria and aggregate reuse is range-owned, not formula-owned
- structural edits are transformer operations, not generic mutation loops
- post-write lookup maintenance is mutation-owned and narrow
- headless consumes direct engine change payloads
- numeric and mask-heavy hot loops run through WASM only where the typed-memory contract is clean

## Phase Graph

```mermaid
flowchart TD
  P0["Phase 0<br/>Measurement and invariants<br/>done"] --> P1["Phase 1<br/>Direct change payload substrate<br/>mostly delivered"]
  P1 --> P2["Phase 2<br/>Compiled plan arena and formula slots<br/>partial"]
  P2 --> P3["Phase 3<br/>Rebuild execution policy<br/>partial"]
  P3 --> P4["Phase 4<br/>Formula template normalization<br/>planned"]
  P4 --> P5["Phase 5<br/>Suspended bulk mutation lane<br/>partial"]
  P5 --> P6["Phase 6<br/>Range and criterion caches<br/>planned"]
  P6 --> P7["Phase 7<br/>Structural transform service<br/>planned"]
  P7 --> P8["Phase 8<br/>Post-write lookup maintenance cutover<br/>planned"]
  P8 --> P9["Phase 9<br/>RuntimeColumnStore authority<br/>partial"]
  P9 --> P10["Phase 10<br/>WASM criteria and search kernels<br/>planned"]
  P10 --> P11["Phase 11<br/>Delete displaced paths and lock gates<br/>planned"]
```

## Current Files And Target Ownership

### Files that will be reshaped further

- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/runtime-state.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/live.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-binding-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-evaluation-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/lookup-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/mutation-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/mutation-support-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/operation-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/recalc-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/read-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/headless/src/work-paper-runtime.ts`
- `/Users/gregkonush/github.com/bilig2/packages/headless/src/tracked-engine-event-refs.ts`
- `/Users/gregkonush/github.com/bilig2/packages/headless/src/initial-sheet-load.ts`

### New engine modules that must exist

- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-template-normalization-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/range-aggregate-cache-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/criterion-range-cache-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/structural-transform-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/rebuild-execution-policy-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/suspended-bulk-mutation-service.ts`

### Existing new modules that must become sole owners

- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/exact-column-index-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/sorted-column-search-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/compiled-plan-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/change-set-emitter-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/runtime-column-store-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/dirty-frontier-scheduler-service.ts`

### New WASM-facing modules that must exist

- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/wasm-kernel-dispatch-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/wasm-kernel/assembly/dispatch-criteria-mask-numeric.ts`
- `/Users/gregkonush/github.com/bilig2/packages/wasm-kernel/assembly/dispatch-criteria-mask-string-id.ts`
- `/Users/gregkonush/github.com/bilig2/packages/wasm-kernel/assembly/dispatch-aggregate-numeric-contiguous.ts`
- `/Users/gregkonush/github.com/bilig2/packages/wasm-kernel/assembly/dispatch-lookup-exact-column.ts`
- `/Users/gregkonush/github.com/bilig2/packages/wasm-kernel/assembly/dispatch-lookup-sorted-column.ts`

## Core Interfaces

These interfaces should be created early and stabilized. They are intentionally boring.

```ts
export interface FormulaTemplateNormalizationService {
  internTemplate(boundFormula: BoundFormulaShape): TemplateBinding;
  resolvePlan(template: TemplateBinding): PlanId;
}

export interface RebuildExecutionPolicyService {
  chooseMode(request: RebuildRequest): "recalculateAll" | "rebuildRuntimeFromSnapshot" | "rebuildFromPersistence";
}

export interface SuspendedBulkMutationService {
  beginBatch(kind: "local" | "undoable"): void;
  recordCellMutation(mutation: DeferredCellMutation): void;
  flushBatch(): readonly WorkPaperChange[];
}

export interface RangeAggregateCacheService {
  getOrBuild(request: AggregateCacheRequest): AggregateCacheHandle;
  invalidateRange(rangeHandle: RangeHandle): void;
}

export interface CriterionRangeCacheService {
  getOrBuild(request: CriterionCacheRequest): CriterionCacheHandle;
  invalidateRange(rangeHandle: RangeHandle): void;
}

export interface StructuralTransformService {
  insertRows(request: InsertRowsRequest): TransformResult;
  removeRows(request: RemoveRowsRequest): TransformResult;
  moveRows(request: MoveRowsRequest): TransformResult;
  insertColumns(request: InsertColumnsRequest): TransformResult;
  removeColumns(request: RemoveColumnsRequest): TransformResult;
  moveColumns(request: MoveColumnsRequest): TransformResult;
}
```

## Phase 0: Measurement And Invariants

### Goal

Freeze the semantic and performance guardrails before deeper storage and ownership changes.

### Work

- keep the expanded benchmark artifact runnable and reproducible
- keep workload-specific correctness tests around:
  - exact indexed lookup after column write
  - approximate sorted lookup after column write
  - criteria-function correctness on reused ranges
  - structural row insert and move correctness
  - rebuild-config equivalence

### Exit gate

- benchmark artifacts are stable enough to compare before and after phase results
- invariants tests cover every workload that the remaining red lanes exercise

## Phase 1: Direct Change Payload Substrate

### Goal

Keep ordinary change ownership in the engine so headless never reconstructs it by diffing state.

### Benchmark proof gate

- single explicit plus recalculated change emission remains green and faster than the old snapshot
  diff route

### Delete target

- any remaining ordinary headless before and after workbook diff logic

### Exit gate

- `WorkPaper` ordinary edit flows consume engine-emitted `WorkPaperChange[]` directly

## Phase 2: Compiled Plan Arena And Formula Slots

### Goal

Make shared compiled plans the primary runtime unit instead of heavyweight formula-local state.

### Work

- finish moving formula runtime identity to `planId` plus lightweight binding records
- stop storing primary lookup and criteria ownership on formula instances

### Benchmark proof gate

- no regression on current green build, recalc, and lookup workloads

### Delete target

- formula-local primary lookup and criteria ownership fields once their engine services are ready

### Exit gate

- identical repeated shapes resolve to shared compiled plans and lightweight bindings only

## Phase 3: Rebuild Execution Policy

### Goal

Make rebuild choose the cheapest valid mode and delete the generic expensive default.

### Work

- formalize `recalculateAll`, `rebuildRuntimeFromSnapshot`, and `rebuildFromPersistence`
- keep in-place full recalc as the primary path for rebuild-and-recalculate
- keep snapshot import as the primary path when config changes preserve function and language surface

### Files

- `/Users/gregkonush/github.com/bilig2/packages/headless/src/work-paper-runtime.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/rebuild-execution-policy-service.ts`

### Benchmark proof gate

- `rebuild-config-toggle` must flip green against the current HyperFormula reference on the same
  harness
- current reference from the baseline artifact is `14.454667 ms`

### Delete target

- generic persistence reconstruction as the default config-toggle rebuild path

### Exit gate

- rebuild mode selection is explicit, covered by tests, and the old default path is no longer hot

## Phase 4: Formula Template Normalization Service

### Goal

Normalize repeated row and column template formulas so mixed build and parser-cache workloads stop
paying per-cell compile cost.

### Work

- create `formula-template-normalization-service.ts`
- canonicalize repeated row-shifted and column-shifted formula shapes
- make `CompiledPlanService` intern by template family plus lightweight offset bindings
- build mixed-content sheets through template-normalized plan creation instead of repeated compile

### Files

- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-template-normalization-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-binding-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/compiled-plan-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/headless/src/initial-sheet-load.ts`

### Benchmark proof gate

- `build-parser-cache-row-templates` must flip green against the current HyperFormula reference
- `build-mixed-content` must also flip green
- current references from the baseline artifact are `50.341541 ms` and `16.372333 ms`

### Delete target

- repeated per-cell compilation of identical relative formula shapes in build and rebuild paths

### Exit gate

- mixed-content build and row-template parser-cache workloads are green and stable on rerun

## Phase 5: Suspended Bulk Mutation Lane

### Goal

Make suspended and ordinary multi-edit batches pay one mutation frontier and one emission pass.

### Work

- promote deferred suspended cell mutation batching into a first-class engine service
- make multi-column batches use one transaction scaffold, one dirty-frontier setup, and one change
  emission
- remove per-edit transaction or history scaffolding from the hot suspended path

### Files

- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/suspended-bulk-mutation-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/mutation-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/operation-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/headless/src/work-paper-runtime.ts`

### Benchmark proof gate

- `batch-edit-multi-column` must flip green
- `batch-suspended-multi-column` must flip green
- current references are `1.133042 ms` and `0.799417 ms`

### Delete target

- per-edit transaction scaffolding on the suspended and batched multi-column hot path

### Exit gate

- batch and suspended multi-column edits are green and the old per-edit hot path is displaced

## Phase 6: Range And Criterion Caches

### Goal

Make overlapping aggregates and criteria functions reuse range-owned state instead of rescanning.

### Work

- create `range-aggregate-cache-service.ts`
- create `criterion-range-cache-service.ts`
- create range cache roots in the `RangeEntityStore`
- support prefix-cache reuse and dependent-cache invalidation
- cut `COUNTIF`, `COUNTIFS`, `SUMIF`, `SUMIFS`, `AVERAGEIF`, and `AVERAGEIFS` over to the new
  service

### Files

- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/range-aggregate-cache-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/criterion-range-cache-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-evaluation-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-binding-service.ts`

### Benchmark proof gate

- `aggregate-overlapping-ranges` must flip green
- `conditional-aggregation-reused-ranges` must flip green
- current references are `3.015083 ms` and `1.188792 ms`

### Delete target

- repeated evaluator-time aggregate rescans on overlapping ranges
- repeated evaluator-time criteria rescans over reused identical ranges

### Exit gate

- aggregate and criteria reuse is range-owned and the old per-formula rescan path is gone for the
  supported shapes

## Phase 7: Structural Transform Service

### Goal

Make row and column insert, remove, and move operations dedicated engine transforms.

### Work

- create `structural-transform-service.ts`
- move row and column insert, delete, and move logic out of generic mutation loops
- update:
  - row indirection
  - range handles
  - formula address rewrites
  - exact and sorted lookup services
  - range and criterion caches
  - dependency graph generations

### Files

- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/structural-transform-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/operation-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/mutation-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-binding-service.ts`

### Benchmark proof gate

- `structural-insert-rows` must flip green
- current reference is `6.222416 ms`

### Delete target

- structural insert or move implemented as large ordinary cell-mutation loops

### Exit gate

- structural row and column operations are transformer-owned and the old generic loop path is gone

## Phase 8: Post-Write Lookup Maintenance Cutover

### Goal

Make post-write exact and approximate lookup maintenance fully mutation-owned and narrow.

### Work

- make `ExactColumnIndexService` the sole owner for post-write exact lookup maintenance
- make `SortedColumnSearchService` the sole owner for post-write approximate lookup maintenance
- stop paying evaluator refresh, rebinding, or broad invalidation on simple column writes
- narrow dirty-frontier wakeups to direct lookup subscribers

### Files

- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/exact-column-index-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/sorted-column-search-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/operation-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-evaluation-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/lookup-service.ts`

### Benchmark proof gate

- `lookup-with-column-index-after-column-write` must flip green
- `lookup-approximate-sorted-after-column-write` must flip green
- current references are `0.098500 ms` and `0.116708 ms`

### Delete target

- evaluator-owned or formula-owned post-write lookup refresh logic
- broad column-write invalidation on the exact and sorted lookup hot path

### Exit gate

- after-write exact and sorted lookup workloads are green and the old mixed-ownership path is gone

## Phase 9: RuntimeColumnStore Authority

### Goal

Make typed runtime storage authoritative for all hot-path lookup, criteria, and aggregate reads.

### Work

- complete migration of hot-path reads and writes onto typed column-native storage
- make criteria mask generation and aggregate reuse consume typed slices directly
- displace cell-object-centric hot reads from all competitive workloads

### Benchmark proof gate

- no regression on newly green lookup, criteria, aggregate, or structural lanes

### Delete target

- cell-object-centric hot reads for lookup, criteria, and aggregate workloads

### Exit gate

- typed runtime storage is the sole hot-path storage for the competitive workloads

## Phase 10: WASM Criteria And Search Kernels

### Goal

Push closed numeric and criteria-mask kernels into `packages/wasm-kernel` only after ownership is
correct.

### Work

- create kernels for:
  - numeric aggregate
  - numeric criteria mask generation
  - string-id criteria mask generation
  - exact numeric lookup
  - sorted numeric lookup
- feed kernels typed slices and explicit descriptors only
- keep criterion parsing and string wildcard semantics in JS

### Files

- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/wasm-kernel-dispatch-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-evaluation-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/wasm-kernel/assembly/dispatch-criteria-mask-numeric.ts`
- `/Users/gregkonush/github.com/bilig2/packages/wasm-kernel/assembly/dispatch-criteria-mask-string-id.ts`
- `/Users/gregkonush/github.com/bilig2/packages/wasm-kernel/assembly/dispatch-aggregate-numeric-contiguous.ts`
- `/Users/gregkonush/github.com/bilig2/packages/wasm-kernel/assembly/dispatch-lookup-exact-column.ts`
- `/Users/gregkonush/github.com/bilig2/packages/wasm-kernel/assembly/dispatch-lookup-sorted-column.ts`

### Benchmark proof gate

- criteria, aggregate, and lookup workloads remain green after kernel dispatch
- marshaling cost does not erase the gain on the target lanes

### Delete target

- JS object materialization before criteria, aggregate, or search kernel dispatch

### Exit gate

- WASM kernels accelerate the intended lanes without semantic regressions or fake wins

## Phase 11: Delete Displaced Paths And Lock Gates

### Goal

End with one clean architecture.

### Work

- delete formula-local primary lookup ownership
- delete formula-local primary criteria ownership
- delete public-surface rebuild as a default hot path
- delete structural edit implementation through generic mutation loops
- delete remaining lookup and criteria fallback ownership from evaluator hot paths
- tighten CI and benchmark gates so the new architecture is defended

### Benchmark proof gate

- `24/24` directly comparable workload wins on the expanded suite on a clean committed tree
- rerun on the same machine with `--sample-count 3 --warmup-count 1`

### Required repo gates

- `pnpm run ci`
- focused engine, headless, and WASM suites for touched subsystems
- expanded benchmark rerun on the committed tree

### Exit gate

- there is one primary runtime architecture, not two
- every red workload from the current scoreboard is green
- deleted paths are not still present for convenience

## Detailed Runtime Flow

```mermaid
sequenceDiagram
  participant API as WorkPaper API
  participant COL as RuntimeColumnStore
  participant EXI as ExactColumnIndexService
  participant SSI as SortedColumnSearchService
  participant CRC as CriterionRangeCacheService
  participant ST as StructuralTransformService
  participant DGS as DependencyGraph
  participant DFS as DirtyFrontierScheduler
  participant EXE as ExecutionTier
  participant WKS as WASM Kernel Tier
  participant CSE as ChangeSetEmitter

  API->>COL: write explicit cell or apply transform
  COL->>EXI: maintain exact buckets
  COL->>SSI: maintain sorted descriptor
  COL->>CRC: invalidate affected criteria caches
  API->>ST: apply row or column transform when structural
  ST->>EXI: apply transform metadata
  ST->>SSI: apply transform metadata
  ST->>CRC: invalidate or rewrite cache generations
  EXI->>DGS: dirty exact-lookup subscribers
  SSI->>DGS: dirty sorted-lookup subscribers
  CRC->>DGS: dirty criteria-cache subscribers
  DGS->>DFS: enqueue affected plan ids only
  DFS->>EXE: execute dirty frontier
  EXE->>WKS: run aggregate, mask, or search kernels when eligible
  WKS-->>EXE: typed result payload
  EXE->>CSE: record old and new payloads
  CSE-->>API: WorkPaperChange[]
```

## Benchmarks And Test Gates

### Required tests

- exact lookup correctness after column writes
- approximate lookup correctness after column writes
- criteria-function correctness with reused range caches
- structural row and column transform correctness
- rebuild-config equivalence
- JS and WASM parity for every new kernel

### Required benchmark command

`pnpm exec tsx packages/benchmarks/src/benchmark-workpaper-vs-hyperformula-expanded.ts --sample-count 3 --warmup-count 1`

### Required benchmark rule

No phase is done if it wins only its target workload while turning a previously green workload red.

## Stop Conditions

The program must stop and correct course if any phase produces one of these outcomes:

- benchmark win with semantic drift
- benchmark win that still routes common hot work through the legacy path
- persistent mixed ownership between evaluator code and an engine service
- WASM kernel whose marshaling cost erases the gain
- structural transform phase that still leaves generic mutation loops on the hot path
- criterion cache phase that still leaves per-formula criteria rescans for supported shapes

These are not acceptable tradeoffs.

## Delivery Sequence

The execution order is fixed:

1. direct change payload substrate
2. compiled plan arena and formula slots
3. rebuild execution policy
4. formula template normalization
5. suspended bulk mutation lane
6. range and criterion caches
7. structural transform service
8. post-write lookup maintenance cutover
9. runtime column store authority
10. WASM criteria and search kernels
11. delete displaced paths

This order matters because:

- rebuild and template normalization are the direct fixes for the current build and rebuild reds
- range and criterion caches are the direct fix for the current worst red lane
- structural transforms are the direct fix for the next worst red lane
- post-write lookup maintenance should not be revisited until broader mutation ownership is clean
- WASM is last because it amplifies the right architecture; it does not rescue the wrong one

## Done Definition

This delivery program is complete only when all of the following are true:

1. `WorkPaper` wins all directly comparable workloads in the expanded benchmark suite
2. `conditional-aggregation-reused-ranges` is range-cache-owned and green
3. `structural-insert-rows` is transformer-owned and green
4. `lookup-with-column-index-after-column-write` is exact-index-owned and green
5. `lookup-approximate-sorted-after-column-write` is sorted-service-owned and green
6. `build-parser-cache-row-templates` and `build-mixed-content` are template-normalization-owned and green
7. `rebuild-config-toggle` is rebuild-policy-owned and green
8. displaced legacy paths are deleted

That is the bar. Not “faster in some cases.” Not “cleaner but still red.” Green across the full
competitive suite, with one architecture and no fallback-first sludge.
