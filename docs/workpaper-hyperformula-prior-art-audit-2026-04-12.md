# WorkPaper HyperFormula Prior-Art Audit

Date: `2026-04-12`

Related documents:

- `/Users/gregkonush/github.com/bilig2/docs/workpaper-ultra-performance-engine-delivery-2026-04-12.md`
- `/Users/gregkonush/github.com/bilig2/docs/workpaper-hyperformula-closeout-plan-2026-04-12.md`
- `/Users/gregkonush/github.com/bilig2/docs/workpaper-ultra-performance-engine-architecture-2026-04-12.md`
- `/Users/gregkonush/github.com/bilig2/docs/workpaper-performance-acceleration-plan.md`
- `/Users/gregkonush/github.com/bilig2/docs/workpaper-engine-leadership-program.md`

## Purpose

This document records the deeper prior-art audit of HyperFormula relevant to the last-mile
competitive benchmarks in this repo.

It exists for two reasons:

1. to explain how HyperFormula actually works in the red benchmark areas instead of relying on
   vague recollection
2. to identify where `WorkPaper` still carries structural overhead, and where order-of-magnitude
   wins are still realistic

This is not a generic HyperFormula review. It is narrowly scoped to:

- `lookup-with-column-index`
- `lookup-approximate-sorted`
- first-mutation cost after build
- fresh-sheet mixed-content build
- local mutation and post-mutation overhead

## Scope Of Audit

Reviewed HyperFormula files:

- `/Users/gregkonush/github.com/hyperformula/src/BuildEngineFactory.ts`
- `/Users/gregkonush/github.com/hyperformula/src/GraphBuilder.ts`
- `/Users/gregkonush/github.com/hyperformula/src/Evaluator.ts`
- `/Users/gregkonush/github.com/hyperformula/src/Operations.ts`
- `/Users/gregkonush/github.com/hyperformula/src/Lookup/SearchStrategy.ts`
- `/Users/gregkonush/github.com/hyperformula/src/Lookup/ColumnIndex.ts`
- `/Users/gregkonush/github.com/hyperformula/src/Lookup/ColumnBinarySearch.ts`
- `/Users/gregkonush/github.com/hyperformula/src/Lookup/AdvancedFind.ts`
- `/Users/gregkonush/github.com/hyperformula/src/interpreter/binarySearch.ts`

Compared against current `bilig` files:

- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-binding-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-evaluation-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/lookup-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/mutation-support-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/headless/src/work-paper-runtime.ts`

## What HyperFormula Actually Does

### 1. Search is an engine subsystem, not evaluator policy

HyperFormula constructs its search strategy once, during engine build:

- `/Users/gregkonush/github.com/hyperformula/src/BuildEngineFactory.ts:76`
- `/Users/gregkonush/github.com/hyperformula/src/BuildEngineFactory.ts:77`
- `/Users/gregkonush/github.com/hyperformula/src/BuildEngineFactory.ts:107`
- `/Users/gregkonush/github.com/hyperformula/src/BuildEngineFactory.ts:116`

The strategy switch itself is simple:

- `/Users/gregkonush/github.com/hyperformula/src/Lookup/SearchStrategy.ts:56`

Meaning:

- they do not let lookup shape emerge late in formula evaluation
- they choose a durable engine-owned search subsystem and everything else consumes that

This is the single most important architectural observation from the audit.

### 2. Exact indexed lookup is persistent column state

When `useColumnIndex` is on, `ColumnIndex.find()` tries exact indexed lookup first:

- `/Users/gregkonush/github.com/hyperformula/src/Lookup/ColumnIndex.ts:110`
- `/Users/gregkonush/github.com/hyperformula/src/Lookup/ColumnIndex.ts:115`

That exact path is:

- normalize lookup key
- look up `value -> row[]` in a column-local map
- pick first/last row within the requested range

See:

- `/Users/gregkonush/github.com/hyperformula/src/Lookup/ColumnIndex.ts:119`
- `/Users/gregkonush/github.com/hyperformula/src/Lookup/ColumnIndex.ts:132`
- `/Users/gregkonush/github.com/hyperformula/src/Lookup/ColumnIndex.ts:138`
- `/Users/gregkonush/github.com/hyperformula/src/Lookup/ColumnIndex.ts:142`

This is not a prepared formula-local descriptor. It is persistent engine-owned column state.

### 3. Approximate sorted lookup is a different mechanism

If the exact column index cannot answer, HyperFormula falls back to ordered search:

- `/Users/gregkonush/github.com/hyperformula/src/Lookup/ColumnIndex.ts:116`
- `/Users/gregkonush/github.com/hyperformula/src/Lookup/ColumnBinarySearch.ts:54`

That ordered search is implemented as direct binary search over the dependency graph’s range view:

- `/Users/gregkonush/github.com/hyperformula/src/Lookup/AdvancedFind.ts:44`
- `/Users/gregkonush/github.com/hyperformula/src/interpreter/binarySearch.ts:25`

Meaning:

- `lookup-with-column-index` and `lookup-approximate-sorted` are not one problem
- exact indexed lookup and approximate sorted lookup should not share the same primary abstraction

### 4. Mutation owns search maintenance

HyperFormula mutates search state on workbook writes and graph updates:

- formula writes:
  - `/Users/gregkonush/github.com/hyperformula/src/Operations.ts:663`
  - `/Users/gregkonush/github.com/hyperformula/src/Operations.ts:671`
  - `/Users/gregkonush/github.com/hyperformula/src/Operations.ts:673`
- value writes:
  - `/Users/gregkonush/github.com/hyperformula/src/Operations.ts:681`
  - `/Users/gregkonush/github.com/hyperformula/src/Operations.ts:684`
  - `/Users/gregkonush/github.com/hyperformula/src/Operations.ts:686`
- empty writes:
  - `/Users/gregkonush/github.com/hyperformula/src/Operations.ts:695`
  - `/Users/gregkonush/github.com/hyperformula/src/Operations.ts:702`
  - `/Users/gregkonush/github.com/hyperformula/src/Operations.ts:704`
- structural edits:
  - `/Users/gregkonush/github.com/hyperformula/src/Operations.ts:790`
  - `/Users/gregkonush/github.com/hyperformula/src/Operations.ts:791`
  - `/Users/gregkonush/github.com/hyperformula/src/Operations.ts:792`
  - `/Users/gregkonush/github.com/hyperformula/src/Operations.ts:854`
  - `/Users/gregkonush/github.com/hyperformula/src/Operations.ts:855`
  - `/Users/gregkonush/github.com/hyperformula/src/Operations.ts:856`

The takeaway is straightforward:

- the search layer is not refreshed opportunistically from formula evaluation
- it is updated as part of the mutation pipeline

### 5. Build-time work is front-loaded

During fresh workbook build, HyperFormula indexes literal cells as it constructs the graph:

- `/Users/gregkonush/github.com/hyperformula/src/GraphBuilder.ts:78`
- `/Users/gregkonush/github.com/hyperformula/src/GraphBuilder.ts:91`
- `/Users/gregkonush/github.com/hyperformula/src/GraphBuilder.ts:102`
- `/Users/gregkonush/github.com/hyperformula/src/GraphBuilder.ts:121`
- `/Users/gregkonush/github.com/hyperformula/src/GraphBuilder.ts:122`

This matters because it eliminates “first query builds the missing lookup state” behavior.

### 6. Their lazy path is narrow

`ColumnIndex.ensureRecentData()` only replays row add/remove transforms for one value bucket:

- `/Users/gregkonush/github.com/hyperformula/src/Lookup/ColumnIndex.ts:215`
- `/Users/gregkonush/github.com/hyperformula/src/Lookup/ColumnIndex.ts:221`
- `/Users/gregkonush/github.com/hyperformula/src/Lookup/ColumnIndex.ts:225`

That is a very different shape from:

- “rebuild the descriptor if the column version changed”
- “re-scan the range if the formula’s cached view may be stale”

HyperFormula’s lazy work is narrowly scoped and mutation-history-aware.

## What `bilig` Currently Does

### 1. Direct lookup is still primarily formula-local

We parse direct lookup candidates during formula binding and build descriptors onto
`RuntimeFormula`:

- candidate collection:
  - `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-binding-service.ts:300`
- descriptor construction:
  - `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-binding-service.ts:357`
- runtime formula attachment:
  - `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-binding-service.ts:973`

This is materially better than building descriptors on first evaluation, but it still means:

- lookup state is attached to formulas instead of columns
- multiple formulas referencing the same column can still carry redundant descriptor state
- mutation-time ownership is incomplete

### 2. Evaluation still owns the last-mile fast path

The remaining direct lookup fast paths live in the evaluator:

- exact and approximate prepared lookup resolution:
  - `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-evaluation-service.ts:300`
  - `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-evaluation-service.ts:339`
- direct lookup execution:
  - `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-evaluation-service.ts:385`
- uniform numeric shortcuts:
  - `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-evaluation-service.ts:391`
  - `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-evaluation-service.ts:494`

This means we are still paying:

- formula-local descriptor dispatch
- descriptor refresh checks
- evaluator-local result shaping

That is the wrong ownership layer for beating the final two lookup benchmarks.

### 3. Lookup service is halfway to the right shape

Our `lookup-service` already has the seeds of the correct architecture:

- prepared exact lookup descriptors
- prepared approximate lookup descriptors
- uniform numeric detection
- numeric/text specialization

But it is still consumed mainly as formula-owned prepared state instead of a shared engine service.

### 4. Headless still does too much post-mutation reconstruction

The headless layer translates tracked engine events into `WorkPaperChange[]` by diffing current
cell-store values against cached visibility maps:

- `/Users/gregkonush/github.com/bilig2/packages/headless/src/work-paper-runtime.ts:3110`
- `/Users/gregkonush/github.com/bilig2/packages/headless/src/work-paper-runtime.ts:3193`
- `/Users/gregkonush/github.com/bilig2/packages/headless/src/work-paper-runtime.ts:3261`
- `/Users/gregkonush/github.com/bilig2/packages/headless/src/work-paper-runtime.ts:3327`

That overhead is still large enough to dominate small-mutation benchmarks, especially:

- exact lookup with one edited input and one recalculated formula
- approximate sorted lookup with one edited input and one recalculated formula

### 5. Event change composition is improved but still not the dominant issue

We added a small-path fast branch for `1 explicit + 1 recalculated` in:

- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/mutation-support-service.ts:851`

That is correct, but it is not sufficient. The remaining lookup deficit is larger than this
particular union cost.

## Empirical Findings In This Checkout

These measurements came from local instrumentation on this checkout during the audit.

### First mutation cost is still much higher than steady state

For `lookup-with-column-index` on a fresh workbook:

- first mutation: roughly `0.80ms`
- second mutation: roughly `0.17ms`
- warmed steady state: roughly `0.05ms`

For `lookup-approximate-sorted`:

- first mutation: roughly `0.24-0.33ms`
- second mutation: roughly `0.15-0.17ms`
- warmed steady state: still materially closer to HyperFormula than the first-mutation number

Meaning:

- our steady-state lookup math is already close
- the remaining benchmark gap is heavily influenced by first-mutation path cost

### Headless change materialization is a major part of small mutation cost

For exact indexed lookup, one representative breakdown on this checkout showed:

- `recalculateNowSync`: about `0.09ms`
- `computeCellChangesFromTrackedEvents`: about `0.20ms`
- `captureChanges`: about `0.73ms`

Meaning:

- the engine math is no longer the main story
- headless post-mutation work can exceed the recalc cost itself

## Where 10-50x Is Actually Available

10-50x is not realistic on the whole benchmark totals anymore. The remaining workloads are around
`0.05ms` to `0.10ms`, so total-workload 10x gains are not grounded in reality.

Where 10-50x is still realistic is on specific surplus overhead buckets.

### 1. Headless event-to-change reconstruction

The two red lookup workloads currently trigger tiny logical mutations:

- one explicit input cell
- one recalculated formula cell

But headless still:

- drains tracked events
- translates changed cell indices to row/col/sheet
- looks up before-values from visibility maps
- materializes `WorkPaperChange[]`
- sorts/fixes ordering when needed

If the engine emitted direct cell changes with old/new values, most of that path disappears.

That is a real place where `10x` or more is plausible for the surplus overhead, not the whole
workload.

### 2. Formula-local lookup descriptor refresh/dispatch

Even after the current improvements, a direct lookup still starts from the formula’s
`directLookup` descriptor and runs evaluator-side branching.

A true engine-owned exact index service can reduce:

- descriptor dispatch
- refresh checks
- formula-local duplication

Again, this is a `10x` opportunity on the remaining lookup-management overhead, not on the
entire benchmark.

### 3. First-mutation post-build setup

The warmed steady-state exact indexed lookup path is already very close to HyperFormula. The big
remaining delta is the first mutation after build.

A design that fully finishes interactive lookup state during build can remove most of that cliff.

That is the clearest remaining “order of magnitude on the leftover cost” opportunity.

## The Architectural Gap, Stated Plainly

HyperFormula:

- search state is engine-owned
- exact indexed lookup and approximate sorted lookup are separate mechanisms
- mutation owns maintenance
- build owns initial construction
- evaluator consumes search services

`bilig` today:

- search state is still largely formula-owned
- exact and approximate direct lookups still share too much abstraction
- evaluator still owns too much of the last-mile fast path
- headless still reconstructs post-mutation changes too expensively

That is why we are close but not ahead.

## What We Need To Do To Actually Beat It

### 1. Replace formula-owned exact descriptors with a shared exact column index service

Requirements:

- key by sheet id and column
- maintain numeric and string-id indexes
- update on:
  - literal writes
  - formula result changes
  - clears
  - row/column structural transforms
- serve exact-match lookups directly from mutation-owned state

This should replace the formula-owned exact prepared descriptor as the primary path.

### 2. Build a dedicated sorted column search service

Requirements:

- track monotonic numeric/text columns explicitly
- store sortedness and uniform arithmetic progression metadata in engine-owned state
- update incrementally from mutation paths
- serve approximate sorted `MATCH` directly from that state

This should not be layered on top of the exact index service.

### 3. Change the engine-to-headless contract

For small local mutations, the engine should emit:

- changed cell addresses
- new values
- ideally old values too

That lets `WorkPaper` stop reconstructing tiny change lists from cached visibility maps.

This is likely the highest-leverage remaining change for the last two benchmarks.

### 4. Front-load interactive lookup state during build

For workbook build paths:

- exact index state
- sorted column descriptors
- lookup-relevant normalization state

should be ready before the first user mutation.

The first timed mutation should not finish initializing lookup infrastructure.

## Non-Goals

These are not the right next moves:

- more parser micro-optimizations
- more formula-local prepared descriptor variants
- more benchmark-specific branching
- generic “make the evaluator faster” work without changing search ownership

Those approaches may move noise, but they do not close the structural gap identified by the audit.

## Recommended Execution Path

1. Introduce an engine-owned `ExactColumnIndexService`.
2. Introduce an engine-owned `SortedColumnSearchService`.
3. Update mutation paths so both services are maintained eagerly.
4. Add a new engine event shape for direct changed-cell payloads.
5. Change `WorkPaper` to consume direct cell-change payloads instead of visibility-map diffing when
   available.
6. Re-run the unchanged competitive suite after each phase.

If the goal is specifically “beat HyperFormula on the last two benchmarks,” this is the only
execution path from the audit that is both technically coherent and likely to finish the job.
