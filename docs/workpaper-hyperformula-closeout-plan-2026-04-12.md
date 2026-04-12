# WorkPaper HyperFormula Closeout Plan

Date: `2026-04-12`

Commit baseline: `ef63195` (`perf(core): tighten benchmark hot paths`)

## Purpose

This document is the narrow design plan for the next engine pass needed to beat HyperFormula on
all directly comparable competitive benchmarks in this repo.

It is intentionally narrower than:

- `/Users/gregkonush/github.com/bilig2/docs/workpaper-performance-acceleration-plan.md`
- `/Users/gregkonush/github.com/bilig2/docs/workpaper-engine-leadership-program.md`

Those documents describe the broader performance and leadership program. This document describes
the last-mile engineering plan after the `ef63195` tranche landed.

## Current State

The pushed `ef63195` tranche already landed the following real performance work:

- synchronous wasm initialization for Node/Bun benchmark runs in
  `/Users/gregkonush/github.com/bilig2/packages/core/src/engine.ts`
- direct exact and approximate lookup hot paths in
  `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-evaluation-service.ts`
  and
  `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/lookup-service.ts`
- typed exact and approximate lookup caches, including numeric and text-specialized paths, in
  `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/lookup-service.ts`
- lower-allocation mutation history and local mutation paths in
  `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/mutation-history-fast-path.ts`
  and
  `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/mutation-service.ts`
- faster fresh-sheet literal and mixed-sheet initialization in
  `/Users/gregkonush/github.com/bilig2/packages/core/src/literal-sheet-loader.ts`,
  `/Users/gregkonush/github.com/bilig2/packages/headless/src/initial-sheet-load.ts`, and
  `/Users/gregkonush/github.com/bilig2/packages/headless/src/work-paper-runtime.ts`
- cheaper topo ordering in
  `/Users/gregkonush/github.com/bilig2/packages/core/src/scheduler.ts`
- a topology-preserving formula rewrite path so dependency-equivalent formula edits no longer force
  a full topo rebuild in
  `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-binding-service.ts`
  and
  `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/operation-service.ts`

That work materially improved the benchmark surface, but it did not finish the job.

## Latest Measured Position

One confirmed local run on the `ef63195` tree from
`/tmp/workpaper-vs-hf-final.json` produced:

- `8` wins
- `5` losses

The directly comparable losses in that run were:

| Workload | WorkPaper mean | HyperFormula mean | Result |
| --- | ---: | ---: | --- |
| `build-mixed-content` | `9.776ms` | `9.619ms` | HyperFormula `1.02x` faster |
| `single-edit-chain` | `2.285ms` | `1.722ms` | HyperFormula `1.33x` faster |
| `batch-edit-single-column` | `0.680ms` | `0.644ms` | HyperFormula `1.06x` faster |
| `lookup-with-column-index` | `0.119ms` | `0.053ms` | HyperFormula `2.24x` faster |
| `lookup-approximate-sorted` | `0.078ms` | `0.049ms` | HyperFormula `1.59x` faster |

The important nuance is benchmark variance:

- the best observed run during the same optimization tranche reached `10/13` wins with only three
  losses left
- the remaining losses are therefore not all equal
- some are stable structural deficits
- some are mostly tail-latency or benchmark-variance problems

Current classification:

- `red`
  - `lookup-with-column-index`
  - `lookup-approximate-sorted`
  - `single-edit-chain`
- `yellow`
  - `build-mixed-content`
  - `batch-edit-single-column`
  - `single-formula-edit-recalc` because it oscillates across reruns even after the
    dependency-equivalent rewrite optimization

## What The Current Data Means

### 1. Lookup math is no longer the main problem

The engine floor for exact and approximate lookup is already close to HyperFormula. The remaining
gap is dominated by local mutation, dirty-set bookkeeping, and post-mutation overhead around the
lookup itself.

That means additional lookup work should focus on:

- removing remaining validation and refresh overhead from prepared descriptors
- explicit invalidation instead of repeated revalidation
- avoiding generic local mutation overhead on local-only engines

It does not mean writing another custom exact-match algorithm from scratch.

### 2. Formula-edit performance is now mostly topology and mutation overhead

The new dependency-equivalent formula rewrite path moved `single-formula-edit-recalc` much closer
to parity and in some runs into a `WorkPaper` win.

The remaining gap is not “compile faster in the abstract.” It is:

- avoiding unnecessary reverse-edge churn when dependency shape is unchanged
- avoiding unnecessary mutation-event work after the rewrite

### 3. Mixed-sheet build is still paying fresh-sheet work like an incremental mutation

`build-mixed-content` should be treated as a binder/initialization problem, not a generic build
problem.

The remaining cost is likely coming from:

- per-formula binding overhead that still assumes an incremental existing-sheet world
- repeated dependency and runtime-program plumbing while the destination sheet is still empty

### 4. Chain outliers matter more than chain median

`single-edit-chain` is frequently competitive on median and still loses on mean because the long
sample is too expensive.

That is a signal that the remaining problem is not basic arithmetic evaluation. It is likely one of:

- post-recalc event/change materialization
- local mutation bookkeeping that occasionally allocates or branches badly
- scheduler or changed-set composition tails

## Design Principles

This plan keeps the following constraints:

1. No benchmark hacks.
   The harness stays the same. No workload-specific branching in benchmark code.

2. No regressions traded for narrow wins.
   A change that wins one microbenchmark but regresses mixed-sheet build or ordinary workbook
   behavior is rejected.

3. Local-only fast paths must be explicit.
   If a path is only valid when sync/version tracking is disabled, the code must make that
   condition explicit.

4. Preserve correctness first.
   Every phase must keep `pnpm typecheck` and the affected focused tests green before measuring.

5. Use the current benchmark artifact as the truth source.
   Claims of progress are only accepted after rerunning the same competitive suite.

## The Next Design

### Phase 1: Split local-only mutation from collaborative mutation

This is the highest-value change.

Observed evidence:

- when `trackReplicaVersions=false`, the exact indexed lookup mutation floor drops materially
- the remaining indexed and approximate lookup deficits are now close to the cost of local mutation
  bookkeeping itself

Design:

- add an explicit local-only fast path inside
  `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/operation-service.ts`
  for:
  - `source === "local"`
  - no sync client connected
  - no replica-version tracking
  - no event watchers that require the full watched-cell fanout
- in that path:
  - skip entity-version bookkeeping entirely
  - avoid string-based entity key construction
  - avoid batch/send scaffolding that only matters for sync
  - emit the cheapest possible local event path

Expected workloads improved:

- `lookup-with-column-index`
- `lookup-approximate-sorted`
- `batch-edit-single-column`
- `single-edit-chain`

Files:

- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/operation-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/events.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/mutation-service.ts`

### Phase 2: Make lookup descriptors explicitly invalidated

Current prepared lookup paths are good but still not cheap enough.

Design:

- introduce a persistent lookup-descriptor registry keyed by:
  - sheet id
  - column
  - row start
  - row end
  - lookup mode family
- descriptors are built once and invalidated directly by writes to the same column/range span
- formula evaluation consumes descriptors directly with no refresh check in the hot path

This is the correct follow-on to the existing prepared lookup work in:

- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/lookup-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-evaluation-service.ts`

Expected workloads improved:

- `lookup-with-column-index`
- `lookup-approximate-sorted`
- `lookup-no-column-index` secondarily through cheaper exact direct paths

### Phase 3: Build a true fresh-sheet bulk formula bind path

Current mixed-sheet initialization is better than before, but it is still not a real bulk binder.

Design:

- add a dedicated fresh-sheet formula bind path in
  `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-binding-service.ts`
  that assumes:
  - empty destination sheet
  - no existing formula state to clear
  - row-major insertion
- preallocate dependency/program/range storage once for the full fresh-sheet batch
- avoid any reverse-edge and old-runtime cleanup logic that only exists for incremental rebinding
- have headless initialization call that path directly from:
  - `/Users/gregkonush/github.com/bilig2/packages/headless/src/initial-sheet-load.ts`
  - `/Users/gregkonush/github.com/bilig2/packages/headless/src/work-paper-runtime.ts`

Expected workloads improved:

- `build-mixed-content`

### Phase 4: Finish dependency-equivalent rewrite reuse

The current topology-preserving rewrite optimization is the right direction, but it should go
further.

Design:

- when the dependency entity list and range dependency list are unchanged:
  - skip topo rebuild
  - skip cycle detection
  - skip reverse-edge teardown and recreation where possible
  - reuse existing dependency arena slices when the graph shape is unchanged

Expected workloads improved:

- `single-formula-edit-recalc`
- `single-edit-chain` secondarily, because chain edits also pay for graph-adjacent local work

Files:

- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-binding-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/edge-arena.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/operation-service.ts`

### Phase 5: Remove chain-tail post-recalc overhead

This phase exists specifically for the remaining `single-edit-chain` mean loss.

Design:

- trace the long-sample path through:
  - changed-set composition
  - event emission
  - visibility/change materialization in headless
- remove any rare slow path that is still running even though the mutation is:
  - one existing input cell
  - one long downstream chain
  - no structural invalidation

Expected workloads improved:

- `single-edit-chain`

Files:

- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/operation-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/mutation-support-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/headless/src/work-paper-runtime.ts`

## Execution Order

The next pass should run in this order:

1. local-only mutation fast path
2. persistent invalidated lookup descriptors
3. fresh-sheet bulk formula bind
4. dependency-equivalent rewrite reuse
5. chain-tail cleanup

This order is deliberate:

- phases 1 and 2 attack the remaining red lookup workloads
- phase 3 closes the fresh-sheet mixed-content gap
- phase 4 stabilizes formula-edit wins without compromising the other work
- phase 5 is the final outlier cleanup pass

## Measurement Plan

After each phase:

1. run:
   - `pnpm typecheck`
   - the focused core/headless Vitest suites covering the touched files
2. rerun:
   - `pnpm exec tsx packages/benchmarks/src/benchmark-workpaper-vs-hyperformula-expanded.ts`
3. record:
   - workload means
   - which workloads changed
   - whether the phase improved median only or mean as well

Do not move the benchmark goalposts.

The primary scoreboard remains the current directly comparable competitive suite. Additional
profiling scripts are diagnostic only.

## Acceptance Criteria

This closeout plan is complete only when all of the following are true on the unchanged directly
comparable suite:

1. `WorkPaper` wins every directly comparable workload in the expanded competitive suite
2. no prior `WorkPaper` win regresses back to a `HyperFormula` win
3. `pnpm typecheck` is clean
4. the focused core/headless test suites for the touched engine paths are green
5. the result can be reproduced in at least two consecutive reruns without a contradictory outcome

Until all five are true, the correct statement is that the engine is closer, not finished.

## Immediate Next Move

The next concrete implementation step is:

1. add a true local-only fast path in
   `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/operation-service.ts`
2. make the path conditional on:
   - `source === "local"`
   - `trackReplicaVersions === false`
   - no sync client
3. rerun the competitive suite before touching lookup structures again

If phase 1 does not materially reduce the indexed and approximate lookup gaps, then phase 2 becomes
the primary path. If it does, phase 2 can be smaller and more surgical.
