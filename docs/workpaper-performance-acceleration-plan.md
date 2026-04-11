# WorkPaper Performance Acceleration Plan

Date: `2026-04-10`

## Status

This document is the concrete performance design for making `WorkPaper` faster than HyperFormula
on the workloads that currently matter in this repo.

Current checkpoint on `main`:

- phase 1 is materially executed for the exact-vector lookup hot path
- exact `MATCH(..., range, 0)` / `XMATCH(..., range, 0[, 1|-1])` and approximate
  `MATCH(..., range, 1|-1)` / `XMATCH(..., range, 1|-1)` vector shapes now route through a
  core-owned lookup service without generic `push-range` materialization on the hot path
- `WorkPaper` ordinary mutation diffing no longer snapshots the whole workbook before and after
  each edit; it now uses engine-reported changed cells and falls back only for structural/full
  invalidations
- visibility snapshot capture now reads directly from workbook storage instead of routing through
  `engine.getCellValue()` for every cached cell
- mixed-sheet imports now stage literals before formulas and eagerly prime exact column indexes
  during formula binding, so indexed lookup no longer pays first-build cache construction inside
  the timed mutation path
- the internal engine runtime now uses synchronous service methods inside hot mutation and recalc
  loops instead of paying `Effect.runSync(...)` on every step of a batch or recalculation pass
- literal-only workbook initialization now hydrates directly into fresh core workbook storage
  instead of paying restore-style op execution overhead
- the checked-in competitive artifact now shows:
  - `build-from-sheets`: improved from `6.95x` slower to a `WorkPaper` win at `4.68x` faster
  - `single-edit-recalc`: improved from `1.42x` slower to a `WorkPaper` win at `1.61x` faster
  - `batch-edit-recalc`: improved to `1.70x` slower
  - `lookup-no-column-index`: improved from `68.76x` slower to a `WorkPaper` win at `1.51x` faster
  - `lookup-with-column-index`: improved from `124.42x` slower to `1.21x` slower
  - `lookup-approximate-sorted`: improved from `6.67x` slower to a `WorkPaper` win at `1.07x` faster
  - `range-read`: a `WorkPaper` win at `1.01x` faster

So the remaining performance gap is no longer “lookup is completely fake.” The remaining gap is
that recalculation overhead is still red. Lookup is now close enough that it no longer dominates
the program.

It complements:

- `docs/workpaper-platform-design.md`
- `docs/workpaper-engine-leadership-program.md`

Those documents define the product contract and the top-level leadership program. This document
defines the engine architecture, mutation-path changes, native-boundary changes, and proof gates
required to close the measured performance gap.

## Goal

Make `bilig` beat HyperFormula on directly comparable headless-engine workloads without giving up
its current semantic advantages.

The target claim is narrower and stricter than “faster in general”:

1. `bilig` wins the majority of directly comparable benchmark workloads in
   `packages/benchmarks/baselines/workpaper-vs-hyperformula.json`
2. `bilig` loses none of the remaining directly comparable workloads by more than `1.25x`
3. `bilig` preserves current advantages in dynamic arrays, structured references, tables, and
   publishable OSS packaging
4. every claimed win is backed by a checked-in artifact, not by anecdotal profiling

The repo now tracks two benchmark artifacts on purpose:

- `packages/benchmarks/baselines/workpaper-vs-hyperformula-expanded.json`
  This is the default competitive matrix. It expands the proof surface across more build,
  mutation, and lookup shapes.
- `packages/benchmarks/baselines/workpaper-vs-hyperformula.json`
  This is the narrow control suite. It stays small and stable so regression signals are easy to
  trust.

## Source Corpus

This design is grounded in the local HyperFormula checkout and the current `bilig` implementation.

Reviewed HyperFormula implementation:

- `/Users/gregkonush/github.com/hyperformula/src/BuildEngineFactory.ts`
- `/Users/gregkonush/github.com/hyperformula/src/Lookup/SearchStrategy.ts`
- `/Users/gregkonush/github.com/hyperformula/src/Lookup/ColumnIndex.ts`
- `/Users/gregkonush/github.com/hyperformula/src/Lookup/ColumnBinarySearch.ts`
- `/Users/gregkonush/github.com/hyperformula/src/Lookup/AdvancedFind.ts`
- `/Users/gregkonush/github.com/hyperformula/src/interpreter/plugin/LookupPlugin.ts`
- `/Users/gregkonush/github.com/hyperformula/src/Operations.ts`
- `/Users/gregkonush/github.com/hyperformula/src/ConfigParams.ts`
- `/Users/gregkonush/github.com/hyperformula/docs/guide/performance.md`

Reviewed current `bilig` implementation:

- `/Users/gregkonush/github.com/bilig/packages/core/src/engine/services/formula-evaluation-service.ts`
- `/Users/gregkonush/github.com/bilig/packages/core/src/engine/services/formula-binding-service.ts`
- `/Users/gregkonush/github.com/bilig/packages/core/src/workbook-store.ts`
- `/Users/gregkonush/github.com/bilig/packages/formula/src/js-evaluator.ts`
- `/Users/gregkonush/github.com/bilig/packages/formula/src/builtins/lookup.ts`
- `/Users/gregkonush/github.com/bilig/packages/benchmarks/baselines/workpaper-vs-hyperformula.json`
- `/Users/gregkonush/github.com/bilig/packages/benchmarks/baselines/workpaper-vs-hyperformula-expanded.json`

## Current Measured Gap

Current control-suite direct-comparison results from
`packages/benchmarks/baselines/workpaper-vs-hyperformula.json`:

| Workload | `WorkPaper` mean | HyperFormula mean | Current result |
| --- | ---: | ---: | --- |
| `build-from-sheets` | `0.833ms` | `3.897ms` | `WorkPaper` `4.68x` faster |
| `single-edit-recalc` | `1.373ms` | `2.211ms` | `WorkPaper` `1.61x` faster |
| `batch-edit-recalc` | `1.580ms` | `0.927ms` | HyperFormula `1.70x` faster |
| `range-read` | `0.217ms` | `0.220ms` | `WorkPaper` `1.01x` faster |
| `lookup-no-column-index` | `0.240ms` | `0.362ms` | `WorkPaper` `1.51x` faster |
| `lookup-with-column-index` | `0.232ms` | `0.192ms` | HyperFormula `1.21x` faster |

This is the important reading:

- `build-from-sheets` is now a real `WorkPaper` strength, not a remaining cleanup item
- `range-read` remains near parity
- recalculation is now the largest directly comparable red lane
- lookup is no longer structurally broken in the original way; both directly comparable lookup
  workloads are now near parity relative to where this program started
- the remaining lookup gap is now a narrower engine-quality problem:
  - persistent indexed search is still slower than HyperFormula’s core search subsystem
  - non-indexed exact lookup still needs a dedicated fast path

The broader matrix now makes the remaining risk clearer:

- `WorkPaper` leads `7/13` directly comparable expanded workloads
- HyperFormula still leads `6/13`
- the worst broader deficits are now:
  - `build-mixed-content`: HyperFormula `3.80x` faster
  - `batch-edit-single-column`: HyperFormula `1.54x` faster
  - `batch-edit-multi-column`: HyperFormula `1.54x` faster
  - `single-formula-edit-recalc`: HyperFormula `1.42x` faster

## Root Cause

HyperFormula is fast here because lookup is an engine subsystem. `bilig` is slower because lookup
is still mostly an evaluator behavior.

### 1. HyperFormula chooses search strategy at engine construction

HyperFormula builds `columnSearch` once in
`/Users/gregkonush/github.com/hyperformula/src/BuildEngineFactory.ts` and
`/Users/gregkonush/github.com/hyperformula/src/Lookup/SearchStrategy.ts`.

That strategy becomes either:

- `ColumnIndex` when `useColumnIndex` is enabled
- `ColumnBinarySearch` otherwise

This means lookup policy is part of the core engine graph, not an optional helper attached later.

### 2. HyperFormula maintains a persistent per-column index

`/Users/gregkonush/github.com/hyperformula/src/Lookup/ColumnIndex.ts` stores per-sheet,
per-column maps from normalized value to ordered row positions.

That index is incrementally maintained through:

- single-cell changes
- array result changes
- moved values
- row insertion and removal
- column insertion and removal
- sheet removal

The key point is placement: the index is updated on mutation paths in
`/Users/gregkonush/github.com/hyperformula/src/Operations.ts`, not rebuilt inside lookup
evaluation.

### 3. HyperFormula lookup plugins call search directly

`/Users/gregkonush/github.com/hyperformula/src/interpreter/plugin/LookupPlugin.ts` dispatches
`MATCH`, `VLOOKUP`, and `XLOOKUP` directly to `searchStrategy.find(...)`.

For exact vertical search, the engine can answer from a column index or a specialized binary
search path without first materializing a generic `CellValue[]`.

### 4. `bilig` originally materialized ranges too early

This was the starting problem for exact lookup.

At the start of this program:

- `push-range` in `/Users/gregkonush/github.com/bilig/packages/formula/src/js-evaluator.ts`
  eagerly resolves the full range
- `readRangeValues()` in
  `/Users/gregkonush/github.com/bilig/packages/core/src/engine/services/formula-evaluation-service.ts`
  loops cell-by-cell and allocates a JS array
- `resolveIndexedExactMatch()` only runs after that array already exists

That exact hot path is now closed for direct `MATCH` / `XMATCH` vector shapes. The remaining
materialization problem is broader:

- non-indexed lookup shapes still fall back to generic range materialization
- exact vertical lookup families other than the current direct `MATCH` / `XMATCH` slice are not yet
  lowered to dedicated engine ops
- the old whole-workbook before/after diff in `WorkPaper` was hiding lookup wins until that facade
  overhead was removed

So the remaining costs are:

- range traversal
- repeated value boxing/unboxing
- JS array allocation
- generic evaluator setup for fallback lookup paths
- index/query overhead that is still slower than HyperFormula’s `ColumnIndex`

### 5. `bilig` still uses the wrong boundary for native acceleration

Current WASM lookup acceleration helps computation, but not data access placement.

If a lookup call still has to:

1. resolve a generic range in JS
2. build an array in JS
3. cross the JS/WASM boundary
4. scan or partially accelerate in the kernel

then the kernel is starting too late.

The correct boundary is:

1. resolve search shape during binding
2. hand the kernel or core service a column handle and row bounds
3. read from persistent indexed data structures directly
4. materialize only the final answer

## Design Principles

1. Move search earlier, not just lower.
   Native code is only useful if it starts before generic range materialization.
2. Keep indexes persistent and mutation-driven.
   Rebuilding an index per lookup call is not an optimization.
3. Specialize only proven hot shapes.
   Exact vertical lookups should get dedicated paths first.
4. Preserve semantic source of truth in JS.
   Kernel acceleration is allowed only for closed, differential-tested paths.
5. Optimize for end-to-end workload time.
   A faster probe is irrelevant if binder, mutation, or change materialization dominates the total.

## Target Architecture

### A. Core-Owned Lookup Service

Introduce a first-class lookup subsystem in `packages/core`.

Proposed surface:

- `EngineLookupService`
- `SheetLookupIndex`
- `ColumnExactIndex`
- `ColumnSortedView`

Responsibilities:

- own persistent per-sheet, per-column exact-match indexes
- optionally own sorted-range metadata for binary-searchable columns
- expose direct lookup operations by sheet id, column, and row bounds
- receive mutation notifications from workbook and engine services
- avoid any dependency on JS-evaluator range arrays

Primary repo surfaces to change:

- `packages/core/src/workbook-store.ts`
- `packages/core/src/engine/runtime-state.ts`
- `packages/core/src/engine.ts`
- new service under `packages/core/src/engine/services/`

### B. Column-Oriented Storage Views

Current `WorkbookStore` has cell-addressed access and a minimal `columnVersions` array. That is not
enough to compete on lookup-heavy workloads.

Add lightweight column-oriented access views:

- `getColumnCellIndexes(sheetId, col, rowStart, rowEnd)`
- `getColumnValueView(sheetId, col, rowStart, rowEnd)`
- `getColumnNormalizedKeyView(sheetId, col, rowStart, rowEnd)`

Requirements:

- no intermediate `CellValue[]` allocation for lookup probes
- normalized string keys cached once, not recomputed on every lookup
- numeric and boolean keys mapped to stable comparable representations
- invalidation tied to cell mutation and sheet structural operations

This does not require a second storage engine. It requires adding cheap column projections and
stable normalization caches to the existing one.

### C. Direct Lookup Binding Path

Binding must detect and lower hot lookup shapes to engine-owned operations before JS evaluation.

Priority shapes:

1. `MATCH(value, single_column_range, 0)`
2. `XMATCH(value, single_column_range, 0[, 1 | -1])`
3. exact vertical `VLOOKUP(..., FALSE)`
4. exact vertical `XLOOKUP` with supported match/search modes

Binding behavior:

- compile these shapes to dedicated lookup opcodes or engine call descriptors
- pass:
  - sheet id
  - lookup column
  - row start/end
  - return column or return-range descriptor
  - exact/ordered mode
- bypass `push-range` for the search side entirely

Primary repo surfaces to change:

- `packages/core/src/engine/services/formula-binding-service.ts`
- `packages/formula/src/binder-wasm-rules.ts`
- `packages/formula/src/compiler.ts`
- `packages/formula/src/js-plan-lowering.ts`
- `packages/protocol`

### D. Native Boundary Redesign

WASM should accelerate closed lookup operations, not generic range materialization.

Required change:

- stop treating lookup acceleration as “ship a JS-built array to the kernel”

Preferred design:

- add dedicated opcodes such as `LookupExactColumnIndex` and `LookupOrderedColumn`
- kernel receives numeric handles or slices into engine-owned column buffers
- kernel returns row offset or not-found result
- JS performs only shape selection, fallback dispatch, and result wrapping

Allowed fallback cases:

- wildcards
- regex
- mixed-type unsupported shapes
- non-vector search ranges
- unsupported search-mode variants

Primary repo surfaces to change:

- `packages/wasm-kernel/assembly/`
- `packages/formula/src/compiler.ts`
- `packages/core/src/engine/services/formula-evaluation-service.ts`

### E. Recalc and Mutation Overhead

Lookup is the biggest red zone, but build and recalculation are still too slow to claim leadership.

Required changes:

1. reduce change-object churn during bulk edit paths
2. keep dirty-frontier tracking in compact engine-owned structures
3. avoid materializing public `WorkPaperChange[]` until a public API call actually needs them
4. cache formula binding products by normalized source where possible
5. batch structural invalidation and reverse-edge maintenance more aggressively

Primary repo surfaces to change:

- `packages/core/src/engine/services/mutation-service.ts`
- `packages/core/src/engine/services/recalc-service.ts`
- `packages/core/src/engine/services/formula-binding-service.ts`
- `packages/headless/src/work-paper-runtime.ts`

### F. Build-From-Sheets Fast Path

The original `build-from-sheets` gap is closed for literal-only fresh-workbook construction. The
broader build/import problem is now preserving that lead while extending the same low-overhead
staging to mixed-content rebuilds and more complex import shapes.

Required changes:

- preserve the literal-only direct-storage fast path
- bulk-import mixed literals and formulas without per-cell public mutation semantics
- build dependency and lookup indexes in one staged pass
- delay event/change materialization until engine warm-up completes
- avoid repeated address parse/format churn during import

Primary repo surfaces to change:

- `packages/headless/src/work-paper-runtime.ts`
- `packages/core/src/engine.ts`
- `packages/core/src/workbook-store.ts`
- `packages/core/src/engine/services/formula-binding-service.ts`

## Explicit Non-Goals

This program does not include:

- cloning HyperFormula internals line by line
- rewriting `WorkbookStore` from scratch
- pushing generic evaluator execution wholesale into WASM
- claiming speed wins before the benchmark artifact shows them
- optimizing wildcard or regex lookup before exact vector lookups are competitive

## Phased Execution

### Phase 0: Instrument First

Add workload-local counters so we can prove where time moves.

Required instrumentation:

- range materialization count and total cells materialized
- lookup probe count by shape
- lookup index hit rate vs fallback rate
- binder shape-lowering count
- exact-match direct-path count
- public-change materialization count and total emitted cells

Acceptance:

- artifact or stats dump exists for the current benchmark suite
- we can explain lookup time as a composition of:
  - materialization
  - probe
  - fallback
  - result wrapping

Status:

- partially satisfied
- the combination of `engine.getLastMetrics()` and the competitive artifact was enough to prove
  that ordinary `WorkPaper` change materialization, not recalc itself, was dominating the lookup
  benchmark after the direct exact core path landed
- the literal-only initialization fast path also proved that restore/mutation scaffolding, not raw
  cell storage, was dominating the original `build-from-sheets` workload
- dedicated workload counters are still desirable, but they are no longer blocking the next patch

### Phase 1: Move Exact Lookup Into Core

Implement `EngineLookupService` and route exact vector lookups through it without range
materialization.

Acceptance:

- `lookup-with-column-index` no longer touches generic `readRangeValues()` on the hot path
- exact `MATCH` and `XMATCH` return directly from core-owned lookup service
- correctness suite remains green
- benchmark target:
  - `lookup-with-column-index` improves from `124.42x` slower to under `10x` slower

Status:

- completed on `main`
- current checked-in result:
- `lookup-with-column-index` is now `1.39x` slower
- `lookup-no-column-index` is now a `WorkPaper` win at `1.49x` faster
- the remaining red work is now phase-2 quality work, not phase-1 plumbing

### Phase 2: Native Direct Lookup Ops

Add dedicated lookup opcodes or kernel entrypoints for closed lookup shapes.

Acceptance:

- exact indexed lookup path can run without JS-side range array allocation
- no per-call index rebuild
- benchmark target:
  - `lookup-with-column-index` under `3x` slower
  - `lookup-no-column-index` under `5x` slower

Status:

- completed on `main`
- exact indexed lookup no longer rebuilds its column index on the first post-build mutation when
  the formula was bound against a fully populated column
- approximate sorted lookup now also routes through a direct binary-search path and has improved
  from `6.67x` slower to `1.05x` slower
- `lookup-no-column-index` has crossed into a `WorkPaper` win at `1.49x` faster
- `lookup-with-column-index` is under the `3x` target at `1.39x` slower
- the remaining work is no longer to make lookup plausible; it is to turn the remaining lookup
  losses into actual wins while the program focus shifts to recalculation and mixed-content build

### Phase 3: Batch/Recalc and Build Path Reduction

Reduce mutation churn and rebinding overhead, and keep the new import lead intact.

Acceptance:

- `batch-edit-recalc` lands under `2x` slower, `single-edit-recalc` becomes a `WorkPaper` win,
  and `build-from-sheets` remains a strong `WorkPaper` win
- `range-read` remains within `1.1x`
- no regressions in external consumer smoke

Status:

- partially satisfied on `main`
- `build-from-sheets` has already crossed the finish line and is now a `WorkPaper` win at `4.37x`
  faster
- `range-read` is now a `WorkPaper` win at `1.03x` faster
- `single-edit-recalc` is now a `WorkPaper` win at `1.06x` faster after large fresh-workbook
  formula sets began synchronously initializing the kernel on Node/Bun instead of missing the
  compiled WASM path on the first real edit, coordinate-native simple-cell mutation stopped
  reparsing `sheetName + A1`, local-only headless engines stopped paying replica-version
  bookkeeping, and existing plain literal inputs gained a dedicated overwrite fast path
- `batch-edit-recalc` improved to `1.61x` slower after internal hot mutation paths stopped
  crossing the `Effect` boundary per operation, batch-literal undo stopped using the generic
  engine history builder, coordinate-native simple-cell mutation paths stopped reparsing
  `sheetName + A1`, existing plain literal inputs gained a dedicated overwrite fast path, and
  column-version invalidation started batching across local mutation bursts
- `lookup-no-column-index` is now a `WorkPaper` win at `1.49x` faster
- the remaining work in this phase is now `batch-edit-recalc`, `lookup-with-column-index`, and
  the mixed-content import path

### Phase 4: Majority Wins

Tune the remaining red workloads until the comparable benchmark suite is mostly green.

Acceptance:

- `bilig` wins at least `4/6` directly comparable workloads
- no remaining loss exceeds `1.25x`
- at least one lookup workload and one recalculation workload are clear `bilig` wins

## Benchmark and Proof Discipline

Every performance patch in this program must follow this loop:

1. add or tighten workload-specific tests first
2. run the affected correctness suites
3. run the competitive benchmark generator
4. update the checked-in artifact only if the change is understood
5. update this document or `docs/workpaper-engine-leadership-program.md` if the measured state
   changes materially

Required commands:

- `pnpm workpaper:bench:competitive:generate`
- `pnpm workpaper:bench:competitive:check`
- `pnpm workpaper:bench:competitive:control:generate`
- `pnpm workpaper:bench:competitive:control:check`
- targeted `vitest` suites for lookup, engine, and headless paths

## Risks

### Risk: faster lookup, same total workload

If change materialization, recalc scheduling, or binder churn dominates the benchmark, a much
faster lookup probe may barely move the end-to-end number.

Mitigation:

- instrument workload composition first
- measure end-to-end time after every patch

### Risk: native lookup path becomes semantically narrower than JS path

Mitigation:

- keep JS as the semantic oracle
- differential-test native path against JS fallback
- fall back explicitly for unsupported shapes

### Risk: index maintenance becomes more expensive than lookup wins

Mitigation:

- record mutation-side index maintenance cost separately
- use lazy construction only where it wins empirically
- avoid indexing columns that never participate in lookup workloads

## Completion Criteria

This document is complete only when all of the following are true:

1. the direct comparable benchmark suite shows a majority `bilig` win
2. no directly comparable workload is worse than `1.25x`
3. `lookup-with-column-index` is no longer the worst red workload
4. the implementation no longer relies on generic JS range materialization for the hot exact
   lookup path
5. the changes are shipped on `main` with:
   - updated benchmark artifact
   - green correctness tests
   - green external smoke
   - updated leadership document

## Immediate Next Patch

The next correct patch is not “celebrate phase 1.”

It is:

1. push indexed lookup deeper into engine-owned column/index structures so `useColumnIndex`
   materially outperforms the current non-indexed direct exact path
2. extend dedicated direct lookup lowering beyond the current exact `MATCH` / `XMATCH` slice
3. reduce the remaining recalculation overhead that still keeps `batch-edit-recalc` red and turn
   `lookup-with-column-index` into a real `WorkPaper` win
4. keep the direct literal initialization path fast while extending the same strategy to broader
   mixed-content imports and rebuilds
5. rerun the competitive benchmark and keep updating the artifact after each structural win

That is the shortest path from “phase 1 landed” to “majority benchmark wins.”
