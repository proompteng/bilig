# Symbolic Ownership Performance Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the current remap and rematerialize ownership model with symbolic owners that can move `bilig` from narrow benchmark wins to broad performance leadership against HyperFormula.

**Architecture:** The plan replaces physical structural remap, window-owned lookup caches, range-member ownership, global topo rebuilds, replay-first restore, and eager changed-cell materialization with logical axis maps, persistent column indexes, formula families, symbolic regions, dynamic topo plus calc chain, runtime-image restore, and typed patches. The migration is phased so each cut ships on one authoritative substrate instead of keeping long-lived dual truth.

**Tech Stack:** TypeScript, Effect, Vitest, fast-check fuzz suites, Playwright, `pnpm`, existing `@bilig/formula` rewrite helpers, existing benchmark harness under `packages/benchmarks`.

---

## Scope Lock

This plan is only for engine-performance architecture.

It explicitly includes:

- structural ownership
- lookup ownership
- build and rebuild ownership
- range and dependency ownership
- dirty execution ownership
- history and patch ownership

It explicitly excludes:

- UI redesign
- protocol expansion unrelated to engine patches
- unrelated infra work
- new formula families unless they are required to keep parity while migrating ownership

## Delivery Rules

1. Do not attempt full cutover in one commit.
2. Do not run old and new ownership models in parallel longer than one phase.
3. Every phase needs:
   - one primary owner
   - targeted counters
   - correctness gates
   - competitive benchmark gates
4. No phase is complete if it improves one benchmark family by adding more cache layers on top of the old primary owner.
5. Frequent commits are required. Each workstream checkpoint must land on `main` only after the listed verification passes.

## Benchmark Truth Model

Before changing architecture, make the benchmark story trustworthy.

Required reporting dimensions:

- structural rows and structural columns reported separately
- steady-state indexed lookup and after-write lookup reported separately
- overlapping-range and sliding-window aggregates reported separately
- cold build, warm rebuild, and runtime-image restore reported separately
- compute cost and patch-materialization cost reported separately

Do not use the following as evidence of broad victory:

- `dynamic-array-filter` leadership-only support
- `rebuild-config-toggle`

## Workstream Order

The safest production order is:

1. benchmark truth and counters
2. persistent column index ownership
3. logical structural ownership
4. template bank and runtime image
5. region graph and dependency compression
6. dynamic topo and calc chain
7. typed history and typed patches

This order is not the same as theoretical purity. It is chosen for production risk:

- lookup ownership has the clearest bounded surface and best near-term ROI
- structural ownership is the most important broad family, but also the most dangerous migration
- runtime image and template ownership should not land before the storage substrate is stable
- typed patch and history work should come after the primary owners exist

## Task 1: Fix The Scorecard And Counters

**Files:**
- Create: `packages/benchmarks/src/report-competitive-families.ts`
- Create: `packages/core/src/perf/engine-counters.ts`
- Modify: `packages/benchmarks/src/benchmark-workpaper-vs-hyperformula-expanded.ts`
- Modify: `packages/benchmarks/src/__tests__/expanded-workloads.test.ts`
- Modify: `scripts/bench-contracts.ts`

**Success Criteria:**
- the competitive report groups workloads by family
- benchmark output separates structural rows, structural columns, steady-state lookup, after-write lookup, overlapping-range aggregate, sliding-window aggregate, build, rebuild, runtime restore, dirty execution, and patch emission
- engine counters exist for:
  - `cellsRemapped`
  - `rangesMaterialized`
  - `rangeMembersExpanded`
  - `formulasParsed`
  - `formulasBound`
  - `columnSliceBuilds`
  - `exactIndexBuilds`
  - `approxIndexBuilds`
  - `topoRebuilds`
  - `changedCellPayloadsBuilt`
  - `snapshotOpsReplayed`
  - `wasmFullUploads`

- [ ] Add the reporting module and counter types.
- [ ] Thread counters through the current engine hot paths without changing semantics.
- [ ] Extend the benchmark runner to emit family summaries and counters.
- [ ] Add tests that assert every expanded workload is assigned to exactly one family.
- [ ] Run:
  - `pnpm exec vitest run packages/benchmarks/src/__tests__/expanded-workloads.test.ts`
  - `pnpm bench:workpaper:competitive`
- [ ] Commit: `chore(bench): add performance family reporting`

## Task 2: Make ColumnIndexStore The Primary Lookup Owner

**Files:**
- Create: `packages/core/src/indexes/column-index-store.ts`
- Create: `packages/core/src/indexes/column-pages.ts`
- Create: `packages/core/src/indexes/criteria-index.ts`
- Modify: `packages/core/src/engine/services/runtime-column-store-service.ts`
- Modify: `packages/core/src/engine/services/lookup-column-owner.ts`
- Modify: `packages/core/src/engine/services/exact-column-index-service.ts`
- Modify: `packages/core/src/engine/services/sorted-column-search-service.ts`
- Modify: `packages/core/src/engine/runtime-state.ts`
- Test: `packages/core/src/__tests__/exact-column-index-service.test.ts`
- Test: `packages/core/src/__tests__/sorted-column-search-service.test.ts`
- Test: `packages/core/src/__tests__/lookup-service.test.ts`

**Success Criteria:**
- exact and approximate lookup stop treating window caches as the primary owner
- one `(sheetId, colId)` owner is patched on writes
- prepared lookup descriptors become lightweight views over one owner
- after-write exact and approximate lookup families move materially

- [ ] Introduce `ColumnIndexStore` and move owner state out of service-local maps.
- [ ] Rehost exact first/last positions, row lists, sortedness summaries, and monotone metadata in the store.
- [ ] Update `exact-column-index-service.ts` to use owner-backed summaries first and only keep bounded fallback logic for unsupported cases.
- [ ] Update `sorted-column-search-service.ts` to read approximate-search metadata from the owner, not request-window vectors.
- [ ] Remove `getColumnSlice()` as a primary lookup dependency from both services.
- [ ] Add tests for:
  - single write updates exact lookup without rebuilding the request window
  - batch writes update exact lookup without clearing the owner
  - approximate sorted lookup after writes reads updated owner summaries
- [ ] Run:
  - `pnpm exec vitest run packages/core/src/__tests__/exact-column-index-service.test.ts packages/core/src/__tests__/sorted-column-search-service.test.ts packages/core/src/__tests__/lookup-service.test.ts`
  - `pnpm bench:workpaper:competitive -- --sample-count 3 --warmup-count 1`
- [ ] Gate:
  - `lookup-with-column-index-after-column-write`
  - `lookup-with-column-index-after-batch-write`
  - `lookup-approximate-sorted-after-column-write`
  must all improve
- [ ] Commit: `perf(core): make column indexes owner-backed`

## Task 3: Introduce Logical Axis Maps Behind Workbook Storage

**Files:**
- Create: `packages/core/src/storage/axis-map.ts`
- Create: `packages/core/src/storage/sheet-axis-map.ts`
- Create: `packages/core/src/storage/logical-sheet-store.ts`
- Create: `packages/core/src/storage/cell-page-store.ts`
- Modify: `packages/core/src/sheet-grid.ts`
- Modify: `packages/core/src/workbook-store.ts`
- Modify: `packages/core/src/engine/structural-transaction.ts`
- Modify: `packages/core/src/engine/services/structure-service.ts`
- Test: `packages/core/src/__tests__/sheet-grid.test.ts`
- Test: `packages/core/src/__tests__/workbook-store.test.ts`
- Test: `packages/core/src/__tests__/structure-service.test.ts`

**Success Criteria:**
- row and column insert/delete/move become axis-map segment operations
- visible coordinate remap is no longer the primary structural operation
- cell storage keys are routed through logical row and column ownership
- `cellsRemapped` trends toward zero on structural edits

- [ ] Add `AxisMap` and `SheetAxisMap` types with segment-splice and move semantics.
- [ ] Introduce a logical sheet store wrapper that resolves visible `(row, col)` through axis maps.
- [ ] Move `WorkbookStore` structural operations to axis-map mutation first, with `SheetGrid` as a derived physical index rather than the authority.
- [ ] Keep the old structural path available only as a short-lived fallback during migration.
- [ ] Add unit tests for:
  - insert rows preserves stable row IDs outside the inserted span
  - delete columns removes visible positions without rewriting every stored cell coordinate
  - move rows and move columns update visible order through axis maps
- [ ] Run:
  - `pnpm exec vitest run packages/core/src/__tests__/sheet-grid.test.ts packages/core/src/__tests__/workbook-store.test.ts packages/core/src/__tests__/structure-service.test.ts`
  - `pnpm exec vitest run packages/core/src/__tests__/engine-structure.fuzz.test.ts`
- [ ] Commit: `refactor(core): add logical axis ownership`

## Task 4: Cut Structural Operations Over To Logical Ownership

**Files:**
- Modify: `packages/core/src/engine/services/structure-service.ts`
- Modify: `packages/core/src/engine/services/operation-service.ts`
- Modify: `packages/core/src/engine/services/formula-binding-service.ts`
- Modify: `packages/core/src/range-registry.ts`
- Modify: `packages/core/src/engine/services/mutation-service.ts`
- Modify: `packages/core/src/engine/live.ts`
- Test: `packages/core/src/__tests__/engine-correctness.test.ts`
- Test: `packages/core/src/__tests__/mutation-service.test.ts`
- Test: `packages/core/src/__tests__/engine-history.fuzz.test.ts`
- Test: `packages/core/src/__tests__/engine-structure.fuzz.test.ts`

**Success Criteria:**
- structure service consumes one transaction substrate
- row and column structural ops no longer widen into generic remap and repair loops
- undo/redo remain correct
- structural rows and columns move sharply toward parity

- [ ] Rehost structural transaction planning on the new axis-map substrate.
- [ ] Replace broad cell-coordinate remap with logical invalidation spans plus narrow retargeting.
- [ ] Keep existing preservation heuristics only as a compatibility layer while the new owner lands.
- [ ] Update undo history to record typed structural inverses over axis-map segments.
- [ ] Run:
  - `pnpm test:correctness:core`
  - `pnpm exec vitest run packages/core/src/__tests__/engine-history.fuzz.test.ts packages/core/src/__tests__/engine-structure.fuzz.test.ts`
  - `pnpm bench:workpaper:competitive -- --sample-count 3 --warmup-count 1`
- [ ] Gate:
  - all six structural lanes improve
  - structural columns stop being catastrophic
- [ ] Commit: `perf(core): route structure through axis maps`

## Task 5: Promote TemplateBank And FormulaInstanceTable

**Files:**
- Create: `packages/core/src/formula/template-bank.ts`
- Create: `packages/core/src/formula/formula-instance-table.ts`
- Create: `packages/core/src/formula/structural-retargeting.ts`
- Modify: `packages/core/src/engine/services/formula-template-normalization-service.ts`
- Modify: `packages/core/src/engine/services/compiled-plan-service.ts`
- Modify: `packages/core/src/engine/services/formula-binding-service.ts`
- Modify: `packages/core/src/engine/services/formula-graph-service.ts`
- Test: `packages/core/src/__tests__/formula-binding-service.test.ts`
- Test: `packages/core/src/__tests__/engine-formula-initialization.test.ts`
- Test: `packages/core/src/__tests__/formula-runtime-correctness.test.ts`

**Success Criteria:**
- repeated formulas bind as family instances, not only compiled-plan cache hits
- family retargeting metadata is explicit
- parser-template build lanes and rebuild lanes improve

- [ ] Introduce `TemplateBank` and `FormulaInstanceTable`.
- [ ] Rehost relative-template normalization as the producer for template IDs and family metadata.
- [ ] Move formula binding to instance creation over a shared family record.
- [ ] Keep direct compatibility with current compiled plans until the new family owner is complete.
- [ ] Run:
  - `pnpm exec vitest run packages/core/src/__tests__/formula-binding-service.test.ts packages/core/src/__tests__/engine-formula-initialization.test.ts packages/core/src/__tests__/formula-runtime-correctness.test.ts`
  - `pnpm bench:workpaper:competitive -- --sample-count 3 --warmup-count 1`
- [ ] Gate:
  - `build-parser-cache-row-templates`
  - `build-parser-cache-mixed-templates`
  must improve
- [ ] Commit: `perf(core): add formula family ownership`

## Task 6: Replace Snapshot Replay With Runtime Image Restore

**Files:**
- Create: `packages/core/src/snapshot/runtime-image.ts`
- Create: `packages/core/src/snapshot/runtime-image-codec.ts`
- Modify: `packages/core/src/engine/services/snapshot-service.ts`
- Modify: `packages/core/src/engine/services/formula-graph-service.ts`
- Modify: `packages/core/src/wasm-facade.ts`
- Test: `packages/core/src/__tests__/snapshot-service.test.ts`
- Test: `packages/core/src/__tests__/engine-snapshot.fuzz.test.ts`
- Test: `packages/core/src/__tests__/snapshot-wire-parity.fuzz.test.ts`

**Success Criteria:**
- snapshot restore no longer goes through generic op replay as the hot path
- runtime image sections carry axis maps, strings, cell pages, template bank, formula instances, and index metadata
- rebuild-from-snapshot improves without semantic drift

- [ ] Add runtime-image section format and versioning.
- [ ] Add runtime-image load and save paths beside the current snapshot codec.
- [ ] Keep JSON snapshot import for compatibility, but route hot restore through runtime image when valid.
- [ ] Run:
  - `pnpm exec vitest run packages/core/src/__tests__/snapshot-service.test.ts packages/core/src/__tests__/engine-snapshot.fuzz.test.ts packages/core/src/__tests__/snapshot-wire-parity.fuzz.test.ts`
  - `pnpm bench:workpaper:competitive -- --sample-count 3 --warmup-count 1`
- [ ] Gate:
  - `rebuild-runtime-from-snapshot` improves materially
- [ ] Commit: `perf(core): add runtime image restore`

## Task 7: Replace RangeRegistry Primary Ownership With RegionGraph

**Files:**
- Create: `packages/core/src/deps/region-graph.ts`
- Create: `packages/core/src/deps/region-node-store.ts`
- Create: `packages/core/src/deps/aggregate-state-store.ts`
- Create: `packages/core/src/deps/dep-pattern-store.ts`
- Modify: `packages/core/src/range-registry.ts`
- Modify: `packages/core/src/engine/services/range-aggregate-cache-service.ts`
- Modify: `packages/core/src/engine/services/criterion-range-cache-service.ts`
- Modify: `packages/core/src/engine/services/traversal-service.ts`
- Modify: `packages/core/src/engine/services/formula-binding-service.ts`
- Test: `packages/core/src/__tests__/range-aggregate-cache-service.test.ts`
- Test: `packages/core/src/__tests__/criterion-range-cache-service.test.ts`
- Test: `packages/core/src/__tests__/structure-service.test.ts`

**Success Criteria:**
- region nodes become symbolic owners
- range-member materialization is fallback only
- sliding-window aggregate reuse improves
- dirty-frontier discovery over repeated families improves

- [ ] Introduce symbolic region node kinds and canonical range composition.
- [ ] Move current prefix-reuse and criteria caches onto region-owned aggregate state.
- [ ] Add dependency-pattern compression for repeated family-to-region relationships.
- [ ] Run:
  - `pnpm exec vitest run packages/core/src/__tests__/range-aggregate-cache-service.test.ts packages/core/src/__tests__/criterion-range-cache-service.test.ts`
  - `pnpm bench:workpaper:competitive -- --sample-count 3 --warmup-count 1`
- [ ] Gate:
  - `aggregate-overlapping-sliding-window`
  - `conditional-aggregation-reused-ranges`
  - `conditional-aggregation-criteria-cell-edit`
  - `partial-recompute-mixed-frontier`
  must all improve
- [ ] Commit: `perf(core): add symbolic region ownership`

## Task 8: Replace Global Topo Rebuild With Dynamic Topo And Calc Chain

**Files:**
- Create: `packages/core/src/scheduler/dynamic-topo.ts`
- Create: `packages/core/src/scheduler/calc-chain.ts`
- Create: `packages/core/src/scheduler/dirty-frontier.ts`
- Modify: `packages/core/src/scheduler.ts`
- Modify: `packages/core/src/engine/services/formula-graph-service.ts`
- Modify: `packages/core/src/engine/live.ts`
- Test: `packages/core/src/__tests__/recalc-service.test.ts`
- Test: `packages/core/src/__tests__/engine-correctness.test.ts`
- Test: `packages/core/src/__tests__/engine-replica.fuzz.test.ts`

**Success Criteria:**
- dirty discovery is not reverse-graph BFS plus global topo rank as the common path
- local topo repair replaces most full rank rebuilds
- calc chain executes the dirty slice
- chain and fanout lanes stay green while mixed frontier improves

- [ ] Add dynamic topo labels and calc-chain storage.
- [ ] Rehost dirty execution on `DirtyFrontier + DynamicTopo + CalcChain`.
- [ ] Keep full cycle rebuild as explicit fallback only.
- [ ] Run:
  - `pnpm exec vitest run packages/core/src/__tests__/recalc-service.test.ts packages/core/src/__tests__/engine-correctness.test.ts packages/core/src/__tests__/engine-replica.fuzz.test.ts`
  - `pnpm bench:workpaper:competitive -- --sample-count 3 --warmup-count 1`
- [ ] Gate:
  - `partial-recompute-mixed-frontier` must improve
  - `single-edit-chain` and `single-edit-fanout` must not regress
- [ ] Commit: `perf(core): add dynamic topo and calc chain`

## Task 9: Replace Generic History And Changed-Cell Emission With Typed Deltas

**Files:**
- Create: `packages/core/src/history/typed-history.ts`
- Create: `packages/core/src/patches/patch-types.ts`
- Create: `packages/core/src/patches/patch-emitter.ts`
- Create: `packages/core/src/patches/materialize-changed-cells.ts`
- Modify: `packages/core/src/engine/services/history-service.ts`
- Modify: `packages/core/src/engine/services/mutation-history-fast-path.ts`
- Modify: `packages/core/src/engine/services/change-set-emitter-service.ts`
- Modify: `packages/headless/src/tracked-engine-event-refs.ts`
- Modify: `packages/headless/src/work-paper-runtime.ts`
- Test: `packages/core/src/__tests__/history-service.test.ts`
- Test: `packages/headless/src/__tests__/work-paper-runtime.test.ts`

**Success Criteria:**
- undo/redo stores typed inverse deltas
- public and headless consumers receive typed patches first
- changed-cell materialization becomes lazy
- batch plus undo families improve

- [ ] Add typed history record families and typed patch classes.
- [ ] Rehost headless and runtime event plumbing on typed patches.
- [ ] Keep lazy changed-cell materialization for compatibility APIs.
- [ ] Run:
  - `pnpm exec vitest run packages/core/src/__tests__/history-service.test.ts packages/headless/src/__tests__/work-paper-runtime.test.ts`
  - `pnpm bench:workpaper:competitive -- --sample-count 3 --warmup-count 1`
- [ ] Gate:
  - batch/undo families improve
  - no event-surface semantic regressions appear in browser or headless tests
- [ ] Commit: `perf(core): add typed history and patches`

## Cross-Phase Gates

Run full `pnpm run ci` at the end of:

- Task 2
- Task 4
- Task 6
- Task 9

Do not proceed to the next phase on a dirty tree.

## Performance Exit Criteria

The program is not complete until all of these are true:

- structural columns are no longer catastrophic
- after-write lookup families are green or near-green
- parser-template build and runtime-from-snapshot are green or near-green
- mixed-frontier recompute is green or near-green
- no broad family depends on window-owned caches or physical structural remap as primary ownership
- `pnpm run ci` is green on the final committed tree
- competitive benchmark reporting shows family-by-family movement, not one blended score

## What Not To Do

Do not spend time on these as primary strategies while this plan is active:

- more `buildDirectLookupDescriptor()` work
- more `buildDirectAggregateDescriptor()` or criteria-direct descriptor work as the main strategy
- more `recordLiteralWrite()` surgery on the current per-window lookup services
- more `resolveCellRangeDependencyReuse()` tweaks as the primary aggregate strategy
- more `structuralRewritePreserves*()` heuristics as the main structural strategy
- more snapshot-op replay tuning as the rebuild story
- more WASM full-upload polishing before JS ownership is fixed

## Execution Handoff

Plan complete and saved to `/Users/gregkonush/github.com/bilig2/docs/superpowers/plans/2026-04-18-symbolic-ownership-performance-plan.md`. Two execution options:

1. Subagent-Driven (recommended) - I dispatch a fresh subagent per task, review between tasks, fast iteration
2. Inline Execution - Execute tasks in this session using executing-plans, batch execution with checkpoints

Which approach?
