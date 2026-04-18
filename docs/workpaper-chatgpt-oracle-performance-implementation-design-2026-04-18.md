# Bilig2 Oracle Performance Implementation Design

Date: `2026-04-18`

Status: `design capture, not implemented`

Source:

- Captured from the focused `ChatGPT` macOS app thread on `GPT-5.4 Pro`
- Grounded against current `bilig2` repo structure and the already-landed substrate work in this checkout

Related documents:

- `docs/workpaper-symbolic-ownership-architecture-2026-04-18.md`
- `docs/superpowers/plans/2026-04-18-symbolic-ownership-performance-plan.md`
- `docs/workpaper-ultra-performance-engine-architecture-2026-04-12.md`
- `docs/workpaper-performance-acceleration-plan.md`
- `docs/workpaper-hyperformula-prior-art-audit-2026-04-12.md`

## Purpose

This document captures the strongest current repo-specific performance design line from the
`ChatGPT 5.4 Pro` oracle thread and turns it into an implementation-oriented architecture document
for `bilig2`.

It is not a benchmark summary and it is not a generic performance wishlist.

It is a design statement about what the engine must become if the goal is:

- broad competitive wins against `HyperFormula`
- multi-x movement on today’s red benchmark families
- resilience against future benchmark families built around the same ownership failures

## Executive Verdict

The oracle answer is blunt and the current repo still supports it:

- `bilig2` will not beat `HyperFormula` by multiples across the broad suite if it stays on the
  current architectural line
- the engine can keep finding narrow wins on the current line
- it cannot become broadly dominant on the current line

The reason is not that the engine lacks fast kernels.

The reason is that too many hot paths are still owned by the wrong substrate:

- structure is still primarily owned by visible-coordinate remap and repair
- ranges are still primarily owned by materialized members and expansion logic
- lookups are improved, but not yet durably column-owned end to end
- build/rebuild still stop at parser or compile reuse instead of runtime ownership reuse
- dirty execution still depends on reverse-graph discovery and global ordering
- restore/history/public emission still pay replay and materialization costs too early

The current line can produce isolated green benchmarks.
It cannot plausibly produce broad multi-x leadership while these remain the primary owners.

## Current Repo Reality

The current repo already has useful seeds that should be kept:

- benchmark family reporting direction
- owner-backed column view direction
- exact-vs-approximate lookup split
- structural transaction concepts
- correctness and fuzz discipline
- wasm acceleration as a leaf execution path

Those are seeds, not the destination.

The repo is still fundamentally constrained by the following current owners:

### Structural ownership is still too physical

Primary files still carrying the wrong ownership line:

- `packages/core/src/workbook-store.ts`
- `packages/core/src/sheet-grid.ts`
- `packages/core/src/engine/services/structure-service.ts`
- `packages/core/src/engine/structural-transaction.ts`
- `packages/core/src/range-registry.ts`
- `packages/core/src/engine/services/formula-binding-service.ts`

The critical path is still:

1. remap visible grid positions
2. rebuild key and index mappings
3. refresh or retarget ranges
4. rebind formulas
5. preserve semantics with heuristics

That is still physical repair around the wrong owner.

### Range ownership is still too cell-oriented

Primary files:

- `packages/core/src/range-registry.ts`
- `packages/core/src/engine/services/traversal-service.ts`
- `packages/core/src/engine/services/range-aggregate-cache-service.ts`
- `packages/core/src/engine/services/criterion-range-cache-service.ts`
- `packages/core/src/engine/services/formula-binding-service.ts`

The engine still expands, refreshes, and materializes too much range membership in the common path.
That prevents general wins on sliding-window and criteria-heavy aggregate families.

### Lookup ownership is improved, but not yet end-state correct

Primary files:

- `packages/core/src/engine/services/runtime-column-store-service.ts`
- `packages/core/src/engine/services/lookup-column-owner.ts`
- `packages/core/src/engine/services/exact-column-index-service.ts`
- `packages/core/src/engine/services/sorted-column-search-service.ts`

The oracle answer aligns with the repo:

- exact and approximate lookup should stay split
- prepared descriptors should become lightweight views over one durable owner
- writes must patch one primary column owner instead of invalidating request-window caches

### Build and rebuild ownership is still not a runtime owner

Primary files:

- `packages/core/src/engine/services/formula-template-normalization-service.ts`
- `packages/core/src/engine/services/compiled-plan-service.ts`
- `packages/core/src/engine/services/formula-binding-service.ts`
- `packages/core/src/engine/services/formula-graph-service.ts`
- `packages/core/src/engine/services/snapshot-service.ts`
- `packages/core/src/wasm-facade.ts`

Parser and compile reuse exist, but they are still not the main runtime owner.
The red build and rebuild lanes are telling us the runtime still binds and restores on the wrong line.

### Dirty execution still uses the wrong common path

Primary files:

- `packages/core/src/scheduler.ts`
- `packages/core/src/engine/services/traversal-service.ts`
- `packages/core/src/engine/services/formula-graph-service.ts`
- `packages/core/src/engine/services/formula-binding-service.ts`

The common path is still reverse-graph discovery plus global topological ordering.
That remains acceptable for friendly chain/fanout cases and wrong for mixed frontier ownership.

### Restore, history, and public emission still materialize too much

Primary files:

- `packages/core/src/engine/services/snapshot-service.ts`
- `packages/core/src/engine/services/history-service.ts`
- `packages/core/src/engine/services/change-set-emitter-service.ts`
- `packages/core/src/engine/services/mutation-history-fast-path.ts`
- `packages/headless/src/work-paper-runtime.ts`
- `packages/headless/src/tracked-engine-event-refs.ts`

The replay-first restore and eager changed-cell materialization model hides engine wins behind
public-runtime tax.

## Required End-State Architecture

The oracle answer points to one coherent line:

- logical axis ownership
- durable column ownership
- formula-family ownership
- symbolic region and dependency ownership
- dirty-frontier ownership with local ordering repair
- runtime-image restore
- typed delta history and patch emission
- wasm synchronization hanging off the right owners

### 1. Logical structural ownership

New primary storage layer:

- `packages/core/src/storage/axis-map.ts`
- `packages/core/src/storage/sheet-axis-map.ts`
- `packages/core/src/storage/logical-sheet-store.ts`
- `packages/core/src/storage/cell-page-store.ts`

Design:

- visible row and column coordinates stop being the storage owner
- each sheet owns row and column axis maps that translate visible positions to stable logical IDs
- structural insert/delete/move become segment splice and move operations
- cell pages are keyed by logical row and column ownership, not visible coordinates

Consequence:

- `SheetGrid.remapAxis()` and `WorkbookStore.applyStructuralAxisTransform()` stop being the main path
- structural rows and columns stop scaling with broad coordinate rewrite

### 2. Durable lookup ownership

New primary lookup layer:

- `packages/core/src/indexes/column-index-store.ts`
- `packages/core/src/indexes/column-pages.ts`
- `packages/core/src/indexes/criteria-index.ts`

Design:

Each `(sheetId, logicalColId)` owns:

- typed column pages
- exact-match hash/index structures
- row-position summaries
- sortedness metadata
- monotone-run summaries
- approximate-search summaries
- criteria-oriented secondary structures where needed

Consequence:

- `exact-column-index-service.ts` and `sorted-column-search-service.ts` become query layers over one
  durable owner
- writes patch one owner
- after-write lookup families stop paying request-window rebuild cost

### 3. Formula-family ownership

New runtime formula layer:

- `packages/core/src/formula/template-bank.ts`
- `packages/core/src/formula/formula-instance-table.ts`
- `packages/core/src/formula/structural-retargeting.ts`

Design:

`TemplateBank` owns:

- canonical normalized formula families
- parsed AST and compiled template artifacts
- relative binding schema
- reusable aggregate/lookup/criteria family schema
- reusable runtime kernel template metadata

`FormulaInstanceTable` owns, per instance:

- logical anchor row/column IDs
- `templateId`
- family/fill metadata
- per-instance overrides when required
- stable linkage to dependency owners
- execution metadata used by the scheduler

Consequence:

- templates stop being cache entries and become runtime owners
- repeated formulas stop paying per-cell binding and descriptor construction tax

### 4. Symbolic region and dependency ownership

New dependency layer:

- `packages/core/src/deps/region-graph.ts`
- `packages/core/src/deps/region-node-store.ts`
- `packages/core/src/deps/aggregate-state-store.ts`
- `packages/core/src/deps/dep-pattern-store.ts`

Design:

`RegionGraph` owns symbolic regions instead of materialized members:

- bounded rectangles
- row bands
- column bands
- reusable parent-plus-strip compositions
- aggregate partials
- delta nodes

`DepPatternStore` compresses repeated family-to-region relationships.

Consequence:

- `RangeRegistry` leaves the hot path as the primary range owner
- aggregate and dirty-discovery work stop expanding back into cells as the common path

### 5. Dirty-frontier ownership and local ordering repair

New execution layer:

- `packages/core/src/scheduler/dynamic-topo.ts`
- `packages/core/src/scheduler/calc-chain.ts`
- `packages/core/src/scheduler/dirty-frontier.ts`

Design:

- mutations produce typed invalidation deltas
- dependency owners identify affected formula instances
- dirty bits are marked on instances
- dynamic ordering repair updates only the affected frontier
- calc-chain execution runs the dirty slice
- global topo rebuild becomes fallback, not the common path

Consequence:

- mixed-frontier partial recompute stops depending on reverse-graph BFS as the main substrate

### 6. Runtime-image restore

New restore layer:

- `packages/core/src/snapshot/runtime-image.ts`
- `packages/core/src/snapshot/runtime-image-codec.ts`

Design:

The runtime image should serialize the real owners:

- workbook and sheet metadata
- strings
- axis maps
- cell pages
- template bank
- formula instances
- region graph
- column indexes
- execution metadata
- version stamps

Consequence:

- `snapshot-service.ts#importWorkbook()` stops being replay-first in the performance path
- rebuild-from-snapshot becomes owner restore, not generic op replay

### 7. Typed delta history and typed public patches

New history and patch layer:

- `packages/core/src/history/typed-history.ts`
- `packages/core/src/patches/patch-types.ts`
- `packages/core/src/patches/patch-emitter.ts`
- `packages/core/src/patches/materialize-changed-cells.ts`

Design:

History stores typed inverses over the real owners:

- cell value delta
- formula definition delta
- structural axis delta
- metadata delta

Public runtime emission produces compact typed patches first, and only materializes changed cells
for consumers that explicitly require them.

Consequence:

- undo/redo stop depending on generic transaction replay as the primary model
- public runtime output stops paying eager materialization tax in the common path

### 8. WASM synchronization on the correct owners

WASM remains a keeper, but not as a top-level owner.

It should hang off:

- `TemplateBank`
- `FormulaInstanceTable`
- `RegionGraph`
- owner deltas from structural and patch layers

Consequence:

- full reset or upload behavior becomes fallback only
- JS stays the semantic owner and WASM becomes the incremental compute leaf

## Delete / Demote / Keep Matrix

### Delete as primary path

- `WorkbookStore.applyStructuralAxisTransform()` plus `SheetGrid.remapAxis()` as the main structural path
- `RangeRegistry` materialized-member ownership as the main range path
- request-window lookup caching as the main lookup path
- reverse-graph BFS plus global topo ordering as the common scheduler path
- replay-first restore in `snapshot-service.ts#importWorkbook()`
- generic transaction replay as the primary undo/history path
- eager changed-cell payload construction as the common public emission path

### Demote to fallback or compatibility

- `structuralRewritePreserves*()` preservation heuristics
- `resolveCellRangeDependencyReuse()` narrow reuse tricks
- current `RuntimeColumnOwner` and `LookupColumnOwner` as bridge owners derived from the old grid
- parser-template normalization and compiled-plan services as the full build story
- `syncWasmProgramsNow()`-style reset/upload behavior
- tracked changed-cell reconstruction in headless compatibility layers

### Keep and expand

- exact vs approximate lookup split
- owner-backed column-view direction
- structural transaction concepts, but rehosted on axis-map ownership
- benchmark family reporting
- correctness and fuzz gates
- wasm fast path, hanging from the right owners

## Implementation Program

The oracle answer supports an eight-stage line.
The existing repo-local plan is directionally right, but this is the stronger design framing.

### Stage 0: Benchmark truth and counters

Goal:

- make the performance scorecard honest
- make ownership costs visible

First files:

- `packages/benchmarks/src/report-competitive-families.ts`
- `packages/core/src/perf/engine-counters.ts`
- `packages/benchmarks/src/benchmark-workpaper-vs-hyperformula-expanded.ts`
- `scripts/bench-contracts.ts`

Required counters:

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

Exit criteria:

- benchmark families are reported separately
- compute and materialization costs are distinguishable
- fake progress becomes harder to hide

### Stage 1: Durable column ownership

Goal:

- make lookup-after-write owner-correct before tackling full structural cutover

First files:

- `packages/core/src/indexes/column-index-store.ts`
- `packages/core/src/indexes/column-pages.ts`
- `packages/core/src/indexes/criteria-index.ts`
- `packages/core/src/engine/services/runtime-column-store-service.ts`
- `packages/core/src/engine/services/exact-column-index-service.ts`
- `packages/core/src/engine/services/sorted-column-search-service.ts`

Exit criteria:

- exact and approximate lookup read durable owner state
- writes patch one owner
- after-write lookup families move materially

Expected movement:

- `2x–10x` range on exact and approximate after-write families depending on current fallback rate

### Stage 2: Logical structural substrate

Goal:

- make logical row and column ownership real behind workbook storage

First files:

- `packages/core/src/storage/axis-map.ts`
- `packages/core/src/storage/sheet-axis-map.ts`
- `packages/core/src/storage/logical-sheet-store.ts`
- `packages/core/src/storage/cell-page-store.ts`
- `packages/core/src/workbook-store.ts`
- `packages/core/src/sheet-grid.ts`

Exit criteria:

- logical row/column IDs exist and are authoritative
- cell storage resolves through axis maps
- physical remap is no longer the substrate

### Stage 3: Structural cutover

Goal:

- route structure service and structural transactions over logical ownership end to end

First files:

- `packages/core/src/engine/structural-transaction.ts`
- `packages/core/src/engine/services/structure-service.ts`
- `packages/core/src/engine/services/formula-binding-service.ts`
- `packages/core/src/range-registry.ts`

Exit criteria:

- structural row/column ops stop widening into broad remap and repair loops
- structural columns cease being catastrophic

Expected movement:

- structural rows and columns become one of the highest multi-x opportunities in the suite

### Stage 4: TemplateBank and FormulaInstanceTable

Goal:

- make repeated formulas and build/rebuild ownership real runtime owners

First files:

- `packages/core/src/formula/template-bank.ts`
- `packages/core/src/formula/formula-instance-table.ts`
- `packages/core/src/formula/structural-retargeting.ts`
- `packages/core/src/engine/services/formula-template-normalization-service.ts`
- `packages/core/src/engine/services/compiled-plan-service.ts`
- `packages/core/src/engine/services/formula-binding-service.ts`

Exit criteria:

- templates are runtime owners, not parser cache entries
- formula instances stop paying per-cell binding and descriptor cost in the common path

Expected movement:

- `2x–5x` on parser-template build/import families

### Stage 5: Runtime-image restore

Goal:

- stop replay-first restore from dominating rebuild-from-snapshot

First files:

- `packages/core/src/snapshot/runtime-image.ts`
- `packages/core/src/snapshot/runtime-image-codec.ts`
- `packages/core/src/engine/services/snapshot-service.ts`

Exit criteria:

- runtime image restores real owners directly
- snapshot replay becomes fallback or compatibility only

Expected movement:

- `3x–10x` on rebuild-from-snapshot families

### Stage 6: RegionGraph and dependency compression

Goal:

- replace materialized range ownership and narrow reuse tricks with symbolic region ownership

First files:

- `packages/core/src/deps/region-graph.ts`
- `packages/core/src/deps/region-node-store.ts`
- `packages/core/src/deps/aggregate-state-store.ts`
- `packages/core/src/deps/dep-pattern-store.ts`
- `packages/core/src/range-registry.ts`
- `packages/core/src/engine/services/traversal-service.ts`

Exit criteria:

- range-heavy families stop depending on cell expansion in the common path
- sliding-window and criteria-heavy aggregate lanes move substantially

Expected movement:

- `2x–6x` on range-heavy families depending on reuse shape

### Stage 7: Dynamic topo and calc chain

Goal:

- replace reverse-graph traversal plus global ordering as the common dirty-execution path

First files:

- `packages/core/src/scheduler/dynamic-topo.ts`
- `packages/core/src/scheduler/calc-chain.ts`
- `packages/core/src/scheduler/dirty-frontier.ts`
- `packages/core/src/scheduler.ts`
- `packages/core/src/engine/services/formula-graph-service.ts`

Exit criteria:

- dirty execution runs affected slices with local order repair
- mixed-frontier recompute stops paying global scheduling tax

Expected movement:

- `3x–8x` total on mixed-frontier partial recompute families

### Stage 8: Typed history, patches, and WASM delta synchronization

Goal:

- stop replay/materialization tax from hiding engine improvements

First files:

- `packages/core/src/history/typed-history.ts`
- `packages/core/src/patches/patch-types.ts`
- `packages/core/src/patches/patch-emitter.ts`
- `packages/core/src/patches/materialize-changed-cells.ts`
- `packages/core/src/engine/services/history-service.ts`
- `packages/core/src/engine/services/change-set-emitter-service.ts`
- `packages/headless/src/work-paper-runtime.ts`
- `packages/core/src/engine/services/formula-graph-service.ts`

Exit criteria:

- typed deltas are primary
- eager changed-cell materialization is no longer the common path
- WASM receives owner deltas instead of reset-style uploads

Expected movement:

- `1.5x–3x` on undo/history/public-runtime overhead families

## Expected Performance Movement

High-confidence movers:

- structural rows and columns: `2x–10x`
- lookup after writes: `2x–10x`
- approximate sorted lookup after writes: `4x–10x`
- runtime-image restore: `3x–10x`
- mixed-frontier partial recompute: `3x–8x`

Moderate but still meaningful movers:

- parser-template build/import: `2x–5x`
- range-heavy aggregates: `2x–6x`
- undo/history/patch emission: `1.5x–3x`

Two explicit caveats from the oracle line still apply:

- not every already-near-green lane will move by multiples
- the strongest multi-x opportunity is concentrated in the currently red ownership families

## Migration Rules

The oracle answer is clear about what would sabotage this effort.

Do not:

- add more cache layers on top of the current primary owners and call it progress
- keep old and new owners as dual truth for multiple phases
- keep request-window lookup caches as the real owner under a new interface name
- extend structural preservation heuristics instead of cutting over structural ownership
- invest in replay-first restore tuning when runtime-image restore is the real destination
- polish full WASM upload behavior before JS owner cutovers exist

Do:

- keep one authoritative owner per phase
- keep compatibility layers short-lived
- gate each stage on the benchmark family it is supposed to move
- preserve correctness and fuzz gates while changing owners

## Primary Risks

### Semantic risks

- structural reference translation across mixed refs, named expressions, tables, spills, styles,
  validations, pivots, and metadata-only regions
- dependency mis-modeling for non-associative or criteria-heavy functions inside symbolic region ownership
- local ordering repair hiding scheduler or cycle bugs if fallback rebuild paths are removed too early
- runtime-image version skew if sectioning and compatibility are sloppy

### Delivery risks

- long-lived dual truth between physical and logical owners
- improving one benchmark family while entrenching the wrong primary owner
- structural migration landing before lookup and storage ownership boundaries are clear
- patch/history work landing before typed owner deltas exist

## Immediate Recommendation

If the team is serious about beating `HyperFormula` by multiples, the execution line should be:

1. finish making the scorecard honest
2. make `ColumnIndexStore` the real lookup owner
3. make `AxisMap` and logical storage the real structural owner
4. promote `TemplateBank` and `FormulaInstanceTable`
5. replace replay restore with `RuntimeImage`
6. replace range-member ownership with `RegionGraph` and dependency compression
7. replace reverse-graph common-path scheduling with `DynamicTopo` + `CalcChain`
8. finish with typed history, typed patches, and owner-based WASM delta sync

That is the strongest production-safe design line currently supported by both the repo and the
`ChatGPT 5.4 Pro` oracle response.
