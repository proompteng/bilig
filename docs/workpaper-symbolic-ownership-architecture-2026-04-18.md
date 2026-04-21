# WorkPaper Symbolic Ownership Architecture

Date: `2026-04-18`

Status: `design capture, not implemented`

Related documents:

- `/Users/gregkonush/github.com/bilig2/docs/workpaper-ultra-performance-engine-architecture-2026-04-12.md`
- `/Users/gregkonush/github.com/bilig2/docs/workpaper-performance-acceleration-plan.md`
- `/Users/gregkonush/github.com/bilig2/docs/workpaper-hyperformula-prior-art-audit-2026-04-12.md`
- `/Users/gregkonush/github.com/bilig2/docs/workpaper-hyperformula-targeted-reread-2026-04-13.md`
- `/Users/gregkonush/github.com/bilig2/docs/workpaper-engine-leadership-program.md`
- `/Users/gregkonush/github.com/bilig2/docs/workpaper-sota-performance-whitepaper-roadmap-2026-04-16.md`

## Purpose

This document captures the strongest current architecture critique of `bilig`'s engine performance
line and turns it into a repo-local design reference.

The core claim is blunt:

- `bilig` will not beat `HyperFormula` by multiples across the broad competitive suite if it stays
  on the current ownership model
- current wins prove that some kernels are fast, but the broad red families show that structure,
  ranges, lookups, rebuild, and dirty execution still sit on the wrong substrate
- the path forward is not “more remap or cache tuning”; it is a replacement of the primary
  ownership model with symbolic owners

This document is intentionally stronger and more architectural than the current implementation
notes. It is a design target, not a statement that the work is already done.

## Executive Verdict

The engine is not fundamentally slow. The current benchmark wins prove that.

The ownership model is still wrong.

The current engine still treats:

- structure as coordinate rewrite plus repair
- ranges as member materialization plus cache reuse
- lookup columns as request-window caches over copied views
- rebuild as replay and rebinding work
- dirty execution as reverse-graph traversal plus global topo order
- history and public change emission as generic replay plus cell materialization

That architecture can produce narrow wins. It cannot produce broad, repeatable multi-x wins across
the full benchmark family while those ownership rules remain primary.

The required architectural line is:

- stop optimizing remap/rematerialize paths
- replace them with first-class symbolic owners:
  - logical axis maps
  - formula families
  - persistent column indexes
  - compressed region or dependency ownership
  - calc-chain based dirty execution
  - typed delta history and patch emission

## Current Benchmark Reality

The benchmark pattern that motivated this document is not random.

Recent competitive results discussed in this repo show:

- `WorkPaper` wins `13` comparable workloads
- `HyperFormula` wins `22`

The important shape of the loss matters more than the raw count:

- steady-state indexed lookup can win
- one overlapping-range aggregate lane can win
- corresponding after-write lookup lanes still lose
- sliding-window aggregate relatives still lose
- all structural row and column lanes are still red
- both parser-template build lanes are still red
- rebuild-from-snapshot is still red
- mixed-frontier partial recompute is still red

Representative red lanes from that benchmark set:

| Workload | WorkPaper | HyperFormula | Reading |
| --- | ---: | ---: | --- |
| `structural-insert-columns` | `22.606 ms` | `1.076 ms` | column structure is still on the wrong substrate |
| `structural-delete-columns` | `46.661 ms` | `13.828 ms` | delete is broad repair, not narrow ownership mutation |
| `structural-move-columns` | `18.755 ms` | `11.986 ms` | column move still behaves like generic mutation |
| `lookup-with-column-index-after-column-write` | `0.477 ms` | `0.138 ms` | steady-state owner work does not carry across writes |
| `lookup-approximate-sorted-after-column-write` | `0.982 ms` | `0.092 ms` | sorted lookup still rebuilds too much per request window |

The engine should not read those results as “tune a few hotspots.” It should read them as a map of
ownership failures.

## Root-Cause Diagnosis By Workload Family

### Structural row and column operations

Current structural edits still behave like:

1. physically rewrite visible positions
2. repair workbook metadata
3. retarget formulas or ranges
4. decide preservation heuristically

That is still the wrong primary model.

Main files:

- `/Users/gregkonush/github.com/bilig2/packages/core/src/sheet-grid.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/workbook-store.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/structure-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/structural-transaction.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/range-registry.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-binding-service.ts`

Current local heuristics like `structuralRewritePreservesValue()` and
`structuralRewritePreservesBinding()` can reduce some waste, but they still sit on top of physical
coordinate rewrite. They are repair logic around the wrong owner.

### Lookup after writes

The current runtime has the beginning of the right direction:

- `RuntimeColumnOwner`
- owner-backed column views

But lookup state is still mostly request-window owned rather than column-owner owned.

Main files:

- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/runtime-column-store-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/exact-column-index-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/sorted-column-search-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-binding-service.ts`

This is why steady-state lookup can improve while after-write lookup still loses: invalidation and
rebuild still happen at the window-cache layer, not one durable column owner.

### Approximate sorted lookup

Approximate sorted lookup still materializes request vectors and recomputes sortedness or numeric
text state for the requested window.

Main files:

- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/sorted-column-search-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/runtime-column-store-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-binding-service.ts`

Uniform-step detection helps only a narrow perfect-case family. It does not replace persistent
column-level monotone ownership.

### Build, import, and rebuild

Template normalization and compiled plan reuse are real, but they are not yet first-class runtime
owners.

Main files:

- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-template-normalization-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/compiled-plan-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-binding-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/snapshot-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-graph-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/wasm-kernel/src/index.ts`

Current rebuild still pays too much per-formula binding, graph upload, and generic restore work.
The template cache is a compiler convenience, not yet the runtime owner.

### Partial recompute, scheduler, and dirty frontier

Dirty execution is still identified through reverse-graph traversal and then ordered by a global
topological rank view.

Main files:

- `/Users/gregkonush/github.com/bilig2/packages/core/src/scheduler.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-graph-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/traversal-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-binding-service.ts`

That is good enough for friendly chain and fanout cases. It is not the right substrate for
mixed-frontier dirty execution over repeated formulas and reused regions.

### Range-heavy aggregates

The current range and aggregate system has real reuse, but too much of it still depends on
materialized members and narrow prefix tricks.

Main files:

- `/Users/gregkonush/github.com/bilig2/packages/core/src/range-registry.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/range-aggregate-cache-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/criterion-range-cache-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/traversal-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-binding-service.ts`

The overlapping-range win is real. The sliding-window loss shows that the engine still has a good
special case, not a general region owner.

### Undo, history, replay, and emitted changes

Undo and public runtime output still over-materialize too much generic cell-level state.

Main files:

- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/history-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/mutation-history-fast-path.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/change-set-emitter-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/headless/src/tracked-engine-event-refs.ts`
- `/Users/gregkonush/github.com/bilig2/packages/headless/src/work-paper-runtime.ts`

Fast paths help some simple edit lanes. They do not replace the need for typed deltas.

## Required Architecture Replacement

### High-level direction

The engine should move from:

- visible coordinate ownership
- range-member ownership
- window-cache lookup ownership
- reverse-graph plus global-rank scheduling
- restore by replay
- public changes as eagerly materialized changed cells

To:

- logical axis ownership
- stable row and column identity
- formula family ownership
- persistent column index ownership
- symbolic region ownership
- dynamic topo plus calc-chain execution
- runtime-image restore
- typed patch and delta emission

### Structural ownership

Add a storage layer under `packages/core/src/storage/`:

- `axis-map.ts`
- `sheet-axis-map.ts`
- `logical-sheet-store.ts`
- `cell-page-store.ts`

Each sheet should own:

- one row axis map
- one column axis map
- stable logical row IDs
- stable logical column IDs

Visible `(row, col)` stops being the storage key. It becomes a lookup through the axis maps into
stable logical ownership.

Result:

- insert/delete/move rows and columns become segment operations
- structure no longer rewrites hot-path cell coordinates

### Row and column address ownership

Cells should be stored in pages keyed by stable IDs, not by visible coordinates.

Existing sparse block ideas can survive, but the block key must move from visible row or column
position to stable logical ownership.

### Dependency graph ownership

Replace the current range-primary plus reverse-edge model with two central owners:

- `FormulaInstanceTable`
- `RegionGraph`

Formulas become lightweight instances over template families.
Regions become symbolic nodes, not primarily expanded member lists.

### Calc chain and topological maintenance

Replace recurring global rank rebuilds with:

- dynamic label-based topological maintenance
- calc-chain based execution of the dirty slice

Add under `packages/core/src/scheduler/`:

- `dynamic-topo.ts`
- `calc-chain.ts`
- `dirty-frontier.ts`

Normal flow:

1. typed mutation delta arrives
2. structural or cell owners apply it
3. dependency owners mark affected formula instances dirty
4. dynamic topo repairs local order if needed
5. calc chain executes the dirty frontier
6. cycle fallback remains a rare fallback, not the common path

### Formula-family compression

Promote relative-template normalization into a first-class runtime subsystem.

Add under `packages/core/src/formula/`:

- `template-bank.ts`
- `formula-instance-table.ts`
- `structural-retargeting.ts`

Each formula family should own:

- `templateId`
- relative compiled IR
- retargeting metadata
- runtime-image serialization

Each formula cell becomes:

- sheetId
- rowId
- colId
- templateId
- dependency handles
- calc-chain position
- minimal instance flags

### Lookup ownership

Replace window-owned lookup caches with a real `ColumnIndexStore`.

Add under `packages/core/src/indexes/`:

- `column-index-store.ts`
- `column-pages.ts`
- `criteria-index.ts`

Each `(sheetId, colId)` should own:

- typed column pages
- exact-match hash index
- first and last row-position lists
- sortedness metadata
- monotone-run summaries
- optional uniform-step hints
- optional criteria bitmaps or row sets

Lookup requests become lightweight views over a persistent column owner.

### Range ownership

Delete materialized range membership as the primary abstraction.

Add under `packages/core/src/deps/`:

- `region-graph.ts`
- `region-node-store.ts`
- `aggregate-state-store.ts`

Region node kinds should include:

- bounded rectangle
- row band
- column band
- composition node
- delta node
- aggregate partial node

For associative operations, canonicalize larger regions into reused parents plus strips or blocks
instead of rescanning or rematerializing overlapping windows.

### Dependency compression

Add under `packages/core/src/deps/`:

- `dep-pattern-store.ts`

This store should compress repeated formula-to-region relationships using family metadata such as:

- fixed or relative heads and tails
- owner range
- repeated offsets
- repeated band or region shapes

### History and undo

Replace generic forward or inverse replay with typed delta history.

Add:

- `packages/core/src/transactions/typed-structural-transactions.ts`
- `packages/core/src/history/typed-history.ts`

Transaction families should include:

- `CellValueDelta`
- `FormulaDefinitionDelta`
- `InsertRowsDelta`
- `DeleteRowsDelta`
- `MoveRowsDelta`
- `InsertColumnsDelta`
- `DeleteColumnsDelta`
- `MoveColumnsDelta`
- `MetadataDelta`

### Snapshot, import, and rebuild

Kill op replay as the hot restore path.

Add:

- `packages/core/src/snapshot/runtime-image.ts`
- `packages/core/src/snapshot/runtime-image-codec.ts`

Runtime image sections should include:

- workbook and sheet metadata
- strings
- axis maps
- cell pages
- template bank
- formula instances
- region graph
- column indexes
- calc-chain and topo labels
- version stamps

### Public and headless patches

Move from changed-cell materialization as the primary API to typed patches first.

Add:

- `packages/core/src/patches/patch-types.ts`
- `packages/core/src/patches/patch-emitter.ts`
- `packages/core/src/patches/materialize-changed-cells.ts`

Patch kinds:

- structural axis patch
- cell value patch
- formula-definition patch
- metadata patch
- optional viewport hint patch later

Public A1-addressed payloads should be lazy materialization for consumers that explicitly need
them.

## Delete / Demote / Keep Matrix

| Delete as primary path | Demote to fallback only | Keep and expand |
| --- | --- | --- |
| `packages/core/src/sheet-grid.ts::remapAxis` | `structure-service` preservation heuristics | workbook row or column metadata semantics, rehosted behind axis maps |
| `packages/core/src/workbook-store.ts::applyStructuralAxisTransform` | current structural rewrite preserve heuristics | relative formula translation and symbolic range parsing in `packages/formula` |
| `packages/core/src/range-registry.ts` as primary range owner | `traversal-service` range-to-cell expansion | current aggregate and criteria logic as temporary lowerings into a future `RegionGraph` |
| `runtime-column-store-service.ts::getColumnSlice` as primary API | per-window prepared lookup descriptors | existing owner-backed column page work as the seed of `ColumnIndexStore` |
| snapshot restore by generic op replay | `compiled-plan-service` and `formula-template-normalization-service` as only template owners | parser and compiler in `packages/formula` |
| change-set emission as sole public invalidation model | current scheduler/global topo rebuild path | existing benchmark harness and public runtime surfaces |

## Implementation Program

### Phase 0 — Fix the scorecard first

Goal:

- stop claiming progress from blended or misleading metrics

Required work:

- family geomeans, not one blended score
- structural rows and structural columns reported separately
- steady-state lookup and after-write lookup reported separately
- overlapping-range and sliding-window aggregate families reported separately
- cold build, warm rebuild, and runtime-image restore reported separately
- compute cost and patch-materialization cost reported separately

Suggested modules:

- `packages/benchmarks/src/report-competitive-families.ts`
- `packages/core/src/perf/engine-counters.ts`

### Phase 1 — Axis maps and persistent column indexes

Goal:

- kill physical structural remap
- kill window-owned lookup caches

Required benchmark movement before continuing:

- all structural column lanes move materially
- row structural lanes move toward parity
- exact lookup after single and batched writes moves materially
- approximate sorted after write moves materially

Suggested modules:

- `packages/core/src/storage/*`
- `packages/core/src/indexes/*`

### Phase 2 — TemplateBank, FormulaInstanceTable, and runtime image seed

Goal:

- make repeated formulas and rebuild first-class

Required benchmark movement before continuing:

- `build-parser-cache-row-templates`
- `build-parser-cache-mixed-templates`
- `rebuild-runtime-from-snapshot`

Suggested modules:

- `packages/core/src/formula/template-bank.ts`
- `packages/core/src/formula/formula-instance-table.ts`
- `packages/core/src/snapshot/runtime-image.ts`

### Phase 3 — RegionGraph and DepPatternStore

Goal:

- replace materialized ranges and explicit repeated dependency expansion

Required benchmark movement before continuing:

- `aggregate-overlapping-sliding-window`
- `partial-recompute-mixed-frontier`
- conditional aggregate reuse families

Suggested modules:

- `packages/core/src/deps/region-graph.ts`
- `packages/core/src/deps/dep-pattern-store.ts`

### Phase 4 — Dynamic topo, calc chain, and typed structural transactions

Goal:

- replace recurring global topo rebuilds
- replace structural repair heuristics

Required benchmark movement before continuing:

- `partial-recompute-mixed-frontier` turns green
- structural lanes are green or at least near-parity
- simple edit lanes stay green

Suggested modules:

- `packages/core/src/scheduler/dynamic-topo.ts`
- `packages/core/src/scheduler/calc-chain.ts`
- `packages/core/src/transactions/typed-structural-transactions.ts`

### Phase 5 — Typed history, typed patches, and WASM delta sync

Goal:

- remove replay and materialization tax from undo, redo, batch, and public runtime output

Required benchmark movement before continuing:

- batch plus undo lanes move
- suspended batch lanes move
- new patch-emission benchmark shows typed patch wins over eager changed-cell materialization

Suggested modules:

- `packages/core/src/history/typed-history.ts`
- `packages/core/src/patches/patch-types.ts`
- `packages/core/src/patches/patch-emitter.ts`

## Performance Hypotheses

| Phase | Primary families that should move | Honest expectation |
| --- | --- | --- |
| Phase 1 | structural rows and columns, indexed lookup after write, approximate sorted after write | strongest multiple-x candidate because the current base costs are obviously wrong |
| Phase 2 | parser-template build, rebuild-from-snapshot, rebuild-and-recalculate | strongest multiple-x candidate because replay and per-formula binding still dominate |
| Phase 3 | sliding-window aggregates, conditional reuse, mixed dirty frontier | multiple-x for some aggregate families, meaningful wins for mixed frontier |
| Phase 4 | partial recompute, residual structural overhead, chain or fanout consistency | parity-to-2x type work, foundational rather than flashy |
| Phase 5 | batch history, undo or redo, public patches, WASM sync overhead | 1.5x-3x likely, potentially more where patch materialization dominates |

## Risks and Failure Modes

### Semantic risks

- axis-map migration can break reference translation across named ranges, tables, spills, filters,
  validations, pivots, and cross-sheet formulas
- `RegionGraph` can mis-handle non-associative functions, criteria semantics, or error propagation
- dynamic topo plus calc chain can hide cycle bugs if the fallback rebuild path is removed too
  early
- runtime-image sections can create compatibility problems if they are not aggressively versioned

### Operational risks

- dual ownership models living too long in parallel will create two sources of truth
- keeping old direct fast paths as primary will bury the new architecture under the same special
  cases
- typed patch migration can break consumers if lazy materialization is not introduced carefully

### Ways to accidentally optimize the wrong thing

- more `buildDirectLookupDescriptor()` work can improve one benchmark and make lookup ownership more
  fragmented
- more `resolveCellRangeDependencyReuse()` tweaks can entrench `RangeRegistry`
- more `structuralRewritePreserves*()` heuristics can slightly improve rows while preserving the
  fatal coordinate-remap model
- more `recordLiteralWrite()` tuning can move one after-write lane while keeping window-owned caches
- more snapshot replay tuning can improve import slightly without solving rebuild ownership

## Current Partial Alignment

Some current work in tree is directionally aligned with this design, but does not yet satisfy it:

- owner-backed column views in `runtime-column-store-service.ts`
- lookup owner substrate in `lookup-column-owner.ts`
- render-commit fast-path and undo reductions in `mutation-service.ts`
- structural ownership cleanup work already in `structure-service.ts`

These are seeds, not the destination.

This document should not be read as “already implemented.” It should be read as:

- the current local optimizations can survive if they are rehosted behind the new owners
- the ownership replacement itself still remains to be built

## Bottom Line

The next architectural line should be:

- logical axis maps
- stable logical row and column ownership
- formula family ownership
- persistent column indexes
- symbolic region and dependency ownership
- dynamic topo plus calc-chain execution
- runtime-image restore
- typed history and typed patch emission

The engine should stop spending design energy on:

- more direct lookup descriptor tuning
- more direct aggregate descriptor tuning as a primary strategy
- more request-window cache invalidation surgery
- more structural preserve heuristics as the main structural strategy
- more snapshot replay tuning as the rebuild story

That line can plausibly beat `HyperFormula` by multiples.

The current line cannot.
