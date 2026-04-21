# WorkPaper Oracle SOTA Performance Design

Date: `2026-04-21`
Source: ChatGPT native app, model selector showing `5.4 Pro`
Input archive: `/private/tmp/codex-share/bilig2-codebase-da133b5a-oracle.zip`
Repo snapshot: `/Users/gregkonush/github.com/bilig2`, `main` at `da133b5a`
Status: `captured oracle design memo for execution planning`

This document preserves the full oracle response captured from ChatGPT 5.4 Pro after it reviewed the attached codebase archive and the structural/symbolic ownership design docs. The text below is the copied oracle memo; local repo context is limited to this header.

---

Executive verdict

The current symbolic ownership plan is directionally right but not sufficient as an execution plan to get multiple-x performance over HyperFormula. It names the right destinations—logical axes, persistent indexes, formula families, symbolic regions, dynamic topo, runtime-image restore, typed patches—but it does not yet replace the live bottleneck: structural operations still enumerate and physically remap scoped cells.

The dominant problem is not “formula rewrite needs one more heuristic.” The uploaded code still has a physical-remap core:

* SheetGrid.collectAxisRemapEntries() scans blocks and emits one entry per moved/deleted cell, including column-path scans across every row in each touched block: packages/core/src/sheet-grid.ts:120-200.
* WorkbookStore.planStructuralAxisTransform() calls that collector before the axis edit: packages/core/src/workbook-store.ts:967-994.
* WorkbookStore.applyPlannedStructuralTransaction() deletes and rewrites cellKeyToIndex, SheetGrid, and cellStore.rows/cols for every remapped survivor: packages/core/src/workbook-store.ts:996-1022.
* Delete undo capture still scans sheet cells through captureAxisRangeCellState(): packages/core/src/engine/services/structure-service.ts:286-304, called from delete-column undo construction at packages/core/src/engine/services/mutation-service.ts:914-929.

That architecture cannot produce multiple-x structural wins over HyperFormula. The right replacement is a Structural Edit Kernel: cells bind to stable row/column identities; axis edits mutate an order/index structure; surviving cells are not rewritten; SheetGrid becomes a cache, not the owner. Then formula dependencies must move from per-cell/per-range expansion to formula-family and region arrangements. HyperFormula’s public docs explicitly call out dense/sparse address mapping, column indexes, dependency graph range nodes, and range composition to avoid n² arcs; WorkPaper has partial analogues, but they are not yet authoritative.  ￼

I did not rerun pnpm run ci or the benchmark suite in this environment. This is a source/doc audit of the uploaded checkout, using your provided recent benchmark numbers as the current performance baseline.

⸻

Findings against current docs/code

1. 2026-04-16-finish-structural-ownership.md

Still correct

The doc’s central thesis is correct: structural ownership must become one authoritative transaction substrate consumed by WorkbookStore, RangeRegistry, StructureService, MutationService, and event emission. The doc’s own done criteria still describe the missing state well: structural inverse replay must stop capturing unrelated sheet state, structural insert must stop flooding changed cells, and delete/move must stop being dominated by broad structural bookkeeping: docs/superpowers/plans/2026-04-16-finish-structural-ownership.md:929-936.

Obsolete

The benchmark baselines are stale. That doc quotes column insert around 21–23 ms vs HyperFormula 0.55–0.64 ms and other older lane numbers: docs/superpowers/plans/2026-04-16-finish-structural-ownership.md:43-50. Your current facts are worse for column insert on the 2-sample gate: 33.278 ms vs 0.580 ms, HyperFormula 57.33x.

The “first finish pass” gates are also obsolete for the target you set. A gate allowing structural-insert-columns <= 4x HyperFormula is not compatible with “multiple-x over HyperFormula”: docs/superpowers/plans/2026-04-16-finish-structural-ownership.md:884-893.

Partially implemented but not authoritative

StructuralTransaction exists, but it is still a physical cell-remap payload, not a structural edit command. Its core fields are remappedCells, removedCellIndices, and invalidationSpans: packages/core/src/engine/structural-transaction.ts:20-28. That is better than ad hoc remapping, but still proportional to moved cells.

WorkbookStore already has axisMap, logicalAxisMap, and LogicalSheetStore: packages/core/src/workbook-store.ts:130-144, packages/core/src/workbook-store.ts:200-213. This is the correct substrate. It is not yet the authority because cellStore.rows/cols, cellKeyToIndex, and SheetGrid still get physically updated for every survivor.

Dangerous

Treating this doc as “almost done” is dangerous. The code still has both a planned remap path and legacy remap helpers:

* remapSheetCells(): packages/core/src/workbook-store.ts:928-965
* applyStructuralAxisTransform(): packages/core/src/workbook-store.ts:1025-1066

Those should be treated as fallback/legacy code slated for removal, not as the final ownership model.

2. 2026-04-18-symbolic-ownership-performance-plan.md

Still correct

The high-level architecture is the right direction: logical axis maps, persistent column indexes, formula families, symbolic regions, dynamic topo plus calc chain, runtime-image restore, and typed patches: docs/superpowers/plans/2026-04-18-symbolic-ownership-performance-plan.md:5-8.

Its warning is also correct: do not stack new caches on top of old owners; each phase needs a primary owner, counters, correctness gates, and competitive benchmark gates: docs/superpowers/plans/2026-04-18-symbolic-ownership-performance-plan.md:33-43.

Obsolete or out of order

The listed order puts persistent column index ownership before logical structural ownership: docs/superpowers/plans/2026-04-18-symbolic-ownership-performance-plan.md:64-72. With your current structural-column numbers, structural must move first or in parallel with instrumentation. A lookup-first plan cannot fix a 57x loss on column insert.

Partially implemented but not authoritative

The plan’s Task 1 counter set exists in packages/core/src/perf/engine-counters.ts:1-14, and family grouping exists in packages/benchmarks/src/report-competitive-families.ts:4-103. But the additional workload helper file still returns a legacy BenchmarkSample without engineCounters: packages/benchmarks/src/benchmark-workpaper-vs-hyperformula-expanded-support.ts:12-16, :66-83. That matches your note that structural helpers lack counters.

Dangerous

The plan is too broad to hand to a coding agent as-is. It names many subsystems but does not specify the invariant migration that prevents stale dual truth between logicalAxisMap, SheetGrid, cellStore.rows/cols, and cellKeyToIndex.

3. workpaper-pre-stage1-performance-state-2026-04-18.md

Still correct

The doc correctly states that counter plumbing is not universal. It records 14 workloads with counters and 22 without, and says additional-workload helper paths still emit legacy payloads: docs/workpaper-pre-stage1-performance-state-2026-04-18.md:166-199.

Obsolete

Its scorecard is stale: it records 15 WorkPaper wins vs 19 HyperFormula wins: docs/workpaper-pre-stage1-performance-state-2026-04-18.md:201-209. Your current checkout facts say 9 vs 25 over 34 eligible comparable workloads.

4. workpaper-ultra-performance-engine-architecture-2026-04-12.md

Still useful

The doc’s conceptual goals—ultra-performance, first-class structural transforms, range/index ownership, runtime state reuse—are consistent with the architecture below.

Obsolete

Its benchmark claims are older than the current checkout facts. Do not use it as the current baseline. Use the family report plus your current 9/25 scorecard.

Dangerous

It should not justify early WASM/GPU work. The current bottleneck is ownership and representation. WASM over an O(n) physical remap is still O(n), and GPU transfer overhead is the wrong target until WorkPaper has stable columnar/region arrangements.

5. workpaper-hyperformula-targeted-reread-2026-04-13.md

Still correct

This is the best prior-art doc among the set. HyperFormula’s docs support the same conclusions: it uses address mapping strategies, column indexes for lookup, dependency graph nodes for ranges, and range composition to reduce dependency explosion.  ￼

Missing

It should be updated to say WorkPaper now has partial analogues, but they are not authoritative:

* LogicalSheetStore exists.
* RangeRegistry has prefix reuse only for one-dimensional ranges: packages/core/src/range-registry.ts:626-677.
* ColumnIndexStore exists but is capped and structural-version invalidated: packages/core/src/indexes/column-index-store.ts:39-98, packages/core/src/engine/services/lookup-column-owner.ts:514-517.
* RegionGraph is only single-column: packages/core/src/deps/region-graph.ts:24-38.

6. workpaper-sota-performance-whitepaper-roadmap-2026-04-16.md

Still correct

Its priority—finish structural transform ownership before deeper graph/topo work—is sound.

Needs replacement as an execution plan

The next plan must be narrower and more measurable. “TACO-style family-compressed dependency ownership” should become concrete FormulaFamilyStore and RegionArrangementStore work with tests, counters, and family-level acceptance thresholds. TACO’s key idea is directly relevant: exploit tabular locality to compress formula graphs, query compressed graphs without decompression, and incrementally maintain them; the paper reports very large dependent/precedent query speedups, but also notes build/maintenance tradeoffs that must be benchmarked.  ￼

⸻

Live code findings

Structural storage is dual-truth and still physically remapped

SheetGrid is a block grid with BLOCK_ROWS = 128 and BLOCK_COLS = 32: packages/core/src/sheet-grid.ts:1-2. That is a reasonable sparse/dense hybrid cache, but it is currently treated as an authority during structural edits.

The column remap path scans each scoped local column and every local row in each block: packages/core/src/sheet-grid.ts:177-198. remapAxis() then clears old positions, writes new positions, and deletes empty blocks: packages/core/src/sheet-grid.ts:256-290. That explains why column insert is catastrophic: inserting at column 1 makes every populated cell to the right a remap candidate.

The logical store is the promising part. LogicalSheetStore.getVisibleCell() resolves the current visible row/column to row/column IDs and then looks up (sheetId,rowId,colId): packages/core/src/storage/logical-sheet-store.ts:48-73. That means a structural insert or move can theoretically mutate only the axis map and leave all surviving cells in place. The current code does not yet exploit that.

AxisMap is array-backed and uses Array.splice() for structural changes: packages/core/src/storage/axis-map.ts:93-128. That is acceptable for short-term correctness, but not credible for future million-row/high-index benchmarks.

Undo capture is a hidden structural bottleneck

Delete-row and delete-column inverse construction captures metadata, deleted axis entries, cell state, formula state, and structural workbook metadata: packages/core/src/engine/services/mutation-service.ts:884-900, :914-929.

captureAxisRangeCellState() scans every sheet-grid cell and filters by row/column interval: packages/core/src/engine/services/structure-service.ts:286-304. Delete-column benchmarks will keep paying for broad scans until deleted-axis resident cells can be read directly.

Formula/range/dependency layers are partially optimized but still too materialized

TemplateBank interns templates and has a special anchored prefix aggregate path: packages/core/src/formula/template-bank.ts:81-172. This reduces parsing/translation, but it is not formula-family graph compression. Every formula still has its own runtime formula/dependency identity.

RangeRegistry has HyperFormula-like one-dimensional prefix reuse: packages/core/src/range-registry.ts:113-147, :626-677. But refresh() rematerializes members and dependency sources: packages/core/src/range-registry.ts:277-342, and structural retargeting calls refresh() for touched ranges: packages/core/src/range-registry.ts:344-377.

Dynamic row/column ranges are especially risky: materializeDynamicMembers() scans all sheet cells twice: packages/core/src/range-registry.ts:524-557.

RegionGraph is a useful beginning. It interns single-column regions and uses interval trees for row containment queries: packages/core/src/deps/region-graph.ts:24-38, :45-103, :251-273. It is not yet a full symbolic region graph for rectangles, row spans, range compositions, or formula families.

Columnar/indexing exists but is not persistent enough

RuntimeColumnStoreService builds typed-array column pages from SheetGrid.blocks: packages/core/src/engine/services/runtime-column-store-service.ts:191-255. It caches by columnVersion and structureVersion: :258-355. Because every structural edit bumps structure version, structural operations invalidate owners broadly.

ColumnIndexStore is the right kind of owner for lookup lanes, but it rebuilds owners from runtime column owners and deletes stale owners by registry key: packages/core/src/indexes/column-index-store.ts:39-98.

LookupColumnOwner has a hard MAX_COLUMN_OWNER_SPAN = 65_536 and declines to build for larger populated spans: packages/core/src/engine/services/lookup-column-owner.ts:4, :514-517. That will fail credible future large-column benchmarks.

Scheduler/topo has good pieces but sparse dirty execution still scans too much

RecalcScheduler.collectDirty() uses DirtyFrontier and CalcChain: packages/core/src/scheduler.ts:25-44.

CalcChain.orderDirty() scans the full ordered formula chain for sparse dirty sets unless the dirty count equals the full chain count: packages/core/src/scheduler/calc-chain.ts:90-131. That is a future benchmark trap for “edit one input in a 100k-formula sheet.”

DynamicTopo exists and computes an affected closure and local topological repair: packages/core/src/scheduler/dynamic-topo.ts:27-139. But FormulaGraphService.repairTopoRanksNow() rebuilds the calc chain after a successful repair: packages/core/src/engine/services/formula-graph-service.ts:162-175. That limits the benefit.

Dynamic topological ordering is the right idea, but it needs affected-region counters and rank-bucket calc-chain maintenance. Pearce/Kelly’s dynamic topological sort work is directly relevant because it maintains DAG order under edge insertions/deletions and emphasizes affected-region work rather than full static recomputation.  ￼

Benchmark instrumentation is incomplete

The main expanded benchmark file has counter-aware WorkPaper helpers: packages/benchmarks/src/benchmark-workpaper-vs-hyperformula-expanded.ts:905-967.

The additional workload support file does not. Its BenchmarkSample has only elapsedMs, memory, and verification: packages/benchmarks/src/benchmark-workpaper-vs-hyperformula-expanded-support.ts:12-16, and measureMutationSample() disposes without reading counters: :66-83.

The family report is good and should become the main performance truth model: packages/benchmarks/src/report-competitive-families.ts:4-103.

⸻

Proposed architecture

Core invariant

WorkPaper needs one rule:

Visible coordinates are derived from stable sheet/row/column identities. Structural edits mutate axis order and metadata; they do not move surviving cells.

That rule replaces the current dual-truth model:

* logicalAxisMap and LogicalSheetStore should become the authority.
* cellStore.rows/cols should become cached visible coordinates, refreshed lazily or maintained only for touched/deleted cells.
* SheetGrid should become a derived lookup/cache, not the structural owner.
* cellKeyToIndex should either become a derived cache or be keyed through logical IDs.

HyperFormula’s address-mapping strategy docs are a useful external anchor: dense vs sparse address maps should be chosen per sheet, and dependency graph maintenance should not force every structural CRUD into per-cell updates.  ￼

A. Structural Edit Kernel

Purpose

Make insert/move structural edits O(axis-edit + impacted-formula/range/metadata), not O(cells shifted).

Repo anchors

* Replace physical remap authority in SheetGrid: packages/core/src/sheet-grid.ts:120-200, :256-290.
* Replace remap-list transaction in StructuralTransaction: packages/core/src/engine/structural-transaction.ts:20-28.
* Rehost WorkbookStore.planStructuralAxisTransform() and applyPlannedStructuralTransaction(): packages/core/src/workbook-store.ts:967-1022.
* Replace broad undo capture in StructureService/MutationService: packages/core/src/engine/services/structure-service.ts:286-304, packages/core/src/engine/services/mutation-service.ts:914-929.

Data structures

Create a new storage layer, likely under packages/core/src/storage/:

* axis-piece-map.ts
    * Piece/chunked axis order for row/column IDs.
    * Supports splice, move, delete, snapshot, idAt(index), indexOfId(id).
    * First implementation can be chunked arrays with chunk size 256/512; later replace with a balanced order-stat tree if needed.
* cell-axis-identity-store.ts
    * cellIndex -> rowId/colId identity mapping.
    * Use numeric internal IDs if possible; keep string compatibility at API boundary.
* axis-resident-cell-index.ts
    * rowId -> cellIndex[]/paged set
    * colId -> cellIndex[]/paged set
    * Delete-row/column captures only resident cells in deleted IDs.

StructuralTransaction v2

Replace “list of remapped survivors” with an edit command:

interface StructuralAxisEdit {
  sheetId: number
  sheetName: string
  axis: 'row' | 'column'
  kind: 'insert' | 'delete' | 'move'
  start: number
  count: number
  target?: number
  deletedAxisIds: readonly AxisId[]
  insertedAxisIds: readonly AxisId[]
  movedAxisIds?: readonly AxisId[]
}
interface StructuralTransactionV2 {
  edit: StructuralAxisEdit
  removedCellIndices: readonly number[]
  invalidationSpans: readonly StructuralInvalidationSpan[]
  formulaImpactSummary: StructuralFormulaImpactSummary
  metadataImpactSummary: StructuralMetadataImpactSummary
}

remappedCells should become a compatibility field only during migration, and counters must prove it trends to zero for insert/move.

Expected behavior

* Insert row/column: create new axis IDs; no surviving cell changes.
* Move row/column: reorder axis IDs; no surviving cell changes.
* Delete row/column: remove axis IDs; delete only resident cells whose row/column ID was removed.
* Undo delete: reinsert deleted axis IDs and restore only deleted resident cells.

Why this is necessary

HyperFormula’s docs describe address mapping strategies and structural CRUD that updates references without treating every shifted cell as a changed value. WorkPaper’s current logical store already makes this feasible, but the implementation falls back to physical survivor updates.  ￼

B. Formula-family graph

Purpose

Reduce graph size and structural rewrite work by grouping formulas with the same template and relative dependency pattern.

Repo anchors

* Build on TemplateBank: packages/core/src/formula/template-bank.ts:104-172.
* Integrate with formula binding service and graph service.
* Replace parts of collectStructuralFormulaImpacts(): packages/core/src/engine/services/structure-service.ts:969-1085.

External anchor

TACO compresses spreadsheet formula graphs by exploiting tabular locality—nearby cells often have similar formula structures—and supports querying compressed graphs without decompression and incremental maintenance. This directly maps to WorkPaper’s template bank and repeated formula workloads.  ￼

Implementation

Add packages/core/src/formula/formula-family-store.ts.

Core concepts:

type FormulaFamilyId = number
interface FormulaFamily {
  id: FormulaFamilyId
  sheetId: number
  templateId: number
  ownerRegion: RegionId
  depSpecs: readonly FamilyDependencySpec[]
  members: FormulaMemberRun[]
}
interface FamilyDependencySpec {
  kind: 'relative-cell' | 'relative-range' | 'absolute-cell' | 'absolute-range' | 'name' | 'table' | 'volatile'
  // relative offsets or symbolic region handles
}

Binding flow:

1. TemplateBank.resolve() returns templateId.
2. Formula binding computes FamilyDependencySpec[].
3. Adjacent compatible formulas are merged into a FormulaFamily.
4. Graph edges point from region/arrangement nodes to formula family nodes, not only to individual cells.
5. Evaluation still materializes individual values, but dirty discovery and structural retargeting operate at family/run granularity.

Acceptance

* A 100k-row copied formula column should bind to O(1–O(number of runs)) family records.
* Structural insert/move crossing the sheet but not splitting a family should not iterate 100k formula cells.
* Counter target: formulaFamilyCount << formulaCount, formulaFamilySplits bounded by edit intersections.

C. Region and range arrangements

Purpose

Stop expanding ranges into cell membership for dependency discovery and common aggregate evaluation.

Repo anchors

* Extend RegionGraph: packages/core/src/deps/region-graph.ts:24-38, :45-103, :251-273.
* Replace RangeRegistry.refresh() materialization: packages/core/src/range-registry.ts:277-342.
* Replace dynamic range scans: packages/core/src/range-registry.ts:524-557.

External anchors

HyperFormula explicitly models ranges as dependency graph nodes and composes large ranges to avoid n² arcs; Anti-Freeze/DataSpread highlights how cell-granularity dependency graphs blow up for formulas such as SUM(A1:A1000).  ￼

Differential Dataflow’s “arrangements” are the right systems metaphor: maintain a shared indexed representation once, then let multiple operators reuse it instead of rebuilding private indexes.  ￼

Implementation

Add region node kinds:

type RegionNode =
  | SingleColumnRegionNode
  | SingleRowRegionNode
  | RectangleRegionNode
  | AxisSpanRegionNode
  | ComposedRegionNode

Add arrangements:

* RegionArrangementStore
    * interval indexes by (sheetId, col) for single-column regions
    * interval indexes by (sheetId, row) for single-row regions
    * rectangle overlap index for rectangular formulas
    * composition edges for prefix/sliding ranges
* AggregateArrangementStore
    * prefix/Fenwick/segment-tree structures for SUM/COUNT/MIN/MAX where semantics allow
    * criteria-keyed indexes for COUNTIF/SUMIF families
    * invalidates by column/page/version, not by whole structure version

Acceptance

* rangeMembersExpanded must be zero for dependency discovery on large aggregate formulas.
* rangesMaterialized should not increase on structural insert/move unless a formula source actually changes semantics.
* Sliding-window aggregate benchmarks should be represented as region/family computations, not repeated full range expansion.

D. Persistent column/index ownership

Purpose

Make lookup and columnar aggregate lanes robust beyond current benchmark sizes.

Repo anchors

* RuntimeColumnStoreService: packages/core/src/engine/services/runtime-column-store-service.ts:191-377.
* ColumnIndexStore: packages/core/src/indexes/column-index-store.ts:39-98.
* LookupColumnOwner: packages/core/src/engine/services/lookup-column-owner.ts:492-629.

External anchor

HyperFormula exposes useColumnIndex for VLOOKUP/MATCH, trading memory for faster lookup on unsorted/large datasets; WorkPaper should own this lane rather than rebuild per window.  ￼

Implementation

* Remove the 65,536-span cliff by making LookupColumnOwner paged.
* Keep exact index as Map<normalizedKey, RowPostingList>, where row posting lists are sorted Uint32Array pages or small-array inline sets.
* Keep approximate lookup summaries per sorted run/page.
* Add criteria arrangements for SUMIF/COUNTIF.
* Make structural edits update axis mapping, not rebuild column owners wholesale.
* Split structureVersion into finer versions:
    * axisOrderVersion
    * columnValueVersion[col]
    * columnMembershipVersion[col]
    * columnMetadataVersion[col]

Acceptance

* Exact lookup after a single column write updates O(posting list delta), not rebuild owner.
* Approximate lookup after write refreshes local break/run summaries only.
* Future 1M-row lookup benchmarks should build indexes paged, not decline due to MAX_COLUMN_OWNER_SPAN.

E. Demand-driven scheduler, topo, and calc chain

Purpose

Avoid full dirty-chain scans and unnecessary recomputation when no observer needs values immediately.

Repo anchors

* RecalcScheduler: packages/core/src/scheduler.ts:25-44.
* CalcChain.orderDirty(): packages/core/src/scheduler/calc-chain.ts:90-131.
* DynamicTopo: packages/core/src/scheduler/dynamic-topo.ts:27-139.
* FormulaGraphService.repairTopoRanksNow(): packages/core/src/engine/services/formula-graph-service.ts:162-175.
* EngineOperationService recalc/event path: packages/core/src/engine/services/operation-service.ts:1913-1999.

External anchors

Salsa and Adapton are relevant because they model incremental computation as memoized, dependency-tracked, on-demand queries; Salsa’s docs frame the goal as reusing prior results after input modifications, and Adapton’s publications are explicitly about composable demand-driven incremental computation.  ￼

Excel’s recalculation model also matters: Build Systems à la Carte describes Excel using a dirty bit per cell and the previous calculation chain, with reordering/defer behavior during recalculation.  ￼

Implementation

* Add FormulaValueMemo:
    * value
    * dependency stamp/version vector summary
    * dirty flag
    * observed flag
* setCellContents() marks dependents dirty, but evaluation is demand-triggered by:
    * getCellValue
    * event listeners requiring changed values
    * snapshot/export
    * explicit recalculation
* Replace full-chain sparse dirty scan with:
    * dirty rank buckets or min-heap keyed by topo rank
    * bitset/epoch membership
    * no scan of clean formulas
* DynamicTopo.repair() should update rank buckets incrementally, not rebuild the full calc chain after every repair.

Acceptance

* Sparse edit in a 100k-formula sheet must have calcChainFullScans = 0.
* Topo repair must report topoRepairAffectedFormulas; full rebuild only for sheet delete, cycle recovery, or massive graph rewrite.
* partial-recompute-mixed-frontier becomes a family win, not a one-off.

F. Runtime image restore and typed patches

Purpose

Make rebuild/restore and event emission independent of replaying every op and materializing every changed cell.

Repo anchors

* Current counters include snapshotOpsReplayed and changedCellPayloadsBuilt: packages/core/src/perf/engine-counters.ts:1-14.
* Event path already has invalidated row/column spans: packages/core/src/engine/services/operation-service.ts:1399-1413, :1964-1999.

Implementation

* Add packages/core/src/snapshot/runtime-image.ts.
* Snapshot:
    * axis piece maps
    * cell identity/resident indexes
    * formula templates/families
    * region arrangements
    * column index metadata and optional persisted pages
    * topo ranks/calc-chain buckets
* Typed patches:
    * CellValuePatch
    * AxisInsertPatch
    * AxisDeletePatch
    * AxisMovePatch
    * RangeInvalidationPatch
    * FormulaFamilyInvalidationPatch
* Structural insert/move should emit axis patches, not changed-cell floods.

Acceptance

* Runtime restore does not call parse/bind for unchanged formula templates.
* snapshotOpsReplayed is zero or bounded by unsupported legacy sections.
* Structural insert emits invalidated columns/rows plus axis patch, not all shifted cells.

G. WASM/GPU boundary

Use WASM only after the JS ownership model is fixed. Good candidates:

* numeric vector scans over typed RuntimeColumnOwner pages
* segment tree/Fenwick bulk rebuilds
* criteria bitmap operations
* large range aggregate reductions

Do not start GPU work until the engine has stable columnar pages and a benchmark where transfer cost is amortized. GPU cannot save structural remap or formula graph ownership problems.

⸻

Benchmark and instrumentation plan

Family-level scorecard

The existing family taxonomy is good and should be kept as the top-level truth model: packages/benchmarks/src/report-competitive-families.ts:4-103.

Use three scorecards:

1. Current comparable scorecard
    * Existing expanded suite.
    * Current target baseline: 9 WorkPaper wins vs 25 HyperFormula wins over 34 eligible comparable workloads.
2. Family scorecard
    * Each eligible family has:
        * comparable count
        * WorkPaper wins
        * HyperFormula wins
        * family geomean speedup
        * worst workload ratio
3. Holdout/adversarial scorecard
    * Not checked into “trainable” baselines until the architecture is stable.
    * Used to prevent tuning only the current suite.

Final acceptance thresholds

For “multiple-x over HyperFormula,” do not accept a blended geomean win. Final done means:

* WorkPaper wins every current eligible comparable workload, or any exception has a documented semantic mismatch rather than a performance miss.
* Every eligible family has geomean speedup ≥ 2.0x over HyperFormula.
* No individual current workload is slower than HyperFormula.
* Holdout/adversarial scorecard has family geomean ≥ 1.5x and no catastrophic outlier > 1.25x slower than HyperFormula.
* Structural insert/move counters show zero survivor remaps:
    * cellsRemapped = 0 for insert/move rows/columns after Structural Edit Kernel.
    * Delete remaps/touches only deleted resident cells.
* Range dependency discovery on large ranges has rangeMembersExpanded = 0 unless evaluating a value genuinely requires materialization.
* Sparse dirty edits do not scan the full calc chain.

Counters to add

Current counters are too coarse: packages/core/src/perf/engine-counters.ts:1-14.

Add these groups.

Structural

* structuralTransactions
* structuralPlannedCells
* structuralSurvivorCellsRemapped
* structuralRemovedCells
* structuralAxisIdsInserted
* structuralAxisIdsDeleted
* structuralAxisIdsMoved
* structuralUndoCapturedCells
* structuralUndoCapturedMetadataOps
* structuralFormulaImpactCandidates
* structuralFormulaRebindInputs
* structuralRangeRetargets

Storage/cache

* sheetGridBlockScans
* sheetGridCacheRebuilds
* axisMapSplices
* axisMapMoves
* axisIndexLookups
* cellResidentIndexLookups

Range/region

* regionNodesInterned
* regionQueryIndexBuilds
* regionDependentsQueries
* regionDependentsReturned
* rangeDescriptorRetargets
* rangeFullRefreshes
* rangeLazyRefreshes

Formula families

* formulaFamiliesCreated
* formulaFamilyMembers
* formulaFamilySplits
* formulaFamilyStructuralRetargets
* formulaFamilyGraphQueries

Column/index

* columnOwnerBuilds
* columnOwnerPageBuilds
* lookupOwnerBuilds
* lookupOwnerPagedBuilds
* lookupPostingListUpdates
* approxSummaryRefreshes
* criteriaIndexBuilds
* criteriaIndexUpdates

Scheduler/topo/patch

* dirtyRoots
* dirtyFormulaCount
* dirtyRangeNodeVisits
* calcChainFullScans
* calcChainBucketReads
* topoRepairs
* topoRepairAffectedFormulas
* topoRepairFailures
* topoRebuilds
* changedCellPayloadsBuilt
* typedPatchesBuilt

Microbenchmarks

Add focused microbenchmarks under packages/benchmarks/src/ or scripts/bench-contracts.ts:

* axis-map-splice-middle-100k
* axis-piece-map-splice-middle-100k
* sheet-grid-collect-column-remap-100k
* structural-transaction-plan-insert-column-100k
* structural-delete-column-resident-index-100k
* range-registry-retarget-prefix-100k
* region-graph-query-single-column-100k
* calc-chain-order-dirty-1-of-100k
* dynamic-topo-repair-affected-100-of-100k
* lookup-owner-build-1m-paged
* lookup-owner-write-update-1m-paged
* formula-family-bind-100k-template

Macrobenchmarks

Extend expanded-competitive-workloads.ts only after the current suite has universal counters:

* structural-insert-column-wide-dense
* structural-insert-column-wide-sparse
* structural-delete-column-with-undo-wide
* formula-family-build-100k-row-template
* formula-family-structural-insert-row
* region-overlap-rectangular-ranges
* lookup-exact-1m-paged
* lookup-approx-1m-after-write
* criteria-aggregation-high-cardinality
* runtime-image-restore-with-indexes
* dirty-demand-unobserved-edit
* dirty-demand-observed-cell-edit

Adversarial benchmarks

Use these as holdouts:

* High-index sparse cells: one cell at row 900,000 and column 16,000; insert column 1.
* Dense rectangle: 10k x 100 values and formulas; insert column 1.
* Repeated copied formulas with a structural edit crossing the owner region.
* Sliding windows with varying sizes, not only one fixed window.
* Mixed volatile/non-volatile formulas to ensure demand-driven scheduling does not skip volatile semantics.
* Lookup with wildcards/regex disabled/enabled, to ensure index fallback behavior is explicit.
* Metadata-heavy structural ops: tables, filters, comments, validations, charts, images, shapes.
* Undo/redo after structural delete on dense and sparse sheets.

⸻

Staged implementation plan

Phase 0 — Benchmark truth and counters

Goal

Make every WorkPaper sample emit counters and add enough counters to identify full scans/remaps.

Files

* packages/core/src/perf/engine-counters.ts
* packages/core/src/__tests__/engine-counters.test.ts
* packages/benchmarks/src/benchmark-workpaper-vs-hyperformula-expanded-support.ts
* packages/benchmarks/src/__tests__/expanded-workloads.test.ts
* Hot-path files for counter increments:
    * workbook-store.ts
    * sheet-grid.ts
    * structure-service.ts
    * mutation-service.ts
    * range-registry.ts
    * region-graph.ts
    * runtime-column-store-service.ts
    * column-index-store.ts
    * scheduler/calc-chain.ts
    * scheduler/dynamic-topo.ts
    * formula-graph-service.ts

Tests

* Counters have stable keys and zero/reset behavior.
* Additional workload WorkPaper helpers emit engineCounters.
* Structural helpers emit structural counters.
* Existing family mapping tests still pass.

Commands

git status --short
pnpm exec vitest run packages/core/src/__tests__/engine-counters.test.ts packages/benchmarks/src/__tests__/expanded-workloads.test.ts
pnpm exec tsx packages/benchmarks/src/benchmark-workpaper-vs-hyperformula-expanded.ts --sample-count 2 --warmup-count 1

Expected performance movement

0–5%. This phase is not a win phase. Its success is trustworthy attribution.

Rollback

Revert only counter increments that perturb semantics or create measurable overhead. Keep the helper counter emission.

Phase 1 — Structural Edit Kernel

Goal

Make structural insert/move stop remapping survivor cells. Make delete touch only deleted resident cells.

Files

Create:

* packages/core/src/storage/axis-piece-map.ts
* packages/core/src/storage/cell-axis-identity-store.ts
* packages/core/src/storage/axis-resident-cell-index.ts
* packages/core/src/__tests__/axis-piece-map.test.ts
* packages/core/src/__tests__/cell-axis-identity-store.test.ts

Modify:

* packages/core/src/storage/axis-map.ts
* packages/core/src/storage/logical-sheet-store.ts
* packages/core/src/storage/cell-page-store.ts
* packages/core/src/sheet-grid.ts
* packages/core/src/workbook-store.ts
* packages/core/src/engine/structural-transaction.ts
* packages/core/src/engine/services/structure-service.ts
* packages/core/src/engine/services/mutation-service.ts
* packages/core/src/engine/services/operation-service.ts

TDD tests

* Insert column preserves visible values/formulas without survivor cell remap.
* Move column preserves visible values/formulas without survivor cell remap.
* Delete column removes only cells resident in deleted column ID.
* Delete-column undo restores only deleted resident cells and metadata.
* getCellIndex, getAddress, getQualifiedAddress, snapshot/export, and event invalidations remain correct after insert/move/delete.
* Fuzz structural row/column operations against current behavior for small sheets.

Commands

pnpm exec vitest run \
  packages/core/src/__tests__/axis-piece-map.test.ts \
  packages/core/src/__tests__/cell-axis-identity-store.test.ts \
  packages/core/src/__tests__/structural-transaction.test.ts \
  packages/core/src/__tests__/workbook-store.test.ts \
  packages/core/src/__tests__/structure-service.test.ts \
  packages/core/src/__tests__/mutation-service.test.ts \
  packages/core/src/__tests__/engine-structure.fuzz.test.ts
pnpm exec tsx packages/benchmarks/src/benchmark-workpaper-vs-hyperformula-expanded.ts --sample-count 2 --warmup-count 1

Expected performance movement

* structural-insert-columns: 33.278 ms → first target < 3 ms, final target < 0.5x HyperFormula.
* structural-delete-columns: 57.422 ms → first target < 10 ms, final target < 0.75x HyperFormula.
* structural-move-columns: 24.330 ms → first target < 5 ms, final target < 0.75x HyperFormula.
* Rows should become green once undo capture and survivor remap are removed.

Rollback

If correctness fails in address/snapshot/undo invariants, roll back only the latest migration slice. Keep Phase 0 counters. Do not restore broad physical remap as the primary path.

Phase 2 — Symbolic region and range arrangements

Goal

Make large ranges and overlapping/sliding aggregates symbolic owners, not materialized member lists.

Files

Create:

* packages/core/src/deps/region-arrangement-store.ts
* packages/core/src/formula/range-arrangement-store.ts
* packages/core/src/indexes/numeric-range-index.ts

Modify:

* packages/core/src/deps/region-graph.ts
* packages/core/src/deps/region-node-store.ts
* packages/core/src/range-registry.ts
* packages/core/src/engine/services/formula-binding-service.ts
* packages/core/src/engine/services/formula-graph-service.ts
* aggregate evaluator/direct aggregate paths

TDD tests

* Single-column region behavior preserved.
* Rectangular range dependency queries return the same dependents as expanded graph.
* Structural retarget changes descriptors without member expansion.
* Overlapping prefix aggregate range creates composed regions.
* Sliding-window aggregate does not expand every range member for dependency discovery.

Expected performance movement

This should flip overlapping/sliding aggregate and conditional aggregation families, and reduce structural formula retarget overhead.

Rollback

If region query false positives create unnecessary recalculation but correct values, keep behind a feature flag for that region kind. If false negatives occur, revert the new region kind immediately.

Phase 3 — Formula-family graph

Goal

Compress repeated formula dependencies and structural rebind work.

Files

Create:

* packages/core/src/formula/formula-family-store.ts
* packages/core/src/formula/formula-family-deps.ts
* packages/core/src/__tests__/formula-family-store.test.ts

Modify:

* packages/core/src/formula/template-bank.ts
* packages/core/src/engine/services/formula-binding-service.ts
* packages/core/src/engine/services/formula-graph-service.ts
* packages/core/src/engine/services/structure-service.ts
* packages/core/src/scheduler/dirty-frontier.ts

TDD tests

* 10k copied formulas form one or few families.
* Dependents/precedents queries match per-cell baseline on small sheets.
* Insert row/column splits only intersected families.
* Formula edit removes one member or splits a run without corrupting neighbors.

Expected performance movement

Build/parser-template workloads, partial-recompute mixed frontier, structural formula rebind, and graph traversal should move sharply.

Rollback

If family compression complicates correctness, retain TemplateBank and disable family graph for unsupported formulas. Do not delete per-cell graph until parity tests are green.

Phase 4 — Persistent paged column/index arrangements

Goal

Make lookup and criteria aggregate performance stable for large datasets and after-write workloads.

Files

Create:

* packages/core/src/indexes/paged-posting-list.ts
* packages/core/src/indexes/paged-lookup-column-owner.ts
* packages/core/src/indexes/criteria-arrangement-store.ts

Modify:

* packages/core/src/indexes/column-index-store.ts
* packages/core/src/engine/services/lookup-column-owner.ts
* packages/core/src/engine/services/runtime-column-store-service.ts
* direct lookup/criteria evaluator paths

TDD tests

* Exact lookup after single write updates posting list only.
* Approximate lookup after write updates local summaries only.
* Owner builds beyond 65,536 rows.
* Wildcard/regex fallback behavior is explicit and correct.

Expected performance movement

Lookup-after-write, approximate-after-write, text lookup, and criteria aggregation should turn green.

Rollback

If paged owner causes lookup correctness drift, retain old dense owner below 65,536 rows and use paged owner only for new benchmark-gated large spans until fixed.

Phase 5 — Demand-driven scheduler and calc-chain buckets

Goal

Stop scanning full calc chains for sparse dirty edits and support lazy value recomputation.

Files

Create:

* packages/core/src/scheduler/dirty-rank-buckets.ts
* packages/core/src/scheduler/formula-value-memo.ts

Modify:

* packages/core/src/scheduler.ts
* packages/core/src/scheduler/calc-chain.ts
* packages/core/src/scheduler/dynamic-topo.ts
* packages/core/src/engine/services/formula-graph-service.ts
* packages/core/src/engine/services/operation-service.ts
* packages/core/src/engine/services/evaluator-service.ts

TDD tests

* Sparse dirty edit in long chain orders only dirty formulas.
* Demand read computes before returning value.
* Event listener requiring changed payload forces evaluation only for observed/tracked cells.
* Volatile formulas still recompute as required.

Expected performance movement

Dirty-execution, batch-edit, and event/patch workloads improve. This also protects credible future interactive workloads.

Rollback

If lazy evaluation breaks public synchronous semantics, keep eager recalc default but retain dirty-rank buckets. Demand mode can be enabled only for no-listener or explicit config paths.

Phase 6 — Runtime image and typed patches

Goal

Make restore/rebuild and event emission owned by runtime artifacts, not replay/materialization.

Files

Create:

* packages/core/src/snapshot/runtime-image.ts
* packages/core/src/snapshot/runtime-image-codec.ts
* packages/core/src/events/typed-patches.ts

Modify:

* snapshot restore/export paths
* packages/core/src/engine/services/operation-service.ts
* headless runtime event plumbing

TDD tests

* Runtime image restore preserves values, formulas, families, regions, indexes, and topo ranks.
* Version mismatch falls back to safe replay.
* Structural insert emits typed axis patch plus invalidated spans.
* Browser/headless event consumers continue receiving compatible data.

Expected performance movement

Runtime restore and changed-payload workloads become green or near-green; structural events no longer regress UI paths.

Rollback

On codec mismatch or patch consumer breakage, fall back to replay for that snapshot version while preserving typed patch code behind compatibility checks.

⸻

Immediate first tranche with exact tests and commands

This first tranche should start immediately. It is intentionally narrow: make benchmark/counter truth universal and expose structural full scans. It should not be sold as the performance fix.

Files to change

1. packages/core/src/perf/engine-counters.ts

Add the first structural/scheduler/index counters:

'structuralTransactions',
'structuralPlannedCells',
'structuralSurvivorCellsRemapped',
'structuralRemovedCells',
'structuralUndoCapturedCells',
'structuralFormulaImpactCandidates',
'structuralFormulaRebindInputs',
'structuralRangeRetargets',
'sheetGridBlockScans',
'axisMapSplices',
'axisMapMoves',
'regionQueryIndexBuilds',
'columnOwnerBuilds',
'lookupOwnerBuilds',
'calcChainFullScans',
'topoRepairs',
'topoRepairFailures',
'topoRepairAffectedFormulas',

2. packages/core/src/__tests__/engine-counters.test.ts

Add assertions that all new keys exist, initialize to zero, clone, add, and reset correctly.

3. packages/benchmarks/src/benchmark-workpaper-vs-hyperformula-expanded-support.ts

Change its BenchmarkSample to include:

engineCounters?: Record<string, number>

Make measureWorkPaperBuildFromSheets() and measureMutationSample() mirror the counter-aware helpers in the main expanded benchmark file: reset counters before mutation, read counters after verification, then dispose. Current reference implementation is packages/benchmarks/src/benchmark-workpaper-vs-hyperformula-expanded.ts:905-967.

4. packages/benchmarks/src/__tests__/expanded-workloads.test.ts

Add a small direct test for additional WorkPaper helpers:

it('emits engine counters for additional WorkPaper workload helpers', () => {
  const samples = [
    measureWorkPaperStructuralInsertRowsSample(32),
    measureWorkPaperStructuralDeleteRowsSample(32),
    measureWorkPaperStructuralMoveRowsSample(32),
    measureWorkPaperStructuralInsertColumnsSample(32),
    measureWorkPaperStructuralDeleteColumnsSample(32),
    measureWorkPaperStructuralMoveColumnsSample(32),
  ]
  for (const sample of samples) {
    expect(sample.engineCounters).toBeDefined()
    for (const key of ENGINE_COUNTER_KEYS) {
      expect(sample.engineCounters![key]).toEqual(expect.any(Number))
    }
  }
})

Import the six helpers from benchmark-workpaper-vs-hyperformula-expanded-additional-workloads.ts and ENGINE_COUNTER_KEYS from core.

5. packages/core/src/sheet-grid.ts

Increment:

* sheetGridBlockScans for each block inspected in collectAxisRemapEntries().
* Keep existing behavior unchanged.

6. packages/core/src/workbook-store.ts

Increment:

* structuralTransactions in planStructuralAxisTransform().
* structuralPlannedCells by remappedEntries.length.
* structuralSurvivorCellsRemapped for entries with toRow/toCol.
* structuralRemovedCells for removed entries.
* Keep existing cellsRemapped for compatibility.

7. packages/core/src/engine/services/structure-service.ts

Increment:

* structuralUndoCapturedCells in captureAxisRangeCellState() by captured count.
* structuralFormulaImpactCandidates by candidate set size in collectStructuralFormulaImpacts().
* structuralFormulaRebindInputs by rebindInputs.length.

8. packages/core/src/range-registry.ts

Increment:

* structuralRangeRetargets in applyStructuralTransaction().

9. packages/core/src/deps/region-graph.ts

Increment:

* regionQueryIndexBuilds when ensureIntervalTree() rebuilds a dirty column.

10. packages/core/src/engine/services/runtime-column-store-service.ts

Increment:

* columnOwnerBuilds in buildColumnOwner().

11. packages/core/src/indexes/column-index-store.ts

Increment:

* lookupOwnerBuilds when buildLookupColumnOwner() is called.

12. packages/core/src/scheduler/calc-chain.ts

Add an optional counters dependency or callback. Increment calcChainFullScans when orderDirty() scans orderedChain for a sparse dirty set: packages/core/src/scheduler/calc-chain.ts:117-125.

If threading counters into CalcChain is too large for the first tranche, skip the increment but add the counter key now and wire it in Phase 5.

13. packages/core/src/engine/services/formula-graph-service.ts

Increment:

* topoRepairs
* topoRepairFailures
* topoRepairAffectedFormulas

around repairTopoRanksNow(): packages/core/src/engine/services/formula-graph-service.ts:162-175.

First-tranche commands

git status --short
corepack enable
pnpm install --frozen-lockfile
pnpm exec vitest run \
  packages/core/src/__tests__/engine-counters.test.ts \
  packages/benchmarks/src/__tests__/expanded-workloads.test.ts
pnpm exec vitest run \
  packages/core/src/__tests__/structural-transaction.test.ts \
  packages/core/src/__tests__/workbook-store.test.ts \
  packages/core/src/__tests__/structure-service.test.ts \
  packages/core/src/__tests__/mutation-service.test.ts
pnpm exec tsx packages/benchmarks/src/benchmark-workpaper-vs-hyperformula-expanded.ts --sample-count 2 --warmup-count 1
pnpm run ci

First-tranche acceptance

* All additional WorkPaper workload helpers emit engineCounters.
* Structural workloads show nonzero structural counters.
* No correctness tests regress.
* Competitive benchmark output can explain the six structural lanes by:
    * planned/remapped cells
    * undo captured cells
    * formula impact candidates
    * range retargets
    * topo repair/rebuild activity
* No benchmark family is allowed to regress by more than 5% median from the current baseline without a clear reason.

⸻

Risks and rollback criteria

Risk 1 — stale coordinate invariants

The biggest risk is breaking the relationship between logical IDs, visible coordinates, cellStore.rows/cols, SheetGrid, and cellKeyToIndex.

Mitigation

Before changing performance behavior, add invariant tests:

* every visible coordinate resolves to at most one cell
* every mapped cell has row/column identity
* getCellIndex(row,col) and getAddress(cellIndex) agree after structural edits
* snapshot/export round-trips after structural edits
* undo/redo round-trips after structural edits

Rollback

If stale coordinates appear, roll back only the latest structural migration step. Keep counter work.

Risk 2 — formula rewrite correctness

Relative and absolute references under insert/delete/move are easy to get subtly wrong.

Mitigation

Keep current formula rewrite path until formula-family retargeting has exhaustive tests. Use small-sheet differential tests against current behavior and HyperFormula where semantics match.

Rollback

Disable family retargeting for formula patterns that fail parity. Do not reintroduce broad full-sheet formula rebinding as the default.

Risk 3 — range graph false negatives

Region compression can create false positives safely, but false negatives are correctness bugs.

Mitigation

For each new region kind, test against expanded baseline on small sheets.

Rollback

If false negatives occur, revert that region kind. If false positives only, keep behind a performance gate while tightening.

Risk 4 — event payload compatibility

Typed patches and invalidation spans can break headless/browser consumers.

Mitigation

Emit compatibility payloads during one migration phase. Keep invalidatedRows/invalidatedColumns stable.

Rollback

Temporarily emit legacy changed-cell payloads for affected consumers, but keep typed patches in parallel only as a short-lived bridge.

Risk 5 — premature acceleration

WASM/GPU or more direct-descriptor heuristics can hide architecture debt.

Mitigation

Do not start acceleration until counters show:

* structural survivor remaps are zero
* range dependency discovery is symbolic
* sparse dirty edits avoid full calc-chain scans

Rollback

Remove acceleration code that improves one benchmark while counters still show old-owner full scans.

Risk 6 — multi-agent local changes

The executor must not overwrite unrelated local changes.

Mitigation

Start each tranche with:

git status --short
git diff --stat

Only edit listed files. If unrelated changes exist in a target file, inspect before patching and preserve them.

⸻

Validation strategy and definition of done

Validation ladder

1. Unit tests
    * Storage invariants.
    * Structural transaction semantics.
    * Formula family grouping.
    * Region query parity.
    * Column index update parity.
    * Scheduler/topo sparse dirty ordering.
2. Focused integration tests
    * Existing six focused structural/core test files.
    * Undo/redo structural tests.
    * Snapshot/export/restore structural tests.
    * Event patch tests.
3. Fuzz
    * Small-sheet structural fuzz against current semantics.
    * Formula dependency fuzz for relative/absolute references.
    * Undo/redo replay fuzz.
4. Microbenchmarks
    * Must show counter movement before macro scorecard claims.
    * Structural insert/move must show zero survivor remaps.
    * Sparse dirty must show zero full-chain scans.
5. Expanded competitive benchmarks
    * Run 2-sample smoke after every performance tranche.
    * Run 5-sample suite before committing a green checkpoint.
    * Save/report family summaries.
6. Full gate
    * Final: pnpm run ci.

Definition of done

This architecture program is done only when all of the following are true:

* WorkPaper wins all current eligible comparable workloads, or any remaining exception is documented as a semantic mismatch.
* Every eligible family has geomean speedup ≥ 2.0x over HyperFormula.
* Structural insert/move rows/columns have structuralSurvivorCellsRemapped = 0.
* Structural delete touches only deleted resident cells and required metadata/formula owners.
* Range dependency discovery does not expand large ranges.
* Repeated formula columns are represented by formula families/runs, not only per-cell graph nodes.
* Lookup owners are paged and work beyond 65,536 populated rows.
* Sparse dirty edits avoid full calc-chain scans.
* Runtime restore does not replay unchanged formulas/indexes as ordinary ops.
* Event emission uses typed patches/invalidation spans and does not flood shifted cells.
* pnpm run ci is green on the final committed tree.
* Benchmark artifacts include counters for every WorkPaper workload helper, including structural/additional workloads.
