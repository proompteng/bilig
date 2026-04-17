# WorkPaper Family-Compressed Dependency Graph Design

Date: `2026-04-16`

Status: `proposed`

Related documents:

- `/Users/gregkonush/github.com/bilig2/docs/workpaper-ultra-performance-engine-architecture-2026-04-12.md`
- `/Users/gregkonush/github.com/bilig2/docs/workpaper-hyperformula-targeted-reread-2026-04-13.md`
- `/Users/gregkonush/github.com/bilig2/docs/workpaper-performance-acceleration-plan.md`
- `/Users/gregkonush/Downloads/2302.05482v1.pdf`
- [Efficient and Compact Spreadsheet Formula Graphs](https://arxiv.org/abs/2302.05482)

## Purpose

This document defines the `bilig` integration plan for TACO-style compressed dependency ownership.
It is intentionally not a paper-faithful port.

The paper’s useful idea is narrow and important:

- repeated tabular formulas create repeated dependency edges
- those edges can be represented as compact patterns
- dependent and precedent queries can run directly on the compressed representation

That idea is real. It can improve `bilig`.

The paper’s full architecture is not the right direct transplant for this repo because:

- our current broad benchmark weakness is not mainly dependency-visualization latency
- full graph compression adds build overhead, and our build lanes are already red
- we already have partial template-family, range, lookup, and aggregate descriptors in-tree
- a second independent formula graph would duplicate ownership and make structural correctness worse

So the design target here is:

1. use TACO’s compression insight where it directly helps this engine
2. reuse the existing WorkPaper architecture instead of replacing it
3. improve dirty-frontier discovery, dependency traversal, and memory footprint without making
   build, structural maintenance, or correctness worse

## Decision

`bilig` should adopt a `family-compressed dependency graph` layer.

It should not adopt a full standalone TACO graph.

Specifically:

- compress repeated dependency edges only for formula families already detected by
  `FormulaTemplateNormalizationService`
- query these compressed families directly from traversal and scheduler code
- keep current cell/range/lookup reverse edges as fallback for irregular formulas and dynamic cases
- delete duplicated per-cell reverse edges only after the compressed family path is proven correct and
  faster on the real benchmark suite

This is the version of the design that is production-grade for this repo.

## Why The Paper Matters

The paper identifies a real spreadsheet property: tabular locality. Nearby formulas often have the
same relative dependency structure because users fill or copy formulas.

That overlaps directly with existing `bilig` hot spots:

- `build-parser-cache-row-templates`
- `build-parser-cache-mixed-templates`
- `partial-recompute-mixed-frontier`
- dependency trace and explain flows
- topo dirty-frontier collection
- snapshot rebuild for repeated formula regions

The paper does not directly solve the current largest misses:

- structural row or column transforms
- exact or approximate lookup after write
- public-path structural overhead

So this design is not a substitute for structural rearchitecture. It is a separate architecture cut
that should improve the family-compressible dependency portion of the engine.

## Current Repo Fit

The following pieces already exist and make this integration substantially easier than the paper’s
generic graph-compression setup:

### Existing template-family normalization

- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-template-normalization-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/formula/src/formula-template-key.ts`

We already compute relative template keys. That means we do not need the paper’s generic greedy
pattern search as the primary mechanism for repeated formulas.

### Existing compiled formula metadata

- `/Users/gregkonush/github.com/bilig2/packages/formula/src/compiler.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-binding-service.ts`

We already detect and materialize:

- parsed cell and range references
- direct aggregate candidates
- direct criteria candidates
- direct lookup candidates

That is enough to derive a dependency-shape key per formula family.

### Existing range and descriptor ownership

- `/Users/gregkonush/github.com/bilig2/packages/core/src/range-registry.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/range-aggregate-cache-service.ts`

The engine already accepts the idea that not all dependencies should be represented as literal
cell-to-cell edges.

### Existing traversal and scheduler choke points

- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/traversal-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-graph-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/scheduler.ts`

These are the exact places where compressed family ownership can replace repeated reverse-edge
walking.

## What We Should Not Do

1. Do not build a second fully separate dependency graph with its own truth model.
2. Do not run paper-style greedy graph compression across every edge in every workbook.
3. Do not make dynamic arrays, spills, tables, pivots, names, or structural rewrites depend on the
   compressed family path before the static repeated-family path is proven.
4. Do not optimize dependency visualization while leaving scheduler and recalc on the old hot path.
5. Do not accept build regressions in already-red build lanes just to shrink graph memory.

## Architecture Overview

Introduce a new engine subsystem:

- `FormulaFamilyDependencyService`

Proposed file:

- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-family-dependency-service.ts`

This service owns compressed dependency patterns for repeated formula families.

It does not replace:

- formula compilation
- direct aggregate descriptors
- direct criteria descriptors
- lookup descriptors
- structural transforms

It does replace repeated per-cell dependency ownership for the narrow class of formulas where:

- the formula template is stable
- the dependency shape is stable
- the owner cells form a contiguous tabular run
- the references can be represented as a small set of fixed or relative dependency patterns

## Core Model

### Formula Family

A formula family is a run of formulas that share:

- the same relative template key
- the same dependency-shape key
- the same owner axis orientation
- the same owner sheet
- contiguous ownership along one axis

The family should be a first-class runtime object, not just a build-time cache hit.

Proposed runtime shape:

```ts
interface FormulaDependencyFamily {
  readonly id: number;
  readonly sheetId: number;
  readonly axis: "row" | "col";
  readonly orthogonalIndex: number;
  ownerStart: number;
  ownerEnd: number;
  readonly templateKey: string;
  readonly shapeKey: string;
  readonly planId: number;
  readonly memberCellIndices: Uint32Array;
  readonly patternSlices: readonly FormulaDependencyPattern[];
}
```

### Dependency Pattern

Each family contains one or more compressed dependency patterns.

These correspond to the paper’s useful basic patterns:

- `rr`
- `rf`
- `fr`
- `ff`
- `rr-chain`

Proposed shape:

```ts
type FormulaDependencyPattern =
  | RelativeRelativePattern
  | RelativeFixedPattern
  | FixedRelativePattern
  | FixedFixedPattern
  | RelativeRelativeChainPattern;
```

Each pattern stores:

- precedent sheet
- owner span
- reference head/tail metadata
- optional chain direction
- optional range entity linkage where the precedent is already range-owned

### Dependency Shape Key

`templateKey` alone is not enough.

Two formulas may share the same rendered source pattern while having different dependency ownership
requirements once names, spills, direct lookups, criteria ranges, or non-cell references are
included.

So the family service also needs a `shapeKey` built from:

- parsed symbolic cell refs
- parsed symbolic ranges
- direct aggregate candidate metadata
- direct criteria candidate metadata
- direct lookup candidate metadata
- presence of names, tables, spills, or dynamic ranges

If the shape is not compressible, the family path is skipped.

## Compression Scope

### In Scope For Phase 1

- repeated single-sheet formula fills
- contiguous row-wise or column-wise family runs
- formulas whose references reduce to stable `rr/rf/fr/ff` cell or cell-range patterns
- chain formulas that match the paper’s `RR-Chain` special case

Examples:

- `=A1+B1` filled downward
- `=SUM($B$1:B1)` filled downward
- `=SUM(A1:B3)` copied across a stable row/column pattern
- `=A2+1` style predecessor chains

### Explicitly Out Of Scope For Phase 1

- formulas with sheet-local dynamic spills as primary family dependencies
- formulas whose dependency shape changes across the run
- structural transforms rewriting families in place
- cross-sheet families
- formulas dominated by names, tables, pivots, or metadata references

These can still use the existing per-cell path.

## Why This Is Better Than A Direct TACO Port

The paper compresses arbitrary formula graphs using a greedy candidate-edge search.

That is reasonable in a generic graph setting, but in this repo it would be the wrong first move.

We already know much more than the paper’s generic loader:

- formulas are compiled in cell order
- relative template keys already exist
- parsed dependency metadata already exists
- direct descriptor families already exist

That means we can build families in near-linear local passes instead of:

- building a full raw dependency graph
- searching adjacent candidate edges
- greedily recompressing them afterward

So the WorkPaper-specific design is:

- derive families during binding and rebuild
- intern them directly from existing normalization metadata
- avoid a second compression pass

This is a materially better fit for this engine.

## Runtime Integration

### Build And Bind Path

Primary integration point:

- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-binding-service.ts`

Add a family registration step after prepared dependency derivation but before final reverse-edge
installation.

New flow:

1. compile or translate formula as today
2. derive `templateKey`
3. derive `shapeKey`
4. ask `FormulaFamilyDependencyService` whether this cell extends the current family run
5. if yes, record family membership and compressed patterns
6. if not, fall back to ordinary per-cell dependency storage

Important rule:

- phase 1 keeps the current dependency arrays for correctness comparison and debugging
- phase 2 removes duplicated reverse-edge writes for family-covered formulas on the hot path

### Traversal And Dependency Trace

Primary integration points:

- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/traversal-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/read-service.ts`

Current state:

- traversal reads reverse edges and synthetic range/lookup entities
- dependency explanation walks direct precedents and dependents through these entities

Change:

- add a new dependency source queried alongside reverse edges:
  - `getCompressedFamilyDependents(entityId)`
  - `forEachCompressedFamilyDependencyCell(cellIndex, fn)`

This should allow:

- direct dependent expansion from a precedent cell or precedent range into a family member span
- direct precedent explanation for a family formula without expanding every repeated edge

### Scheduler And Dirty Frontier

Primary integration points:

- `/Users/gregkonush/github.com/bilig2/packages/core/src/scheduler.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-graph-service.ts`

Current state:

- scheduler BFS walks reverse edges and range/lookup entities
- topo rank rebuild iterates dependency cells per formula

Change:

- scheduler should query family-compressed dependents directly
- topo indegree rebuild should be able to count family-compressed formula precedents without
  materializing each repeated edge

This is the highest-value performance consumer of the new service.

### Snapshot Rebuild

Primary integration points:

- snapshot import and initial build paths under core and headless runtime

Change:

- when importing repeated formula families, build family descriptors directly from normalized order
- do not first expand every family into ordinary reverse edges and then compress them again

This is where the design can help `rebuild-runtime-from-snapshot` without taking the paper’s full
build-overhead penalty.

## Pattern Semantics

### `rr`

Owner cells move with stable relative head and tail offsets.

Best for:

- sliding windows
- same-shape fills
- local cell references

### `rf`

Relative head, fixed tail.

Best for:

- shrinking windows
- mixed fixed/relative cumulative formulas

### `fr`

Fixed head, relative tail.

Best for:

- running totals such as `SUM($B$1:B1)`

### `ff`

Fixed head and tail.

Best for:

- formulas that all point at one fixed cell or fixed range

### `rr-chain`

Special case of `rr` for direct predecessor/successor chains.

This is worth a first-class implementation because:

- the paper found it specifically addresses repeated edge access during BFS
- chainy column formulas exist in real spreadsheets
- our scheduler and dependency tracing code can suffer from the same repeated-hop behavior

## Structural Interaction

This design must not block the structural rearchitecture.

The family layer should integrate with structure work as follows:

- phase 1: structural edits invalidate affected families and fall back to ordinary rebinding
- phase 2: shape-stable structural rewrites can retarget family spans in place
- phase 3: family transforms and structural transforms share ownership rules

Important constraint:

- the family-compressed graph is not allowed to become the reason structural edits get slower or
  semantically riskier

So the default structural behavior is conservative invalidation until the transform model is ready.

## Correctness Rules

1. A family is only valid while every member preserves the same dependency shape.
2. Any local edit that changes template or dependency shape ejects that cell from the family.
3. Any operation that splits a contiguous run may split the family into multiple families.
4. Undo, redo, snapshot restore, and replica replay must rebuild the same family partition.
5. Direct aggregate, criteria, and lookup descriptors remain source-of-truth owners for those
   optimized paths.
6. A family-compressed dependency must produce the same direct dependents and direct precedents as
   the literal per-cell graph.
7. No compressed-family path may hide explicit empty-cell semantics or dependency placeholders.

## Proposed Service API

```ts
interface FormulaFamilyDependencyService {
  clear(): void;

  tryRegisterFormula(args: {
    cellIndex: number;
    sheetId: number;
    row: number;
    col: number;
    templateKey: string;
    shapeKey: string;
    compiled: CompiledFormula;
    directAggregate: RuntimeDirectAggregateDescriptor | undefined;
    directLookup: RuntimeDirectLookupDescriptor | undefined;
    rangeDependencies: Uint32Array;
    dependencyIndices: Uint32Array;
  }): FormulaFamilyMembership | undefined;

  unregisterFormula(cellIndex: number): void;

  collectDependentFormulas(entityId: number, out: Uint32Array): number;

  forEachDependencyCell(cellIndex: number, fn: (dependencyCellIndex: number) => void): void;

  invalidateSheet(sheetId: number): void;

  applyStructuralInvalidation(args: {
    sheetId: number;
    axis: "row" | "col";
    start: number;
    end: number;
  }): void;
}
```

Phase 1 should bias toward invalidation over in-place repair.

## Storage And Indexes

The service needs three cheap indexes:

1. `familyByCellIndex`
   - maps a formula cell to its family membership

2. `familiesByTemplateAndLane`
   - used to extend contiguous runs during build

3. `familiesByPrecedentLane`
   - used to answer dependent queries quickly from a changed cell or range

`lane` here means:

- sheet id
- axis orientation
- orthogonal row or column

This is enough for the common fill-down and fill-right cases without a generic spatial index.

## Performance Hypothesis

This design should materially improve:

- `partial-recompute-mixed-frontier`
- dependency tracing and workbook explain flows
- topo rebuild and dirty-frontier discovery on repeated formula regions
- snapshot rebuild for repeated template families
- memory use for large repeated formula blocks

This design may moderately improve:

- `build-parser-cache-row-templates`
- `build-parser-cache-mixed-templates`

This design is unlikely to materially improve by itself:

- structural row or column workloads
- exact lookup after write
- approximate sorted lookup after write
- direct aggregate sliding-window reuse

That is acceptable. This design is not pretending to be the whole performance program.

## Risks

### Build Overhead Risk

The paper shows compression can make build slower.

Mitigation:

- derive families opportunistically from the existing normalized build path
- do not run a second greedy compression pass
- measure build lanes after each phase

### Correctness Drift Risk

Compressed families can diverge from per-cell truth after edits, restore, or replay.

Mitigation:

- keep dual-path validation in phase 1
- add property tests comparing family-compressed traversal against literal traversal
- prefer family invalidation over speculative repair

### Structural Interaction Risk

Family spans and structural transforms can fight each other.

Mitigation:

- structural edits invalidate family ownership first
- no in-place structural family retargeting until structure rearchitecture is further along

### API Sprawl Risk

A sloppy family service could become a second graph engine.

Mitigation:

- keep it strictly as a dependency ownership source consumed by existing traversal and scheduler
- do not let it own formula evaluation, range caches, or lookup policy

## Rollout Plan

### Phase 1: Build The Family Service And Validate In Parallel

Files:

- add `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-family-dependency-service.ts`
- wire from `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/live.ts`
- register families from `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-binding-service.ts`

Behavior:

- family service builds compressed family metadata
- current reverse edges still remain active
- tests compare family-derived dependents and precedents against the literal graph

Success gate:

- no semantic regressions
- no meaningful build regressions on current benchmark suite

### Phase 2: Traversal And Read Path Consumption

Files:

- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/traversal-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/read-service.ts`

Behavior:

- traversal and dependency explanation consult family-compressed dependents
- repeated formula regions stop paying literal reverse-edge walks first

Success gate:

- direct dependency tracing remains exact
- dirty-frontier and trace workloads improve

### Phase 3: Scheduler And Topo Consumption

Files:

- `/Users/gregkonush/github.com/bilig2/packages/core/src/scheduler.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-graph-service.ts`

Behavior:

- scheduler and topo rebuild consume family-compressed dependents and dependency counts

Success gate:

- `partial-recompute-mixed-frontier` improves
- no cycle-detection or topo-order regressions

### Phase 4: Remove Duplicated Reverse Edges For Family-Covered Formulas

Behavior:

- family-covered formulas no longer install ordinary repeated reverse edges on the hot path

Success gate:

- memory footprint improves
- traversal and scheduler remain correct

### Phase 5: Snapshot And Rebuild Integration

Behavior:

- snapshot import and initial load create families directly

Success gate:

- `rebuild-runtime-from-snapshot` improves
- no rebuild correctness regressions

## Benchmark Gates

The feature is only successful if at least one of these materially improves without creating new
red regressions:

- `build-parser-cache-row-templates`
- `build-parser-cache-mixed-templates`
- `partial-recompute-mixed-frontier`
- `rebuild-runtime-from-snapshot`

Additional repo-local gates:

- dependency trace flows in workbook agent tools must remain exact
- `formula-graph-service` cycle and topo tests remain green
- undo, redo, snapshot restore, and replica fuzz stay green

## Test Plan

### Unit Tests

- pattern matching for `rr`, `rf`, `fr`, `ff`, and `rr-chain`
- family split and merge behavior after direct edits
- family invalidation after dependency-shape change
- direct dependent queries from precedent cells and ranges
- direct precedent enumeration from family members

### Property Tests

- for compressible generated formula grids:
  - family-compressed direct dependents equal literal direct dependents
  - family-compressed transitive dependents equal literal transitive dependents

### Correctness Regression Tests

- undo/redo through family creation and breakup
- snapshot export/import preserving visible workbook state
- replica replay preserving visible workbook state
- explicit blanks and dependency placeholders remain semantically identical

### Benchmark Tests

- current competitive suite
- focused family-heavy synthetic benchmarks
- dependency trace workload microbenchmarks

## Success Criteria

This design is successful when all of the following are true:

1. repeated formula families no longer require one reverse-edge expansion per member on the hot
   traversal path
2. scheduler dirty-frontier discovery is narrower and cheaper for family-heavy workbooks
3. dependency trace and explain flows remain exact
4. snapshot rebuild gets cheaper for repeated formula regions
5. build lanes do not regress enough to offset the wins
6. structural rearchitecture remains unblocked

## Bottom Line

The paper is worth integrating, but only as a `family-compressed dependency graph` layer that is
built on top of WorkPaper’s existing template normalization and descriptor architecture.

The right move is not:

- “implement TACO as a second graph”

The right move is:

- “use TACO’s compression model to stop representing repeated formula dependencies one cell at a
  time”

That is a real state-of-the-art improvement for this codebase, and it is the version of the design
that is worth building.
