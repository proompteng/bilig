# WorkPaper Family-Compressed Dependency Graph Execution Program

Date: `2026-04-16`

Status: `proposed`

Design document:

- `/Users/gregkonush/github.com/bilig2/docs/workpaper-family-compressed-dependency-graph-design-2026-04-16.md`

Primary source:

- `/Users/gregkonush/Downloads/2302.05482v1.pdf`

## Why this exists

The design doc explains why TACO-style family-compressed dependency ownership is the right
adaptation for this repo.

This document turns that design into an execution program that can be implemented without:

- creating a second formula graph
- regressing build lanes blindly
- destabilizing structural or replay correctness

## Problem to solve

The current engine still pays too much repeated dependency cost for repeated formulas.

The main red families this program is supposed to move are:

- `build-parser-cache-row-templates`
- `build-parser-cache-mixed-templates`
- `partial-recompute-mixed-frontier`
- `rebuild-runtime-from-snapshot`

Secondary wins are expected in:

- dependency trace and explanation flows
- memory footprint of repeated formula regions

## Non-goals

- replacing direct aggregate, criteria, or lookup descriptors
- fixing structural transforms directly
- adding a second standalone compressed graph
- greedily compressing every workbook edge post-hoc

## Entry conditions

Before phase 1 starts, the repo must already have:

- stable relative template keys in
  - `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-template-normalization-service.ts`
  - `/Users/gregkonush/github.com/bilig2/packages/formula/src/formula-template-key.ts`
- parsed dependency metadata in
  - `/Users/gregkonush/github.com/bilig2/packages/formula/src/compiler.ts`
  - `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-binding-service.ts`
- current traversal and scheduler choke points identified in
  - `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/traversal-service.ts`
  - `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-graph-service.ts`
  - `/Users/gregkonush/github.com/bilig2/packages/core/src/scheduler.ts`

Those conditions are already true on the current tree.

## Exit conditions

This program is complete only when all of the following are true:

1. repeated family-covered formulas no longer rely on one reverse-edge insertion per member on the
   common traversal hot path
2. dirty-frontier discovery for repeated family regions uses compressed family ownership
3. snapshot rebuild creates family ownership directly for repeated runs
4. the family path is covered by deterministic tests and fuzz or property checks
5. at least one primary benchmark family materially improves without new red regressions

## Main write set

Core runtime:

- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-family-dependency-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/live.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-binding-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/traversal-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-graph-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/scheduler.ts`

Formula metadata support:

- `/Users/gregkonush/github.com/bilig2/packages/formula/src/compiler.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-template-normalization-service.ts`

Tests:

- `/Users/gregkonush/github.com/bilig2/packages/core/src/__tests__/formula-family-dependency-service.test.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/__tests__/traversal-service.test.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/__tests__/formula-graph-service.test.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/__tests__/engine.test.ts`
- add a property or fuzz suite if the deterministic matrix is not enough

Benchmarks:

- add focused family-heavy dependency-trace and dirty-frontier microbench workloads if the current
  expanded suite is too coarse to isolate wins

## Phase 0: Measurement and Guard Rails

### Goal

Establish exact current costs and correctness expectations before changing ownership.

### Work

1. Add a narrow benchmark or diagnostic for:
   - repeated fill-down family dirty-frontier discovery
   - repeated-family direct dependent trace
   - repeated-family snapshot rebuild
2. Add a literal-vs-family equivalence harness:
   - given a generated repeated formula region
   - compare direct dependents
   - compare transitive dependents
   - compare dependency-cell enumeration

### Required tests

- deterministic generated family matrices
- mixed rows with non-family breakpoints
- formulas with:
  - `rr`
  - `rf`
  - `fr`
  - `ff`
  - `rr-chain`

### Stop criteria

- do not change runtime ownership until the equivalence harness exists

## Phase 1: Family Detection In Shadow Mode

### Goal

Build family metadata without letting runtime behavior depend on it yet.

### Work

1. Add `FormulaFamilyDependencyService`
2. Derive and store:
   - `templateKey`
   - `shapeKey`
   - owner axis and lane
   - member run boundaries
   - dependency patterns per family
3. Register families during formula bind in
   - `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-binding-service.ts`
4. Unregister or split families when:
   - formula is removed
   - dependency shape changes
   - direct descriptors make the shape ineligible

### Behavior rule

- existing reverse edges remain the source of truth
- family metadata is advisory only

### Required tests

- family formation for fill-down rows
- family formation for fill-right runs
- family split on one edited member
- family breakup on shape mismatch
- chain recognition for predecessor columns

### Benchmark gate

- no measurable regression on:
  - `build-parser-cache-row-templates`
  - `build-parser-cache-mixed-templates`

If shadow-mode registration already hurts those lanes materially, stop and simplify the metadata.

## Phase 2: Traversal Consumption

### Goal

Make dependency queries consume family-compressed dependents before touching literal repeated edges.

### Work

1. Extend traversal with:
   - `getCompressedFamilyDependents(entityId)`
   - `forEachCompressedFamilyDependencyCell(cellIndex, fn)`
2. Teach `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/traversal-service.ts`
   to merge:
   - literal reverse edges
   - family-compressed dependents
   - existing range and lookup synthetic entities
3. Teach `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/read-service.ts`
   to use family-compressed direct dependents in explanation and trace surfaces

### Required tests

- family-covered direct dependents equal literal direct dependents
- family-covered transitive dependents equal literal transitive dependents
- dependency explanation stays exact across family and non-family borders

### Benchmark gate

- improve at least one of:
  - repeated dependency-trace benchmark
  - `partial-recompute-mixed-frontier`

without regressing explanation correctness

## Phase 3: Scheduler and Topo Consumption

### Goal

Make the recalculation frontier consume family-compressed ownership.

### Work

1. Update `/Users/gregkonush/github.com/bilig2/packages/core/src/scheduler.ts`
   so family-compressed dependents are traversed directly
2. Update `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-graph-service.ts`
   so topo work counts family-owned dependents and dependencies correctly
3. Keep current dense rank model for now

### Required tests

- no change in dirty formula set for generated family cases
- no change in execution order for acyclic family-covered graphs
- cycle detection still finds the same members

### Benchmark gate

- `partial-recompute-mixed-frontier` must move in the right direction

If it does not, stop and inspect whether traversal still expands family-covered formulas too early.

## Phase 4: Reverse Edge Pruning For Family-Covered Formulas

### Goal

Remove duplicated per-member reverse-edge ownership from the hot path.

### Work

1. For family-covered formulas, stop installing ordinary repeated reverse edges where the family
   representation is exact
2. Keep fallback literal edges for:
   - non-family members
   - dynamic references
   - ranges and synthetic entities not represented by the family

### Required tests

- literal-vs-family equivalence suite still passes
- undo, redo, restore, and replay preserve family shape or safely fall back

### Benchmark gate

- no regression in:
  - `lookup-*`
  - direct aggregate lanes
  - structure-adjacent correctness suites

## Phase 5: Snapshot Rebuild Integration

### Goal

Create family ownership directly during rebuild and import.

### Work

1. update snapshot and initial load paths so repeated formula runs build families directly
2. avoid:
   - full literal reverse-edge expansion first
   - later family compression second

### Benchmark gate

- `rebuild-runtime-from-snapshot` must improve or at minimum not regress

## Rollback Criteria

Rollback the family-consumption phase if any of the following happen:

- structural undo or redo drift appears
- replay or replica parity breaks
- direct aggregate or lookup ownership becomes ambiguous
- build-template lanes regress significantly without runtime wins

## Validation Matrix

Required before merge of each major phase:

- `pnpm exec tsc -p packages/core/tsconfig.json --noEmit`
- focused Vitest on:
  - family service
  - traversal
  - graph service
  - engine regressions
- full `pnpm run ci` before the phase is declared complete

## What “done” looks like

This program is done when:

- repeated family regions are dependency-owned as compressed families
- traversal and scheduler consume them directly
- snapshot rebuild creates them directly
- at least one primary benchmark family flips materially
- the path is still semantically exact
