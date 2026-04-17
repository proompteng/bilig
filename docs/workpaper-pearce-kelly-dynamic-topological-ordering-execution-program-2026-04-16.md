# WorkPaper Pearce-Kelly Dynamic Topological Ordering Execution Program

Date: `2026-04-16`

Status: `proposed`

Design document:

- `/Users/gregkonush/github.com/bilig2/docs/workpaper-pearce-kelly-dynamic-topological-ordering-design-2026-04-16.md`

Primary source:

- `/Users/gregkonush/Downloads/pearce-kelly-dynamic-topological-sort.pdf`

## Why this exists

The design doc explains why Pearce/Kelly is the right first dynamic-topo paper for this repo.

This execution program defines the implementation order that keeps the change practical and safe.

## Problem to solve

The engine still does too much broad topological work after graph changes.

The current code still relies on:

- full topo rank rebuilds in
  - `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-graph-service.ts`
- scheduler ordering over ranks rebuilt too broadly in
  - `/Users/gregkonush/github.com/bilig2/packages/core/src/scheduler.ts`

That amplifies:

- structural row or column edits
- mixed dependency graph rewrites
- dirty-frontier collection after graph changes

## Non-goals

- replacing the scheduler entirely
- building a generic graph library
- optimizing every delete case first
- introducing label-based order maintenance yet

## Entry conditions

Do not start this program unless:

1. structural services can produce or infer concrete edge-diff sets for changed formulas
2. existing cycle detection remains green
3. current topo rebuild tests are stable

The edge-diff story does not need to be perfect at phase 1, but it must be clear enough for:

- formula bind
- formula clear
- formula rebind after structural edits

## Exit conditions

This program is complete only when:

1. full topo rebuild is no longer the common hot path after ordinary graph edits
2. local edge insertions and removals repair order incrementally
3. scheduler consumes incrementally maintained order
4. cycle detection still agrees with the previous authoritative path
5. at least one structural or mixed-frontier lane improves materially

## Main write set

- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/incremental-topo-order-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-binding-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/structure-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-graph-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/scheduler.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/__tests__/incremental-topo-order-service.test.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/__tests__/formula-graph-service.test.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/__tests__/engine.test.ts`

## Phase 0: Baseline and Invariants

### Goal

Make the current graph order assumptions explicit.

### Work

1. add tests that define current invariants:
   - rank order respects all dependency edges
   - cycle creation blocks rank repair and reports the cycle
   - deleting an edge never introduces a rank violation
2. add a small benchmark or trace for:
   - edge insertion count per structural edit
   - full topo rebuild time after local formula rewrites

### Stop criteria

- do not implement the new service until the current invariants are checked in

## Phase 1: Add the Topo Service in Shadow Mode

### Goal

Compute an incremental topo order beside the current rebuilt order.

### Work

1. add `IncrementalTopoOrderService`
2. feed it graph mutations from:
   - formula bind
   - formula clear
3. compare its order against the rebuilt order in test-only validation

### Runtime rule

- the existing `rebuildTopoRanksNow()` path stays authoritative

### Tests

- sparse graph insertion sequences
- edge deletion sequences
- cycle-forming insertions
- graph edits across cell and range-entity boundaries

## Phase 2: Formula Bind And Clear Integration

### Goal

Make ordinary formula lifecycle mutations emit edge diffs.

### Work

1. in `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-binding-service.ts`
   emit:
   - removed dependency edges
   - inserted dependency edges
2. batch those updates into the topo service

### Tests

- bind same formula twice
- rewrite formula preserving bindings
- clear formula with direct aggregate or lookup descriptors nearby

## Phase 3: Structural Integration

### Goal

Make structural edits pay local topo repair instead of broad rebuild when possible.

### Work

1. update `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/structure-service.ts`
   to produce batched edge diffs for:
   - rewritten formulas
   - removed formulas
   - newly rebound formulas
2. apply those diffs to the topo service
3. keep a fallback to full rebuild for:
   - catastrophic invalidation
   - sheet deletes
   - unsupported transform cases

### Tests

- insert, delete, move rows
- insert, delete, move columns
- shape-stable structural rewrite cases
- structural undo and redo

### Benchmark gate

- at least one structural lane must improve or get measurably narrower in graph work

## Phase 4: Make Incremental Order Authoritative

### Goal

Switch the runtime to trust the incremental order path.

### Work

1. use incremental ranks or order projection as the source for scheduler ordering
2. demote full rebuild to:
   - initialization
   - snapshot import fallback
   - severe invalidation recovery

### Tests

- dirty formula ordering matches prior behavior on stable acyclic cases
- no cycle false negatives
- no order inversions under replay or restore

## Phase 5: Simplify The Old Rebuild Path

### Goal

Remove dead assumptions that topo rebuild is always expected.

### Work

1. strip unnecessary unconditional rebuild calls from common paths
2. keep one audited fallback path only

## Benchmark gates

Primary:

- `partial-recompute-mixed-frontier`
- at least one of:
  - `structural-insert-rows`
  - `structural-delete-rows`
  - `structural-move-rows`
  - `structural-insert-columns`
  - `structural-delete-columns`
  - `structural-move-columns`

Must not regress:

- cycle detection tests
- undo or redo correctness
- replica and snapshot replay

## Rollback Criteria

Rollback to full rebuild if:

- cycle detection disagrees with the rebuilt graph
- structural edits produce order violations
- sparse local edits become slower due to bookkeeping overhead

## Validation Matrix

Per phase:

- `pnpm exec tsc -p packages/core/tsconfig.json --noEmit`
- focused Vitest on topo, graph, structure, and engine slices

Program completion:

- full `pnpm run ci`

## What “done” looks like

- the common graph-mutation path uses local topo repair
- full rank rebuild is fallback-only
- structural edits no longer force broad topo work by default
- scheduler order remains correct
