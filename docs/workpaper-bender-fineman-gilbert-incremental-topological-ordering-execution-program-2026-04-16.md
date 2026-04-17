# WorkPaper Bender-Fineman-Gilbert Incremental Topological Ordering Execution Program

Date: `2026-04-16`

Status: `deferred until prerequisites exist`

Design document:

- `/Users/gregkonush/github.com/bilig2/docs/workpaper-bender-fineman-gilbert-incremental-topological-ordering-design-2026-04-16.md`

Primary source:

- `/Users/gregkonush/Downloads/bender-fineman-gilbert-incremental-topological-ordering.pdf`

## Why this exists

This paper is likely useful, but not first.

The point of this execution doc is to prevent premature work on a stronger topo model before the
practical prerequisites are present.

## Do not start until

All of the following must already be true:

1. the Pearce/Kelly execution program is implemented enough that local topo repair is already the
   common path
2. a family-compressed dependency graph exists for repeated formulas
3. benchmark evidence shows rank churn is still a measurable bottleneck after those two cuts

If those conditions are not true, this program should not start.

## Problem this program is supposed to solve

Only start this program if the current bottleneck is specifically:

- too much relabeling or rank churn under many edge insertions
- unstable ordering under repeated family split and merge operations
- snapshot rebuild or repeated family maintenance still paying too much order maintenance after the
  practical topo cut has landed

If the red lane is still structural ownership, do that first instead.

## Non-goals

- replacing the graph or scheduler before the simpler topo path lands
- building label math without a benchmark-driven reason
- using this as a substitute for family compression

## Main write set

Only if started:

- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/label-topo-order-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/scheduler.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-graph-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-family-dependency-service.ts`
- tests under `packages/core/src/__tests__/`

## Phase 0: Proof that we need this

### Goal

Prove rank churn is the remaining order bottleneck.

### Required measurements

Add metrics for:

- rank relabel count per graph mutation batch
- affected order span per family insertion or split
- topo maintenance time after family-compressed graph updates

If these metrics do not show a real order-maintenance bottleneck, stop.

## Phase 1: Label Store Prototype

### Goal

Add a label-based order store without changing scheduler behavior.

### Work

1. add `LabelTopoOrderService`
2. maintain labels beside current dense ranks
3. provide conversion:
   - label order -> stable comparison
   - label order -> projected dense ranks for compatibility

### Rule

- current dense-rank scheduler remains authoritative

## Phase 2: Family-Aware Label Allocation

### Goal

Use labels to absorb family insertions with less churn.

### Work

1. allocate label ranges per family block
2. support:
   - family split
   - family merge
   - family removal
3. fall back to compaction only when local label space is exhausted

### Tests

- repeated family insertions
- family splits caused by local edits
- family rebuild after replay or restore

## Phase 3: Scheduler Consumption

### Goal

Allow the scheduler to consume order from labels rather than assuming contiguous ranks.

### Work

1. either:
   - project dense ranks on demand
   - or update the scheduler to compare labels directly
2. keep dense-rank compatibility until all callers are migrated

## Phase 4: Cut Over

### Goal

Make label ordering authoritative where it shows a measurable win.

### Gate

Only do this if measured gains exist on still-red family-heavy lanes.

## Benchmark gates

Only meaningful if these are still red after earlier programs land:

- `partial-recompute-mixed-frontier`
- `rebuild-runtime-from-snapshot`
- family-heavy rebuild or trace microbenches

## Rollback Criteria

Rollback if:

- labels complicate family ownership without measurable wins
- scheduler complexity grows faster than the order-maintenance benefit
- replay and restore invariants become harder to prove

## What “done” looks like

This program is done only if label ordering measurably outperforms practical local repair in a
real remaining bottleneck.

If it never clears that bar, the correct outcome is to not land it.
