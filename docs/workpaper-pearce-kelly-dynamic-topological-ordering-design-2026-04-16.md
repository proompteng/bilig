# WorkPaper Pearce-Kelly Dynamic Topological Ordering Design

Date: `2026-04-16`

Status: `proposed`

Primary source:

- `/Users/gregkonush/Downloads/pearce-kelly-dynamic-topological-sort.pdf`
- [A Dynamic Topological Sort Algorithm for Directed Acyclic Graphs](https://whileydave.com/publications/pk07_jea/)

## Purpose

This document defines how the Pearce/Kelly dynamic topological sort work should influence
`bilig`.

The paper matters because the current engine still does broad topo and dependency work after graph
changes. The practical fit is stronger than the asymptotically fancier alternatives because our
graph is still sparse in the most important paths:

- formulas depend on a bounded number of direct cells or range entities
- direct aggregate, criteria, and lookup descriptors already collapse many edges
- the remaining dominant structural misses are caused by rebuild breadth, not by giant dense graphs

## Current Repo Fit

Primary target files:

- `/Users/gregkonush/github.com/bilig2/packages/core/src/scheduler.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-graph-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/traversal-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/structure-service.ts`

Current state:

- `RecalcScheduler.collectDirty()` walks outward from changed roots and then bucket-sorts by stored
  `topoRanks`
- `EngineFormulaGraphService.rebuildTopoRanksNow()` recomputes indegrees and rebuilds ranks from
  scratch
- structural operations still create too much dependency churn before the scheduler ever runs

This makes the following current red lanes worse:

- `structural-insert-rows`
- `structural-delete-rows`
- `structural-move-rows`
- `structural-insert-columns`
- `structural-delete-columns`
- `structural-move-columns`
- `partial-recompute-mixed-frontier`

## Decision

`bilig` should adopt Pearce/Kelly style local topo repair for dependency graph edits.

It should not replace the whole scheduler with a generic academic graph package.

The engine should:

- maintain topological order incrementally after edge insertions and deletions
- repair only the invalid interval around affected vertices
- continue using the current entity model:
  - cells
  - ranges
  - exact lookup columns
  - sorted lookup columns

## What To Build

Introduce a new service:

- `IncrementalTopoOrderService`

Proposed file:

- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/incremental-topo-order-service.ts`

Responsibilities:

- own rank maintenance after local graph edits
- expose:
  - `insertDependencyEdge`
  - `deleteDependencyEdge`
  - `insertDependencyEdges`
  - `deleteDependencyEdges`
  - `currentTopoRank`
  - `repairOrder`
- provide cycle-detection handoff when an inserted edge crosses the current order illegally

## Engine Integration

### Formula Bind And Clear

When formula dependencies change in:

- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-binding-service.ts`

the engine should emit edge-level mutations into `IncrementalTopoOrderService` rather than relying
on later full rank rebuilds.

### Structural Transform

When structural edits rewrite formulas in:

- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/structure-service.ts`

the structural service should batch the removed and inserted dependency edges and hand them to the
topo-order service.

This is important: Pearce/Kelly only pays off if structural edits stop hiding the graph change set.

### Scheduler

`/Users/gregkonush/github.com/bilig2/packages/core/src/scheduler.ts` should stop assuming that
`topoRanks` are rebuilt globally. It should trust incrementally maintained ranks and only fall back
to full rebuild on explicit invalidation.

## Data Model

The current `topoRanks[cellIndex]` storage can stay.

What changes is how it is maintained:

- today:
  - full recompute after broad graph change
- target:
  - local relabel of affected interval only

We do not need a separate graph data structure for this design.

We already have:

- reverse edges via edge slices
- direct dependency enumeration per formula
- a finite set of synthetic dependency entities

That is enough.

## Algorithm Shape

For an inserted dependency edge `u -> v`:

1. if `rank(u) < rank(v)`, accept without relabel
2. otherwise:
   - find the affected forward and backward search windows
   - detect whether a cycle was created
   - if acyclic, relabel only the invalid interval

For a deleted edge:

- do not eagerly compact global ranks
- mark the relevant region as eligible for local relaxation only when needed

This keeps the implementation practical and avoids over-optimizing deletions before the insertion
path is healthy.

## Why Pearce/Kelly First

The paper explicitly prioritizes practical performance on sparse graphs over the best asymptotic
bound. That matches this repo right now.

`bilig` does not yet need the more complicated label-ordering design as the first topo change.
It first needs a reliable local repair model that slots into the current graph and structural code.

## What Not To Copy

- do not build a fully generic standalone DAG layer
- do not represent the workbook as a separate topo-specific graph apart from existing runtime state
- do not optimize deletes before inserts
- do not make range and lookup descriptors lose their current ownership boundaries

## Benchmark Gates

This design is only successful if it materially improves at least one of:

- `partial-recompute-mixed-frontier`
- one or more structural row or column lanes

and does not regress:

- cycle detection
- dirty formula ordering
- batch and replay correctness

## Rollout

### Phase 1

- add `IncrementalTopoOrderService`
- wire formula bind and clear edge mutations into it
- retain full topo rebuild fallback behind an explicit safety flag

### Phase 2

- structural services emit batched edge diffs instead of only broad invalidation
- scheduler consumes incrementally maintained ranks

### Phase 3

- remove full rank rebuild from the common hot path
- keep it only for reset, snapshot import fallback, or severe invalidation

## Expected Outcome

This design should not single-handedly win the suite.

It should:

- reduce topo maintenance cost after local graph changes
- shrink structural fallout after dependency rewrites
- lower dirty-frontier overhead for mixed repeated formulas

It is a multiplier on the structural and family-compression work, not a replacement for them.
