# WorkPaper SOTA Performance Whitepaper Roadmap

Date: `2026-04-16`

Status: `planning docs aligned to current source state`

Related documents:

- `/Users/gregkonush/github.com/bilig2/docs/workpaper-ultra-performance-engine-architecture-2026-04-12.md`
- `/Users/gregkonush/github.com/bilig2/docs/workpaper-family-compressed-dependency-graph-design-2026-04-16.md`
- `/Users/gregkonush/github.com/bilig2/docs/workpaper-pearce-kelly-dynamic-topological-ordering-design-2026-04-16.md`
- `/Users/gregkonush/github.com/bilig2/docs/workpaper-bender-fineman-gilbert-incremental-topological-ordering-design-2026-04-16.md`
- `/Users/gregkonush/github.com/bilig2/docs/workpaper-adapton-demand-driven-incremental-computation-design-2026-04-16.md`
- `/Users/gregkonush/github.com/bilig2/docs/workpaper-differential-dataflow-design-2026-04-16.md`

Primary sources:

- `/Users/gregkonush/Downloads/2302.05482v1.pdf`
- `/Users/gregkonush/Downloads/pearce-kelly-dynamic-topological-sort.pdf`
- `/Users/gregkonush/Downloads/bender-fineman-gilbert-incremental-topological-ordering.pdf`
- `/Users/gregkonush/Downloads/adapton-demand-driven-incremental-computation.pdf`
- `/Users/gregkonush/Downloads/mcsherry-differential-dataflow-cidr2013.pdf`

## Purpose

This document ranks the downloaded whitepapers by expected value for the current `bilig` engine.
It is not a literature review. It is the decision document for what to build, what to defer, and
what to reject.

The benchmark reality in `/Users/gregkonush/github.com/bilig2/docs/workpaper-ultra-performance-engine-architecture-2026-04-12.md`
still says the dominant misses are:

- structural row and column transforms
- parser or template build and snapshot rebuild
- dirty-frontier discovery over repeated formulas
- exact and approximate lookup freshness after writes
- sliding-window aggregate reuse
- batch plus undo scaffolding

So the papers only matter if they move those families.

## Ranked Value

### Tier 1: Build Soon

1. `TACO`
   - doc: `/Users/gregkonush/github.com/bilig2/docs/workpaper-family-compressed-dependency-graph-design-2026-04-16.md`
   - why: best direct match for repeated formula dependency ownership and dirty-frontier discovery
   - target families:
     - `build-parser-cache-row-templates`
     - `build-parser-cache-mixed-templates`
     - `partial-recompute-mixed-frontier`
     - `rebuild-runtime-from-snapshot`

2. `Pearce/Kelly`
   - doc: `/Users/gregkonush/github.com/bilig2/docs/workpaper-pearce-kelly-dynamic-topological-ordering-design-2026-04-16.md`
   - why: best practical fit for dynamic topo repair on a sparse changing dependency graph
   - target families:
     - structural edits
     - dirty-frontier ordering after graph changes
     - broad topo rebuild cost

### Tier 2: Build After Tier 1

3. `Bender/Fineman/Gilbert`
   - doc: `/Users/gregkonush/github.com/bilig2/docs/workpaper-bender-fineman-gilbert-incremental-topological-ordering-design-2026-04-16.md`
   - why: stronger label-based incremental ordering model, but more invasive than Pearce/Kelly
   - target families:
     - high-churn family-compressed graph maintenance
     - persistent order maintenance after many edge insertions

4. `Adapton`
   - doc: `/Users/gregkonush/github.com/bilig2/docs/workpaper-adapton-demand-driven-incremental-computation-design-2026-04-16.md`
   - why: strong fit for observer-facing recomputation and visible-surface demand control
   - target families:
     - viewport-only work
     - workbook agent traces and explanation flows
     - visible worker patch and commit paths

### Tier 3: Selective, Not Core Engine First

5. `Differential Dataflow`
   - doc: `/Users/gregkonush/github.com/bilig2/docs/workpaper-differential-dataflow-design-2026-04-16.md`
   - why: powerful ideas for versioned delta maintenance, but too large for a direct engine rewrite
   - target families:
     - sync replay
     - reconnect catch-up
     - materialized projection maintenance on event streams

## Recommended Execution Order

1. finish structural transform ownership already in progress
2. build TACO-style family-compressed dependency ownership
3. add Pearce/Kelly-style dynamic topo repair
4. if the family graph is still too rebuild-heavy, add Bender/Fineman/Gilbert label ordering
5. add Adapton-style demand control for visible and agent-facing derived surfaces
6. use Differential Dataflow ideas only for sync and worker event projections

## What Not To Do

- do not attempt a full TACO port with a second standalone graph
- do not rewrite the whole engine into Adapton
- do not rewrite the whole engine into Differential Dataflow
- do not add theory-heavy topo machinery before finishing the current structural transform service
- do not let any paper-derived subsystem fight the current direct aggregate, criteria, or lookup
  descriptors

## SOTA Target

The realistic SOTA stack for this repo is:

- HyperFormula-style transform ownership for structural edits
- TACO-style family-compressed dependency ownership for repeated formulas
- Pearce/Kelly or Bender-style incremental topo maintenance
- Excel-style calc-chain reuse from the existing prior-art notes
- Adapton-style demand control only where observers decide what must stay hot
- Differential Dataflow-style versioned delta maintenance only for sync projections

That combination is the highest-value path to leadership from the current codebase.
