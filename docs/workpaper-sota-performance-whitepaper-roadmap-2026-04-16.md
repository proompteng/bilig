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

The current benchmark reality in `packages/benchmarks/baselines/workpaper-vs-hyperformula.json`
is much narrower than the original paper-ranking context:

- generated at `2026-04-29T14:47:16.831Z`
- WorkPaper wins `44/46` scorecard-eligible comparable workloads
- public lane is `36/38`
- holdout lane is `8/8`
- current HyperFormula mean rows are `build-mixed-content` and
  `structural-delete-rows`, both with overlapping confidence intervals
- `lookup-text-exact` is the current p95 tail-risk row

So the papers only matter now if they move the remaining production rows without
regressing the current green holdout and public rows.

## Ranked Value

### Tier 1: Build Soon

1. `TACO`
   - doc: `/Users/gregkonush/github.com/bilig2/docs/workpaper-family-compressed-dependency-graph-design-2026-04-16.md`
   - why: useful as a guardrail for formula-family ownership, but no longer the
     first implementation step because parser-cache and holdout build rows are
     green
   - target families:
     - preserve `build-parser-cache-unique-formulas`
     - reduce `build-mixed-content` cold-build allocation without changing
       formula semantics

2. `Pearce/Kelly`
   - doc: `/Users/gregkonush/github.com/bilig2/docs/workpaper-pearce-kelly-dynamic-topological-ordering-design-2026-04-16.md`
   - why: still relevant for structural row-delete tail overhead and dependency
     update locality
   - target families:
     - `structural-delete-rows`
     - structural preservation checks for insert/move/delete rows and columns
     - dirty-ordering only where profiling shows it is still on the hot path

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

1. finish `build-mixed-content` cold-build allocation and initialization
   hardening
2. finish `structural-delete-rows` metadata/result-collection hardening
3. harden `lookup-text-exact` p95 tail latency
4. apply Pearce/Kelly-style dynamic topo repair only if profiling proves topo
   order maintenance is still the row-delete bottleneck
5. build TACO-style family-compressed dependency ownership only where it reduces
   real mixed-build allocation without regressing current build holdouts
6. if the family graph is still too rebuild-heavy, add Bender/Fineman/Gilbert
   label ordering
7. add Adapton-style demand control for visible and agent-facing derived surfaces
8. use Differential Dataflow ideas only for sync and worker event projections

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
