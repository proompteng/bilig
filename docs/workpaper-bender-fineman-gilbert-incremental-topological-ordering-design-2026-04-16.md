# WorkPaper Bender-Fineman-Gilbert Incremental Topological Ordering Design

Date: `2026-04-16`

Status: `proposed`

Primary source:

- `/Users/gregkonush/Downloads/bender-fineman-gilbert-incremental-topological-ordering.pdf`
- [A New Approach to Incremental Topological Ordering](https://people.csail.mit.edu/jfineman/topsort.pdf)

## Purpose

This document defines the narrower `bilig` fit for the Bender/Fineman/Gilbert work.

Unlike Pearce/Kelly, this paper is not the first topo design to build here. It is the stronger
follow-on option once the engine has:

- better structural transform ownership
- family-compressed dependency ownership
- a clearer edge-diff story for graph edits

## Decision

`bilig` should treat this paper as the advanced topo-ordering phase, not the immediate one.

The part worth adopting is:

- label-based incremental topological ordering for repeated edge insertions on large changing DAGs

The part not worth adopting directly is:

- a theory-shaped generic implementation divorced from the current entity model

## Current Repo Fit

Primary target files:

- `/Users/gregkonush/github.com/bilig2/packages/core/src/scheduler.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-graph-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-family-dependency-service.ts`
  when that service lands

Why this is not phase 1:

- the current graph still rebuilds too much around structural edits
- the current family-compressed dependency graph is not landed yet
- the main immediate wins are practical sparse repairs, not sophisticated label systems

Why it will matter later:

- once repeated formulas compress into families, graph edits can become fewer but more structured
- a label-based order can maintain long-lived stable orderings across many insertions
- snapshot rebuild and repeated family maintenance may benefit more from stable labels than from
  repeated interval repairs

## Architecture Target

Introduce a second-stage topo subsystem:

- `LabelTopoOrderService`

This service should:

- maintain monotone labels rather than densely packed rebuilt ranks
- support many insertions before global compaction
- expose a total order view for the scheduler
- cooperate with cycle detection

This service should not exist until:

- `IncrementalTopoOrderService` or equivalent local repair is already in place
- family-compressed dependency ownership exists

## Why Labels Matter Here

The current `scheduler.ts` assumes:

- dense integer ranks
- counting sort over a contiguous rank span

That is good enough today, but it becomes brittle when:

- many repeated formulas share similar dependency regions
- families split and merge frequently
- we want to avoid rank churn from repeated insertions

Label-based order maintenance would allow:

- stable rank identities for large formula regions
- cheaper insertion without dense relabeling
- a cleaner fit for persistent family ownership

## Proposed Repo Shape

### Phase 1: Keep Dense Ranks Publicly

Even after label ordering lands internally, the rest of the engine can still consume a projected
dense order buffer when needed.

This avoids rewriting every caller at once.

### Phase 2: Scheduler Consumes Labels Directly

Later, `scheduler.ts` can stop assuming contiguous ranks and instead bucket or compare by labels.

### Phase 3: Family-Aware Ordering

Once `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/formula-family-dependency-service.ts`
exists, family insertions should allocate order labels in blocks rather than one cell at a time.

## What Not To Do

- do not build this before Pearce/Kelly style local repair
- do not rebuild the whole engine around generic label mathematics
- do not use this as an excuse to avoid fixing structural transform ownership first
- do not make family compression depend on this design landing first

## Benchmark Gates

This design is only worth landing if the earlier work already exists and one of these is still red:

- `partial-recompute-mixed-frontier`
- `rebuild-runtime-from-snapshot`
- repeated-family build or rebind lanes after family compression lands

## Bottom Line

This paper is likely worth using, but later.

It is the right design to reach for if the first practical topo improvements still leave too much
rank churn under compressed repeated-formula families. It is not the first topo paper to implement.
