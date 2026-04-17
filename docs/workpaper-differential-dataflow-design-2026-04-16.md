# WorkPaper Differential Dataflow Design

Date: `2026-04-16`

Status: `proposed`

Primary source:

- `/Users/gregkonush/Downloads/mcsherry-differential-dataflow-cidr2013.pdf`
- [Differential dataflow (CIDR 2013)](https://www.cidrdb.org/cidr2013/Papers/CIDR13_Paper111.pdf)

## Purpose

This document defines the narrow `bilig` use for Differential Dataflow ideas.

This is not a proposal to turn the spreadsheet engine into Naiad.

It is a proposal to use the paper’s strongest ideas where `bilig` already has:

- versioned event streams
- reconnect catch-up
- pending local mutations
- repeatedly materialized projections over ordered revisions

## Decision

Use Differential Dataflow ideas for sync and projection maintenance only.

Do not use Differential Dataflow as the primary formula evaluation architecture.

## Current Repo Fit

Primary target files:

- `/Users/gregkonush/github.com/bilig2/packages/zero-sync/src/workbook-events.ts`
- `/Users/gregkonush/github.com/bilig2/packages/zero-sync/src/mutators.ts`
- `/Users/gregkonush/github.com/bilig2/packages/headless/src/work-paper-runtime.ts`

Current fitting benchmark family:

- `workerReconnectCatchUp100Pending`

Current fitting product behavior:

- replaying authoritative event batches
- rebasing local pending changes
- maintaining projections over revisioned workbook event streams

## The Useful Idea

The paper’s useful idea is not “dataflow everywhere.”

It is:

- keep indexed deltas, not just collapsed current state
- let different logical times combine prior work differently
- reuse prior partial results across both input changes and iterative progress

That is directly useful for:

- reconnect catch-up
- stale-draft reconciliation
- event-derived projections such as dirty regions, visible patches, and lightweight summaries

## What To Build

Introduce a sync-side arrangement layer:

- `RevisionedProjectionArrangementService`

Proposed files:

- `/Users/gregkonush/github.com/bilig2/packages/zero-sync/src/revisioned-projection-arrangements.ts`
- `/Users/gregkonush/github.com/bilig2/packages/headless/src/revisioned-catch-up-service.ts`

Responsibilities:

- maintain indexed deltas by revision
- answer:
  - projection at revision `r`
  - projection delta from revision `r1` to `r2`
  - projection delta from authoritative head plus local pending queue

Candidate projections:

- changed dirty regions by sheet
- changed visible cells by viewport
- compact workbook event bundles for reconnect replay

## Why This Fits Better Than A Full Port

The formula engine already has a strong single-threaded semantic source of truth.

Differential Dataflow would be overkill and destabilizing as a replacement.

But the sync layer already behaves like:

- a versioned stream of event batches
- local speculative batches plus authoritative batches
- repeated projection maintenance over revisions

That is much closer to the paper’s natural habitat.

## Concrete Design

### Arrangement Keys

Key arrangements by:

- workbook id
- revision
- sheet id
- cell key or dirty region key

### Delta Storage

Store deltas as:

- inserted cell changes
- cleared cell changes
- structural range mutations
- metadata mutations

Then derive compact projections from those deltas instead of replaying whole snapshots.

### Catch-Up

For reconnect:

1. find authoritative revisions missing locally
2. pull arranged deltas for those revisions
3. merge with locally pending deltas
4. materialize only the projections needed by the current observer

That is the Differential Dataflow part worth stealing.

## What Not To Do

- do not express formula evaluation as dataflow operators
- do not introduce partially ordered time into the core workbook semantics
- do not push direct aggregate or lookup services behind a general dataflow runtime

## Benchmark Gates

This design is only worth landing if it improves:

- `workerReconnectCatchUp100Pending`
- visible patch replay or reconciliation latency

without regressing:

- local mutation latency
- replay correctness
- sync determinism

## Bottom Line

Differential Dataflow is a strong paper for this repo, but only on the sync and projection side.

Its best reusable idea here is revisioned delta arrangements for reconnect and replay, not a
replacement for the spreadsheet engine core.
