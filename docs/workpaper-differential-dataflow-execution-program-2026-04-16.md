# WorkPaper Differential Dataflow Execution Program

Date: `2026-04-16`

Status: `proposed`

Design document:

- `/Users/gregkonush/github.com/bilig2/docs/workpaper-differential-dataflow-design-2026-04-16.md`

Primary source:

- `/Users/gregkonush/Downloads/mcsherry-differential-dataflow-cidr2013.pdf`

## Why this exists

The design doc limits Differential Dataflow ideas to sync and projection maintenance.

This execution program makes that concrete and prevents accidental engine-wide scope creep.

## Problem to solve

The sync and worker side still has a real versioned delta problem:

- authoritative batches arrive over time
- local pending mutations must be rebased
- reconnect catch-up should reuse prior work instead of replaying broad state

The clearest fitting benchmark is:

- `workerReconnectCatchUp100Pending`

## Non-goals

- using dataflow operators for formula calculation
- introducing partially ordered time into workbook truth
- replacing zero-sync event semantics

## Entry conditions

Do not start until:

1. current reconnect and replay semantics are deterministic
2. the target projections are identified
3. we can state exactly which projection is too expensive today

## Main write set

- `/Users/gregkonush/github.com/bilig2/packages/zero-sync/src/revisioned-projection-arrangements.ts`
- `/Users/gregkonush/github.com/bilig2/packages/headless/src/revisioned-catch-up-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/zero-sync/src/workbook-events.ts`
- `/Users/gregkonush/github.com/bilig2/packages/zero-sync/src/mutators.ts`
- `/Users/gregkonush/github.com/bilig2/packages/headless/src/work-paper-runtime.ts`

Tests:

- sync replay tests
- reconnect tests
- projection arrangement tests

## Phase 0: Pick One Projection

### Goal

Do not start with a generalized arrangement system for everything.

Pick one projection first, ideally:

- reconnect-visible changed cells
or
- reconnect dirty regions

### Required output

One explicit target projection with:

- key
- source event fields
- expected output
- current hot path

## Phase 1: Arrangement Storage

### Goal

Store revisioned deltas in a reusable indexed form.

### Work

1. add `RevisionedProjectionArrangementService`
2. store deltas by:
   - document
   - revision
   - sheet
   - projection key
3. answer:
   - projection at revision
   - delta between revisions

### Rule

- this service stores projection deltas, not workbook truth

## Phase 2: Authoritative Replay Consumption

### Goal

Use arrangements for authoritative replay before mixing local pending changes.

### Work

1. build the target projection from arrangement deltas rather than broad replay
2. verify exact agreement with the old replay path

## Phase 3: Merge Local Pending Mutations

### Goal

Apply the differential idea to:

- authoritative revisions
- local speculative revisions

### Work

1. define a merge model for:
   - authoritative deltas
   - pending local deltas
2. materialize the target projection from both

### Safety rule

- if merge ambiguity exists, fall back to the old path

## Phase 4: Reconnect Catch-Up Integration

### Goal

Use the arrangement path during reconnect.

### Work

1. route reconnect catch-up through `revisioned-catch-up-service.ts`
2. reuse indexed deltas where possible
3. keep the old path as a fallback while validating

## Required tests

- projection at revision `r` equals old replay result
- delta from `r1` to `r2` equals replay difference
- merge of authoritative and local pending changes is deterministic
- reconnect leaves no stale projection state behind

## Benchmark gate

Primary:

- `workerReconnectCatchUp100Pending`

Secondary:

- any targeted reconnect projection microbench added in phase 0

## Rollback Criteria

Rollback if:

- replay determinism weakens
- merge semantics become harder to reason about than the saved time is worth
- the arrangement layer starts to duplicate workbook truth

## Validation Matrix

Per phase:

- focused sync and headless tests
- replay parity tests

Program completion:

- full `pnpm run ci`

## What “done” looks like

- one reconnect or replay projection is arrangement-backed
- it is exact
- it materially improves reconnect catch-up or replay latency
- the rest of the engine remains untouched
