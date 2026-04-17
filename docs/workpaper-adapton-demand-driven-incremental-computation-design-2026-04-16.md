# WorkPaper Adapton Demand-Driven Incremental Computation Design

Date: `2026-04-16`

Status: `proposed`

Primary source:

- `/Users/gregkonush/Downloads/adapton-demand-driven-incremental-computation.pdf`
- [Adapton: Composable, Demand-Driven Incremental Computation](https://drum.lib.umd.edu/items/b7401a53-964b-4f78-8b2d-a93ab750bb0a)

## Purpose

This document defines how demand-driven incremental computation should influence `bilig`.

This is not a proposal to rewrite the spreadsheet engine into Adapton.

It is a proposal to use the demand-driven idea where the current codebase actually has observers:

- visible viewport surfaces
- workbook agent traces and explanations
- worker-visible patch generation
- headless or UI caches that do not need full-engine freshness on every mutation

## Decision

Adopt Adapton-style demand-driven invalidation only for observer-facing derived surfaces.

Do not adopt Adapton as the core workbook semantics engine.

The semantic source of truth for workbook calculation should stay the existing engine, not a thunked
general incremental runtime.

## Current Repo Fit

Primary target files:

- `/Users/gregkonush/github.com/bilig2/packages/headless/src/work-paper-runtime.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/read-service.ts`
- `/Users/gregkonush/github.com/bilig2/apps/bilig/src/codex-app/workbook-agent-comprehension.ts`
- `/Users/gregkonush/github.com/bilig2/apps/bilig/src/codex-app/workbook-agent-workflows.ts`

Current symptoms that fit demand control:

- visible and agent-facing work can still over-read or overcompute relative to what is being shown
- the perf contract lanes include observer-facing costs such as worker visible edit and reconnect
  catch-up
- workbook explain and dependency trace work is observer-demanded by nature

## What To Build

Introduce a narrow observer cache layer:

- `DemandedProjectionService`

Proposed files:

- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/demanded-projection-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/headless/src/demanded-surface-cache.ts`

Responsibilities:

- track observer registrations for:
  - visible viewport ranges
  - active workbook agent trace roots
  - selected explicit derived summaries
- hold invalidation state lazily
- recompute only when an observer requests the projection again

## Core Rule

The engine still eagerly maintains workbook truth.

Demand-driven behavior applies only to derived projections layered on top of that truth.

Examples:

- a workbook dependency explanation tree
- a visible patch summary for a worker
- a lazily maintained sheet summary or trace visualization

Non-examples:

- formula result values
- range cache correctness
- lookup freshness
- structural rewrite semantics

## Why This Matters

The Adapton paper is strongest where:

- not every output is always demanded
- different observers request different projections
- preserving previously computed observer results matters

That maps well to:

- workbook agent features
- visible viewport or patch summaries
- reconnect and replay projections

It maps poorly to:

- mandatory recalculation of true spreadsheet values

## Observer Model

Each observer should declare:

- projection kind
- root entity or range
- freshness token
- invalidation status

Example projection kinds:

- `dependency-trace`
- `selection-explanation`
- `visible-sheet-window`
- `worker-visible-patch`

When a mutation occurs:

- mark overlapping projections dirty
- do not recompute immediately
- recompute on next observer demand

## Integration Points

### Workbook Agent

The workbook agent code currently builds traces and explanations directly from engine reads.

Those should become cached demanded projections keyed by:

- workbook revision
- root selection
- direction
- depth budget

### Headless Visible Surfaces

The headless runtime currently carries visibility and named-expression caches. This is the best
place to move from eager refresh to observer-demanded refresh.

### Worker Patch Paths

The `workerVisibleEdit10k` contract suggests observer-specific work still matters. A demanded
surface cache can let visible patch generation stay narrow without changing core recalculation.

## What Not To Do

- do not make core formula recalculation demand-driven
- do not suspend correctness-required workbook updates behind observers
- do not route structural invalidation through lazy observer logic
- do not introduce thunks into the core hot numeric path

## Benchmark Gates

This design is worthwhile if it improves:

- worker visible edit or patch latency
- workbook agent trace latency
- visible-surface projection costs

without worsening:

- full workbook mutation correctness
- replay and undo or redo semantics

## Bottom Line

Adapton is worth using here, but only as an observer-demanded projection architecture.

That is the part of the paper that fits the current repo and can move real user-facing latency
without destabilizing the engine core.
