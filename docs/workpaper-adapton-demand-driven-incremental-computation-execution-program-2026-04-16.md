# WorkPaper Adapton Demand-Driven Incremental Computation Execution Program

Date: `2026-04-16`

Status: `proposed`

Design document:

- `/Users/gregkonush/github.com/bilig2/docs/workpaper-adapton-demand-driven-incremental-computation-design-2026-04-16.md`

Primary source:

- `/Users/gregkonush/Downloads/adapton-demand-driven-incremental-computation.pdf`

## Why this exists

The design doc narrows Adapton to observer-facing projections.

This execution program says how to implement that without contaminating the core formula engine.

## Problem to solve

The current engine eagerly maintains workbook truth, which is correct.

But some derived surfaces are observer-demanded and should not always pay eager recomputation:

- workbook agent dependency traces
- workbook explanation trees
- visible viewport or worker patch summaries
- selected derived summaries maintained above engine truth

## Non-goals

- turning formula evaluation into thunks
- changing workbook truth semantics
- pushing structural invalidation behind lazy observers

## Entry conditions

Before starting:

1. observer-facing surfaces must be enumerated
2. each target projection must have:
   - clear inputs
   - clear output shape
   - clear invalidation story

Do not start by “making reads lazy” globally.

## Main write set

- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/demanded-projection-service.ts`
- `/Users/gregkonush/github.com/bilig2/packages/headless/src/demanded-surface-cache.ts`
- `/Users/gregkonush/github.com/bilig2/packages/core/src/engine/services/read-service.ts`
- `/Users/gregkonush/github.com/bilig2/apps/bilig/src/codex-app/workbook-agent-comprehension.ts`
- `/Users/gregkonush/github.com/bilig2/apps/bilig/src/codex-app/workbook-agent-workflows.ts`
- `/Users/gregkonush/github.com/bilig2/packages/headless/src/work-paper-runtime.ts`

Tests:

- targeted tests for demanded projection invalidation and reuse
- agent flow tests
- headless runtime projection tests

## Phase 0: Surface Inventory

### Goal

Define the exact derived surfaces to be demand-driven.

### Required outputs

Document and implement ids for:

- `dependency-trace`
- `selection-explanation`
- `visible-sheet-window`
- `worker-visible-patch`

Each surface must declare:

- root input
- revision or freshness token
- invalidation condition
- output materialization path

## Phase 1: Demanded Projection Service

### Goal

Add the observer-demanded cache service without changing callers yet.

### Work

1. add `DemandedProjectionService`
2. support:
   - register observer
   - mark dirty
   - resolve projection on demand
   - evict projection when no observer remains

### Runtime rule

- the service caches derived outputs only
- it never mutates workbook truth

## Phase 2: Workbook Agent Integration

### Goal

Use demanded projections for workbook agent traces and explanations.

### Work

1. route dependency trace generation through `DemandedProjectionService`
2. route explanation generation through the same service
3. key caches by:
   - workbook revision
   - root selection
   - trace direction
   - depth or breadth options

### Benchmark gate

- agent trace and explanation latencies should improve or at least stop scaling with unrelated
  observer-invisible work

## Phase 3: Headless Visible Surface Integration

### Goal

Make visible projection caches observer-demanded rather than globally eager.

### Work

1. add `demanded-surface-cache.ts`
2. use it for:
   - visible sheet-window summaries
   - worker-visible patch derivation
3. mark dirty on engine mutation, but do not recompute until demanded

### Benchmark gate

- improve or protect:
  - `workerVisibleEdit10k`

## Phase 4: Projection Eviction and Reuse

### Goal

Make the system reclaim or reuse observer work safely.

### Work

1. evict projections with no observers
2. reuse demand-cached projections when the same observer shape comes back

This is the actual Adapton-like payoff.

## Required tests

- projection invalidates when roots change
- projection does not recompute when unrelated workbook regions change
- projection recomputes on next demand, not immediately
- agent traces stay semantically exact
- visible patch summaries stay exact

## Rollback Criteria

Rollback if:

- projection demand logic leaks into core workbook semantics
- observer caches return stale results
- visible patch latency improves only by hiding correctness work

## Validation Matrix

Per phase:

- relevant unit tests
- focused browser or headless tests where applicable

Program completion:

- `pnpm run ci`

## What “done” looks like

- observer-demanded derived surfaces recompute only when demanded
- workbook truth remains eagerly correct
- at least one observer-facing latency family improves
