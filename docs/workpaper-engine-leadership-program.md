# WorkPaper Engine Leadership Program

This document captures the actual engine work required for `bilig` to beat HyperFormula in ways
that matter, rather than only matching its public headless surface.

Date: `2026-04-10`

## Status

This document is an execution program, not a public-contract spec.

It complements, and does not replace:

- `docs/workpaper-platform-design.md`

`docs/workpaper-platform-design.md` defines the `WorkPaper` API contract, parity gates, external
consumer guarantees, and evidence policy.

This document defines the engine, formula-runtime, and benchmark work required to make
`WorkPaper` competitively superior in reality.

## Problem Statement

`WorkPaper` is now a real, publishable, measured headless package. That is good, but it is not the
same thing as engine leadership.

Current state:

- `WorkPaper` is the canonical top-level API in `@bilig/headless`
- HyperFormula public surface parity by method/config inventory is checked
- external install and consumer smoke paths are checked
- a checked-in comparison artifact exists at
  `packages/benchmarks/baselines/workpaper-vs-hyperformula.json`
- that artifact currently shows HyperFormula faster on the directly comparable microbenchmarks run
  on this host

So the remaining work is not:

- more API renaming
- more parity counting
- more marketing language

The remaining work is:

- actual engine optimization
- actual formula-runtime optimization
- actual semantics expansion where `bilig` can surpass HyperFormula
- actual workload-specific proof that the improvements are real

## Source Corpus

The program is based on:

- local HyperFormula checkout: `/Users/gregkonush/github.com/hyperformula`
- HyperFormula performance and limitation docs reviewed in
  `docs/workpaper-platform-design.md`
- current `bilig` runtime layers:
  - `packages/core/src`
  - `packages/formula/src`
  - `packages/wasm-kernel/src`
  - `packages/headless/src`
- checked-in benchmark evidence:
  - `packages/benchmarks/baselines/workpaper-baseline.json`
  - `packages/benchmarks/baselines/workpaper-vs-hyperformula.json`

## What “Beat HyperFormula” Means

This program does not allow vague claims.

To say that `bilig` beats HyperFormula, at least one of the following must be true:

1. `bilig` supports materially important spreadsheet semantics that HyperFormula does not support,
   with production-grade tests and documentation.
2. `bilig` is measurably faster on named comparable workloads, with checked-in benchmark evidence.
3. `bilig` offers stronger production behavior:
   - better packaging
   - stronger determinism
   - stronger diagnostics
   - better integration surface

To say that `bilig` beats HyperFormula overall, all three must be true in a defensible way.

## Current Reality

The current repository supports these claims:

- `bilig` already leads in some capability areas:
  - dynamic arrays
  - structured references and table-aware semantics
  - multiple independent workbook instances per process
- `bilig` already leads in some product/package areas:
  - MIT runtime packaging
  - verified external consumer install path
  - richer detailed event payloads

The current repository does not support these claims:

- that `bilig` is faster than HyperFormula in general
- that `bilig` is `10x` better overall
- that all meaningful engine gaps are closed

That means the engineering program must focus on real runtime work, not just surface-level parity.

## Non-Goals

This program explicitly does not call for:

- building a second spreadsheet engine
- cloning HyperFormula internals line by line
- destabilizing the `WorkPaper` public contract without necessity
- adding speculative features with no semantics design
- claiming performance wins before benchmarks prove them

## Required Engine Workstreams

### 1. Build and Load Path

Goal:

- make workbook construction materially faster and cheaper

Likely bottlenecks:

- formula parse cost during workbook build
- compilation churn for repeated formulas
- snapshot hydration overhead
- string allocation and duplication
- eager materialization of cells and ranges during setup

Required work:

- introduce formula parse/compile cache keyed by canonical source plus config
- identify repeated-formula patterns and avoid recompiling identical AST/program pairs
- audit workbook restore path for avoidable object churn
- reduce serialization overhead in headless build-from-sheets and restore paths
- add benchmark fixtures for:
  - dense literals
  - repeated formulas
  - mixed literals + formulas
  - large named-sheet workbook restore

Exit criteria:

- `workpaper-vs-hyperformula.json` includes improved build-path results
- no correctness regression in headless parity tests

### 2. Single-Edit Recalc Hot Path

Goal:

- reduce the cost of a single cell edit with downstream recalculation

Likely bottlenecks:

- dirty-node propagation
- dependency traversal fanout
- formula execution overhead
- change-array construction and serialization
- repeated address translation and allocation

Required work:

- profile edit-to-recalc pipeline in `@bilig/core`
- shrink per-node object churn
- avoid recomputing unchanged metadata in returned change arrays
- separate internal dirty traversal cost from external serialization cost in stats
- add a benchmark fixture for long linear dependency chains and medium fanout graphs

Exit criteria:

- lower mean and p95 on named edit workloads
- stats expose enough detail to attribute wins

### 3. Batch Mutation Efficiency

Goal:

- make batched edits scale with the affected graph, not with avoidable per-operation overhead

Likely bottlenecks:

- repeated write-path validation
- repeated invalidation work inside suspended evaluation mode
- overly expensive combined change-array materialization

Required work:

- audit batch path for per-edit repeated work that should be deferred
- aggregate invalidation metadata during `batch()` / suspended evaluation
- deduplicate structural change reporting where the public result permits it
- add a benchmark fixture with many edits to formula-backed rows

Exit criteria:

- batched-edit slope improves materially with edit count
- benchmark artifact records the result

### 4. Lookup and Indexing Workloads

Goal:

- close the current performance gap on lookup-heavy workloads

This is currently one of the clearest directly comparable areas where HyperFormula wins.

Required work:

- profile `MATCH` / lookup path with and without `useColumnIndex`
- audit index construction and invalidation policy
- avoid rebuilding index state more broadly than necessary
- verify search path does not regress on sparse sheets
- add fixtures for:
  - monotonic lookup columns
  - mixed-type lookup columns
  - repeated lookup edits against stable data

Exit criteria:

- benchmark artifact shows improved lookup workloads
- memory growth from indexing is explicitly measured and acceptable

### 5. Range and Spill Graph Performance

Goal:

- keep `bilig`’s leadership areas from becoming performance liabilities

Why this matters:

- dynamic arrays and structured references are already strategic feature advantages
- if they are much slower or much harder to reason about, they do not translate into product
  leadership

Required work:

- profile spill recomputation path
- profile range-node invalidation and expansion
- reduce work for unchanged spill boundaries
- reduce allocation churn in spill materialization and serialization
- add dedicated leadership fixtures for:
  - `FILTER`
  - `UNIQUE`
  - table/structured-reference formulas
  - repeated spill resize under threshold edits

Exit criteria:

- leadership workloads remain deterministic and stable under stress
- leadership benchmarks show acceptable throughput and memory

### 6. Change Emission and Serialization Overhead

Goal:

- stop public result construction from dominating otherwise-fast internal work

Required work:

- measure internal compute vs external change serialization separately
- optimize `HeadlessChange[]` construction order and allocation behavior
- ensure adapters and public reads are detached without excessive copying
- consider fast-path representations internally with stable public conversion at the boundary

Exit criteria:

- stats can distinguish compute from export cost
- export cost decreases on edit/batch workloads

### 7. Memory Discipline

Goal:

- beat HyperFormula not just on time but on memory behavior and stability

Required work:

- record memory deltas for all benchmark scenarios
- add retained-memory checks for repeated workbook create/destroy loops
- audit caches so they have invalidation/eviction policy instead of unbounded growth
- add a long-run smoke benchmark for repeated headless workbook lifecycle use

Exit criteria:

- no obvious retained-memory slope in repeated lifecycle benchmarks
- benchmark artifacts make memory tradeoffs visible

### 8. WASM Fast-Path Expansion

Goal:

- use `@bilig/wasm-kernel` where it gives proven wins on numeric hot paths

Rules:

- JS semantics remain authoritative
- WASM is only allowed after parity and differential tests are green
- WASM adoption must be workload-backed, not aesthetic

Candidate areas:

- arithmetic-heavy aggregates
- vectorizable numeric transforms
- lookup helpers only if semantics remain exact
- date/time kernels only if locale and compatibility semantics are preserved

Exit criteria:

- named workloads show measurable improvement attributable to WASM use
- no divergence in formula correctness tests

## Semantics Work Required For Feature Leadership

These are real engine features, not API polish.

### Async Custom Functions

Needed if `bilig` wants a stronger programmable engine story than HyperFormula.

Must define:

- execution model
- timeout behavior
- cancellation model
- event timing and ordering
- deterministic snapshot/replay semantics

### Relative Named Expressions

Needed if `bilig` wants stronger workbook-programming semantics than HyperFormula.

Must define:

- rebasing rules under insert/delete/move/sort
- scope interaction with sheets and workbook names
- round-trip serialization behavior

### 3D References

Needed for broader workbook compatibility.

Must define:

- parser syntax
- cross-sheet dependency fanout
- mutation semantics under sheet reorder/rename/delete

### UI-Metadata-Aware Functions

Needed only if `bilig` wants product-aware semantics beyond a pure engine.

Must define:

- boundary between UI/app metadata and core engine
- deterministic headless fallback behavior
- isolation from `@bilig/headless` when metadata is absent

## Benchmark Program Requirements

The benchmark program must stay honest.

Rules:

- every direct comparison must use matched semantics
- unsupported workloads must be marked unsupported
- every published ratio must point to a checked-in artifact
- every performance claim must mention workload name, engine versions, host metadata, and sample
  methodology

Required artifacts:

- `packages/benchmarks/baselines/workpaper-baseline.json`
- `packages/benchmarks/baselines/workpaper-vs-hyperformula.json`

Recommended future artifacts:

- workload-family specific artifacts for:
  - build/restore
  - lookup/indexing
  - spill/range leadership
  - memory-retention lifecycle tests

## Execution Order

### Phase 1: Close direct-comparison deficits

- build/load
- single-edit recalc
- batch edit
- lookup/indexing

Reason:

- these are the cleanest places where HyperFormula currently wins on the checked-in artifact

### Phase 2: Strengthen leadership workloads

- dynamic arrays
- structured references
- table-aware recalculation
- multi-workbook process behavior

Reason:

- these are the places where `bilig` already has strategic feature leverage

### Phase 3: Expand engine semantics

- async custom functions
- relative named expressions
- 3D references where justified

Reason:

- these are expensive semantics projects and should not precede direct performance deficit work

## Required Deliverables

This program is only complete when it yields concrete repo outputs:

- code changes in `packages/core`, `packages/formula`, and optionally `packages/wasm-kernel`
- updated benchmark artifacts
- updated docs with precise claims
- tests that directly cover any new semantics
- profiling evidence for major engine changes

## Completion Criteria

This program is complete only when all of the following are true:

- the directly comparable benchmark deficits against HyperFormula have been materially reduced or
  reversed on named workloads
- `bilig` retains its leadership on dynamic arrays and structured references with production-grade
  tests
- at least one additional real engine feature gap has been closed, not just documented
- performance claims in `docs/workpaper-platform-design.md` can be stated more strongly without
  violating its evidence policy

## Bottom Line

If `bilig` wants to beat HyperFormula, the remaining work is not more facade work.

The remaining work is:

- faster core recalculation
- faster lookup/indexing
- cheaper build/restore
- stronger spill/range execution
- selective WASM acceleration where it actually helps
- closing real semantics gaps that HyperFormula still leaves open

That is the engine program this document captures.
