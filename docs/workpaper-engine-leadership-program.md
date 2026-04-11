# WorkPaper Engine Leadership Program

This document captures the actual engine work required for `bilig` to beat HyperFormula in ways
that matter, rather than only matching its public headless surface.

Date: `2026-04-10`

## Status

This document is an execution program, not a public-contract spec.

It complements, and does not replace:

- `docs/workpaper-platform-design.md`
- `docs/workpaper-performance-acceleration-plan.md`

`docs/workpaper-platform-design.md` defines the `WorkPaper` API contract, parity gates, external
consumer guarantees, and evidence policy.

`docs/workpaper-performance-acceleration-plan.md` defines the concrete architecture changes
required to close the measured performance gap.

This document defines the engine, formula-runtime, benchmark, and evidence program required to make
`WorkPaper` competitively superior in reality.

## Empirical Rule

This program uses empirical metrics, not taste.

Every leadership claim must fall into one of these buckets:

1. surface parity
   - API/config coverage against the local HyperFormula checkout
2. feature dominance
   - important semantics supported in `bilig` that HyperFormula explicitly does not support
3. formula dominance
   - formula breadth and production-quality closure
4. performance dominance
   - measured wins on directly comparable workloads
5. operability dominance
   - packaging, consumer installation, licensing, determinism, and diagnostics

No bucket may borrow evidence from another bucket.

Examples:

- API parity does not prove runtime speed.
- Formula count does not prove semantic quality.
- Dynamic-array support does not prove lookup performance.
- npm install success does not prove formula compatibility.

## Claim Vocabulary

This document uses a strict vocabulary so design language cannot outrun the evidence.

- `parity`
  - audited equality on a named surface against the local HyperFormula checkout
- `lead`
  - evidence-backed advantage on a named axis, while other axes may still be behind
- `dominance`
  - lead on an axis family with no red submetric left inside that family
- `overall lead`
  - no red category remains across surface, feature, formula, performance, and operability
- `10x better`
  - allowed only for a named workload with a checked-in artifact that includes mean, p95, host
    metadata, and engine versions

Forbidden shortcuts:

- calling breadth leadership formula dominance without production-closure proof
- calling feature leadership overall leadership while performance remains red
- calling a single benchmark win a package-level speed win

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

## Dominance Scorecard

This is the scorecard that should drive engineering priority.

| Axis | Metric | Current `bilig` evidence | Current HyperFormula evidence | Current position | Next proof needed |
| --- | --- | --- | --- | --- | --- |
| Surface parity | Public method inventory | `132/132` method names present on `HeadlessWorkbook` / `WorkPaper` | local checkout baseline | parity | keep the snapshot artifact green |
| Surface parity | Config-key inventory | `38/38` config keys present on `WorkPaperConfig` | local checkout baseline | parity | keep the snapshot artifact green |
| Formula breadth | Registered breadth against Office list | `487/508` registered in codebase = `95.9%` | docs claim `350/515` Excel functions = `68%` | `bilig` leads on breadth, but denominators are not identical | keep the generated dominance snapshot current |
| Formula breadth | Unified inventory breadth | `487/525` unified tracked functions = `92.8%` | no comparable local unified inventory artifact | `bilig` leads on tracked breadth | keep the unified inventory generated and current |
| Formula production quality | Canonical production closure | `300/300` canonical rows production-closed = `100%` | no matching canonical artifact | `bilig` leads on closure | keep the dominance snapshot current and extend grouped-array coverage beyond the canonical SUM forms |
| Feature dominance | Critical semantics unsupported by HyperFormula but present in `bilig` | dynamic arrays, structured references/tables, multiple workbook instances | HyperFormula docs list all three as unsupported/limited | `bilig` leads | add leadership workload benchmarks and soak tests so the lead is not purely semantic |
| Performance dominance | Directly comparable benchmark workloads | `1/6` wins in `workpaper-vs-hyperformula.json` | `5/6` wins on current host | HyperFormula still leads overall, but the lookup gaps are narrowing and `WorkPaper` now has a measured direct win on range reads | convert the current red workloads into majority `bilig` wins from the new lower-gap baseline |
| Performance dominance | Leadership workloads | `1/1` leadership workload exercised, with HyperFormula marked unsupported | dynamic arrays unsupported | `bilig` leads on capability, not comparable speed | expand leadership artifacts beyond one unsupported workload |
| Operability dominance | Clean external consumer path | packed tarball install and Vite/Node smoke are checked in-repo | no equivalent artifact in this repo | `bilig` leads in current repo evidence | keep smoke and publish paths green on every release path |
| Licensing and packaging | Open-source package posture | MIT publishable packages on npm | GPL license key flow in docs | `bilig` leads for embeddable OSS consumption | preserve the publishable OSS path while adding no hidden runtime requirements |

The important reading is:

- `bilig` already leads on surface completeness, feature breadth, and package operability
- `bilig` does not yet lead on directly comparable runtime speed
- `bilig` does not yet deserve a blanket overall-win claim

## Current Metric Values

Current measured values from local repo artifacts and docs:

- HyperFormula public API coverage in `WorkPaper`: `132/132`
- HyperFormula config coverage in `WorkPaperConfig`: `38/38`
- `bilig` registered formula breadth:
  - `487/508` Office-listed functions = `95.9%`
  - `487/525` unified tracked functions = `92.8%`
- HyperFormula published Excel coverage from the local docs:
  - `350/515` Excel functions = `68%`
- `bilig` canonical formula production closure:
  - `300/300` rows = `100%`
- directly comparable benchmark record:
  - `WorkPaper` wins: `2/6`
  - HyperFormula wins: `4/6`
  - current `WorkPaper` direct wins on this host:
    - build-from-sheets at `4.44x` faster
    - range-read at `1.08x` faster
  - HyperFormula current win range on this host: `1.37x` to `7.28x`
  - notable improvements from the latest tranches:
    - build-from-sheets improved from `6.95x` slower to a `WorkPaper` win at `4.44x` faster
    - batch-edit recalculation improved from `1000.39x` slower to `7.28x` slower
    - single-edit recalculation improved to `4.88x` slower after the `WorkPaper` facade stopped cloning or rescanning the workbook on ordinary edits and mixed-sheet imports stopped rebinding against half-built lookup columns
    - lookup without column indexing improved from `134.35x` slower to `1.37x` slower after the headless full-workbook diff path was removed from ordinary edits
    - lookup with `useColumnIndex` improved from `134.35x` slower to `2.28x` slower after direct exact lookup, event-driven `WorkPaper` change tracking, eager index priming, literal-only direct initialization, and trusted owned-op execution landed
- leadership workload record:
  - dynamic-array benchmark present
  - HyperFormula marked `unsupported`

## Evidence and Artifact Map

Every claim in the scorecard must point at a concrete artifact or command.

Surface parity:

- artifact:
  - `packages/headless/src/__tests__/fixtures/hyperformula-surface.json`
- source generator:
  - `scripts/gen-workpaper-hyperformula-audit.ts`
- check command:
  - `pnpm workpaper:parity:check`

Competitive performance:

- artifact:
  - `packages/benchmarks/baselines/workpaper-vs-hyperformula.json`
- source generator:
  - `scripts/gen-workpaper-vs-hyperformula-benchmark.ts`
- check command:
  - `pnpm workpaper:bench:competitive:check`

Internal regression baseline:

- artifact:
  - `packages/benchmarks/baselines/workpaper-baseline.json`
- source generator:
  - `scripts/gen-workpaper-benchmark-baseline.ts`
- check command:
  - `pnpm workpaper:bench:check`

External-consumer operability:

- smoke harness:
  - `scripts/workpaper-external-smoke.ts`
- check command:
  - `pnpm workpaper:smoke:external`

Publishability:

- package validation command:
  - `pnpm publish:runtime:check`
- release workflow:
  - `.github/workflows/headless-package.yml`

Formula breadth and canonical closure:

- breadth source:
  - `packages/formula/src/generated/formula-inventory.ts`
- canonical closure source:
  - `packages/formula/src/compatibility.ts`
  - `packages/formula/src/__tests__/fixtures/formula-dominance-snapshot.json`

## Program States

This program should be read as a state machine, not a binary done/not-done checklist.

`contract-complete`

- surface parity green
- operability green
- publishability green
- competitive benchmark artifact exists

`engine-catching-up`

- directly comparable performance is still red
- feature or formula leadership may already exist
- work should focus on the measured deficits, not new surface work

`performance-leading`

- `bilig` wins the majority of directly comparable workloads
- `bilig` loses none of the remaining comparable workloads by more than `1.25x`
- at least one hotspot family has both mean and p95 wins

`overall-leading`

- no red scorecard category remains
- formula closure is at or above the threshold in this document
- feature leadership remains intact
- performance-leading is true

The current state is `engine-catching-up`.

## What Counts As “Fully Beat HyperFormula”

This program treats “fully beat HyperFormula” as a multidimensional standard, not a slogan.

To claim that `bilig` fully beats HyperFormula, all of the following must be true at the same
time:

1. surface parity is preserved
   - `132/132` public method parity remains true
   - `38/38` config-key parity remains true
2. feature dominance is preserved
   - `bilig` continues to support dynamic arrays
   - `bilig` continues to support structured references/tables
   - `bilig` continues to support multiple workbook instances per process
3. formula dominance improves beyond breadth-only claims
   - Office-listed formula breadth stays at or above `95%`
   - canonical production closure remains `100%`
   - no canonical rows remain open in names, tables, structured references, lambda, or grouped-array families
   - non-canonical grouped-array follow-ons are tracked as expansion work, not as canonical debt
4. performance dominance becomes true on comparable workloads
   - `bilig` wins a majority of directly comparable benchmark workloads
   - `bilig` loses none of them by more than `1.25x`
   - at least two major hotspot families show a clear `bilig` win
5. operability dominance remains true
   - clean external install remains green
   - release/publish path remains green
   - public diagnostics and deterministic behavior remain stronger than the baseline package story

If any one of those is false, the correct claim is narrower than “fully beat HyperFormula”.

## Benchmark And Profiling Discipline

Performance work in this program is only accepted when the measurement method is as disciplined as
the code change.

Rules:

- all direct comparisons must preserve semantic equivalence and verification output
- both mean and p95 matter; a faster mean with a materially worse p95 is not a clean win
- every benchmark change must identify whether it improved:
  - compute cost
  - export/change serialization cost
  - memory behavior
  - or fixture/setup overhead
- no hotspot claim is accepted without naming the exact workload from
  `benchmark-workpaper-vs-hyperformula.ts`
- no profiling result is accepted without naming the repo surface it implicates

Required measurement loop for performance patches:

1. profile the existing workload and name the hot repo surface
2. patch the engine or formula runtime
3. rerun:
   - `pnpm workpaper:bench:competitive:check`
   - the targeted benchmark generator in write mode when the artifact must change
4. record whether the win came from compute, serialization, or setup
5. reject the change as a dominance improvement if it shifts cost into a different red category

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

For this program, “overall” is made concrete by the scorecard above, not by prose.

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
  - multiple independent workbook instances with no one-instance-one-workbook restriction

The current repository does not support these claims:

- that `bilig` is faster than HyperFormula in general
- that `bilig` is `10x` better overall
- that all meaningful engine gaps are closed
- that formula breadth leadership already implies formula production leadership

That means the engineering program must focus on real runtime work, not just surface-level parity.

## Priority By Deficit Size

The scorecard gives a direct priority order.

### Priority 0: Preserve existing wins

Do not regress:

- method/config parity
- dynamic arrays
- structured references/tables
- external consumer install/publish path

These are existing advantages and must not be traded away while chasing speed.

### Priority 1: Fix the largest measured performance deficits

Based on the checked-in benchmark artifact, the most urgent directly comparable gaps are:

- batch-edit recalculation: HyperFormula currently leads by `7.28x`
- single-edit recalculation: HyperFormula currently leads by `4.88x`
- lookup with column indexing: HyperFormula currently leads by `2.28x`
- lookup without column indexing: HyperFormula currently leads by `1.37x`
- range-read: `WorkPaper` currently leads by `1.08x`
- build from sheets: `WorkPaper` currently leads by `4.44x`

The important trend changes are:

- the triple-digit lookup deficits are gone
- the last tranches proved that `WorkPaper` facade overhead was masking engine progress; removing whole-workbook before/after diffs from ordinary edits cut the lookup workloads down to single-digit territory
- direct reads from workbook storage improved the build and edit workloads again after the event-driven diff path landed
- `useColumnIndex` is now a real direct lookup path and meets the phase-1 threshold, but HyperFormula still has a faster core search subsystem

### Priority 2: Protect formula production-quality leadership

Breadth is already high, and canonical production closure is now finished.

The next formula-quality work is expansion beyond the canonical closure:

- broader `GROUPBY` aggregate support beyond the SUM-only canonical form
- broader `PIVOTBY` aggregate support beyond the SUM-only canonical form
- grouped-array performance and stability under larger spill and mutation workloads

This matters because formula dominance is not just “the parser recognizes the name.”

It means:

- semantically correct
- production-routable
- benchmarkable
- stable under mutation and serialization

## Primary Code Surfaces By Workstream

This program should point directly into the runtime, not stop at high-level package names.

Build and load path:

- `packages/core/src/engine/services/snapshot-service.ts`
- `packages/core/src/engine.ts`
- `packages/formula/src/parser.ts`
- `packages/formula/src/compiler.ts`
- `packages/headless/src/headless-workbook.ts`

Single-edit and batch recalculation:

- `packages/core/src/engine/services/recalc-service.ts`
- `packages/core/src/engine/services/mutation-service.ts`
- `packages/core/src/engine/services/traversal-service.ts`
- `packages/core/src/engine/services/formula-evaluation-service.ts`

Lookup and indexing:

- `packages/formula/src/builtins/lookup.ts`
- `packages/formula/src/builtins/lookup-reference-builtins.ts`
- `packages/formula/src/js-evaluator.ts`
- `packages/core/src/engine/services/read-service.ts`

Spill, arrays, and structured references:

- `packages/formula/src/js-evaluator-array-special-calls.ts`
- `packages/formula/src/js-evaluator-workbook-special-calls.ts`
- `packages/core/src/engine/services/formula-graph-service.ts`
- `packages/core/src/range-registry.ts`
- `packages/core/src/formula-table.ts`

Serialization, export, and public boundary cost:

- `packages/headless/src/headless-workbook.ts`
- `packages/headless/src/types.ts`
- `packages/core/src/engine/services/event-service.ts`
- `packages/core/src/engine/services/history-service.ts`

WASM acceleration:

- `packages/formula/src/binder-wasm-rules.ts`
- `packages/core/src/wasm-facade.ts`
- `packages/wasm-kernel/assembly/*`

## Non-Goals

This program explicitly does not call for:

- building a second spreadsheet engine
- cloning HyperFormula internals line by line
- destabilizing the `WorkPaper` public contract without necessity
- adding speculative features with no semantics design
- claiming performance wins before benchmarks prove them
- confusing directional breadth metrics with apples-to-apples quality metrics

## Required Engine Workstreams

Every workstream below must report its effect back into the scorecard.

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
- preserve the literal-only build-from-sheets fast path while extending the same low-overhead
  strategy to mixed-content restore paths
- add benchmark fixtures for:
  - dense literals
  - repeated formulas
  - mixed literals + formulas
  - large named-sheet workbook restore

Exit criteria:

- `workpaper-vs-hyperformula.json` includes improved build-path results
- build workload no longer exceeds a `2x` loss ratio
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
- single-edit workload no longer shows a material multi-x loss
- stats expose enough detail to attribute wins

### 3. Batch Mutation Efficiency

Goal:

- make batched edits scale with the affected graph, not with avoidable per-operation overhead

Likely bottlenecks:

- repeated write-path validation
- repeated invalidation work inside suspended evaluation mode
- overly expensive combined change-array materialization

Required work:

- keep the already-landed headless fixes intact:
  - suppress per-edit snapshot diffing inside outer `batch()`
  - defer literal cell writes into a single engine-op flush where semantics allow it
- audit batch path for per-edit repeated work that should be deferred
- aggregate invalidation metadata during `batch()` / suspended evaluation
- deduplicate structural change reporting where the public result permits it
- add a benchmark fixture with many edits to formula-backed rows

Exit criteria:

- batched-edit slope improves materially with edit count
- batch-edit loss ratio is reduced to a non-embarrassing range before broader victory claims
- benchmark artifact records the result

### 4. Lookup and Indexing Workloads

Goal:

- close the current performance gap on lookup-heavy workloads

This is currently one of the clearest directly comparable areas where HyperFormula wins.

Required work:

- keep the newly-landed indexed exact-match path intact:
  - `useColumnIndex` now routes exact vector `MATCH` / `XMATCH` candidates onto a persistent column cache with column-local invalidation
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
- at least one lookup workload becomes a `bilig` win before this workstream can be called done
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
- at least one leadership workload has a checked-in advantage claim that is stronger than “unsupported
  in HyperFormula”

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

## Formula Dominance Program

Formula dominance must be measured on three separate layers.

### Layer 1: Breadth

Metric:

- registered runtime formulas / target formula inventory

Current values:

- `487/508` Office-listed = `95.9%`
- `487/525` unified tracked = `92.8%`

Interpretation:

- good breadth
- not sufficient by itself to claim formula leadership

### Layer 2: Production Closure

Metric:

- canonical rows that are production-routable with correct semantics

Current value:

- `300/300` = `100%`

Interpretation:

- full canonical closure
- no JS-only rows remain in the canonical slice
- names, tables, structured references, and lambda canonical rows are now production-closed

### Layer 3: Strategic Semantics

Metric:

- whether `bilig` production-closes important formula families that HyperFormula does not support

Families:

- dynamic arrays
- structured references
- table-aware totals and references
- lambda-family semantics where product value justifies them

Interpretation:

- this is where `bilig` can become obviously stronger, not merely broader

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

Shipped formula artifact:

- `packages/formula/src/__tests__/fixtures/formula-dominance-snapshot.json`

That snapshot now covers:

- formula breadth derived from `formula-inventory.ts`
- canonical production closure derived from `compatibility.ts`
- strategic-family status for arrays, names, tables, structured references, and lambda

## Minimum Patch Shape

No workstream in this document is complete with a code-only patch.

Every substantial patch in this program should produce:

- the code change in the implicated runtime surface
- at least one targeted test in the owning package
- either:
  - an updated benchmark artifact
  - or a deliberate note that the patch was correctness-only and should not change dominance claims
- a short doc update if the public claim surface changes
- clear evidence that no existing leadership area regressed

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

### Phase 4: Convert directional wins into dominance claims

- promote breadth metrics into artifact-backed dominance statements
- publish workload-specific win claims only where the artifact shows real wins
- revise `docs/workpaper-platform-design.md` only after the scorecard actually improves

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
- formula breadth and formula production-closure metrics both support a leadership claim, not just a
  breadth claim
- the scorecard supports “overall lead” without hiding any red category

## What Still Makes Me Unsatisfied

This document is only worth keeping if it remains stricter than a normal roadmap.

I am satisfied with it only if these statements stay true:

- it makes the current losses explicit instead of hiding them behind breadth wins
- it points from every major claim to a concrete repo artifact or runtime surface
- it tells an engineer where to start next without inventing a second design doc
- it blocks vague “10x better” language unless the artifact and workload actually prove it

If those stop being true, this document should be edited again before more leadership claims are
made.

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
