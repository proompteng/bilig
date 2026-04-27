# Production Debt Plan: Core State, Formula/WASM, Browser Runtime

Date: 2026-04-27
Status: proposed
Scope: debt items 4, 5, and 6 from the April 2026 stack audit

This plan covers:

1. Core engine state and mutation orchestration.
2. Formula runtime and WASM fast-path maintainability.
3. Browser worker runtime, projected viewport authority, and sync/E2E coverage.

The goal is not to make files smaller for optics. The goal is to reduce semantic
blast radius, remove duplicate authority, improve replayability, and make the
system safer to evolve under production load.

## Current Evidence

The live `main` checkout is structurally healthy enough to support a serious
production cleanup:

- `pnpm analyze:quality` passes with no dependency-cycle violations.
- `pnpm typecheck` passes.
- `pnpm lint` passes with zero warnings and zero errors.
- Correctness gates already exist for core, formula, browser, and server paths.

That means the next work should not be generic hygiene. It should target the
remaining architectural debt directly.

Primary debt surfaces:

- `packages/core/src/workbook-store.ts`
- `packages/core/src/engine/services/mutation-service.ts`
- `packages/formula/src/js-evaluator.ts`
- `packages/formula/src/binder.ts`
- `packages/formula/src/binder-wasm-rules.ts`
- `packages/wasm-kernel/assembly/*`
- `apps/web/src/worker-runtime.ts`
- `apps/web/src/projected-viewport-store.ts`
- `e2e/tests/web-shell.pw.ts`
- `e2e/tests/web-shell-remote-sync.pw.ts`

## Non-Negotiables

- No semantic rewrite without a failing replay, invariant, differential, or
  contract test first.
- No broad compatibility branch left behind after the replacement path is proven.
- No hidden feature flag that lets the old path silently stay alive in product.
- No benchmark-specific special cases.
- No WASM semantic expansion until the JS path has parity tests and differential
  coverage.
- No removal of skipped E2E coverage unless an equivalent deterministic harness
  is already green.
- No "green" claim based only on targeted tests when the slice changes shared
  semantics. Each tranche must state whether it reached full `pnpm run ci`.
- No unreviewable giant diff. If source changes exceed roughly 1000 lines, land a
  focused checkpoint commit before the final verification pass.

## Program Structure

Run this as three coordinated workstreams with shared gates:

- Stream A: Core state and mutation ownership.
- Stream B: Formula/WASM semantic decomposition.
- Stream C: Browser runtime and sync verification.

The workstreams can proceed in parallel only when their write sets do not
overlap. Shared contracts, generated protocol output, or public package exports
must be coordinated in one tranche.

## Shared Acceptance Gates

Every production tranche must satisfy:

- `pnpm typecheck`
- `pnpm lint`
- focused vitest suites for touched modules
- generated-file checks when protocol/formula inventory changes
- no new `test.skip(...)` on production-critical paths
- no dependency-cycle regression via `pnpm analyze:quality`

Release-level completion requires:

- `pnpm run ci`
- `pnpm test:correctness:core`
- `pnpm test:correctness:formula`
- `pnpm test:correctness:browser`
- `pnpm test:correctness:server`
- `pnpm test:browser`
- `pnpm bench:smoke`
- `CI=1 pnpm bench:contracts`

## Stream A: Core State And Mutation Ownership

### Target Outcome

`WorkbookStore` becomes a composition of explicit state stores. `EngineMutationService`
becomes a workflow coordinator over isolated mutation-policy modules. The public
`SpreadsheetEngine` API remains stable.

### Invariants

- Snapshot import/export preserves workbook meaning.
- Applying an op stream from the same base produces the same workbook.
- Undo/redo restores semantic state, not incidental object identity.
- Structural insert/delete/move rewrites formulas, metadata, tables, filters,
  sorts, spills, pivots, charts, images, and shapes consistently.
- Style/format interning remains deterministic across snapshot round trips.
- Axis identity and logical row/column identity do not diverge after structural
  operations.

### Phase A0: Lock The Baseline

Add focused characterization tests before extraction:

- `workbook-store` tests for style/format interning, axis metadata, freeze panes,
  tables, filters, sorts, comments, pivots, charts, images, and shapes.
- `mutation-service` tests for inverse op creation, structural delete undo,
  structural insert undo, range move/copy/fill, render commit, and remote restore
  paths.
- Property tests for edit plus undo, structural mutation plus snapshot round
  trip, and op replay parity.

Exit gate:

- Existing behavior is captured well enough that extraction failures are obvious.

### Phase A1: Split WorkbookStore By Ownership

Create focused modules under `packages/core/src/workbook-store/` or
`packages/core/src/storage/`:

- `sheet-registry-store.ts`: sheet name/id/order lifecycle.
- `cell-record-store.ts`: cell storage access and cell id ownership wrappers.
- `style-format-store.ts`: style and number-format interning.
- `axis-metadata-store.ts`: row/column metadata and axis materialization.
- `workbook-object-store.ts`: tables, validations, conditional formats,
  protections, comments, notes, spills, pivots, charts, images, and shapes.
- `structural-axis-store.ts`: structural axis transforms and resident cell
  remapping.

Keep `WorkbookStore` as the facade while moving implementation behind those
stores.

Exit gate:

- `WorkbookStore` no longer directly owns every metadata family.
- Each extracted store has direct tests.
- Public core exports and snapshot shapes are unchanged.

### Phase A2: Split Mutation Policy From Mutation Execution

Extract from `mutation-service.ts`:

- `mutation-inverse-ops.ts`: inverse op construction.
- `mutation-canonicalization.ts`: forward op normalization.
- `mutation-structural-undo.ts`: deleted cell/formula/metadata capture.
- `mutation-batch-planner.ts`: transaction record creation, potential-new-cell
  accounting, and batch creation.
- `mutation-render-commit.ts`: render-commit conversion and undo capture.

`createEngineMutationService` remains the orchestration boundary, but its local
helpers move into tested modules.

Exit gate:

- Inverse construction can be tested without instantiating the full engine.
- Structural undo can be tested against narrow workbook fixtures.
- Mutation service line count drops because responsibility moved, not because
  code was hidden in untested helpers.

### Phase A3: Add Mutation Counters And Failure Diagnostics

Add counters for:

- inverse ops created
- structural metadata records captured
- cells remapped
- changed formulas rebound
- batch listeners notified
- undo records merged
- render commit cell mutations converted

Expose counters through existing benchmark/reporting paths where appropriate.

Exit gate:

- Batch edit and structural mutation workloads explain where time is spent.
- Future performance work has evidence instead of speculation.

## Stream B: Formula Runtime And WASM Fast Path

### Target Outcome

Formula semantics are organized by responsibility. The JS evaluator remains the
semantic source of truth. WASM accelerates only closed, proven-equal formula
families.

### Invariants

- Parser, binder, compiler, JS evaluator, and WASM eligibility agree on the
  same formula meaning.
- Error propagation and coercion stay Excel-compatible for covered formulas.
- JS/WASM parity is mandatory for every accelerated family.
- Unsupported formulas fail visibly and deterministically.
- No benchmark-only fast path is introduced.

### Phase B0: Build The Semantic Safety Net

Before extraction:

- Add direct tests for evaluator special calls by family.
- Add binder tests that distinguish dependency collection from WASM safety.
- Add generated fixture coverage for lookup, text, date/time, finance, aggregate,
  array, and database families.
- Add JS/WASM differential tests for every currently accelerated family.
- Add negative tests for formulas that must remain JS-only.

Exit gate:

- A formula-family extraction can fail one small test instead of only a giant
  builtins suite.

### Phase B1: Split Binder Responsibilities

Separate:

- dependency collection
- scope/local-name handling
- builtin availability and encoding
- WASM eligibility rules
- top-level lowering policy

Proposed files:

- `formula-dependency-collector.ts`
- `formula-builtin-encoding.ts`
- `formula-wasm-eligibility.ts`
- `formula-lowering-policy.ts`

Exit gate:

- Dependency extraction can run without WASM policy.
- WASM eligibility can be tested with synthetic AST fixtures.
- `binder.ts` becomes a composition layer.

### Phase B2: Split JS Evaluator Special Calls

Extract special-call families from `js-evaluator.ts`:

- lambda/apply helpers
- dynamic array helpers
- lookup dispatch helpers
- range/matrix coercion helpers
- broadcast shape and array result helpers
- scalar coercion/comparison helpers

Proposed files:

- `evaluator-stack.ts`
- `evaluator-coercion.ts`
- `evaluator-arrays.ts`
- `evaluator-lookup.ts`
- `evaluator-lambda.ts`
- `evaluator-special-calls.ts`

Exit gate:

- `evaluatePlanResult` remains the public evaluator entrypoint.
- Extracted modules have direct tests.
- No behavior change unless a replay test proves the old behavior was wrong.

### Phase B3: Organize Builtin Registries By Semantic Family

Continue decomposing lookup and text builtins so registries are registries, not
semantic kitchens.

Targets:

- lookup reference/search
- lookup criteria/database
- lookup matrix/regression/statistics
- text byte functions
- text formatting
- text search/regex
- locale/Japanese width functions

Exit gate:

- Builtin map files primarily compose family modules.
- Shared helpers live in explicit utility modules with direct tests.

### Phase B4: Make WASM Eligibility And AssemblyScript Dispatch Auditable

For each WASM family:

- name the JS owner function
- name the WASM dispatch entrypoint
- name the differential test file
- name unsupported shape boundaries
- record mismatch policy

Split AssemblyScript files only by semantic family and dispatch boundary, not by
arbitrary line count.

Exit gate:

- No accelerated formula family exists without a JS owner and parity suite.
- `packages/wasm-kernel/assembly/builtins.ts` and broad dispatch files are
  reduced by moving real semantic families into focused modules.

### Phase B5: Add Formula Performance Counters After Correctness

Add counters for:

- formulas parsed
- formulas bound
- formulas lowered to WASM
- formulas rejected from WASM with reason
- JS evaluator special-call counts
- WASM program upload bytes
- full vs delta WASM upload count
- formula family install count

Exit gate:

- Build-template, runtime-restore, and WASM upload costs are measurable before
  optimization begins.

## Stream C: Browser Runtime And Sync Verification

### Target Outcome

The browser worker runtime becomes a service-based runtime with one local
authority model. The projected viewport store stops being a broad truth layer.
Skipped sync E2E paths are replaced with deterministic harnesses or unskipped.

### Invariants

- Worker runtime is the mounted browser session authority.
- UI consumes projected patches and render tiles; it does not invent workbook
  truth.
- Pending local mutations survive crash/restart where designed.
- Authoritative revisions absorb submitted local mutations exactly once.
- Reconnect catch-up preserves local/authoritative parity.
- Projection rows match engine state.
- Browser-visible cells never split from formula bar/engine truth after commit.

### Phase C0: Build Deterministic Runtime Harnesses

Create harnesses before decomposition:

- fake worker transport
- fake local persistence
- fake Zero bridge
- fake clock/id source
- in-memory runtime session
- authoritative revision feed simulator

Test scenarios:

- boot from empty state
- boot from persisted local state
- authoritative snapshot install
- pending mutation journal replay
- local edit then authoritative absorption
- reconnect with drift
- viewport patch after stale local overlay
- agent preview/apply interaction with runtime state

Exit gate:

- Browser-runtime correctness can be tested without Playwright timing.

### Phase C1: Split WorkerRuntime Services

Extract from `worker-runtime.ts`:

- `worker-runtime-bootstrap.ts`: boot, hydration, initial snapshot selection.
- `worker-runtime-authority.ts`: authoritative snapshot/revision install.
- `worker-runtime-journal.ts`: pending mutation journal and absorption.
- `worker-runtime-projection.ts`: projection rebuild and overlay reconciliation.
- `worker-runtime-viewport-service.ts`: viewport patch and render-tile
  publication.
- `worker-runtime-agent-preview.ts`: agent command preview and applied bundle
  support.
- `worker-runtime-persistence-service.ts`: local persistence queue and recovery.

Keep `WorkerRuntime` as the facade until each service is independently tested.

Exit gate:

- Persistence, sync, projection, viewport, and agent preview can be tested
  independently.
- `worker-runtime.ts` delegates instead of owning all workflows directly.

### Phase C2: Replace ProjectedViewportStore As Broad Authority

Split `projected-viewport-store.ts` into:

- `projected-cell-cache.ts`: bounded projected cell cache.
- `projected-axis-cache.ts`: row/column metadata cache.
- `projected-damage-index.ts`: dirty range and invalidation tracking.
- `projected-render-tile-store.ts`: render tile deltas and tile snapshots.
- `projected-overlay-state.ts`: narrow local-only visual overlays.

Remove or downgrade `MAX_CACHED_CELLS_PER_SHEET = 6000` from an architecture
constraint to an implementation detail with tests and counters.

Exit gate:

- The projected store no longer acts as a broad local truth layer.
- Reconnect/invalidation tests prove stale projected data cannot override
  authoritative worker state.

### Phase C3: Replace Skipped Sync E2E With Deterministic Coverage

For each skipped remote-sync browser test:

1. Identify the product guarantee.
2. Add deterministic server/browser-runtime harness coverage.
3. Keep or unskip Playwright only for behavior that truly requires a browser.
4. Delete the skip only after replacement coverage is green.

Targets:

- `e2e/tests/web-shell-remote-sync.pw.ts`
- remote-sync sections in `e2e/tests/web-shell.pw.ts`
- remote-sync sections in `e2e/tests/web-shell-scroll-performance.pw.ts`

Exit gate:

- No critical sync behavior is hidden behind environment-only `test.skip(...)`
  without a deterministic test covering the same guarantee.

### Phase C4: Add Resource And Soak Gates

Add release/nightly checks for:

- reconnect catch-up with pending ops
- local journal replay volume
- viewport patch memory ceiling
- projected cache eviction correctness
- event-loop latency during large workbook edits
- browser scroll/render behavior after authoritative drift

Exit gate:

- The browser runtime is not only logically correct; it is bounded under load.

## Production Rollout

### Tranche Order

1. A0, B0, C0: characterization and harnesses.
2. A1: split `WorkbookStore` ownership.
3. A2/A3: split mutation policy and add counters.
4. B1/B2: split binder and evaluator policy.
5. B3/B4/B5: builtin family cleanup, WASM audit map, formula counters.
6. C1: split worker runtime services.
7. C2: narrow projected viewport authority.
8. C3/C4: remove skipped sync blind spots and add soak gates.

### Commit Discipline

Each tranche should be small enough to review. Suggested commit scopes:

- `test(core): characterize workbook store ownership`
- `refactor(core): split workbook style and axis stores`
- `refactor(core): isolate mutation inverse planning`
- `test(formula): gate wasm eligibility policy`
- `refactor(formula): split evaluator special calls`
- `refactor(wasm): split aggregate dispatch family`
- `test(web): add worker runtime reconnect harness`
- `refactor(web): split worker runtime persistence service`
- `test(e2e): replace remote sync skip with deterministic harness`

### Final Definition Of Done

This program is done only when all of the following are true:

- `WorkbookStore` is a facade over focused ownership stores.
- Mutation canonicalization, inverse op planning, and structural undo are tested
  without the full engine.
- Binder dependency collection and WASM eligibility are independent modules.
- JS evaluator special-call families are isolated and directly tested.
- Every WASM accelerated family has JS owner, dispatch owner, and differential
  parity coverage.
- `WorkerRuntime` is no longer a god object.
- `ProjectedViewportStore` is no longer a broad truth layer.
- Critical remote-sync skips are gone or backed by deterministic tests.
- `pnpm run ci` is green on the committed tree.

