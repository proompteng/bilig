# Tech Debt Remediation Program
## Date: 2026-04-09
## Status: proposed

## Purpose

This document turns the current codebase audit into an execution plan.

The goal is not to chase file-size metrics in isolation. The goal is to reduce
semantic blast radius, replace temporary authority layers, tighten reliability,
and make future bug fixes safer and cheaper.

This plan is based on the tracked `main` checkout as of `d83cc9b1`.
It intentionally excludes untracked local scratch work such as
`/Users/gregkonush/github.com/bilig/packages/headless`.

It is aligned with:

- [correctness-gates-program-2026-04-09.md](/Users/gregkonush/github.com/bilig/docs/correctness-gates-program-2026-04-09.md)
- [spreadsheet-engine-effect-service-refactor-2026-04-08.md](/Users/gregkonush/github.com/bilig/docs/spreadsheet-engine-effect-service-refactor-2026-04-08.md)
- [production-stability-remediation-2026-04-02.md](/Users/gregkonush/github.com/bilig/docs/production-stability-remediation-2026-04-02.md)
- [05-06-next-phase.md](/Users/gregkonush/github.com/bilig/docs/05-06-next-phase.md)

## Current state

The repo is materially healthier than it was:

- correctness gates exist and are wired into CI
- dead-code and cycle analysis are in place and currently green
- the old giant `engine.ts` hotspot has been reduced significantly

The remaining debt is now concentrated in a smaller set of structurally risky
modules and in a few intentionally temporary architectural paths.

## Operating constraints

This program should be executed under the following rules:

- use the shared repository checkout at
  [/Users/gregkonush/github.com/bilig](/Users/gregkonush/github.com/bilig)
- keep the checked out branch on `main`
- do not use temporary worktrees or side branches unless there is an explicit
  operational reason and it is documented in the current task
- do not overwrite or revert unrelated local changes in the shared checkout
- do not treat untracked local scratch areas as part of this program

Current examples that are out of scope unless they become intentionally tracked:

- [/Users/gregkonush/github.com/bilig/packages/headless](/Users/gregkonush/github.com/bilig/packages/headless)
- [/Users/gregkonush/github.com/bilig/docs/core-product-quality-release.md](/Users/gregkonush/github.com/bilig/docs/core-product-quality-release.md)

If an execution slice is blocked by unrelated local state, the fix is to work
around that state without destroying it.

## Ownership model

This program is organized by subsystem ownership rather than by one generic
"refactor" bucket.

- Formula runtime owner surface:
  [/Users/gregkonush/github.com/bilig/packages/formula](/Users/gregkonush/github.com/bilig/packages/formula)
- Browser runtime owner surface:
  [/Users/gregkonush/github.com/bilig/apps/web](/Users/gregkonush/github.com/bilig/apps/web)
- Core engine/store owner surface:
  [/Users/gregkonush/github.com/bilig/packages/core](/Users/gregkonush/github.com/bilig/packages/core)
- Grid/rendering owner surface:
  [/Users/gregkonush/github.com/bilig/packages/grid](/Users/gregkonush/github.com/bilig/packages/grid)
- Sync/server owner surface:
  [/Users/gregkonush/github.com/bilig/apps/bilig/src](/Users/gregkonush/github.com/bilig/apps/bilig/src)
- WASM fast-path owner surface:
  [/Users/gregkonush/github.com/bilig/packages/wasm-kernel](/Users/gregkonush/github.com/bilig/packages/wasm-kernel)

Each wave below names the primary owning surface so the work can be tracked and
reviewed by responsibility.

## What the audit found

### 1. Formula runtime remains the largest semantic hotspot

Highest-risk files:

- [lookup.ts](/Users/gregkonush/github.com/bilig/packages/formula/src/builtins/lookup.ts)
- [text.ts](/Users/gregkonush/github.com/bilig/packages/formula/src/builtins/text.ts)
- [js-evaluator.ts](/Users/gregkonush/github.com/bilig/packages/formula/src/js-evaluator.ts)
- [binder.ts](/Users/gregkonush/github.com/bilig/packages/formula/src/binder.ts)

Why this is debt:

- `lookup.ts` still mixes lookup behavior, database criteria logic, matrix math,
  statistics helpers, and builtin registration in one file.
- `text.ts` still mixes text builtins, byte-oriented behavior, regex behavior,
  number/date formatting semantics, Japanese width transforms, and locale-style
  text helpers in one file.
- `js-evaluator.ts` still contains a large special-call interpreter.
- `binder.ts` still combines dependency extraction with WASM-safety policy and
  formula-lowering decisions.

Risk:

- this is still the biggest semantic blast radius in the product
- bugs here are customer-visible and often subtle
- changes are expensive to verify because responsibilities are blended

### 2. The browser still contains a temporary truth layer

Primary file:

- [projected-viewport-store.ts](/Users/gregkonush/github.com/bilig/apps/web/src/projected-viewport-store.ts)

Why this is debt:

- it still acts as a local authority for visible cell/state overlays
- it still relies on `MAX_CACHED_CELLS_PER_SHEET = 6000`
- cache eviction is still heuristic rather than architecture-driven

The docs already describe this as temporary:

- [05-06-next-phase.md](/Users/gregkonush/github.com/bilig/docs/05-06-next-phase.md#L1216)
- [05-06-next-phase.md](/Users/gregkonush/github.com/bilig/docs/05-06-next-phase.md#L1225)
- [correctness-gates-program-2026-04-09.md](/Users/gregkonush/github.com/bilig/docs/correctness-gates-program-2026-04-09.md#L187)

Risk:

- visible split-brain under reconnect, invalidation, or stale overlays
- memory and pruning edge cases
- expensive reasoning because authoritative state is still partially mirrored in
  the browser

### 3. Core workbook state and mutation logic are still too centralized

Primary files:

- [workbook-store.ts](/Users/gregkonush/github.com/bilig/packages/core/src/workbook-store.ts)
- [mutation-service.ts](/Users/gregkonush/github.com/bilig/packages/core/src/engine/services/mutation-service.ts)

Why this is debt:

- `workbook-store.ts` still owns sheets, cells, style/format interners, axis
  state, metadata, filters, sorts, spills, pivots, tables, and freeze panes
- `mutation-service.ts` still owns a large portion of mutation canonicalization,
  inverse-op creation, and transaction orchestration

Risk:

- correctness fixes touch too many concerns at once
- the state model is harder to isolate for tests and service boundaries
- metadata/range/axis/state interactions are still dense

### 4. The worker-side browser runtime is still a god object

Primary file:

- [worker-runtime.ts](/Users/gregkonush/github.com/bilig/apps/web/src/worker-runtime.ts)

Why this is debt:

- bootstrap, persistence, authoritative hydration, projection rebuild, pending
  mutation journaling, snapshot export, viewport patch production, and agent
  preview still live in one class

Risk:

- behavior is coupled across boot, sync, persistence, and rendering
- failures at runtime boundaries are harder to isolate
- any change in one path risks regressions in unrelated worker behavior

### 5. Grid interaction and rendering are still oversized controller surfaces

Primary files:

- [useWorkbookGridInteractions.ts](/Users/gregkonush/github.com/bilig/packages/grid/src/useWorkbookGridInteractions.ts)
- [gridGpuScene.ts](/Users/gregkonush/github.com/bilig/packages/grid/src/gridGpuScene.ts)

Why this is debt:

- `useWorkbookGridInteractions.ts` still combines pointer selection, drag
  behavior, fill handle behavior, resize behavior, clipboard handling, keyboard
  handling, and context menu behavior in one hook
- `gridGpuScene.ts` still mixes scene tokens, geometry generation, selection
  overlays, hover visuals, resize guides, and border rendering

Risk:

- input bugs are hard to localize
- rendering regressions are hard to review
- performance work competes with readability because the module boundary is too broad

### 6. End-to-end sync coverage is still partly skipped

Primary file:

- [web-shell.pw.ts](/Users/gregkonush/github.com/bilig/e2e/tests/web-shell.pw.ts)

Why this is debt:

- several Zero-backed browser sync tests still use `test.skip(...)`
- the file remains large and multi-purpose

Risk:

- critical cross-tab behaviors are still not fully enforced by the browser suite
- skipped tests normalize blind spots over time

### 7. The WASM kernel is still difficult to evolve safely

Primary files:

- [builtins.ts](/Users/gregkonush/github.com/bilig/packages/wasm-kernel/assembly/builtins.ts)
- [date-finance.ts](/Users/gregkonush/github.com/bilig/packages/wasm-kernel/assembly/date-finance.ts)
- various `dispatch-*` files under
  [/Users/gregkonush/github.com/bilig/packages/wasm-kernel/assembly](/Users/gregkonush/github.com/bilig/packages/wasm-kernel/assembly)

Why this is debt:

- the accelerated formula fast path still has very large AssemblyScript modules
- parity and dispatch logic are still broad, even though the correctness gates
  are stronger now

Risk:

- fast-path changes remain expensive and review-heavy
- kernel parity fixes are harder than they need to be

## What is not the main problem anymore

These are not currently the dominant debt signals:

- dead code drift
- dependency cycles in the tracked workspaces
- the old `engine.ts` monolith shape
- lack of correctness gates

This matters because the next work should target semantic and architectural
surfaces, not just run more hygiene tooling.

## Remediation principles

### 1. Fix bug-positive slices first

Every refactor wave should start by either:

- landing a failing replay/invariant test, or
- tightening a correctness gap that the current code already makes hard to reason about

### 2. Reduce authority duplication

Temporary browser-side state stores should move toward:

- worker-side hot state
- persisted local authoritative state
- narrower render overlays

### 3. Keep pure logic pure

Use `Effect` only for boundaries:

- storage
- clock
- retries
- timeouts
- worker messaging
- sync transport
- resource lifecycle

Do not wrap pure formula math, binding logic, or address rewriting in `Effect`
just to match a style.

### 4. Split by responsibility, not by arbitrary line count

The correct target is not "smaller files" in the abstract.
The correct target is:

- one module per concern
- explicit boundaries
- smaller blast radius
- easier replay and invariant testing

### 5. Keep test debt visible

Skipped tests and giant multi-purpose suites are debt, even when the repo is green.

## Execution order

### Wave 1: Formula runtime decomposition

Primary owner surface:

- [/Users/gregkonush/github.com/bilig/packages/formula](/Users/gregkonush/github.com/bilig/packages/formula)

Targets:

- [lookup.ts](/Users/gregkonush/github.com/bilig/packages/formula/src/builtins/lookup.ts)
- [text.ts](/Users/gregkonush/github.com/bilig/packages/formula/src/builtins/text.ts)
- [js-evaluator.ts](/Users/gregkonush/github.com/bilig/packages/formula/src/js-evaluator.ts)
- [binder.ts](/Users/gregkonush/github.com/bilig/packages/formula/src/binder.ts)

Work:

- split lookup/database/statistics/matrix helpers from builtin registries
- split text formatting, byte behavior, regex behavior, locale text helpers, and builtin registries
- extract special-call evaluation families from `js-evaluator.ts`
- separate dependency collection from WASM-safety policy in `binder.ts`
- add direct tests to each extracted helper family
- keep oracle/differential coverage in the correctness gates

Exit bar:

- no formula source file over roughly `1500` lines without a clear documented exception
- no giant mixed registry/helper file for lookup or text families
- binder and evaluator policies are testable as separate modules

Verification:

- `pnpm test:correctness:formula`
- focused `vitest` for extracted helper families
- `pnpm exec oxlint --config .oxlintrc.json --type-aware --deny-warnings packages/formula packages/core/src/__tests__/formula-runtime-correctness.test.ts`

Risk notes:

- highest semantic regression risk in the repo
- medium migration risk if helper extraction crosses JS/WASM parity assumptions
- low rollout risk because correctness gates already exist here

### Wave 2: Replace the projected viewport stopgap

Primary owner surface:

- [/Users/gregkonush/github.com/bilig/apps/web](/Users/gregkonush/github.com/bilig/apps/web)

Targets:

- [projected-viewport-store.ts](/Users/gregkonush/github.com/bilig/apps/web/src/projected-viewport-store.ts)
- worker-side viewport and overlay modules in `apps/web/src`

Work:

- move hot visible cell authority further into the worker
- split the current projected store into the components already described in
  [05-06-next-phase.md](/Users/gregkonush/github.com/bilig/docs/05-06-next-phase.md#L1218)
- separate cache storage, damage tracking, and renderer emission
- minimize browser-side cached authority
- keep replay and viewport correctness tests green while shrinking the store

Exit bar:

- `MAX_CACHED_CELLS_PER_SHEET` is gone or reduced to a narrow implementation detail
- the browser no longer acts as a broad local truth store
- reconnect/invalidation behavior is covered by deterministic tests

Verification:

- `pnpm test:correctness:browser`
- focused `vitest` for viewport/runtime reconnect and projected store paths
- targeted browser fuzz replay for any newly discovered viewport bugs

Risk notes:

- medium semantic risk
- high migration risk because authority boundaries are changing
- medium rollout risk because this affects visible browser behavior

### Wave 3: Split workbook state and mutation orchestration

Primary owner surface:

- [/Users/gregkonush/github.com/bilig/packages/core](/Users/gregkonush/github.com/bilig/packages/core)

Targets:

- [workbook-store.ts](/Users/gregkonush/github.com/bilig/packages/core/src/workbook-store.ts)
- [mutation-service.ts](/Users/gregkonush/github.com/bilig/packages/core/src/engine/services/mutation-service.ts)

Work:

- split style/format interning from sheet/cell storage
- split axis mutation and metadata buckets from general workbook storage
- split transaction canonicalization and inverse-op building from execution flow
- preserve snapshot/op/undo correctness gates while extracting services

Exit bar:

- workbook state concerns are decomposed into smaller focused modules
- mutation canonicalization and inverse-op logic are isolated and directly tested

Verification:

- `pnpm test:correctness:core`
- focused `vitest` on workbook-store and mutation-service slices
- targeted `oxlint` and package-scoped `tsc`

Risk notes:

- high semantic risk
- medium migration risk
- low rollout risk because this work sits behind strong core correctness tests

### Wave 4: Decompose worker runtime orchestration

Primary owner surface:

- [/Users/gregkonush/github.com/bilig/apps/web](/Users/gregkonush/github.com/bilig/apps/web)

Target:

- [worker-runtime.ts](/Users/gregkonush/github.com/bilig/apps/web/src/worker-runtime.ts)

Work:

- separate bootstrap/hydration
- separate persistence and mutation journal handling
- separate projection rebuild and authoritative reconciliation
- separate viewport patch publication and runtime state snapshotting
- move orchestration boundaries into concrete `Effect` services where the path crosses IO or lifecycle boundaries

Exit bar:

- worker runtime orchestration is service-based instead of class-centralized
- persistence, sync, and viewport paths can be tested independently

Verification:

- `pnpm test:correctness:browser`
- targeted `vitest` on `worker-runtime`, reconnect, snapshot caches, and runtime-session
- local browser-path replay coverage for persistence and reconcile bugs

Risk notes:

- medium semantic risk
- high migration risk because persistence and sync boundaries are involved
- medium rollout risk

### Wave 5: Grid interaction and rendering split

Primary owner surface:

- [/Users/gregkonush/github.com/bilig/packages/grid](/Users/gregkonush/github.com/bilig/packages/grid)

Targets:

- [useWorkbookGridInteractions.ts](/Users/gregkonush/github.com/bilig/packages/grid/src/useWorkbookGridInteractions.ts)
- [gridGpuScene.ts](/Users/gregkonush/github.com/bilig/packages/grid/src/gridGpuScene.ts)

Work:

- split pointer drag/resize/fill/clipboard/keyboard concerns into separate controllers or hooks
- split scene token definitions from geometry generation
- split border/overlay/selection geometry families into focused modules

Exit bar:

- grid interactions are controller-based rather than one large hook
- GPU scene construction is decomposed by visual family

Verification:

- focused `vitest` on grid interaction and scene modules
- browser correctness smoke coverage for pointer, resize, fill, and keyboard behavior
- type-aware `oxlint` for the affected grid modules

Risk notes:

- medium semantic risk
- medium migration risk
- medium rollout risk because pointer behavior regressions are visible immediately

### Wave 6: Browser sync and E2E debt reduction

Primary owner surfaces:

- [/Users/gregkonush/github.com/bilig/e2e](/Users/gregkonush/github.com/bilig/e2e)
- [/Users/gregkonush/github.com/bilig/apps/web](/Users/gregkonush/github.com/bilig/apps/web)
- [/Users/gregkonush/github.com/bilig/apps/bilig/src](/Users/gregkonush/github.com/bilig/apps/bilig/src)

Target:

- [web-shell.pw.ts](/Users/gregkonush/github.com/bilig/e2e/tests/web-shell.pw.ts)

Work:

- replace skipped Zero-backed flows with deterministic local harnesses where possible
- shrink the file by scenario family
- reserve Playwright for true browser-system behaviors rather than routine semantic coverage

Exit bar:

- the current skipped Zero-backed sync scenarios are removed, replaced, or unskipped
- browser tests are smaller and scenario-focused

Verification:

- `pnpm test:browser`
- deterministic replacement tests must already be green before any skip is removed
- no new `test.skip(...)` on critical sync flows

Risk notes:

- low semantic risk in product code
- medium migration risk in test infrastructure
- medium rollout risk because CI duration and stability can regress if done poorly

### Wave 7: WASM kernel maintainability

Primary owner surfaces:

- [/Users/gregkonush/github.com/bilig/packages/wasm-kernel](/Users/gregkonush/github.com/bilig/packages/wasm-kernel)
- [/Users/gregkonush/github.com/bilig/packages/formula](/Users/gregkonush/github.com/bilig/packages/formula)

Targets:

- [packages/wasm-kernel/assembly](/Users/gregkonush/github.com/bilig/packages/wasm-kernel/assembly)

Work:

- split accelerated builtin families into smaller dispatch modules
- reduce giant kitchen-sink kernel files
- keep JS/WASM differential correctness as the gating truth

Exit bar:

- no single AssemblyScript semantic family remains overly broad without a documented performance reason

Verification:

- `pnpm test:correctness:formula`
- JS/WASM differential suites
- targeted kernel tests under `packages/wasm-kernel/src/__tests__`

Risk notes:

- medium semantic risk
- medium migration risk
- low rollout risk if differential coverage stays green

## Tracking rules

Every wave should be tracked using the same minimum checklist:

- named guarantee or risk being reduced
- target files
- failing or tightened tests first when a bug is present
- verification commands actually run
- landed commit hashes
- residual risk after the slice

If a wave starts producing large commits with mixed concerns, stop and split the
work before continuing.

## Acceptance criteria

This remediation program is complete when:

- the main formula semantic hotspots are decomposed into focused modules
- the projected viewport store is no longer a broad stopgap authority
- workbook state and mutation logic are split into smaller services/modules
- worker runtime orchestration is no longer centralized in one class
- grid interaction/rendering surfaces are decomposed
- skipped Zero-backed browser sync flows are gone or replaced
- WASM kernel modules are split by semantic family rather than convenience

## Recommended first slice

The next highest-value slice is:

1. start with [lookup.ts](/Users/gregkonush/github.com/bilig/packages/formula/src/builtins/lookup.ts)
2. keep TDD strict using existing formula correctness gates
3. extract helper families without changing behavior
4. land small commits while the diff is still easy to review

This is the best first move because it reduces the biggest remaining semantic
blast radius without depending on browser/runtime architecture work.

## First-slice checklist

Use this exact checklist to start Wave 1 safely:

1. inventory the helper families already present in
   [lookup.ts](/Users/gregkonush/github.com/bilig/packages/formula/src/builtins/lookup.ts)
   This should at minimum separate:
   - lookup/search matching
   - database criteria logic
   - statistical helpers
   - matrix helpers
   - builtin registry/export glue
2. add or tighten direct tests for the first extracted helper family before moving code
3. extract one helper family only
4. run:
   - `pnpm test:correctness:formula`
   - focused `vitest` for the touched formula helper tests
   - targeted `oxlint`
5. commit that slice before starting the next helper family

Do not start by moving `lookup.ts` wholesale.
The first slice should prove that the decomposition strategy works with small,
reviewable, behavior-preserving commits.
