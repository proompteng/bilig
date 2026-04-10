# Correctness Gates Program
## Date: 2026-04-09
## Status: implemented

## Why this document exists

The repository now has better hygiene gates:

- dead-code detection
- unused export detection
- dependency cycle checks
- stricter type and lint enforcement

That work reduces drift and accidental complexity, but it does not prove that the product is correct.

The next quality step is not "more tests" in the abstract. It is a repo-wide correctness program with explicit product guarantees, deterministic harnesses, replay coverage for real bugs, and CI gates that fail when the runtime violates those guarantees.

This document is the execution plan for that program.

It is intentionally aligned with:

- [docs/reliability/00-program.md](/Users/gregkonush/github.com/bilig/docs/reliability/00-program.md)
- [docs/spreadsheet-engine-effect-service-refactor-2026-04-08.md](/Users/gregkonush/github.com/bilig/docs/spreadsheet-engine-effect-service-refactor-2026-04-08.md)
- [docs/production-stability-remediation-2026-04-02.md](/Users/gregkonush/github.com/bilig/docs/production-stability-remediation-2026-04-02.md)

## Problem statement

Today the repo has a quality gap between "code shape looks healthier" and "the product is reliably correct."

Examples:

- large logic surfaces still exist in formula and evaluator code:
  - [lookup.ts](/Users/gregkonush/github.com/bilig/packages/formula/src/builtins/lookup.ts)
  - [text.ts](/Users/gregkonush/github.com/bilig/packages/formula/src/builtins/text.ts)
  - [js-evaluator.ts](/Users/gregkonush/github.com/bilig/packages/formula/src/js-evaluator.ts)
  - [binder.ts](/Users/gregkonush/github.com/bilig/packages/formula/src/binder.ts)
- the browser still contains a stopgap projected viewport authority in [projected-viewport-store.ts](/Users/gregkonush/github.com/bilig/apps/web/src/projected-viewport-store.ts)
- the main browser E2E file is still large and historically contains skipped or fragile sync flows in [web-shell.pw.ts](/Users/gregkonush/github.com/bilig/e2e/tests/web-shell.pw.ts)
- production incident history already shows that runtime boundaries can be logically correct in local happy paths yet still fail under memory or event-volume pressure

The repo needs correctness gates in four dimensions:

1. logical correctness
2. contract correctness
3. resource correctness
4. behavioral correctness under integration and concurrency

## Goal

Create a production-grade correctness program that makes the following true:

- every major subsystem has explicit invariants
- every known production bug or regression gets a replay test
- every high-risk boundary has a deterministic integration harness
- the highest-value flows have differential or oracle checks where available
- CI runs correctness gates at fast, medium, and deep tiers
- flaky or skipped tests are treated as defects, not tolerated noise

## Non-goals

- pretending static analysis alone is enough
- adding broad test volume without defined guarantees
- wrapping pure spreadsheet math in Effect for style
- relying on browser-only tests for core semantic coverage
- claiming "there are no bugs left"
- keeping flaky tests around indefinitely under `skip` or retry-only policies

## Principles

### 1. Guarantees first, tests second

Every new test must defend a named guarantee. We do not add tests just to increase counts.

### 2. Replay every real bug

Every incident, regression, or customer-visible spreadsheet correctness bug must become a permanent replay test in the nearest valid layer.

### 3. Deterministic harnesses beat flaky end-to-end coverage

Critical browser and sync behavior should be covered by deterministic local harnesses before it is covered by slow or timing-sensitive E2E.

### 4. Effect belongs at boundaries

Use `Effect` for:

- clock
- randomness and id generation
- storage and db access
- network and process access
- worker messaging
- sync transport
- resource lifecycle
- retries, timeouts, and shutdown

Do not wrap pure formula, dependency, parsing, or recalc kernels in `Effect` unless they cross one of those boundaries.

### 5. TDD is mandatory on bug-positive slices

For correctness work:

1. add the failing replay, invariant, or property test first
2. make the smallest fix or refactor that makes it pass
3. keep the new guarantee in CI

### 6. Resource failures count as correctness failures

If a flow leaks memory, blocks the event loop, or regresses large-sheet latency past budget, it is not "correct enough."

## Product guarantees

These are the first guarantees that must be formalized and gated.

### Core engine guarantees

Target package:

- [packages/core](/Users/gregkonush/github.com/bilig/packages/core)

Required guarantees:

- snapshot round-trip parity
  - `importSnapshot(exportSnapshot(x))` preserves workbook meaning
- recalc determinism
  - the same workbook and volatile context produce the same outputs
- dependency stability
  - rebuilding dependency state does not change computed meaning
- structural rewrite correctness
  - insert/delete/move operations rewrite references consistently
- undo/redo reversibility
  - `redo(undo(x))` returns to the prior semantic state
- op replay parity
  - applying the same op stream from the same base produces the same workbook

### Formula runtime guarantees

Target package:

- [packages/formula](/Users/gregkonush/github.com/bilig/packages/formula)

Required guarantees:

- builtin oracle parity where a stable oracle exists
- error propagation correctness
- coercion and edge-case stability
- binder and evaluator agreement
- JS and WASM parity for accelerated formula families

Highest-risk targets:

- [lookup.ts](/Users/gregkonush/github.com/bilig/packages/formula/src/builtins/lookup.ts)
- [text.ts](/Users/gregkonush/github.com/bilig/packages/formula/src/builtins/text.ts)
- [js-evaluator.ts](/Users/gregkonush/github.com/bilig/packages/formula/src/js-evaluator.ts)
- [binder.ts](/Users/gregkonush/github.com/bilig/packages/formula/src/binder.ts)

### Sync and projection guarantees

Target surfaces:

- [apps/bilig/src/zero](/Users/gregkonush/github.com/bilig/apps/bilig/src/zero)
- [packages/zero-sync](/Users/gregkonush/github.com/bilig/packages/zero-sync)

Required guarantees:

- mutation idempotence where designed
- event ordering and causality
- projection parity between engine state and persisted rows
- migration replay correctness
- snapshot and replica state compatibility

### Browser/runtime guarantees

Target surfaces:

- [apps/web](/Users/gregkonush/github.com/bilig/apps/web)
- [packages/grid](/Users/gregkonush/github.com/bilig/packages/grid)

Required guarantees:

- worker/runtime/UI state parity
- authoritative data path correctness
- viewport subscription correctness
- multiplayer shell consistency under reconnect and invalidation
- no visible split-brain between local caches and authoritative state

Special attention:

- [projected-viewport-store.ts](/Users/gregkonush/github.com/bilig/apps/web/src/projected-viewport-store.ts) remains a correctness risk because it is still a temporary local authority with cache limits and pruning behavior

## Required test categories

### 1. Replay tests

Every regression gets a minimal fixture:

- workbook snapshot
- operation sequence
- expected output or invariant

Replay tests belong in the narrowest layer that can reproduce the bug:

- `packages/core` for semantic engine bugs
- `packages/formula` for evaluation bugs
- `apps/bilig` for projection, migration, or sync/server bugs
- `apps/web` only when the failure is inherently browser-specific

### 2. Property-based tests

Use `fast-check` to generate:

- edit sequences
- structural mutation sequences
- range selections
- formula inputs with coercion edge cases
- sync event sequences

Property-based tests must assert invariants, not snapshots of incidental structure.

Examples:

- edit then undo returns semantic equivalence
- insert row then export/import preserves formulas
- range formatting ops preserve address normalization
- binder plus evaluator yields the same visible value as direct formula execution

### 3. Differential tests

Use when the repo has two valid implementations or an external oracle:

- JS evaluator vs WASM kernel
- engine state vs persisted projection rows
- captured Excel or fixture results vs `@bilig/formula`

### 4. Deterministic integration tests

Build harnesses that stub only the boundary, not the logic:

- fake clock
- fake id/random source
- in-memory or disposable db
- fake worker transport
- fake websocket/sync channel
- effect-managed resource lifecycle

This is the primary way to replace flaky browser and sync tests.

### 5. Resource and soak tests

Must assert:

- bounded memory growth
- acceptable event-loop latency
- stable large-sheet recalc time
- stable sync fanout behavior
- no unbounded queue growth

These belong in nightly or release-level gates, not only ad hoc benchmarking.

## Harness architecture

### Core and formula

Keep the semantic kernels pure.

Add shared test helpers that can:

- load workbook fixtures
- apply generated or recorded op sequences
- export semantic fingerprints
- compare workbook meaning instead of raw object identity

### Bilig server and sync path

Introduce deterministic boundary services for:

- db access
- clock
- UUID generation
- Zero ingress and egress
- runtime session persistence

Use `Effect` layers so tests can substitute:

- disposable Postgres or SQLite-compatible harnesses when valid
- in-memory transport doubles
- deterministic retry and timeout behavior

### Browser and worker path

Add local deterministic harnesses for:

- worker startup
- runtime refresh
- viewport invalidation
- reconnect sequences
- authoritative projection updates

The goal is to move critical correctness away from browser timing and toward repeatable state-machine tests.

## CI structure

The current pipeline should be extended into three correctness tiers.

### Fast gate

Runs on every PR.

Includes:

- lint
- typecheck
- static analysis
- touched-package unit tests
- touched-package replay tests
- touched-package property tests with bounded seed count

Purpose:

- fast semantic signal for local edits

### Medium gate

Runs on protected branches and high-risk PRs.

Includes:

- deterministic integration suites
- JS/WASM differential tests
- engine/projection parity suites
- sync ordering and replay suites
- targeted browser-state harnesses without remote dependencies

Purpose:

- catch orchestration failures and boundary regressions

### Deep gate

Runs nightly and before release.

Includes:

- long fuzz runs
- resource soak tests
- large-workbook replay suites
- browser multiplayer flows
- production acceptance matrix checks

Purpose:

- catch slow-burn and scale-sensitive regressions

## Flake policy

Flake is a bug.

Rules:

- no new `test.skip` for critical correctness flows without a linked issue and replacement plan
- every flaky test needs:
  - owner
  - defect label
  - target removal date
- retries may hide noise temporarily but do not count as a fix
- if a test is flaky because the harness is poor, replace the harness

The current large browser sync flows in [web-shell.pw.ts](/Users/gregkonush/github.com/bilig/e2e/tests/web-shell.pw.ts) should be narrowed, replaced, or de-flaked rather than left as permanent exceptions.

## Execution waves

### Wave 0: Formalize guarantees and bug intake

Deliverables:

- add this document
- create a correctness issue template for replay tests
- define the first invariant list for each subsystem
- tag known incidents and regressions by subsystem

Exit bar:

- every new bug report can be mapped to a replay target and guarantee

### Wave 1: Core semantic gate

Targets:

- [packages/core/src/engine.ts](/Users/gregkonush/github.com/bilig/packages/core/src/engine.ts)
- [packages/core/src/workbook-store.ts](/Users/gregkonush/github.com/bilig/packages/core/src/workbook-store.ts)

Work:

- add snapshot round-trip tests
- add undo/redo property tests
- add structural rewrite invariant tests
- add op replay parity tests

Exit bar:

- `packages/core` has explicit semantic invariants enforced in CI

### Wave 2: Formula differential gate

Targets:

- [lookup.ts](/Users/gregkonush/github.com/bilig/packages/formula/src/builtins/lookup.ts)
- [text.ts](/Users/gregkonush/github.com/bilig/packages/formula/src/builtins/text.ts)
- [js-evaluator.ts](/Users/gregkonush/github.com/bilig/packages/formula/src/js-evaluator.ts)
- [binder.ts](/Users/gregkonush/github.com/bilig/packages/formula/src/binder.ts)

Work:

- add oracle fixtures for high-risk builtins
- add binder/evaluator agreement tests
- add JS/WASM parity where accelerated families exist
- add coercion/error replay tests from known regressions

Exit bar:

- high-risk formula families have differential or replay coverage

### Wave 3: Server projection and sync parity

Targets:

- [apps/bilig/src/zero](/Users/gregkonush/github.com/bilig/apps/bilig/src/zero)
- [packages/zero-sync](/Users/gregkonush/github.com/bilig/packages/zero-sync)

Work:

- add engine-to-projection parity tests
- add mutation idempotence and ordering tests
- add migration replay suites
- add deterministic Effect-backed integration harnesses for sync/server workflows

Exit bar:

- persisted rows and runtime state are checked for parity under replay

### Wave 4: Browser authoritative-path hardening

Targets:

- [apps/web/src/projected-viewport-store.ts](/Users/gregkonush/github.com/bilig/apps/web/src/projected-viewport-store.ts)
- [packages/grid](/Users/gregkonush/github.com/bilig/packages/grid)
- [e2e/tests/web-shell.pw.ts](/Users/gregkonush/github.com/bilig/e2e/tests/web-shell.pw.ts)

Work:

- add deterministic worker/runtime/UI parity harnesses
- replace skipped sync flows with local deterministic tests where possible
- reduce dependence on the stopgap projected viewport path
- add reconnect and invalidation replay coverage

Exit bar:

- critical browser flows do not depend solely on flaky browser timing

### Wave 5: Resource correctness and release bar

Targets:

- core recalc
- large workbook imports
- sync fanout
- browser shell memory behavior

Work:

- add memory and event-loop budget assertions
- add soak runs and large-fixture regression suites
- gate releases on resource budgets

Exit bar:

- performance regressions and memory leaks fail a defined gate

## TDD workflow

Every correctness fix or refactor slice should follow this order:

1. identify the guarantee being defended
2. write or tighten the failing replay, invariant, or differential test
3. make the smallest change that restores correctness
4. if the code is hard to fix safely, refactor behind the test first
5. keep the new test in the permanent gate for that subsystem

This rule applies especially to `Effect` refactors. We do not move orchestration behind services first and hope correctness follows later.

## Effect usage rules for this program

Use `Effect` to make the following deterministic and injectable:

- `Clock`
- random/id generation
- database access
- transport access
- retry policy
- timeout policy
- shutdown semantics
- metrics and logging sinks

Do not use `Effect` to obscure:

- pure value transforms
- formula parsing and evaluation kernels
- dependency graph math
- address and range rewriting
- workbook state diffing

The right structure is:

- pure compute in plain modules
- side-effect orchestration in concrete Effect services
- deterministic integration tests through test layers

## Definition of done

This program is done only when all of the following are true:

- each major subsystem has a written guarantee set
- every production incident and known regression has a replay test
- critical JS/WASM and engine/projection parity checks are automated
- critical skipped or flaky sync/browser tests are removed or replaced
- correctness gates exist at fast, medium, and deep CI tiers
- resource budgets are enforced for large-workbook and sync-heavy flows

Until then, the repo may be healthier, but it is not yet protected against correctness regressions at the level required for a spreadsheet product.

## First execution targets

If work starts immediately, the first sequence should be:

1. `packages/core`
   - add semantic invariants and property-based replay tests
2. `apps/bilig`
   - add engine/projection parity and sync ordering tests
3. `packages/formula`
   - add differential coverage for lookup, text, evaluator, and binder
4. `apps/web`
   - replace flaky or skipped sync flows with deterministic local harnesses

This order is intentional:

- core defines semantic truth
- server projection must match core truth
- formula correctness is the biggest remaining semantic blast radius
- browser correctness should sit on top of stronger lower-layer guarantees
