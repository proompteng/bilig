# WorkPaper Platform Design

This document defines the production headless spreadsheet contract for `bilig` after auditing the local HyperFormula checkout in `/Users/gregkonush/github.com/hyperformula`.

Date: `2026-04-10`

## Status

This document is the canonical design for the headless spreadsheet layer in `bilig`.
It supersedes `docs/hyperformula-headless-api-design.md`.

The follow-on engine program for actually beating HyperFormula on runtime work is tracked in:

- `docs/workpaper-engine-leadership-program.md`

The top-level public interface is:

- `WorkPaper` from `@bilig/headless`

`HeadlessWorkbook` remains exported as a compatibility alias for existing callers, but it is no longer the primary design name.

## Source Corpus

Reviewed HyperFormula checkout:

- `/Users/gregkonush/github.com/hyperformula/package.json`
- `/Users/gregkonush/github.com/hyperformula/README.md`
- `/Users/gregkonush/github.com/hyperformula/src/HyperFormula.ts`
- `/Users/gregkonush/github.com/hyperformula/src/Emitter.ts`
- `/Users/gregkonush/github.com/hyperformula/src/ConfigParams.ts`
- `/Users/gregkonush/github.com/hyperformula/src/errors.ts`
- `/Users/gregkonush/github.com/hyperformula/src/interpreter/FunctionRegistry.ts`
- `/Users/gregkonush/github.com/hyperformula/src/NamedExpressions.ts`
- `/Users/gregkonush/github.com/hyperformula/docs/guide/basic-operations.md`
- `/Users/gregkonush/github.com/hyperformula/docs/guide/arrays.md`
- `/Users/gregkonush/github.com/hyperformula/docs/guide/clipboard-operations.md`
- `/Users/gregkonush/github.com/hyperformula/docs/guide/custom-functions.md`
- `/Users/gregkonush/github.com/hyperformula/docs/guide/dependency-graph.md`
- `/Users/gregkonush/github.com/hyperformula/docs/guide/i18n-features.md`
- `/Users/gregkonush/github.com/hyperformula/docs/guide/known-limitations.md`
- `/Users/gregkonush/github.com/hyperformula/docs/guide/performance.md`
- `/Users/gregkonush/github.com/hyperformula/docs/guide/sorting-data.md`
- `/Users/gregkonush/github.com/hyperformula/docs/guide/undo-redo.md`
- `/Users/gregkonush/github.com/hyperformula/docs/guide/compatibility-with-microsoft-excel.md`

Reviewed `bilig` surfaces:

- `packages/headless/src/headless-workbook.ts`
- `packages/headless/src/types.ts`
- `packages/headless/src/work-paper.ts`
- `packages/core/src/engine.ts`
- `packages/core/src/__tests__/engine.test.ts`
- `packages/formula/src/*`
- `docs/public-api.md`

HyperFormula checkout facts used here:

- version: `3.2.0`
- commit: `6de904b8876f920f287b63a95934c479acf78307`

## HyperFormula Audit Summary

From the local checkout, HyperFormula is known for:

- a headless spreadsheet runtime that assumes no UI
- workbook factories from empty state, arrays, and sheet maps
- CRUD operations over cells, rows, columns, and sheets
- clipboard operations
- undo/redo
- named expressions
- custom functions
- dependency graph inspection
- sorting and index-reordering flows
- localization and 17 built-in languages
- performance controls such as batching, evaluation suspension, address mapping policy, and column indexing

The local HyperFormula docs also list important limitations:

- no multiple workbooks inside one instance
- no 3D references
- no constant arrays
- no dynamic arrays
- no asynchronous functions
- no structured references or tables
- no relative named expressions
- no UI-metadata-aware function behavior
- custom functions do not automatically resize result arrays as dependencies change

## Measured Parity Audit Against `bilig`

The current `bilig` headless class already matches HyperFormula much more closely than its naming suggests.

Measured from the local source trees:

- public class method inventory: `132 / 132` HyperFormula method names are present on `HeadlessWorkbook`
- additional `bilig` method beyond HyperFormula: `dispose`
- config inventory: `38 / 38` HyperFormula config keys are present on `HeadlessConfig`

That means the remaining design gap is not the absence of a headless API. The gap is:

- the top-level contract name
- documentation quality
- proof that the parity claim is durable
- reliability and packaging guarantees for external consumers
- making `bilig` strengths explicit where HyperFormula is still limited

## Design Intent

`WorkPaper` is the headless spreadsheet engine contract for `bilig`.

It must satisfy two constraints at the same time:

1. Be directly usable by teams building their own grid, form builder, workflow engine, or server-side computation layer.
2. Outperform HyperFormula on the axes that matter in production: feature coverage, deterministic behavior, visibility into recalculation, npm packaging, and benchmarked hot paths.

`10x better` in this document is not marketing language. It means the package is held to stronger gates than HyperFormula in the areas below:

- broader feature coverage, especially around dynamic arrays and structured references
- fewer known runtime limitations
- stronger packaging and consumer-install guarantees
- better observability and diagnostics
- benchmarked performance targets instead of generic speed claims
- clearer migration and API stability rules

## Evidence Policy

This document is intentionally stricter than a normal design memo.

It separates four classes of statement:

1. audited fact
   - derived from the local HyperFormula checkout or the current `bilig` source tree
2. shipped capability
   - backed by code, tests, CI wiring, or a checked-in artifact in this repo
3. engineering target
   - an intended outcome with a defined verification path, not yet proven
4. deliberate non-goal
   - explicitly out of scope until a later semantics design exists

Rules:

- method-count and config-key parity may be stated as current fact because they are snapshot-tested
- support beyond HyperFormula known limitations may be stated only when there is code and test evidence in `bilig`
- a `10x better` claim is never allowed as a blanket package-level claim
- a `10x` claim is allowed only for a named workload, against a named HyperFormula version, with a checked-in artifact and exact measurement method
- if a comparison is not apples-to-apples, this document must say so explicitly instead of presenting a numeric win

## Current Proof State

As of this revision, the following are already proved in-repo:

- `WorkPaper` is the canonical public API and `HeadlessWorkbook` remains a compatibility alias
- HyperFormula public method/category and config-key parity by surface name is checked against the local checkout snapshot
- external consumers can install the runtime tarballs into clean Node and Vite projects without monorepo context
- the runtime package workflow verifies publishability, smoke installs, parity tests, and benchmark-baseline shape
- a checked-in competitive benchmark artifact compares directly comparable workloads against HyperFormula `3.2.0` and labels leadership workloads unsupported where apples-to-apples timing is invalid
- rebuild semantics, deterministic change ordering, adapter immutability, and documentation example usage all have direct tests

The following are not yet proved and therefore must not be claimed as current fact:

- that WorkPaper is `10x` faster than HyperFormula in general
- that WorkPaper is better on every axis simultaneously in one aggregate sense
- that deferred feature families such as async custom functions and relative named expressions are solved

The current state is therefore:

- production-grade headless contract: shipped
- HyperFormula surface parity by public API shape: shipped
- stronger feature coverage in selected areas such as dynamic arrays and structured references: shipped
- workload-specific competitive benchmark evidence against HyperFormula: shipped
- on the checked-in comparable microbenchmarks, HyperFormula currently wins raw throughput on this host
- blanket `10x` superiority claim: still disallowed without a named workload reference

## WorkPaper Contract

Canonical public import:

```ts
import { WorkPaper } from "@bilig/headless";
```

Compatibility import retained:

```ts
import { HeadlessWorkbook } from "@bilig/headless";
```

Compatibility rule:

- `WorkPaper` and `HeadlessWorkbook` refer to the same runtime class in v1
- new documentation, examples, and integration guidance use `WorkPaper`
- `HeadlessWorkbook` stays supported until a future major version removes or de-emphasizes it

Primary branded types:

- `WorkPaperConfig`
- `WorkPaperSheet`
- `WorkPaperSheets`
- `WorkPaperCellAddress`
- `WorkPaperCellRange`
- `WorkPaperChange`
- `WorkPaperNamedExpression`
- `WorkPaperFunctionPluginDefinition`
- `WorkPaperLanguagePackage`
- `WorkPaperEventName`
- `WorkPaperEventMap`
- `WorkPaperDetailedEventMap`
- `WorkPaperInternals`

## Feature Matrix

| Area                             | HyperFormula signal from checkout | WorkPaper target                                                  | Current `bilig` position            |
| -------------------------------- | --------------------------------- | ----------------------------------------------------------------- | ----------------------------------- |
| Headless factories               | README, `HyperFormula.ts`         | `WorkPaper.buildEmpty/buildFromArray/buildFromSheets`             | implemented                         |
| Reads and writes                 | basic operations guide            | full workbook CRUD surface                                        | implemented                         |
| Clipboard                        | clipboard guide                   | copy, cut, paste, fill-range helpers                              | implemented                         |
| Undo/redo                        | undo-redo guide                   | history with explicit clear/reset helpers                         | implemented                         |
| Named expressions                | named expressions guide           | workbook and scoped named expressions                             | implemented                         |
| Custom functions                 | custom functions guide            | registry, translations, instance use                              | implemented                         |
| Localization                     | i18n guide                        | static language registration and translation lookup               | implemented                         |
| Dependency graph                 | dependency graph guide            | stable graph and adapter surface                                  | implemented                         |
| Sorting and reordering           | sorting guide                     | row and column order/swap/move flows                              | implemented                         |
| Batch and suspend/resume         | performance guide                 | grouped recalc control and stable change returns                  | implemented                         |
| Dynamic arrays                   | HyperFormula known limitation     | first-class spill support                                         | implemented in `bilig`              |
| Structured references and tables | HyperFormula known limitation     | supported formulas and sheet semantics                            | implemented in `bilig` core/formula |
| Relative named expressions       | HyperFormula known limitation     | explicit validation and future scoped support path                | currently rejected deliberately     |
| Async functions                  | HyperFormula known limitation     | out of scope for v1, design for later                             | not implemented                     |
| Multiple workbooks               | HyperFormula known limitation     | multiple `WorkPaper` instances per process                        | implemented                         |
| UI metadata-aware functions      | HyperFormula known limitation     | keep out of the core package until UI metadata contract is stable | not implemented                     |

## Architecture

`WorkPaper` stays as a facade package rather than becoming a second engine.

Layers:

1. `@bilig/protocol`
   - value tags, workbook snapshots, literal cell payloads
2. `@bilig/formula`
   - parsing, formatting, translation, formula helpers
3. `@bilig/core`
   - authoritative `SpreadsheetEngine`, dependency graph, history, spill behavior, tables
4. `@bilig/headless`
   - `WorkPaper` API, compatibility aliases, event model, adapters, npm-facing contract

This keeps engine semantics in one place and avoids a second spreadsheet runtime that can drift.

## Reliability Standard

WorkPaper is only considered better than HyperFormula if it is easier to trust in production.

Required reliability properties:

- deterministic rebuild: `updateConfig()` and `rebuildAndRecalculate()` preserve workbook semantics
- stable change ordering: returned `WorkPaperChange[]` is deterministic
- no `workspace:*` leakage in published manifests
- fresh external project install succeeds without monorepo context
- Node.js server usage and browser bundling both work
- public adapters do not leak mutable internals
- documentation examples are verified by tests or typechecked snippets

Required quality gates:

- public API parity test for method inventory against the local HyperFormula checkout
- config-key parity test against the local HyperFormula checkout
- regression tests for clipboard, history, named expressions, rebuilds, adapters, and WorkPaper aliasing
- packaging smoke test in a clean external project

Shipped enforcement for this repo:

- generated snapshot: `packages/headless/src/__tests__/fixtures/hyperformula-surface.json`
- local regeneration/check command: `bun scripts/gen-workpaper-hyperformula-audit.ts` and `bun scripts/gen-workpaper-hyperformula-audit.ts --check`
- CI-safe parity test: `packages/headless/src/__tests__/hyperformula-surface-parity.test.ts`
- external-consumer smoke command: `pnpm workpaper:smoke:external`
- runtime-package CI workflow also verifies clean Node and Vite consumers from packed tarballs

Reliability claims earned today:

- deterministic rebuild semantics are covered by direct tests
- stable change ordering is covered by direct tests
- adapter immutability is covered by direct tests
- published usage examples are covered by direct tests

## Performance Standard

HyperFormula documents batching, evaluation suspension, address mapping policy, and column indexing as performance levers. WorkPaper keeps those controls and raises the bar.

Performance requirements:

- no regression relative to current `HeadlessWorkbook` behavior on read, write, batch, and rebuild paths
- benchmark suites for:
  - workbook build from sheets
  - single-cell write with recalculation
  - batched writes
  - range reads
  - lookup-heavy workloads with `useColumnIndex`
  - dynamic-array recalculation
- use the `@bilig/wasm-kernel` fast path when a formula family has a proven numeric hot spot and JS parity is already green
- publish benchmark artifacts or checked-in benchmark baselines before claiming any `10x` improvement

Shipped benchmark harness:

- package benchmark entrypoint: `packages/benchmarks/src/benchmark-workpaper.ts`
- repo command: `pnpm bench:workpaper`
- checked-in baseline artifact: `packages/benchmarks/baselines/workpaper-baseline.json`
- regeneration/check commands:
  - `pnpm workpaper:bench:generate`
  - `pnpm workpaper:bench:check`
- covered scenarios:
  - workbook build from sheets
  - single-cell edit with downstream recalculation
  - batched edits
  - dense range reads
  - lookup workload with `useColumnIndex`
  - dynamic-array recalculation

WorkPaper is allowed to claim a `10x` win only when a benchmark and fixture suite proves it for a defined workload. Until then, the doc treats `10x better` as an engineering target, not a blanket claim.

## Competitive Benchmark Program

This proof program is now implemented in-repo.

The comparison target must be explicit:

- HyperFormula version: `3.2.0`
- HyperFormula commit: `6de904b8876f920f287b63a95934c479acf78307`
- Node version, CPU architecture, OS, and benchmark fixture sizes must be recorded in the artifact

Workloads must be separated into three categories:

1. directly comparable workloads
   - workbook build from arrays or named sheets
   - single-cell edit with downstream recalculation
   - batched edits
   - range reads
   - lookup-heavy workloads with and without column indexing where both engines support the scenario
2. leadership workloads
   - dynamic arrays
   - structured references and tables
   - multiple workbook instances per process
   - these may demonstrate capability leadership, but not apples-to-apples speed comparisons
3. non-comparable workloads
   - any workload where one engine lacks feature support or uses materially different semantics
   - these must be labeled `unsupported`, not silently omitted

Benchmark reporting requirements:

- default artifact path: `packages/benchmarks/baselines/workpaper-vs-hyperformula-expanded.json`
- control-suite artifact path: `packages/benchmarks/baselines/workpaper-vs-hyperformula.json`
- regeneration/check commands:
  - `pnpm workpaper:bench:competitive:generate`
  - `pnpm workpaper:bench:competitive:check`
  - `pnpm workpaper:bench:competitive:control:generate`
  - `pnpm workpaper:bench:competitive:control:check`
- every comparable workload must include:
  - exact fixture definition
  - warmup count and sample count
  - elapsed metrics
  - memory metrics when available
  - engine version metadata
- any `10x` claim must reference:
  - the exact workload name
  - the exact artifact path
  - the exact measured ratio

Because the expanded default artifact now exists, the correct statement is:

- WorkPaper has stronger feature coverage and stronger release/reliability gates than HyperFormula
- WorkPaper has a checked-in cross-engine benchmark artifact for named workloads
- the current checked-in artifact does not prove general speed leadership for WorkPaper
- WorkPaper still must not claim a blanket `10x` win without citing a specific workload ratio from that artifact

## What Exceeds HyperFormula Today

From the local repo state, `bilig` already exceeds the limitations listed in HyperFormula's own docs in these areas:

- dynamic arrays and spill-aware formulas
- structured references and table-aware formula evaluation
- multiple independent workbook instances in one process
- richer detailed events in addition to HyperFormula-style positional listeners
- npm packaging under MIT-licensed packages rather than GPL-only core distribution

## Remaining Deliberate Gaps

These are not hidden gaps. They are deliberate scope decisions:

- 3D references are still out of scope
- async custom functions are out of scope for v1
- UI-metadata-aware functions remain outside the headless package until a stable metadata contract exists
- relative named expressions continue to reject relative references until a complete semantics model is specified

Deferred-feature exit criteria:

| Deferred area | Why it is deferred | What must exist before implementation is considered correct |
| --- | --- | --- |
| Async custom functions | changes evaluation determinism, batching, and event timing | async execution model, cancellation semantics, timeout/error policy, deterministic event contract |
| Relative named expressions | requires clear scope and rebasing semantics | formal semantics for insert/delete/move/reorder operations plus round-trip tests |
| UI-metadata-aware functions | would leak product/UI concerns into the engine | stable metadata contract owned outside `@bilig/headless`, adapter boundary, and isolation tests |
| 3D references | expands parser, dependency graph, and structure-change semantics | workbook-wide reference model, dependency fanout rules, and Excel-compat correctness cases |

## Execution Plan

### Phase 1: Contract stabilization

- make `WorkPaper` the canonical public entrypoint
- retain `HeadlessWorkbook` as a compatibility alias
- publish WorkPaper-branded type aliases for external consumers
- update package docs and public API docs to center `WorkPaper`

### Phase 2: Proof of parity

- keep the parity suite aligned with the local HyperFormula checkout
- test WorkPaper alias behavior directly
- verify dynamic-array workflows through the `WorkPaper` entrypoint

### Phase 3: Reliability hardening

- keep packaging checks in CI
- keep clean external-consumer smoke tests in release verification
- prevent undocumented API drift by checking the public method inventory

### Phase 4: Performance leadership

- add reproducible benchmarks
- close verified hotspots with WASM-backed kernels where justified
- document benchmark deltas instead of making unmeasured speed claims

## Executed In This Tranche

This design is not speculative. The repo work paired with it does the following:

- introduces `WorkPaper` as the canonical `@bilig/headless` entrypoint
- adds WorkPaper-branded public type aliases
- updates package and public API docs to use `WorkPaper`
- adds tests that exercise real headless workflows through `WorkPaper`
- generates and checks a HyperFormula surface snapshot from the local checkout
- adds a CI-safe parity test against that generated snapshot
- adds a WorkPaper benchmark suite and repo command
- adds a WorkPaper-vs-HyperFormula benchmark suite and checked-in artifact
- keeps `HeadlessWorkbook` working for compatibility

## Acceptance Criteria

This design is complete only when all of the following are true:

- `WorkPaper` is the documented primary API
- every HyperFormula public class method category remains covered in `@bilig/headless`
- every HyperFormula config key remains covered in `WorkPaperConfig`
- external consumers can install and use `@bilig/headless` without monorepo context
- the WorkPaper tests and parity suite pass
- the docs do not claim unmeasured performance wins

Completion states:

- platform-complete
  - all acceptance criteria above are true
- competitively-proved
  - platform-complete is true
  - `packages/benchmarks/baselines/workpaper-vs-hyperformula-expanded.json` exists
  - any comparative superiority claims point to workload-specific evidence in that artifact

## Bottom Line

The correct execution path is not to clone HyperFormula's internals. It is to keep `bilig`'s stronger engine semantics, expose them under the clearer `WorkPaper` contract, prove parity against the HyperFormula surface that users actually know, and raise the release bar on correctness, packaging, and measured performance.
