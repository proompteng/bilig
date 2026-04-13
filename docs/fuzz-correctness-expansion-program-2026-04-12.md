# Fuzz Correctness Expansion Program
## Date: 2026-04-12
## Status: implemented

## Completed implementation

The repo now implements this program with the following committed surfaces:

- engine fuzz suites
  - `packages/core/src/__tests__/engine-history.fuzz.test.ts`
  - `packages/core/src/__tests__/engine-structure.fuzz.test.ts`
  - `packages/core/src/__tests__/engine-replica.fuzz.test.ts`
  - `packages/core/src/__tests__/engine-metadata.fuzz.test.ts`
  - `packages/core/src/__tests__/engine-snapshot.fuzz.test.ts`
  - `packages/core/src/__tests__/engine-replay-fixtures.test.ts`
  - `packages/core/src/__tests__/formula-runtime-differential.fuzz.test.ts`
- formula fuzz suites
  - `packages/formula/src/__tests__/formula-parse.fuzz.test.ts`
  - `packages/formula/src/__tests__/formula-translation.fuzz.test.ts`
  - `packages/formula/src/__tests__/formula-rename.fuzz.test.ts`
  - `packages/formula/src/__tests__/formula-evaluation.fuzz.test.ts`
  - `packages/formula/src/__tests__/formula-replay-fixtures.test.ts`
- shared generators and corpora
  - `packages/core/src/__tests__/engine-fuzz-helpers.ts`
  - `packages/core/src/__tests__/engine-fuzz-metadata-helpers.ts`
  - `packages/formula/src/__tests__/formula-fuzz-helpers.ts`
  - `packages/formula/src/__tests__/formula-fuzz-replay-fixtures.ts`
  - `packages/core/src/__tests__/fixtures/fuzz-replays/`
  - `packages/formula/src/__tests__/fixtures/fuzz-replays/`
- deterministic regression coverage for fuzz-found engine bugs
  - `packages/core/src/__tests__/engine-fuzz-regressions.test.ts`

The implementation also fixed the engine bugs the expanded suites exposed:

- structural delete undo now restores rewritten formulas on other sheets, not just the edited sheet
- local range clear/copy/fill/move helpers no longer create no-op history entries
- empty local batches no longer clear redo state
- structural deletes on semantically blank sheets are treated as no-ops
- pivot output cleanup prunes orphaned empty cells
- pivot materialization now uses the JS semantic path until accelerated parity is proven
- snapshot export preserves explicit authored blank cells while excluding uninitialized empties

## Why this document exists

The repo already has a real correctness program:

- [docs/correctness-gates-program-2026-04-09.md](/Users/gregkonush/github.com/bilig/docs/correctness-gates-program-2026-04-09.md)
- [docs/testing-and-benchmarks.md](/Users/gregkonush/github.com/bilig/docs/testing-and-benchmarks.md)

That program is directionally correct, but the fuzzing layer is still too compressed.

Today the repo has strong property coverage in a few places, but too much semantic surface is packed into too few broad fuzz properties. That makes failures harder to isolate, encourages accidental “fix the harness” behavior, and leaves important workbook states under-generated.

This document is the execution plan to expand fuzz coverage without cheating:

- no timeout padding as a substitute for fixing semantics
- no weaker assertions just to get green CI
- no lower run counts on important properties just to hide failures
- no retries or nondeterministic skips

The goal is not “more random tests.” The goal is better semantic coverage with smaller, sharper, deterministic properties and durable replay fixtures.

## Problem statement

Current fuzzing has three structural weaknesses:

1. too many guarantees per property

The current broad suites in:

- [packages/core/src/__tests__/engine.fuzz.test.ts](/Users/gregkonush/github.com/bilig/packages/core/src/__tests__/engine.fuzz.test.ts)
- [packages/formula/src/__tests__/formula.fuzz.test.ts](/Users/gregkonush/github.com/bilig/packages/formula/src/__tests__/formula.fuzz.test.ts)

defend valuable guarantees, but each property still covers too much behavior at once.

2. too much blank-workbook bias

A large share of generated states still begin from nearly empty workbooks. That misses the cases where spreadsheet engines actually break:

- rewritten formulas after structural edits
- sparse style and format metadata
- metadata-only ranges
- named ranges, tables, filters, validations, and protection
- tracked but semantically empty dependency cells

3. replay artifacts are not yet the primary debugging unit

Seeds and paths help, but the ideal artifact is a minimized command sequence tied to a named guarantee. CI failures should be reproducible as compact, committed replays, not only as ad hoc seed values in logs.

## Concrete lessons from recent failures

Recent failures in [packages/core/src/__tests__/engine.fuzz.test.ts](/Users/gregkonush/github.com/bilig/packages/core/src/__tests__/engine.fuzz.test.ts) exposed exactly why the program needs to expand.

### Structural delete plus undo can violate semantic history

Example sequences that failed:

- `formula(A1 = A1+D4) -> deleteRows(2,2) -> undo -> undo`
- `format(A1 = 0.00) -> deleteColumns(0,1) -> undo -> undo`
- `insertRows(0,1) -> deleteColumns(0,1) -> undo -> undo`

These failures showed that:

- structural transforms can be information-losing
- inverse transactions must capture full pre-delete semantic sheet state, not just deleted-band cells
- tracked placeholder cells must not be restored as if they were real workbook content
- formula rewrite correctness depends on parser support for workbook error literals like `#REF!`

That class of bug is exactly what fuzzing should find early and repeatedly.

## Goals

### Primary goal

Increase semantic correctness coverage for engine and formula behavior by splitting broad fuzz suites into smaller properties with stronger state generation and durable replay artifacts.

### Secondary goals

- reduce time-to-root-cause when CI fuzz fails
- make fuzz failures map to one named guarantee
- turn real production and CI counterexamples into permanent replay tests
- increase confidence in structural edit, history, metadata, and formula rewrite semantics

## Non-goals

- proving every subsystem correct via fuzzing alone
- replacing deterministic regression tests with fuzz tests
- using browser fuzzing as the first line of defense for core engine semantics
- increasing CI runtime without increasing correctness signal
- masking expensive properties with blanket timeout inflation

## Principles

### 1. One property, one guarantee

If a property failure could plausibly come from three different subsystems, the property is too broad.

### 2. Real workbook states beat synthetic emptiness

Generated states must cover realistic workbook shapes, not just empty-sheet mutation streams.

### 3. Counterexamples must become assets

Every minimized counterexample should be promotable into a permanent replay fixture.

### 4. CI tiers must reflect semantic value

- `default` should stay fast and sharp
- `main` should cover the highest-value semantic properties
- `nightly` should hold the expansive or expensive properties

### 5. No cheating

If a property is flaky or too expensive:

- split it
- move it to the right tier
- or improve the generator

Do not hide the problem with weaker assertions, retries, or broad timeout padding.

## Program structure

## Phase 1: Split broad fuzz suites into domain suites

### Engine

Replace the current “fat” engine fuzz file with smaller suites such as:

- `packages/core/src/__tests__/engine-history.fuzz.test.ts`
  - guarantee: model-based undo/redo semantics align with replayed workbook state
- `packages/core/src/__tests__/engine-structure.fuzz.test.ts`
  - guarantee: structural row/column operations preserve semantic rewrite correctness and inverse replay
- `packages/core/src/__tests__/engine-replica.fuzz.test.ts`
  - guarantee: local batches replay into replica-equivalent workbook state
- `packages/core/src/__tests__/engine-replay-fixtures.test.ts`
  - guarantee: committed minimized replay fixtures stay green over time
- `packages/core/src/__tests__/formula-runtime-differential.fuzz.test.ts`
  - guarantee: fast-path runtime evaluation stays aligned with recalculation and snapshot restore

### Formula

Split formula fuzzing into suites such as:

- `packages/formula/src/__tests__/formula-parse.fuzz.test.ts`
  - guarantee: parse and serialize canonicalization remains stable
- `packages/formula/src/__tests__/formula-translation.fuzz.test.ts`
  - guarantee: reference translation and structural transforms roundtrip when they should
- `packages/formula/src/__tests__/formula-evaluation.fuzz.test.ts`
  - guarantee: evaluation stays stable across canonicalization and equivalent rewrites
- `packages/core/src/__tests__/formula-runtime-differential.fuzz.test.ts`
  - guarantee: engine-level accelerated runtime paths agree with recalculation where both are valid

## Phase 2: Add seeded workbook-state generators

Create reusable workbook-state builders for fuzz entrypoints.

Minimum required seed families:

- blank workbook
- formula graph workbook
- sparse style/format workbook
- table/named-range workbook
- validation/filter/sort workbook
- structural metadata workbook

These should live close to the fuzz suites or under a shared helper module, for example:

- `packages/core/src/__tests__/engine-fuzz-helpers.ts`
- `packages/formula/src/__tests__/formula-fuzz-helpers.ts`

The important point is that fuzz does not always start from zero.

## Phase 3: Add subsystem-specific state models

The current model-based history test is valuable, but we need distinct models for distinct truths.

Required models:

- history model
  - tracks applied and undone operations
  - asserts semantic snapshot parity after each step
- structural model
  - tracks row/column axis expectations and address rewrites
  - asserts inverse replay parity after insert/delete/move
- metadata model
  - tracks style and format ranges independently from concrete cell values
  - asserts no orphaned or placeholder state survives replay
- formula rewrite model
  - tracks rewritten formulas and expected error-literal survival

## Phase 4: Promote minimized counterexamples into committed replay corpora

The harness in [packages/test-fuzz/src/index.ts](/Users/gregkonush/github.com/bilig/packages/test-fuzz/src/index.ts) already captures artifacts, but the repo should formalize committed corpora for high-value failures.

Recommended structure:

- `packages/core/src/__tests__/fixtures/fuzz-replays/`

Each replay fixture should contain:

- suite name
- guarantee name
- minimized command sequence or value
- expected semantic snapshot or invariant
- origin note if it came from CI or production

## Phase 5: Add differential fuzz where multiple semantic paths exist

### Formula

For formulas supported on more than one runtime path:

- parser/binder/evaluator agreement
- canonicalized source vs original source agreement
- JS plan vs fast path result agreement
- restored workbook vs live workbook agreement after recalculation

### Engine

For engine semantics:

- direct op replay vs snapshot restore parity
- local execution vs replica batch replay parity
- structural inverse replay vs rebuilt workbook parity

## Required new invariants

The following invariants should be expressed explicitly and owned by named suites.

### Engine invariants

- `ops -> snapshot -> restore -> snapshot` preserves semantics
- `ops -> undo all -> initial snapshot`
- `ops -> undo -> redo -> same semantic state`
- structural delete followed by inverse restore preserves formulas, formats, and metadata
- metadata-only ranges survive undo without materializing empty cells
- replica batch replay produces the same workbook as primary execution

### Formula invariants

- `parse -> serialize -> parse -> serialize` is stable
- translation reversal preserves canonical source when mathematically reversible
- rename roundtrip preserves canonical source
- structural rewrites preserve explicit error literals like `#REF!`
- evaluation of equivalent canonical forms is equal
- JS and accelerated execution agree when both are supported

## Generator bias requirements

Uniform random generation is not enough. The generators must intentionally overweight dangerous transitions.

Required bias:

- formulas that reference cells near structural boundaries
- formulas that rewrite into `#REF!`
- sparse metadata-only style and format ranges
- insert/delete at `0`, `last`, and overlapping boundaries
- undo/redo churn directly after structural edits
- range operations over partially empty regions
- row/column operations combined with formulas and formats in the same turn

## CI tiering

### Default

Purpose:

- fast semantic smoke

Allowed properties:

- a small number of high-signal, low-runtime properties

Must include:

- one engine history property
- one engine structural property
- one formula parse/canonicalization property
- one formula structural rewrite property

### Main

Purpose:

- authoritative correctness gate for PRs and branch pushes

Must include:

- engine history suite
- engine structural suite
- engine metadata suite
- engine replica parity suite
- formula canonicalization suite
- formula translation suite
- formula rename suite
- formula evaluation stability suite

### Nightly

Purpose:

- exhaustive search and corpus growth

Must include:

- long-budget versions of the `main` suites
- broader workbook-state generators
- high-cost differentials
- newly captured replay fixtures before promotion to `main`

## Acceptance criteria

This program is successful when all of the following are true:

- engine and formula fuzzing are split into smaller domain suites
- every suite defends one named guarantee
- workbook-state generators cover more than blank workbooks
- replay artifacts are promotable into committed fixtures
- at least one differential fuzz suite exists for formula runtime agreement
- recent structural history regressions remain permanently defended
- CI no longer depends on oversized all-in-one fuzz properties for core confidence

## First implementation tranche

The first tranche should do only the highest-value work:

1. split `engine.fuzz.test.ts` into history, structure, and replica suites
2. split `formula.fuzz.test.ts` into parse, translation, and evaluation suites
3. add workbook seed generators for blank, formula, and sparse-format states
4. add committed replay fixtures for the recent structural undo regressions
5. add one formula differential fuzz suite

That tranche is enough to materially improve correctness coverage without turning the repo into a test swamp.

## What success looks like

When a fuzz failure happens, the output should immediately answer:

- which guarantee failed
- which subsystem owns it
- which minimal command stream or value triggered it
- whether the counterexample should be promoted into the permanent replay corpus

That is the standard. Not “we ran more random stuff,” but “we increased semantic confidence and made failures cheaper to understand.”
