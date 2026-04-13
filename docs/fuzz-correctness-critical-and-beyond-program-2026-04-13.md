# Fuzz Correctness Critical And Beyond Program
## Date: 2026-04-13
## Status: implemented

## Why this document exists

The April 12 program is done:

- [/Users/gregkonush/github.com/bilig/docs/fuzz-correctness-expansion-program-2026-04-12.md](</Users/gregkonush/github.com/bilig/docs/fuzz-correctness-expansion-program-2026-04-12.md>)

That work fixed the largest weakness in the repo’s fuzzing story:

- engine and formula semantics were too compressed into too few broad properties

The repo now has real guarantee-owned fuzz suites in:

- `packages/core/src/__tests__/engine-history.fuzz.test.ts`
- `packages/core/src/__tests__/engine-structure.fuzz.test.ts`
- `packages/core/src/__tests__/engine-replica.fuzz.test.ts`
- `packages/core/src/__tests__/engine-metadata.fuzz.test.ts`
- `packages/core/src/__tests__/engine-snapshot.fuzz.test.ts`
- `packages/core/src/__tests__/engine-replay-fixtures.test.ts`
- `packages/core/src/__tests__/formula-runtime-differential.fuzz.test.ts`
- `packages/formula/src/__tests__/formula-parse.fuzz.test.ts`
- `packages/formula/src/__tests__/formula-translation.fuzz.test.ts`
- `packages/formula/src/__tests__/formula-rename.fuzz.test.ts`
- `packages/formula/src/__tests__/formula-evaluation.fuzz.test.ts`
- `packages/formula/src/__tests__/formula-replay-fixtures.test.ts`

That is the right base. It is not the end state.

The next correctness risk is no longer “can the engine and formula layers survive fuzzing at all?” The next risk is that the repo still has critical product boundaries with little or no fuzz pressure:

- server-side sync and projection
- browser projection and viewport state
- import/export semantic parity
- advanced workbook metadata surfaces under structural edits and replay
- corpus automation so CI failures become durable fixtures instead of one-off seeds

This document is the next-phase design. It is intentionally larger than the April 12 program. It is meant to cover the full critical path and then push beyond it.

## Execution completed

This program is implemented in the repo, not left as planning-only text.

The landed coverage includes:

- import/export parity in:
  - `packages/core/src/__tests__/engine-import-export.fuzz.test.ts`
  - `packages/core/src/__tests__/snapshot-wire-parity.fuzz.test.ts`
  - `packages/core/src/__tests__/literal-loader-parity.fuzz.test.ts`
- server sync and projection fuzz in:
  - `apps/bilig/src/zero/__tests__/projection.fuzz.test.ts`
  - `apps/bilig/src/zero/__tests__/migration-parity.fuzz.test.ts`
  - `apps/bilig/src/zero/__tests__/sync-relay.fuzz.test.ts`
  - `apps/bilig/src/zero/__tests__/reconnect-replay.fuzz.test.ts`
- browser projection and runtime fuzz in:
  - `apps/web/src/__tests__/projected-viewport.fuzz.test.ts`
  - `apps/web/src/__tests__/viewport-cache-pruning.fuzz.test.ts`
  - `apps/web/src/__tests__/worker-workbook-app-model.fuzz.test.ts`
  - `apps/web/src/__tests__/runtime-sync.fuzz.test.ts`
  - `apps/web/src/__tests__/selection-command-parity.fuzz.test.ts`
- advanced workbook metadata fuzz in:
  - `packages/core/src/__tests__/engine-charts.fuzz.test.ts`
  - `packages/core/src/__tests__/engine-media.fuzz.test.ts`
  - `packages/core/src/__tests__/engine-annotations.fuzz.test.ts`
  - `packages/core/src/__tests__/engine-conditional-formats.fuzz.test.ts`
  - `packages/core/src/__tests__/engine-protection.fuzz.test.ts`
  - `packages/core/src/__tests__/engine-pivot.fuzz.test.ts`
  - `packages/core/src/__tests__/engine-data-validation.fuzz.test.ts`
- replay corpora and deterministic replays in:
  - `packages/core/src/__tests__/fixtures/fuzz-replays/`
  - `packages/formula/src/__tests__/fixtures/fuzz-replays/`
  - `apps/bilig/src/zero/__tests__/fixtures/fuzz-replays/`
  - `apps/web/src/__tests__/fixtures/fuzz-replays/`
- replay tooling in:
  - `packages/test-fuzz/src/index.ts`
  - `scripts/promote-fuzz-artifact.ts`

The implementation also fixed product bugs exposed while executing this program, including:

- cycle error propagation drift across direct formula writes vs import/restore
- pivot anchor rewrite and owned-output cleanup under structural edits
- stale pivot materialization after edits inside pivot output footprints
- live/browser selection command targeting drift
- projection/runtime parity and metadata roundtrip defects found by the new suites

The repo’s `default`, `main`, and browser fuzz lanes now exercise these workstreams through the existing `pnpm test:fuzz`, `pnpm test:fuzz:main`, and `pnpm test:fuzz:nightly` entrypoints.

## Relationship to existing correctness docs

This document extends, not replaces:

- [/Users/gregkonush/github.com/bilig/docs/correctness-gates-program-2026-04-09.md](</Users/gregkonush/github.com/bilig/docs/correctness-gates-program-2026-04-09.md>)
- [/Users/gregkonush/github.com/bilig/docs/fuzz-correctness-expansion-program-2026-04-12.md](</Users/gregkonush/github.com/bilig/docs/fuzz-correctness-expansion-program-2026-04-12.md>)
- [/Users/gregkonush/github.com/bilig/docs/testing-and-benchmarks.md](</Users/gregkonush/github.com/bilig/docs/testing-and-benchmarks.md>)

The repo already has correctness tiers:

- `pnpm test:correctness:core`
- `pnpm test:correctness:formula`
- `pnpm test:correctness:server`
- `pnpm test:correctness:browser`
- `pnpm test:correctness:fast`
- `pnpm test:correctness:medium`
- `pnpm test:correctness:deep`
- `pnpm test:fuzz`
- `pnpm test:fuzz:main`
- `pnpm test:fuzz:nightly`

This program uses those tiers and expands them. It does not introduce a parallel test universe.

## Problem statement

The current fuzz program is good for the semantic core. It is not yet good enough for the product boundary.

### The critical gap

The most dangerous remaining correctness failures are no longer pure formula parser bugs or isolated undo bugs. They are cross-boundary failures:

- local mutation batch is semantically valid but replays incorrectly through sync
- server projection diverges from engine state
- browser projected viewport diverges from worker/runtime authority
- imported workbook semantics do not survive export or restore
- metadata surfaces such as charts, comments, notes, protections, conditional formats, and pivots drift under structural edits

Those are product failures, not just package failures.

### The current asymmetry

The repo already fuzzes:

- engine semantics
- formula semantics

The repo does not yet fuzz with comparable force:

- `apps/bilig/src/zero/projection.ts`
- `apps/bilig/src/zero/sync-relay.ts`
- `apps/bilig/src/zero/workbook-mutation-store.ts`
- `apps/bilig/src/zero/workbook-migration-store.ts`
- `apps/web/src/projected-viewport-store.ts`
- `apps/web/src/projected-viewport-patch-application.ts`
- `apps/web/src/projected-viewport-axis-store.ts`
- `apps/web/src/projected-viewport-cell-cache.ts`
- `apps/web/src/use-workbook-sync.ts`
- `apps/web/src/worker-runtime-viewport.ts`
- import/export boundaries and browser-visible parity

That asymmetry is the next correctness hole.

## Goal

Build a production-grade fuzz program that covers:

- the full critical spreadsheet product path
- the surrounding advanced workbook domains beyond the critical path
- the CI artifact pipeline needed to make fuzz failures reproducible, reviewable, and permanent

The result should make this statement true:

“Randomized correctness pressure exists at every high-risk semantic boundary where the product can corrupt, diverge, or lie.”

## Non-goals

- adding random tests to every directory just to inflate fuzz counts
- replacing deterministic regression suites with fuzz suites
- pushing browser fuzz into places where a lower deterministic harness is better
- weakening assertions to make long-running properties less painful
- widening timeouts or lowering run counts instead of splitting the property properly
- pretending charts or media are “nice to have” correctness domains when they mutate workbook state

## Principles

### 1. Fuzz from the lowest authoritative layer that can express the guarantee

Prefer:

- engine-level fuzz for semantic workbook truth
- server-level fuzz for projection and sync truth
- browser-level fuzz only for projection/runtime/browser authority truth

Do not start in the browser when the real bug can be expressed lower.

### 2. One suite defends one named product guarantee

The suite name should tell the reviewer exactly what broke:

- history reversibility
- projection parity
- sync relay causality
- viewport patch coherence
- import/export semantic parity

If a suite can fail for five unrelated reasons, it is too broad.

### 3. State generators are first-class product assets

The quality of fuzzing is capped by the quality of state generation.

This program treats seeded workbook builders, sync event generators, viewport patch generators, and import fixture generators as core infrastructure, not test scaffolding.

### 4. Counterexamples must become committed corpora

Seeds are not enough.

The repo should preserve:

- the minimized command or event stream
- the exact guarantee name
- the expected semantic invariant

for both engine/formula and the new server/browser/import layers.

### 5. Critical path first, then beyond-critical domains

The order matters:

1. engine and formula semantics
2. sync and projection boundaries
3. browser projection and runtime coherence
4. import/export parity
5. advanced workbook objects

“Beyond” means we go further than minimum spreadsheet cell correctness, not that we skip the core path.

## Scope map

### Critical now

These domains are product-critical and should be fuzzed before anything decorative:

- engine history and structure
- formula rewrite and evaluation
- local batch replay and replica parity
- server projection and sync relay semantics
- browser projected viewport and worker/runtime authority
- import/export semantic parity

### Critical next

These domains become critical as soon as the base path is stable:

- migrations and persisted snapshot compatibility
- reconnect and replay ordering
- range protections and permission-aware mutation ordering
- workbook metadata under structural edits

### Beyond critical

These still matter and should be included in the program:

- charts
- media
- comments and notes
- conditional formats
- pivots
- validations, filters, sorts, freeze panes
- workbook agent preview/application semantics where workbook mutation is involved

## Program structure

## Workstream 1: Engine and formula hardening

This workstream keeps the April 12 program strong while extending it on the highest-value missing edges.

### Required additions

- richer interleaving in `packages/core/src/__tests__/engine-history.fuzz.test.ts`
  - more undo/redo churn between structural and metadata edits
- explicit no-op and cancel-out history coverage in `packages/core/src/__tests__/engine-correctness.test.ts`
  - history signal must reflect semantic change across the stream, not only final-state difference
- broader workbook seeds in:
  - `packages/core/src/__tests__/engine-fuzz-helpers.ts`
  - `packages/core/src/__tests__/engine-fuzz-metadata-helpers.ts`
  - `packages/formula/src/__tests__/formula-fuzz-helpers.ts`

### New seed families

Add at minimum:

- workbook with multiple sheets plus cross-sheet formulas and named ranges
- workbook with dense fragmented style and format ranges
- workbook with charts, notes, comments, conditional formats, and protections
- workbook with pivots and blocked overwrite cases
- workbook with imported-looking blank/value/formula mixtures

### New guarantees

- history exists iff semantic state changed at some point in the stream
- structural rewrites preserve metadata surfaces, not just cell formulas
- accelerated evaluation and recalculation stay aligned on expanded formula families

## Workstream 2: Import/export semantic parity fuzz

This is a missing critical layer.

### Why

A spreadsheet product is not correct if import/export silently changes workbook meaning.

### Required suites

Add new suites under `packages/core` or the narrowest valid import layer:

- `import-export-parity.fuzz.test.ts`
  - guarantee: import -> export -> import preserves semantic workbook state
- `snapshot-wire-parity.fuzz.test.ts`
  - guarantee: snapshot serialization and deserialization preserve workbook meaning and metadata
- `literal-loader-parity.fuzz.test.ts`
  - guarantee: literal workbook loaders preserve authored blanks, formulas, and metadata identity

### Required generators

- CSV-like sheets with ambiguous numeric/text coercions
- XLSX-like style/format fragments
- formulas with quoted sheet names, errors, names, and structural refs
- sparse metadata-only ranges

### Required invariants

- imported workbook snapshot equals semantic reference snapshot
- roundtrip does not invent explicit cells or drop authored blanks
- structural metadata survives import/export where supported
- unsupported constructs degrade explicitly and deterministically

## Workstream 3: Server sync and projection fuzz

This is the highest-priority expansion after engine/formula.

### Target files

- `apps/bilig/src/zero/projection.ts`
- `apps/bilig/src/zero/sync-relay.ts`
- `apps/bilig/src/zero/workbook-mutation-store.ts`
- `apps/bilig/src/zero/workbook-migration-store.ts`
- `apps/bilig/src/workbook-runtime/browser-sync-replay.test.ts`
- `apps/bilig/src/workbook-runtime/sync-frame-router.test.ts`
- `apps/bilig/src/workbook-runtime/workbook-sync-session-host.test.ts`

### Required suites

- `projection.fuzz.test.ts`
  - guarantee: projected rows remain semantically equivalent to engine state across randomized mutation streams
- `sync-relay.fuzz.test.ts`
  - guarantee: relay ordering, dedupe, and causality preserve semantic state
- `migration-parity.fuzz.test.ts`
  - guarantee: randomized legacy/current snapshots converge to the same projected meaning after migration
- `reconnect-replay.fuzz.test.ts`
  - guarantee: disconnect/reconnect/replay ordering converges without duplicate side effects

### Required generators

- local and remote mutation batch streams
- delayed, duplicated, and reordered delivery
- reconnect points and replay windows
- schema-versioned snapshots and migration steps

### Required invariants

- projection parity with engine snapshot
- idempotent replay where designed
- convergent end state after duplicate or reordered delivery within contract
- migrations preserve workbook meaning

### Required model layer

Add a model-based sync harness:

- real system: server stores + relay + projection
- model: authoritative workbook snapshot plus ordered event ledger

The model should prove semantic convergence, not implementation shape.

## Workstream 4: Browser projection and runtime fuzz

This is the next critical product boundary.

### Target files

- `apps/web/src/projected-viewport-store.ts`
- `apps/web/src/projected-viewport-patch-application.ts`
- `apps/web/src/projected-viewport-axis-store.ts`
- `apps/web/src/projected-viewport-axis-patches.ts`
- `apps/web/src/projected-viewport-cell-cache.ts`
- `apps/web/src/projected-viewport-cache-pruning.ts`
- `apps/web/src/use-workbook-sync.ts`
- `apps/web/src/worker-runtime-viewport.ts`
- `apps/web/src/worker-workbook-app-model.ts`

### Required suites

- `projected-viewport.fuzz.test.ts`
  - guarantee: viewport state matches authoritative worker/runtime state after randomized patch streams
- `viewport-cache-pruning.fuzz.test.ts`
  - guarantee: pruning never drops authoritative visible state or corrupts patch application
- `runtime-sync-fuzz.test.ts`
  - guarantee: worker/runtime/app-shell state stays coherent across reconnect, invalidation, and selection churn
- `selection-command-parity.fuzz.test.ts`
  - guarantee: browser-issued workbook commands target the same semantic ranges the engine sees

### Required generators

- viewport moves and resizes
- row/column axis patch streams
- selection mutations mixed with sync patches
- patch pruning windows
- reconnect and invalidation bursts

### Required invariants

- rendered viewport content equals worker/runtime authoritative projection for visible bounds
- cache pruning never changes visible semantic content
- browser command targeting never uses stale selection state
- projected state converges after reconnect and replay

## Workstream 5: Advanced workbook metadata fuzz

This is “beyond critical” only in scheduling, not in importance.

### Target domains

- charts
- media
- comments
- notes
- conditional formats
- validations
- protections
- pivots
- filters, sorts, freeze panes

### Required suites

Add domain suites under `packages/core/src/__tests__` such as:

- `engine-charts.fuzz.test.ts`
- `engine-media.fuzz.test.ts`
- `engine-annotations.fuzz.test.ts`
- `engine-conditional-formats.fuzz.test.ts`
- `engine-protection.fuzz.test.ts`
- `engine-pivot.fuzz.test.ts`

### Required guarantees

- structural edits preserve or rewrite metadata correctly
- undo/redo preserves metadata identity and semantic scope
- snapshot roundtrip preserves supported metadata exactly
- replay fixtures cover every discovered metadata corruption bug

### Required generators

- metadata-only ranges without cell content
- overlapping metadata surfaces
- blocked operations due to protection or pivot ownership
- edits that create and then cancel metadata changes

## Workstream 6: Replay corpus automation

This workstream is mandatory. Without it, fuzz failures decay into log spam.

### Required repo structure

Extend committed corpora beyond current engine/formula fixture directories:

- `packages/core/src/__tests__/fixtures/fuzz-replays/`
- `packages/formula/src/__tests__/fixtures/fuzz-replays/`
- `apps/bilig/src/**/__tests__/fixtures/fuzz-replays/`
- `apps/web/src/**/__tests__/fixtures/fuzz-replays/`

### Required tooling

Extend `packages/test-fuzz/src/index.ts` so CI artifacts can promote directly into stable replay fixtures with:

- suite name
- guarantee kind
- minimized command or event stream
- seed and path
- reproduction command
- expected invariant or expected final snapshot

### Required policy

Every real CI fuzz failure that exposes a product bug must end as:

1. a product fix
2. a deterministic replay fixture
3. a narrower regression test if the bug is not worth keeping purely in fuzz form

## Workstream 7: CI tier redesign for the expanded program

The repo’s existing tiers are the right skeleton. The contents need to become more deliberate.

### Default

Must stay fast enough for everyday use.

Include:

- highest-signal engine/formula suites
- one server sync/projection suite
- one browser projection suite
- replay corpora tests

### Main

This is the real correctness gate.

Include:

- full engine/formula fuzz set
- server projection and relay fuzz
- browser projection/runtime fuzz
- import/export parity fuzz
- critical metadata domain fuzz

### Nightly

This is where expensive search belongs.

Include:

- long-run sync ordering fuzz
- browser projection churn fuzz with larger state windows
- import/export differential matrices
- broad metadata permutation fuzz
- corpus-mining passes for new replay fixtures

### Rules

- do not widen timeouts to rescue a broad property
- split or tier properties instead
- do not drop run counts to hide failures
- move properties to the correct tier when they are semantically important but expensive

## Implementation order

## Phase A: Critical boundary expansion

1. import/export semantic parity fuzz
2. server sync/projection fuzz
3. browser projected viewport/runtime fuzz

This phase closes the remaining critical gaps.

## Phase B: Metadata domain expansion

4. charts, media, annotations, conditional formats, protections, pivots
5. broaden engine/formula seed families so those metadata domains appear in mixed streams

## Phase C: Corpus automation and nightly depth

6. replay corpus automation in `packages/test-fuzz`
7. nightly fuzz tier growth
8. CI artifact promotion workflow

## Required engineering standards

### No garbage code

- no giant “fuzz helpers” files that hide unrelated generators forever
- split helpers by domain once they approach the same sprawl the old suites had
- no dead generators left around after suite migrations
- no stale replay fixtures for removed guarantees

### No fake fixes

- no assertion weakening unless the previous assertion was logically wrong
- no timeout inflation as a substitute for splitting suites
- no retries
- no test skips for fuzz failures

### TDD rules

For every bug-positive slice:

1. reproduce with seed or minimized stream
2. promote to replay fixture or narrow regression
3. fix product code
4. keep the fuzz pressure that exposed it

## Acceptance criteria

This program is done only when all of the following are true:

- critical fuzz coverage exists for engine, formula, server projection, browser projection, and import/export
- advanced workbook metadata domains have dedicated fuzz suites
- replay corpora exist across engine, formula, server, and browser where meaningful
- CI tiers map cleanly to semantic value
- recent bug classes can be reproduced by committed fixtures, not only ad hoc seeds
- a fuzz failure tells the reviewer which product guarantee broke

## What “cover everything critical and beyond” means

It does not mean literally every file in the repo gets a fuzz test.

It means:

- every critical semantic boundary is defended by guarantee-owned fuzz
- every advanced workbook domain that can corrupt workbook meaning is part of the program
- deterministic replay fixtures preserve discovered failures permanently
- the test architecture stays sharp enough that future fuzz work keeps finding real bugs instead of creating noise

That is the bar worth aiming for.
