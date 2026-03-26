# `bilig` Canonical Product Design

`bilig` is a local-first spreadsheet system with a browser-native Excel shell, a deterministic semantic core, and a WASM execution engine.

## Current state

- `@bilig/core`, `@bilig/formula`, `@bilig/wasm-kernel`, `apps/web`, `apps/local-server`, and `apps/sync-server` exist and are executable
- `@bilig/core` has a transaction-based workbook engine with metadata, row and column structure, freeze panes, filters, sorts, tables, spills, pivots, and undo/redo
- `apps/web` now boots worker-first by default and consumes worker-derived viewport patches
- `apps/local-server` is a live worksheet host for browser and agent sessions
- `apps/sync-server` handles durable ingress and snapshots, but remote worksheet execution is not closed by default
- `@bilig/agent-api` uses binary framing around JSON payloads rather than a typed binary payload schema
- the canonical registry in code currently contains `101` canonical entries:
  - `92` are `implemented-wasm-production`
  - `6` are `implemented-js`
  - `3` are `blocked`
- the `9` non-production canonical rows are:
  - `dynamic-array:filter-basic`
  - `dynamic-array:unique-basic`
  - `lambda:let-basic`
  - `lambda:lambda-invoke`
  - `lambda:map-basic`
  - `lambda:byrow-basic`
  - `names:defined-name-range`
  - `tables:table-total-row-sum`
  - `structured-reference:table-column-ref`

## Current milestone

- close the current `101`-row canonical worksheet formula corpus as represented in `packages/formula/src/compatibility.ts`
- keep parity proved by checked-in oracle fixtures and differential tests
- keep JS as the semantic oracle until the `9` non-production canonical rows close
- close reference-valued names, table totals and column refs, and the JS-only lambda and dynamic-array rows listed above

## Canonical target

- formula semantics target Excel 365 built-in worksheet parity as of `2026-03-15`
- browser and local-server execution remain local-first
- all supported production formulas execute in WASM
- workbook metadata needed by formulas travels with the workbook model:
  - defined names
  - tables
  - structured references
  - spill metadata
  - volatile recalc context

## Repo boundary

- `bilig` owns:
  - parser, binder, optimizer, oracle harness, WASM kernel, workbook metadata model, dynamic-array runtime, compatibility matrix, browser shell, and acceptance docs
- `lab` owns:
  - deployment manifests, rollout gates, observability wiring, alerts, dashboards, and SLO plumbing

See:

- [formula-canonical-program.md](/Users/gregkonush/github.com/bilig/docs/formula-canonical-program.md)
- [formula-canonical-matrix.md](/Users/gregkonush/github.com/bilig/docs/formula-canonical-matrix.md)
- [formula-oracle-capture.md](/Users/gregkonush/github.com/bilig/docs/formula-oracle-capture.md)
- [wasm-runtime-contract.md](/Users/gregkonush/github.com/bilig/docs/wasm-runtime-contract.md)
- [workbook-metadata-model.md](/Users/gregkonush/github.com/bilig/docs/workbook-metadata-model.md)
- [dynamic-array-runtime.md](/Users/gregkonush/github.com/bilig/docs/dynamic-array-runtime.md)
- [authoritative-workbook-op-model-rfc.md](/Users/gregkonush/github.com/bilig/docs/authoritative-workbook-op-model-rfc.md)
- [workbook-metadata-runtime-rfc.md](/Users/gregkonush/github.com/bilig/docs/workbook-metadata-runtime-rfc.md)
- [worker-runtime-and-viewport-patches-rfc.md](/Users/gregkonush/github.com/bilig/docs/worker-runtime-and-viewport-patches-rfc.md)
- [durable-multiplayer-replication-rfc.md](/Users/gregkonush/github.com/bilig/docs/durable-multiplayer-replication-rfc.md)
- [typed-agent-protocol-rfc.md](/Users/gregkonush/github.com/bilig/docs/typed-agent-protocol-rfc.md)
- [bilig-lab-contract.md](/Users/gregkonush/github.com/bilig/docs/bilig-lab-contract.md)

## Exit gate

- the canonical formula registry is fully decision-complete
- every canonical formula entry has fixture-backed status
- every canonical entry is `implemented-wasm-production`
- JS remains oracle, differential, and debug infrastructure rather than a production requirement for canonical rows
- `lab` contracts exist and match the bilig-side runtime assumptions
