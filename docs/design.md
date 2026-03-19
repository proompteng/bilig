# `bilig` Canonical Product Design

`bilig` is a local-first spreadsheet system with a browser-native Excel shell, a deterministic semantic core, and a WASM execution engine. The current product milestone is the **canonical formula corpus** with a hard repo boundary: `bilig` owns product code and product docs; `lab` owns deployment, rollout, and observability.

## Current state

- `@bilig/core`, `@bilig/formula`, `@bilig/wasm-kernel`, `apps/web`, `apps/local-server`, and `apps/sync-server` exist and are executable
- formula compatibility has moved beyond the original starter set, but full canonical-corpus parity is still open
- JS remains the semantic oracle today
- WASM covers only a proven subset of arithmetic, aggregation, logical, and selected date behavior
- names, tables, structured references, dynamic arrays, and `LET`/`LAMBDA` remain open

## Current milestone

The next product-closing milestone is:

- the **canonical Excel for the web worksheet formula corpus**, which currently contains `100` audited cases
- parity proved by checked-in oracle fixtures
- production routing flipped to WASM for every formula family that closes
- JS retained only for oracle, differential, and debug paths

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
  - parser, binder, optimizer, oracle harness, WASM kernel, workbook metadata model, dynamic-array runtime, compatibility matrix, acceptance docs
- `lab` owns:
  - deployment manifests, rollout gates, observability wiring, alerts, dashboards, SLO plumbing

See:

- [formula-canonical-program.md](/Users/gregkonush/github.com/bilig/docs/formula-canonical-program.md)
- [formula-canonical-matrix.md](/Users/gregkonush/github.com/bilig/docs/formula-canonical-matrix.md)
- [formula-oracle-capture.md](/Users/gregkonush/github.com/bilig/docs/formula-oracle-capture.md)
- [wasm-runtime-contract.md](/Users/gregkonush/github.com/bilig/docs/wasm-runtime-contract.md)
- [workbook-metadata-model.md](/Users/gregkonush/github.com/bilig/docs/workbook-metadata-model.md)
- [dynamic-array-runtime.md](/Users/gregkonush/github.com/bilig/docs/dynamic-array-runtime.md)
- [bilig-lab-contract.md](/Users/gregkonush/github.com/bilig/docs/bilig-lab-contract.md)

## Exit gate

- canonical formula registry is fully decision-complete
- every canonical formula entry has a fixture-backed status
- every closed family has a WASM production route
- `lab` contracts exist and match the bilig-side assumptions
