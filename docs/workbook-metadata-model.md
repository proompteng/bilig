# Workbook Metadata Model

## Required metadata for formula parity

- defined names
- table definitions
- pivot definitions and materialized output regions
- structured reference bindings
- spill ownership and blocked ranges
- workbook calc and volatile context

## Ownership

- `@bilig/core` owns runtime workbook metadata state
- `@bilig/formula` consumes metadata during binding and evaluation
- `@bilig/wasm-kernel` consumes lowered metadata needed by production execution
- Workbook snapshots persist workbook metadata under `workbook.metadata`; this tranche wires scalar `definedNames`, spill ownership, and pivot definitions/output bounds

## Canonical Corpus Relevance

The canonical formula corpus cannot close without workbook metadata because:

- names unlock named formulas
- tables unlock structured references
- spill metadata unlocks dynamic-array correctness
- pivot metadata unlocks stable materialized summary tables across sheets

## Exit gate

- metadata is persisted, transported, and used consistently across parser, engine, and WASM execution
