# Workbook Metadata Model

## Required metadata for formula parity

- defined names
- table definitions
- structured reference bindings
- spill ownership and blocked ranges
- workbook calc and volatile context

## Ownership

- `@bilig/core` owns runtime workbook metadata state
- `@bilig/formula` consumes metadata during binding and evaluation
- `@bilig/wasm-kernel` consumes lowered metadata needed by production execution

## Top 100 relevance

The Top 100 milestone cannot close without workbook metadata because:

- names unlock named formulas
- tables unlock structured references
- spill metadata unlocks dynamic-array correctness

## Exit gate

- metadata is persisted, transported, and used consistently across parser, engine, and WASM execution
