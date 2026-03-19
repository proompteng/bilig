# Dynamic Array Runtime

## Scope

Dynamic-array semantics cover:

- array literals
- spill allocation
- spill blocking
- scalar-vs-array argument behavior
- visible `#SPILL!` and related array errors

## Runtime contract

- the engine must know which cell owns a spill
- blocked spill targets must produce Excel-style visible errors
- array-producing functions must preserve shape through recalculation and copy/fill
- WASM production routing for dynamic-array families requires array value transport, not scalar-only lowering

## Dependency order

Dynamic arrays depend on:

- workbook metadata model
- richer runtime value model
- parser support for spill operators and array syntax

## Exit gate

- SEQUENCE/FILTER/UNIQUE-class behavior is fixture-covered and production-routed through WASM
