# Formula Language

## Current state

- JS is still the semantic oracle
- canonical production routing is closed for `92` of `101` canonical rows
- `6` canonical rows remain `implemented-js`
- `3` canonical rows remain `blocked`

## Canonical corpus milestone

The active milestone is the canonical Excel for the web worksheet formula corpus represented in `packages/formula/src/compatibility.ts`.

The current code-backed canonical registry contains `101` rows.

The `9` non-production canonical rows are:

- `dynamic-array:filter-basic`
- `dynamic-array:unique-basic`
- `lambda:let-basic`
- `lambda:lambda-invoke`
- `lambda:map-basic`
- `lambda:byrow-basic`
- `names:defined-name-range`
- `tables:table-total-row-sum`
- `structured-reference:table-column-ref`

## Full target

The full formula target remains Excel 365 worksheet parity as of `2026-03-15`, including:

- absolute and mixed refs
- quoted sheet refs
- unions and intersections
- `%`, `@`, `#`
- array literals
- defined names
- tables and structured references
- dynamic arrays
- `LET`
- `LAMBDA`

## Host and service function boundary

The worksheet engine keeps non-worksheet Excel surfaces behind the adapter contract in `packages/formula/src/external-function-adapter.ts`.

- native builtins stay focused on the worksheet corpus tracked in this package
- cube, web, host-backed, external-data, and add-in-like surfaces register through `installExternalFunctionAdapter()`
- adapters can expose scalar or range-aware functions to the JS evaluator
- adapted functions do **not** receive a `BuiltinId` and do **not** enter the WASM fast path
- without an installed adapter, those function names continue to resolve as `#NAME?`

## Semantic rules

- Excel for the web is the behavior oracle
- visible error strings follow Excel codes
- JS and WASM must not diverge on coercion, blanks, error precedence, spill blocking, or lookup comparison rules
- a canonical row is not complete until production execution can route it through WASM

## Canonical companions

- [formula-canonical-program.md](/Users/gregkonush/github.com/bilig/docs/formula-canonical-program.md)
- [formula-canonical-matrix.md](/Users/gregkonush/github.com/bilig/docs/formula-canonical-matrix.md)
- [formula-oracle-capture.md](/Users/gregkonush/github.com/bilig/docs/formula-oracle-capture.md)
- [wasm-runtime-contract.md](/Users/gregkonush/github.com/bilig/docs/wasm-runtime-contract.md)
- [dynamic-array-runtime.md](/Users/gregkonush/github.com/bilig/docs/dynamic-array-runtime.md)
- [workbook-metadata-model.md](/Users/gregkonush/github.com/bilig/docs/workbook-metadata-model.md)
