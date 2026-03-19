# Formula Language

## Current state

- JS is still the semantic oracle
- WASM executes only a subset of the implemented surface
- the compatibility registry now tracks a broader starter corpus, but the Top 100 milestone is still open

## Top 100 milestone

The active milestone is **Top 100 Excel for the web worksheet formulas**. That milestone includes:

- arithmetic and aggregate operators
- logical and information functions
- high-usage text functions
- high-usage date/time functions
- high-usage lookup/reference functions
- high-usage conditional statistical functions

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

## Semantic rules

- Excel for the web is the behavior oracle
- visible error strings follow Excel codes
- JS and WASM must not diverge on coercion, blanks, error precedence, spill blocking, or lookup comparison rules
- a formula family is not complete until production execution can route it through WASM

## Canonical companions

- [formula-top100-program.md](/Users/gregkonush/github.com/bilig/docs/formula-top100-program.md)
- [formula-top100-matrix.md](/Users/gregkonush/github.com/bilig/docs/formula-top100-matrix.md)
- [formula-oracle-capture.md](/Users/gregkonush/github.com/bilig/docs/formula-oracle-capture.md)
- [wasm-runtime-contract.md](/Users/gregkonush/github.com/bilig/docs/wasm-runtime-contract.md)
- [dynamic-array-runtime.md](/Users/gregkonush/github.com/bilig/docs/dynamic-array-runtime.md)
- [workbook-metadata-model.md](/Users/gregkonush/github.com/bilig/docs/workbook-metadata-model.md)
