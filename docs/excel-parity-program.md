# Excel Parity Program

## Current milestone

The immediate milestone is the checked-in canonical Excel for the web worksheet formula corpus as represented in `packages/formula/src/compatibility.ts`.

The current code-backed canonical registry contains `101` rows.

That milestone gates formula completion.

## Canonical target

Match Excel 365 worksheet semantics as of `2026-03-15` across:

- grammar
- evaluation semantics
- workbook metadata semantics
- spill and dynamic-array semantics
- lookup and reference semantics
- volatile semantics

## Delivery model

- fixtures are captured from Excel for the web and checked into `@bilig/excel-fixtures`
- JS lands first as oracle behavior
- WASM lands second in differential and then production mode
- production routing flips to WASM only after parity closes

## Milestone docs

- [formula-canonical-program.md](/Users/gregkonush/github.com/bilig/docs/formula-canonical-program.md)
- [formula-canonical-matrix.md](/Users/gregkonush/github.com/bilig/docs/formula-canonical-matrix.md)
- [formula-oracle-capture.md](/Users/gregkonush/github.com/bilig/docs/formula-oracle-capture.md)

## Exit gate

- all canonical formula entries are fixture-backed
- all canonical entries run in WASM production mode
- remaining open work is only outside the current canonical corpus or outside the worksheet formula scope
