# Excel Parity Program

## Current milestone

The immediate milestone is the **canonical Excel for the web worksheet formula corpus**, which currently contains `100` audited cases.

That milestone is the gate between the current starter surface and the full canonical target.

## Canonical target

Match Excel 365 worksheet semantics as of `2026-03-15` across:

- grammar
- evaluation semantics
- workbook metadata semantics
- spill/dynamic array semantics
- lookup/reference semantics
- volatile semantics

## Delivery model

- fixtures are captured from Excel for the web and checked into `@bilig/excel-fixtures`
- JS lands first as oracle behavior
- WASM lands second in differential mode
- production routing flips to WASM only after parity closes

## Milestone docs

- [formula-canonical-program.md](/Users/gregkonush/github.com/bilig/docs/formula-canonical-program.md)
- [formula-canonical-matrix.md](/Users/gregkonush/github.com/bilig/docs/formula-canonical-matrix.md)
- [formula-oracle-capture.md](/Users/gregkonush/github.com/bilig/docs/formula-oracle-capture.md)

## Exit gate

- all canonical formula entries are fixture-backed
- all closed entries run in WASM production mode
- remaining open work is only outside the current canonical corpus or outside the worksheet formula scope
