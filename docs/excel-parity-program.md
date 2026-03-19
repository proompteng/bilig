# Excel Parity Program

## Current milestone

The immediate milestone is **Top 100 Excel for the web worksheet formulas**.

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

- [formula-top100-program.md](/Users/gregkonush/github.com/bilig/docs/formula-top100-program.md)
- [formula-top100-matrix.md](/Users/gregkonush/github.com/bilig/docs/formula-top100-matrix.md)
- [formula-oracle-capture.md](/Users/gregkonush/github.com/bilig/docs/formula-oracle-capture.md)

## Exit gate

- all Top 100 entries are fixture-backed
- all closed entries run in WASM production mode
- remaining open work is only outside the Top 100 milestone or outside the worksheet formula scope
