# Formula Language

Current implementation supports:

- numeric literals
- boolean literals
- string literals
- scalar cell refs such as `A1`
- cross-sheet refs such as `Sheet2!B3`
- quoted cross-sheet refs such as `'My Sheet'!B3`
- bounded ranges such as `A1:B3`
- arithmetic operators `+ - * / ^`
- comparisons `= <> > >= < <=`
- text concat `&`
- builtins including `SUM`, `AVG`, `MIN`, `MAX`, `COUNT`, `COUNTA`, `ABS`, `ROUND`, `FLOOR`, `CEILING`, `MOD`, `IF`, `AND`, `OR`, `NOT`, `LEN`, `CONCAT`

The WASM fast path is intentionally narrower than the full JS evaluator.

## CSV bridge

- `exportSheetCsv(sheetName)` exports a single sheet as CSV.
- `importSheetCsv(sheetName, csv)` replaces one sheet from CSV content.
- Cells beginning with `=` import as formulas.
- `TRUE` and `FALSE` import as booleans.
- Numeric scalars import as numbers.
- All other CSV fields import as strings.
