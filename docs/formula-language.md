# Formula Language

Current implementation supports:

- numeric literals
- boolean literals
- string literals
- scalar cell refs such as `A1`
- cross-sheet refs such as `Sheet2!B3`
- quoted cross-sheet refs such as `'My Sheet'!B3`
- bounded ranges such as `A1:B3`
- full-column ranges such as `A:C` and `Sheet2!A:A`
- full-row ranges such as `1:10`
- arithmetic operators `+ - * / ^`
- comparisons `= <> > >= < <=`
- text concat `&`
- builtins including `SUM`, `AVG`, `MIN`, `MAX`, `COUNT`, `COUNTA`, `ABS`, `ROUND`, `FLOOR`, `CEILING`, `MOD`, `IF`, `AND`, `OR`, `NOT`, `LEN`, `CONCAT`

The WASM fast path is intentionally narrower than the full JS evaluator. It currently covers arithmetic, comparisons, numeric aggregates, and branch-safe boolean/numeric formulas such as `IF(A1>0,A1*2,A2-1)`.

## Range semantics

- Bounded cell ranges materialize their member cells eagerly for dependency tracking.
- Full-row and full-column ranges stay on the JS path in v1.
- Full-row and full-column ranges expand over currently materialized cells on the target sheet.
- When a new cell is materialized later, formulas that reference a matching row or column range are rebound so future edits propagate correctly.

## CSV bridge

- `exportSheetCsv(sheetName)` exports a single sheet as CSV.
- `importSheetCsv(sheetName, csv)` replaces one sheet from CSV content.
- Cells beginning with `=` import as formulas.
- `TRUE` and `FALSE` import as booleans.
- Numeric scalars import as numbers.
- All other CSV fields import as strings.
