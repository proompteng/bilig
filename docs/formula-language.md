# Formula Language

Current implementation supports:

- numeric literals
- boolean literals
- string literals
- scalar cell refs such as `A1`
- cross-sheet refs such as `Sheet2!B3`
- bounded ranges such as `A1:B3`
- arithmetic operators `+ - * / ^`
- comparisons `= <> > >= < <=`
- text concat `&`
- builtins including `SUM`, `AVG`, `MIN`, `MAX`, `COUNT`, `COUNTA`, `ABS`, `ROUND`, `FLOOR`, `CEILING`, `MOD`, `IF`, `AND`, `OR`, `NOT`, `LEN`, `CONCAT`

The WASM fast path is intentionally narrower than the full JS evaluator.
