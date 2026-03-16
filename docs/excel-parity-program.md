# Excel Parity Program

## Target

Match Excel 365 built-in worksheet formula semantics as of `2026-03-15`.

## Scope

- formula grammar parity
- evaluation semantics parity
- function-family parity
- spill and dynamic array behavior
- lookup/reference edge cases
- date serial and coercion parity

## Exclusions

- VBA macros
- external add-in UDFs
- Power Query, Pivot, and non-cell formula systems

## Delivery model

- JS evaluator is the oracle
- WASM accelerates safe overlap sets only
- parity is verified against checked-in Excel golden fixtures
- every new function family lands with fixtures before being considered complete

## First tranche

`@bilig/excel-fixtures` now exists and carries the seed parity suite. The full corpus lands incrementally from there.
