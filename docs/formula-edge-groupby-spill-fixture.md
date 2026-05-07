# GROUPBY Spill Fixture Walkthrough

Status: public formula-edge fixture note for `@bilig/headless`.

This page documents one canonical grouped dynamic-array fixture. It is
intentionally narrow: the claim is that the current SUM-form `GROUPBY` fixture
is represented in the compatibility registry and covered by the checked-in
verifier path. It is not a blanket claim that every Excel `GROUPBY` option,
aggregation function, sorting mode, or coercion edge is complete.

## Fixture

Fixture id: `dynamic-array:groupby-basic`

Source:
[`packages/excel-fixtures/src/canonical-expansion-fixtures.ts`](../packages/excel-fixtures/src/canonical-expansion-fixtures.ts)

Formula:

```excel
=GROUPBY(A1:A5,C1:C5,SUM,3,1)
```

Inputs:

| Cell | Value  |
| ---- | ------ |
| A1   | Region |
| A2   | East   |
| A3   | West   |
| A4   | East   |
| A5   | West   |
| C1   | Sales  |
| C2   | 10     |
| C3   | 7      |
| C4   | 5      |
| C5   | 4      |

Expected spill output:

| Cell | Value  |
| ---- | ------ |
| E1   | Region |
| F1   | Sales  |
| E2   | East   |
| F2   | 15     |
| E3   | West   |
| F3   | 11     |
| E4   | Total  |
| F4   | 26     |

The formula groups the region labels in `A2:A5`, sums the matching sales values
from `C2:C5`, keeps the header row, and emits a total row. The expected spill is
therefore two grouped rows plus the final total.

## Compatibility Status

The registry entry is in
[`packages/formula/src/compatibility.ts`](../packages/formula/src/compatibility.ts):

```ts
entry(
  "dynamic-array:groupby-basic",
  "dynamic-array",
  "=GROUPBY(A1:A5,C1:C5,SUM,3,1)",
  "implemented-wasm-production",
  {
    notes:
      "The canonical SUM-form GROUPBY case now lowers onto an internal native grouped-array builtin, so the canonical spill executes on the wasm path with oracle coverage.",
  },
)
```

That status means this fixture is treated as a production WASM-compatible
formula fixture by the repository metadata. Future `GROUPBY` behavior should add
new fixture ids rather than stretching this one beyond what it proves.

## Verifier Commands

Run the focused verifier path:

```sh
pnpm exec vitest run packages/formula/src/__tests__/fixture-harness.test.ts packages/core/src/__tests__/formula-runtime-correctness.test.ts --reporter=dot
```

Run the generated coverage gate that checks fixture registry alignment:

```sh
pnpm calculation:semantics:check
```

Latest local result while adding this note:

```text
Test Files  2 passed (2)
Tests       9 passed (9)
```

The fixture harness checks the canonical formula fixtures through the evaluator.
The runtime correctness suite keeps the canonical grouped-array SUM fixtures in
engine/oracle parity on the WASM fast path, including
`dynamic-array:groupby-basic`.

## What This Does Not Prove

This fixture does not cover every `GROUPBY` option or Excel edge case. In
particular, it does not prove:

- every aggregation function accepted by Excel
- multiple value arrays
- custom sort order and full optional-argument behavior
- nested dynamic arrays
- empty groups or filtered-out rows
- every text, numeric, date, error, and coercion edge case from Excel

Those cases should land as separate fixtures with their own expected spill
outputs and registry entries. That keeps compatibility claims small enough to
audit and easy for contributors to extend.

## Contribution Shape

To extend this area, add a new canonical fixture, give it a precise
compatibility registry status, and include a focused test path that explains
what the fixture does and does not prove. Prefer one evidence-backed spill
behavior per fixture over a broad dynamic-array compatibility claim.
