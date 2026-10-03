---
title: XLOOKUP Exact Fixture Walkthrough
published: true
description: Narrow public fixture note for the canonical exact-match XLOOKUP case covered by Bilig headless formula tests.
tags: formulas, xlookup, headless, excel, typescript
canonical_url: https://proompteng.github.io/bilig/formula-edge-xlookup-exact-fixture.html
---

# XLOOKUP Exact Fixture Walkthrough

Status: public formula-edge fixture note for `@bilig/workpaper`.

This page documents one canonical lookup fixture. It is intentionally narrow:
the claim is that the current exact-match `XLOOKUP` fixture is represented in
the compatibility registry and covered by the checked-in verifier path. It is
not a blanket claim that every Excel `XLOOKUP` mode is complete.

## Fixture

Fixture id: `lookup-reference:xlookup-exact`

Source:
[`packages/excel-fixtures/src/canonical-foundation-fixtures.ts`](../packages/excel-fixtures/src/canonical-foundation-fixtures.ts)

Formula:

```excel
=XLOOKUP("pear",A1:A3,B1:B3)
```

Inputs and expected output:

| Cell | Value |
| ---- | ----- |
| A1   | apple |
| B1   | 10    |
| A2   | pear  |
| B2   | 20    |
| A3   | plum  |
| B3   | 30    |
| C1   | 20    |

The lookup searches `A1:A3` for `pear` and returns the corresponding value from
`B1:B3`, so the expected result in `C1` is `20`.

## Compatibility Status

The registry entry is in
[`packages/formula/src/compatibility.ts`](../packages/formula/src/compatibility.ts):

```ts
entry('lookup-reference:xlookup-exact', 'lookup-reference', '=XLOOKUP("pear",A1:A3,B1:B3)', 'implemented-wasm-production')
```

That status means this fixture is treated as a production WASM-compatible
formula fixture by the repository metadata. Future `XLOOKUP` behavior should add
new fixture ids rather than stretching this one beyond what it proves.

## Verifier Command

Run the focused verifier path:

```sh
pnpm exec vitest run packages/formula/src/__tests__/fixture-harness.test.ts packages/core/src/__tests__/formula-runtime-correctness.test.ts --reporter=dot
```

Latest local result while adding this note:

```text
Test Files  2 passed (2)
Tests       8 passed (8)
```

The fixture harness checks the canonical formula fixtures through the evaluator.
The runtime correctness suite keeps the canonical text and lookup fixtures in
engine/oracle parity on the WASM fast path, including
`lookup-reference:xlookup-exact`.

## What This Does Not Prove

This fixture does not cover every `XLOOKUP` option or Excel edge case. In
particular, it does not prove:

- `if_not_found` behavior
- approximate match modes
- wildcard matching
- reverse or binary search modes
- multi-dimensional return arrays
- every coercion or error-propagation edge case from Excel

Those cases should land as separate fixtures with their own expected values and
registry entries. That keeps compatibility claims small enough to audit and easy
for contributors to extend.

## Contribution Shape

To extend this area, add a new canonical fixture, give it a precise compatibility
registry status, and include a focused test path that explains what the fixture
does and does not prove. Prefer one evidence-backed behavior per fixture over a
large ambiguous compatibility claim.
