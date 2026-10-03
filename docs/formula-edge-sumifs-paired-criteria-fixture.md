---
title: SUMIFS Paired Criteria Fixture Walkthrough
published: true
description: Narrow public fixture note for the canonical paired-criteria SUMIFS case covered by Bilig headless formula tests.
tags: formulas, sumifs, headless, excel, typescript
canonical_url: https://proompteng.github.io/bilig/formula-edge-sumifs-paired-criteria-fixture.html
---

# SUMIFS Paired Criteria Fixture Walkthrough

Status: public formula-edge fixture note for `@bilig/workpaper`.

This page documents one canonical criteria-aggregate fixture. It is intentionally
narrow: the claim is that the current paired-criteria `SUMIFS` fixture is
represented in the compatibility registry and covered by the checked-in verifier
path. It is not a blanket claim that every Excel `SUMIFS` coercion, wildcard, or
range-shape edge is complete.

## Fixture

Fixture id: `statistical:sumifs-basic`

Source:
[`packages/excel-fixtures/src/canonical-expansion-fixtures.ts`](../packages/excel-fixtures/src/canonical-expansion-fixtures.ts)

Formula:

```excel
=SUMIFS(C1:C4,A1:A4,">0",B1:B4,"x")
```

Inputs and expected output:

| Row | A value | B value | C value | Included |
| --- | ------- | ------- | ------- | -------- |
| 1   | 2       | x       | 10      | yes      |
| 2   | -1      | x       | 20      | no       |
| 3   | 4       | y       | 30      | no       |
| 4   | 7       | x       | 40      | yes      |

The formula sums `C1:C4` only where the matching `A` row is greater than `0` and
the matching `B` row is `x`. Rows `1` and `4` match, so the expected result in
`D1` is `50`.

## Compatibility Status

The registry entry is in
[`packages/formula/src/compatibility.ts`](../packages/formula/src/compatibility.ts):

```ts
entry('statistical:sumifs-basic', 'statistical', '=SUMIFS(C1:C4,A1:A4,">0",B1:B4,"x")', 'implemented-wasm-production')
```

That status means this fixture is treated as a production WASM-compatible
formula fixture by the repository metadata. Future `SUMIFS` behavior should add
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
The runtime correctness suite keeps the canonical criteria aggregate fixtures in
engine/oracle parity on the WASM fast path, including
`statistical:sumifs-basic`.

## What This Does Not Prove

This fixture does not cover every `SUMIFS` option or Excel edge case. In
particular, it does not prove:

- wildcard criteria
- date and time criteria coercion
- mixed-type comparison behavior for every Excel case
- criteria ranges with incompatible shapes
- external workbook references
- every error-propagation edge case from Excel

Those cases should land as separate fixtures with their own expected values and
registry entries. That keeps compatibility claims small enough to audit and easy
for contributors to extend.

## Contribution Shape

To extend this area, add a new canonical fixture, give it a precise compatibility
registry status, and include a focused test path that explains what the fixture
does and does not prove. Prefer one evidence-backed behavior per fixture over a
large ambiguous compatibility claim.
