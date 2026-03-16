# Formula Language

## Canonical target

The formula target is **Excel 365 built-in worksheet parity as of March 15, 2026**.

That includes:

- absolute refs
- quoted sheet refs
- unions and intersections
- postfix `%`
- implicit intersection `@`
- spill operator `#`
- array literals
- defined names
- tables and structured references
- dynamic arrays
- `LET`
- `LAMBDA`
- built-in worksheet function families across logical, math, text, date/time, lookup/reference, statistical, financial, engineering, information, and dynamic array categories

## Semantic rules

- JS remains the semantic oracle.
- WASM only accelerates subsets that preserve exact JS parity.
- Excel coercion rules, blank handling, error precedence, spill behavior, and volatile invalidation are part of the parity contract, not optional implementation details.

## Fixture model

- checked-in goldens live under `@bilig/excel-fixtures`
- parity suites are generated offline from real Excel outputs
- CI validates against checked-in goldens, not live external services

## Current tranche status

The current repo still ships a narrower formula surface than the canonical target. The production docs now freeze the end state, and the first foundation tranche adds the fixture package that will carry the parity corpus as it lands.
