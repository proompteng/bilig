# Formula Top 100 Program

## Goal

Close the Top 100 Excel for the web worksheet formulas as the next hard product milestone.

## Canonical Artifacts

- `formulaCompatibilityRegistry` in `packages/formula/src/compatibility.ts`
- `excelTop100CanonicalFixtures` in `packages/excel-fixtures/src/index.ts`
- `excelTop100SmokeSuite` in `packages/excel-fixtures/src/index.ts`
- milestone key `top100-canonical`

The milestone is fixed at `100` canonical cases:
- keep the current `61` tracked cases unchanged
- add the fixed `39` expansion cases from the product plan
- do not substitute convenience variants for the fixed expansion list

## Fixed Expansion Cases

Text:
- `EXACT`
- `LEFT`
- `RIGHT`
- `MID`
- `TRIM`
- `UPPER`
- `LOWER`
- `FIND`
- `SEARCH`
- `VALUE`

Lookup/reference:
- `XMATCH`
- `HLOOKUP`
- `OFFSET`
- `TAKE`
- `DROP`
- `CHOOSECOLS`
- `CHOOSEROWS`

Statistical/math:
- `SUMIF`
- `SUMIFS`
- `AVERAGEIFS`
- `COUNTIFS`
- `SUMPRODUCT`
- `INT`
- `ROUNDUP`
- `ROUNDDOWN`

Date/time:
- `NOW`
- `TIME`
- `HOUR`
- `MINUTE`
- `SECOND`
- `WEEKDAY`

Dynamic array:
- `SORT`
- `SORTBY`
- `TOCOL`
- `TOROW`
- `WRAPROWS`
- `WRAPCOLS`

`SEQUENCE` remains represented by the existing legacy `sequence-spill` row. The product plan text listed a separate 2D case, but the milestone stays fixed at `100`, so the canonical Top 100 uses the preserved legacy `SEQUENCE` slot instead of adding a second `SEQUENCE` row.

## Canonical Exclusions

The Top 100 canonical export intentionally excludes two legacy probe rows that remain useful for ad hoc regression work but are not part of the audited Top 100 milestone:
- `text:case-insensitive-compare`
- `information:value-error-display`

Names/lambda:
- defined-name range
- `BYROW`

## Execution Order

1. arithmetic, aggregation, and core math
2. logical and information
3. text
4. date/time with JS-level volatile coverage, but excluding native volatile promotion
5. lookup/reference
6. conditional statistical functions
7. names, tables, structured refs, and dynamic arrays required by Top 100 entries

## Delivery Rule

- JS lands first as semantic oracle
- fixtures prove Excel for the web behavior
- WASM lands in shadow mode
- production routing flips only after differential parity is green

## Required Outputs

- canonical compatibility registry
- family fixture packs
- oracle capture metadata
- WASM runtime contract
- workbook metadata contract
- dynamic-array runtime contract

## Exit Gate

- Top 100 unsupported count is zero
- every Top 100 family is `implemented-wasm-production`
- JS remains non-production validation only for closed families
