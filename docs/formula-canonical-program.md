# Formula Canonical Program

## Goal

Close the current canonical Excel for the web worksheet formula corpus as represented in `packages/formula/src/compatibility.ts`.

## Canonical artifacts

- `formulaCompatibilityRegistry` in `packages/formula/src/compatibility.ts`
- `canonicalFormulaFixtures` in `packages/excel-fixtures/src/index.ts`
- `canonicalFormulaSmokeSuite` in `packages/excel-fixtures/src/index.ts`
- scope key `canonical`

## Current code-backed status

The canonical registry in code currently contains `101` canonical rows.

Current status split:

- `92` rows are `implemented-wasm-production`
- `6` rows are `implemented-js`
- `3` rows are `blocked`

The remaining open canonical rows are:

- `dynamic-array:filter-basic`
- `dynamic-array:unique-basic`
- `lambda:let-basic`
- `lambda:lambda-invoke`
- `lambda:map-basic`
- `lambda:byrow-basic`
- `names:defined-name-range`
- `tables:table-total-row-sum`
- `structured-reference:table-column-ref`

The checked-in registry is the authoritative source of corpus size and composition.

## Remaining semantic focus

The remaining non-production rows break down as:

- dynamic-array:
  - `dynamic-array:filter-basic`
  - `dynamic-array:unique-basic`
- lambda:
  - `lambda:let-basic`
  - `lambda:lambda-invoke`
  - `lambda:map-basic`
  - `lambda:byrow-basic`
- names:
  - `names:defined-name-range`
- tables:
  - `tables:table-total-row-sum`
- structured-reference:
  - `structured-reference:table-column-ref`

## Delivery rule

- JS lands first as semantic oracle
- fixtures prove Excel for the web behavior
- WASM lands in shadow or direct production mode depending on maturity
- production routing flips only after differential parity is green

## Required outputs

- canonical compatibility registry
- family fixture packs
- oracle capture metadata
- WASM runtime contract
- workbook metadata contract
- dynamic-array runtime contract

## Exit gate

- canonical corpus unsupported count is zero
- every canonical corpus row is `implemented-wasm-production`
- JS remains non-production validation infrastructure for closed canonical rows
