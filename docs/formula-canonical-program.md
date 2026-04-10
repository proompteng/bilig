# Formula Canonical Program

## Goal

Close the current canonical Excel for the web worksheet formula corpus as represented in `packages/formula/src/compatibility.ts`.

## Canonical artifacts

- `formulaCompatibilityRegistry` in `packages/formula/src/compatibility.ts`
- `canonicalFormulaFixtures` in `packages/excel-fixtures/src/index.ts`
- `canonicalFormulaSmokeSuite` in `packages/excel-fixtures/src/index.ts`
- scope key `canonical`

## Current code-backed status

The canonical registry in code currently contains `300` canonical rows.

Current status split:

- `298` rows are `implemented-wasm-production`
- `2` rows are `implemented-js`
- `0` rows are `blocked`

The remaining open canonical rows are:

- `dynamic-array:groupby-basic`
- `dynamic-array:pivotby-basic`

The checked-in registry is the authoritative source of corpus size and composition.
The generated snapshot at
`packages/formula/src/__tests__/fixtures/formula-dominance-snapshot.json` is the easiest current readout.

## Remaining semantic focus

The remaining non-production rows break down as:

- dynamic-array:
  - `dynamic-array:groupby-basic`
  - `dynamic-array:pivotby-basic`

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
