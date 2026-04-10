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

- `300` rows are `implemented-wasm-production`
- `0` rows are `implemented-js`
- `0` rows are `blocked`

There are no remaining open canonical rows.

The checked-in registry is the authoritative source of corpus size and composition.
The generated snapshot at
`packages/formula/src/__tests__/fixtures/formula-dominance-snapshot.json` is the easiest current readout.

## Remaining semantic focus

Canonical closure is complete.

The next semantic expansion work is outside the canonical corpus:

- broader grouped-array aggregates beyond the canonical SUM forms
- larger grouped-array performance and mutation workloads

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
