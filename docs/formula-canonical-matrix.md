# Formula Canonical Matrix

## Contract

- `packages/formula/src/compatibility.ts` is the authoritative source of canonical row count and status
- the current code-backed canonical registry contains `101` rows
- canonical fixture ids must stay aligned with canonical registry ids
- the checked-in registry is the authoritative source of canonical row count and status
- scope key in code is `canonical`

## Family view

| family | status | wasmStatus | current gap |
| --- | --- | --- | --- |
| `arithmetic` | `implemented-wasm-production` | `production` | no non-production canonical rows |
| `comparison` | `implemented-wasm-production` | `production` | no non-production canonical rows |
| `logical` | `implemented-wasm-production` | `production` | no non-production canonical rows |
| `aggregation` | `implemented-wasm-production` | `production` | no non-production canonical rows |
| `math` | `implemented-wasm-production` | `production` | no non-production canonical rows |
| `text` | `implemented-wasm-production` | `production` | no non-production canonical rows |
| `date-time` | `implemented-wasm-production` | `production` | no non-production canonical rows |
| `lookup-reference` | `implemented-wasm-production` | `production` | no non-production canonical rows |
| `statistical` | `implemented-wasm-production` | `production` | no non-production canonical rows |
| `information` | `implemented-wasm-production` | `production` | no non-production canonical rows |
| `dynamic-array` | mixed | mixed | `FILTER` and `UNIQUE` are JS-only in the canonical slice |
| `names` | mixed | mixed | scalar names are promoted; reference-valued names remain blocked |
| `tables` | `blocked` | `blocked` | `tables:table-total-row-sum` |
| `structured-reference` | `blocked` | `blocked` | `structured-reference:table-column-ref` |
| `volatile` | `implemented-wasm-production` | `production` | no non-production canonical rows |
| `lambda` | mixed | `not-started` | `LET`, `LAMBDA`, `MAP`, and `BYROW` remain JS-only |

## Current remaining open rows

- `dynamic-array:filter-basic` (`implemented-js`)
- `dynamic-array:unique-basic` (`implemented-js`)
- `lambda:let-basic` (`implemented-js`)
- `lambda:lambda-invoke` (`implemented-js`)
- `lambda:map-basic` (`implemented-js`)
- `lambda:byrow-basic` (`implemented-js`)
- `names:defined-name-range` (`blocked`)
- `tables:table-total-row-sum` (`blocked`)
- `structured-reference:table-column-ref` (`blocked`)

## Notes

- the canonical corpus excludes `text:case-insensitive-compare` and `information:value-error-display`
- the `extended` scope exists in the code registry for work that comes after the current canonical corpus
- this matrix is descriptive of current source code, not of older planning snapshots
