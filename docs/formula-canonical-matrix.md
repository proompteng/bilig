# Formula Canonical Matrix

## Contract

- `packages/formula/src/compatibility.ts` is the authoritative source of canonical row count and status
- `packages/formula/src/__tests__/fixtures/formula-dominance-snapshot.json` is the generated source of truth for current counts
- the current code-backed canonical registry contains `300` rows
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
| `dynamic-array` | mixed | mixed | `GROUPBY` and `PIVOTBY` remain JS-only in the canonical slice |
| `names` | `implemented-wasm-production` | `production` | no non-production canonical rows |
| `tables` | `implemented-wasm-production` | `production` | no non-production canonical rows |
| `structured-reference` | `implemented-wasm-production` | `production` | no non-production canonical rows |
| `volatile` | `implemented-wasm-production` | `production` | no non-production canonical rows |
| `lambda` | `implemented-wasm-production` | `production` | no non-production canonical rows |

## Current remaining open rows

- `dynamic-array:groupby-basic` (`implemented-js`)
- `dynamic-array:pivotby-basic` (`implemented-js`)

## Notes

- the canonical corpus excludes `text:case-insensitive-compare` and `information:value-error-display`
- the `extended` scope exists in the code registry for work that comes after the current canonical corpus
- this matrix is descriptive of current source code, not of older planning snapshots
