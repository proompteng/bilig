# Formula Top 100 Matrix

## Contract

- the canonical registry contains exactly `100` rows
- the canonical fixture export contains exactly `100` rows
- registry ids and fixture ids must match exactly
- the first `61` rows are the preserved pre-Top-100 set
- the remaining `39` rows are the fixed expansion list
- milestone key in code is `top100-canonical`

## Family View

| family | status | wasmStatus | blocking gaps |
| --- | --- | --- | --- |
| `arithmetic` | `implemented-wasm-production` | `production` | remaining operator edge parity outside the current tracked slice |
| `comparison` | mixed | mixed | string and lookup collation parity across the broader post-Top-100 surface |
| `logical` | mixed | mixed | exact `IF` branch laziness and error suppression |
| `aggregation` | `implemented-wasm-production` | `production` | broader family coverage outside the current tracked slice |
| `math` | mixed | mixed | remaining promoted functions plus criteria-aware math families |
| `text` | mixed | mixed | `VALUE` plus full string-runtime promotion into WASM |
| `date-time` | mixed | mixed | `NOW` native promotion and volatile epoch normalization |
| `lookup-reference` | mixed | mixed | `OFFSET` and richer reference-returning semantics |
| `statistical` | mixed | `not-started` | criteria/wildcard parity and WASM promotion |
| `information` | `implemented-wasm-production` | `production` | broader information-family coverage outside the current tracked slice |
| `dynamic-array` | `blocked` | `blocked` | spill runtime, array value model, and blocking semantics |
| `names` | `blocked` | `blocked` | workbook metadata model and name rebinding |
| `tables` | `blocked` | `blocked` | table metadata model |
| `structured-reference` | `blocked` | `blocked` | parser support and metadata binding |
| `volatile` | mixed | `not-started` | epoch/provider contract for native promotion |
| `lambda` | `blocked` | `blocked` | callable scope/runtime model |

## Fixed Expansion Rows

Text:
- `text:exact-basic`
- `text:left-basic`
- `text:right-basic`
- `text:mid-basic`
- `text:trim-basic`
- `text:upper-basic`
- `text:lower-basic`
- `text:find-basic`
- `text:search-basic`
- `text:value-basic`

Lookup/reference and dynamic-array selection helpers:
- `lookup-reference:xmatch-basic`
- `lookup-reference:hlookup-basic`
- `lookup-reference:offset-basic`
- `dynamic-array:take-basic`
- `dynamic-array:drop-basic`
- `dynamic-array:choosecols-basic`
- `dynamic-array:chooserows-basic`

Statistical and math:
- `statistical:sumif-basic`
- `statistical:sumifs-basic`
- `statistical:averageifs-basic`
- `statistical:countifs-basic`
- `math:sumproduct-basic`
- `math:int-basic`
- `math:roundup-basic`
- `math:rounddown-basic`

Date/time:
- `date-time:now-volatile`
- `date-time:time-basic`
- `date-time:hour-basic`
- `date-time:minute-basic`
- `date-time:second-basic`
- `date-time:weekday-basic`

Dynamic array:
- `dynamic-array:sort-basic`
- `dynamic-array:sortby-basic`
- `dynamic-array:tocol-basic`
- `dynamic-array:torow-basic`
- `dynamic-array:wraprows-basic`
- `dynamic-array:wrapcols-basic`

Names and lambda:
- `names:defined-name-range`
- `lambda:byrow-basic`

`SEQUENCE` stays represented by the preserved legacy `dynamic-array:sequence-spill` row so the audited Top 100 stays fixed at `100`.

## Notes

- `packages/excel-fixtures/src/index.ts` still imports the legacy seed file `top50.ts`; that is a filename legacy, not a milestone naming contract.
- `post-top100` exists in the code registry for work that comes after the audited Top 100 milestone; it is intentionally outside this matrix.
