# Implementation Ledger

This ledger maps the current Top 100 formula milestone to concrete proof points.

## Closed foundation rows

| Row | Proof |
| --- | --- |
| compatibility registry exists | `packages/formula/src/compatibility.ts` |
| checked-in fixture packs exist | `packages/excel-fixtures/src` |
| JS oracle tests exist | `packages/formula/src/__tests__` |
| WASM kernel tests exist | `packages/wasm-kernel/src/__tests__/kernel.test.ts` |
| browser/product acceptance exists | `e2e/tests/web-shell.pw.ts` |

## Required future docs

- [formula-top100-program.md](/Users/gregkonush/github.com/bilig/docs/formula-top100-program.md)
- [formula-top100-matrix.md](/Users/gregkonush/github.com/bilig/docs/formula-top100-matrix.md)
- [formula-oracle-capture.md](/Users/gregkonush/github.com/bilig/docs/formula-oracle-capture.md)
- [wasm-runtime-contract.md](/Users/gregkonush/github.com/bilig/docs/wasm-runtime-contract.md)
- [workbook-metadata-model.md](/Users/gregkonush/github.com/bilig/docs/workbook-metadata-model.md)
- [dynamic-array-runtime.md](/Users/gregkonush/github.com/bilig/docs/dynamic-array-runtime.md)
- [bilig-lab-contract.md](/Users/gregkonush/github.com/bilig/docs/bilig-lab-contract.md)

## Open rows

- Top 100 registry still contains unsupported entries
- not every closed JS family has a matching WASM production route
- workbook metadata semantics are not fully closed
- dynamic-array runtime is not closed
- `lab` deployment/rollout/observability contracts are not yet linked from the product acceptance flow

## Release rule

No formula family closes until:

1. fixtures exist,
2. JS passes,
3. WASM passes differential parity,
4. production routing flips to WASM,
5. the matching `lab` runtime assumptions are documented where relevant.
