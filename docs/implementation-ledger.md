# Implementation Ledger

This ledger maps the current canonical formula milestone to concrete proof points in the checked-in source.

## Closed foundation rows

| Row | Proof |
| --- | --- |
| compatibility registry exists | `packages/formula/src/compatibility.ts` |
| checked-in fixture packs exist | `packages/excel-fixtures/src` |
| JS oracle tests exist | `packages/formula/src/__tests__` |
| WASM kernel tests exist | `packages/wasm-kernel/src/__tests__/kernel.test.ts` |
| browser and product acceptance exists | `e2e/tests/web-shell.pw.ts` |
| worker-backed browser shell exists | `apps/web/src/WorkerWorkbookApp.tsx`, `apps/web/src/workbook.worker.ts` |
| transaction-based workbook engine exists | `packages/core/src/engine.ts` |
| metadata-aware workbook store exists | `packages/core/src/workbook-store.ts` |

## Current open rows

- `9` canonical rows remain non-production in the registry:
  - `6` are `implemented-js`
  - `3` are `blocked`
- reference-valued defined names, table semantics, and structured-reference production routing are not closed
- sync-server remote worksheet execution returns `NOT_IMPLEMENTED` by default for most worksheet requests
- server-side storage remains in-memory in the checked-in storage package
- agent payloads use JSON inside a binary frame envelope
- viewport patch payloads use JSON inside a byte envelope

## Release rule

No canonical formula family closes until:

1. fixtures exist
2. JS passes
3. WASM passes differential parity
4. production routing flips to WASM
5. the matching runtime assumptions are documented where relevant
