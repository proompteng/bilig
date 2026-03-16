# Performance Budgets

## Browser budgets

- local visible edit response p95 `< 16ms`
- `10k` downstream recalc p95 `< 25ms`
- `100k` workbook restore p95 `< 500ms`
- `250k` preset restore p95 `< 1500ms`
- million-row navigation must keep main-thread work bounded

## Transport budgets

- binary encode/decode p95 `< 1ms` for common edit batches
- remote visibility p95 `< 250ms` intra-region

## Resource ceilings

- browser working set for `100k` materialized workbook stays under the configured release ceiling
- server RSS per active hot document stays under the configured release ceiling
- WASM binary size and frontend bundle size are release gates

## Enforcement

- benchmark contracts
- browser perf smoke
- backend latency tests
- release checks
