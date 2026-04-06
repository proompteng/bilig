# Performance Budgets

## Formula budgets

- closed canonical-corpus families must execute in WASM in production mode
- differential JS-vs-WASM runs are allowed only in test/debug paths
- binary encode/decode p95 `< 1ms` for common formula-driven edit batches

## Browser budgets

- local visible edit response p95 `< 16ms`
- `10k` downstream recalc p95 `< 25ms`
- dense `100k` worker warm-start p95 `< 500ms`
- dense `250k` worker warm-start p95 `< 700ms`
- dense `250k` snapshot restore p95 `< 1000ms`
- richer multisheet corpus cases are benchmark fixtures for investigation and optimization, not hard CI gates yet

## WASM budgets

- WASM kernel startup must remain below the current frontend release ceiling
- WASM binary gzip size remains a release gate
- string/runtime extensions for the canonical corpus must not regress the release budget without an explicit budget update

## Ownership split

- `bilig` owns formula/runtime budgets and release checks
- `lab` owns runtime telemetry, alerts, and rollout enforcement

See:

- [bilig-lab-contract.md](/Users/gregkonush/github.com/bilig/docs/bilig-lab-contract.md)
- [/Users/gregkonush/github.com/lab/docs/bilig-observability-contract.md](/Users/gregkonush/github.com/lab/docs/bilig-observability-contract.md)
