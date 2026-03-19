# `bilig` / `lab` Contract

## Ownership split

### `bilig`

- product code
- parser, engine, WASM kernel
- oracle fixtures
- formula compatibility registry
- workbook metadata and dynamic-array contracts
- product acceptance docs

### `lab`

- deployment manifests
- rollout policy
- observability wiring
- alerts, dashboards, SLO plumbing
- runtime environment assumptions

## Rules

- `bilig` docs must not prescribe Argo implementation details beyond required runtime assumptions
- `lab` docs must not redefine formula semantics or product acceptance logic
- every runtime assumption that affects formula behavior or performance must be documented on both sides with matching language

## Companion docs

- [/Users/gregkonush/github.com/lab/docs/bilig-deployment-contract.md](/Users/gregkonush/github.com/lab/docs/bilig-deployment-contract.md)
- [/Users/gregkonush/github.com/lab/docs/bilig-rollout-gates.md](/Users/gregkonush/github.com/lab/docs/bilig-rollout-gates.md)
- [/Users/gregkonush/github.com/lab/docs/bilig-observability-contract.md](/Users/gregkonush/github.com/lab/docs/bilig-observability-contract.md)
