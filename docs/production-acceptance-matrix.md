# Production Acceptance Matrix

## Formula corpus

- canonical compatibility registry exists and is current
- canonical rows have checked-in fixture coverage
- closed canonical families pass JS oracle checks
- closed canonical families pass WASM differential checks
- closed canonical families route to WASM in production mode

## Metadata and dynamic arrays

- workbook metadata model exists for names, tables, structured refs, spills, pivots, filters, sorts, freeze panes, and row and column metadata
- dynamic-array runtime exists and is fixture-covered
- copy, fill, and reference translation are metadata-aware where implemented

## Runtime split

- `bilig` owns parser, oracle, WASM kernel, workbook model, browser shell, and acceptance tests
- `lab` owns deployment, rollout gates, and observability contracts

## Operations

- CI and release checks exist in-repo
- `lab` rollout gates and observability contracts are linked

## Open rows today

- `9` canonical formula rows are not yet `implemented-wasm-production`
- `names:defined-name-range`, `tables:table-total-row-sum`, and `structured-reference:table-column-ref` are not production
- full WASM-only production routing is not yet closed for the complete canonical worksheet surface
- sync-server durable remote worksheet execution is not yet the default checked-in behavior
- typed binary agent and viewport payload codecs are not yet closed
- row resize, hide and unhide, context-menu actions, and frozen-pane UX are not implemented in the product shell
