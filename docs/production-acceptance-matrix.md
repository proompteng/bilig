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

- canonical formula production routing is closed for the complete worksheet surface
- grouped-array expansion now continues as broader non-canonical aggregate coverage and performance work
- typed binary agent and viewport payload codecs are not yet closed
- row resize, hide and unhide, context-menu actions, and frozen-pane UX are not implemented in the product shell
