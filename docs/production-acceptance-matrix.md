# Production Acceptance Matrix

## Formula Corpus

- canonical compatibility registry exists and is current
- every canonical formula entry links to checked-in fixtures
- every closed family passes JS oracle checks
- every closed family passes WASM differential checks
- every closed family routes to WASM in production mode

## Metadata and dynamic arrays

- workbook metadata model exists for names, tables, structured refs, and spills
- dynamic-array runtime exists and is fixture-covered
- copy/fill/reference translation matches metadata-aware semantics where required

## Runtime split

- `bilig` owns parser, oracle, WASM kernel, workbook model, and acceptance tests
- `lab` owns deployment, rollout gates, and observability contracts

## Operations

- GitHub green
- Forgejo green
- release checks green
- `lab` rollout gates and observability contracts are in place

## Open rows today

- unsupported canonical formula entries remain
- dynamic arrays, names, tables, and `LET`/`LAMBDA` are still open
- full WASM-only production routing is not yet closed for the complete worksheet surface
