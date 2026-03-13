# CRDT Model

The engine uses deterministic op batches with replica clocks.

- local mutations create `EngineOpBatch`
- remote mutations are replayed through the same apply path
- batches are ordered by clock, replica id, and batch id
- duplicate batches are ignored
- transport is intentionally out of scope for v1
