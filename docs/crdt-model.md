# CRDT Model

The engine uses deterministic op batches with replica clocks and last-writer-wins per entity. Transport is intentionally out of scope for v1; the engine only defines the local-first replication contract.

- local mutations create `EngineOpBatch`
- remote mutations are replayed through the same apply path
- batches are ordered by clock, replica id, and batch id
- duplicate batches are ignored
- replica snapshots persist enough metadata to survive restart and ignore stale replays

## Entity model

The replication layer resolves conflicts per entity:

- workbook metadata: `workbook`
- sheet metadata: `sheet:${name}`
- cell source: `cell:${sheetName}!${addr}`

Sheet deletion also records a tombstone barrier for the sheet name so stale cell ops cannot recreate deleted-sheet state out of order.

## Ordering

Op ordering is deterministic:

1. Lamport clock counter
2. replica id
3. batch id
4. op index within the batch

This order is used for both merge sorting and entity-level last-writer-wins checks.

## Persistence

`exportReplicaSnapshot()` and `importReplicaSnapshot()` persist:

- replica id
- Lamport counter
- applied batch ids
- entity version map
- sheet delete tombstones

That keeps restart, replay, and duplicate-delivery behavior deterministic without needing a transport-specific store.

The playground layers a persisted relay queue on top of that engine contract. Paused replica traffic survives reloads and replays through `applyRemoteBatch()` after resume, which makes the local-first behavior visible without baking transport policy into shared packages.

Queued relay batches are compacted with the same entity ordering rules. Repeated offline edits collapse to the latest op per entity, and stale cell ops behind later sheet tombstones are discarded before replay.
