# CRDT and Local-First Model

`bilig` is local-first by default. Collaboration is built on deterministic CRDT batches with last-writer-wins conflict resolution per entity.

## Entity ordering

Batch and op ordering is:

1. `clock.counter`
2. `replicaId`
3. `batchId`
4. `opIndex`

## Entity keys

- workbook metadata
- sheet metadata keyed by sheet name
- cell source keyed by `sheetName!address`
- format metadata keyed by `sheetName!address`

## Durability contract

- local edits apply immediately
- outbound sync batches are persisted locally before reconnect/replay
- remote server acks happen only after durable append
- replay after restart must remain deterministic

## Backend contract

The backend sync layer is now part of the canonical product:

- browsers push CRDT batches over the binary protocol
- the server appends batches durably before ack
- reconnect resumes from the last acknowledged cursor
- compaction preserves dedupe and replay safety

## Current tranche status

The repo already had deterministic local-first replay semantics and persisted replica snapshots. The new production tranche adds binary transport framing and server-side durable-store abstractions so the CRDT model now has a concrete backend target instead of stopping at the engine boundary.

## See also

- [authoritative-workbook-op-model-rfc.md](/Users/gregkonush/github.com/bilig/docs/authoritative-workbook-op-model-rfc.md)
- [durable-multiplayer-replication-rfc.md](/Users/gregkonush/github.com/bilig/docs/durable-multiplayer-replication-rfc.md)
