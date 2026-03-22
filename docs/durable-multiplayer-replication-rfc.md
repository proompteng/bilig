# Durable Multiplayer Replication RFC

## Problem

The local server is already meaningfully ahead of the remote sync service. `apps/local-server` hosts live workbook sessions with an in-memory engine, while `apps/sync-server` is still closer to a durability and relay scaffold than a fully durable multiplayer backend. The repo needs a design that turns local-first collaboration into a durable, replayable, multi-client system without blurring durable history and ephemeral awareness.

## Goals

- define durable workbook replication architecture
- separate durable workbook history from presence and render concerns
- define the roles of local server and sync server clearly
- provide catch-up, snapshot, websocket fanout, and ownership semantics

## Non-goals

- treat ephemeral in-cell editing presence as durable workbook history
- move the entire product to cloud authority before local-first behavior is solid
- redefine the engine as a multi-document database

## Stream separation

The service model should keep three streams separate:

1. durable workbook ops
2. derived viewport/render patches
3. ephemeral presence and awareness

Durable workbook history should remain append-only and replay-safe. Presence should stay lightweight and disposable.

## Proposed service roles

### `apps/local-server`

- per-device workbook authority
- live worksheet execution host
- local browser and local agent session manager
- offline-first replay and reconnect coordinator
- optional upstream relay to the remote sync service

### `apps/sync-server`

- durable append-only log
- snapshot compaction and restore
- cursor-based tail catch-up
- websocket fanout for remote clients
- ownership and lease coordination
- remote agent ingress and optional headless worksheet execution

## Durability model

### Append-before-ack

The remote service should durably append workbook transactions before acknowledging them.

### Snapshot plus tail catch-up

Clients should restore from:

1. latest snapshot
2. durable tail replay from cursor

### Compaction

Compaction should preserve:

- dedupe safety
- sheet delete barriers
- replay determinism
- metadata and structural operation correctness

## Ownership and presence

- workbook ownership/lease semantics should live beside the durable service
- presence and cursor channels should be separate from durable mutation history
- same-cell active editing does not require turning the whole spreadsheet into a text CRDT
- committed cell content can remain authoritative workbook history while active in-cell editing uses a smaller ephemeral collaboration surface

## Local-first behavior target

- offline local edits apply immediately
- local queue persists across reload/restart
- reconnect resumes from durable cursor
- remote service catches clients up from snapshot and tail
- multiple browsers and agents converge on the same workbook transaction history

## Suggested package direction

- `@bilig/crdt`
  - authoritative ordering, compaction, dedupe, and replica clocks
  - conceptually closer to an oplog/replication-model package
- `@bilig/binary-protocol`
  - durable sync wire format
- `@bilig/storage-browser`
  - local queue, snapshot, and cursor persistence
- `@bilig/storage-server`
  - durable oplog, snapshot, presence, and ownership backends

## Migration order

1. align on authoritative workbook transaction model
2. harden local persistence and replay semantics
3. add durable append-before-ack storage path to sync server
4. add snapshot plus tail restore endpoints and websocket fanout
5. add ownership/lease and presence channels
6. add remote worksheet execution or headless host integration

## Suggested PR breakdown

1. durable transaction storage contract
2. snapshot plus cursor catch-up contract
3. websocket fanout and reconnect path
4. ownership and presence channel separation
5. remote worksheet execution path
6. long-lived session replay and compaction hardening

## Exit gate

- remote append-before-ack is proven against durable storage
- snapshot plus tail catch-up works end to end
- websocket fanout serves durable workbook history separately from presence
- two or more browsers and agents can converge through offline replay and reconnect
- long-lived collaboration does not rely on infinite in-memory retention
