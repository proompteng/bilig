# Backend Sync Service

## Current state

- `apps/sync-server` is executable and tested
- ingress is binary over HTTP today, not yet binary websocket on the hot path
- durability is still backed by in-memory store abstractions, not Postgres/object storage
- remote agent ingress exists, but live worksheet execution is not complete

## Product role

`apps/sync-server` is the control plane and realtime ingress for collaborative worksheets.

## Canonical responsibilities

- accept binary browser sync frames
- persist CRDT batches durably before ack
- manage document ownership and cursor state
- serve snapshot restore endpoints
- expose a remote agent ingress
- later: binary websocket fanout and cross-pod routing

## Production target

- Fastify HTTP control plane
- binary websocket gateway
- Postgres metadata and cursor store
- Redis presence and routing
- object-storage snapshot persistence

## Exit gate

- append-before-ack is proven against durable storage
- cursor catch-up, snapshot restore, and reconnect replay are tested end to end
- remote agent ingress executes against live worksheet sessions
- production traffic runs through the websocket gateway rather than the HTTP-only scaffold path
