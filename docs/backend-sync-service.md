# Backend Sync Service

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

## Current tranche

The repo now includes:

- a typed sync-server app
- in-memory durable-store abstractions
- binary HTTP frame ingress
- remote agent ingress skeleton

The websocket gateway, durable Postgres/Redis/object storage adapters, and multi-pod routing remain active implementation work.
