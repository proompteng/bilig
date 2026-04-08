# CRDT and Local-First Model

## Status

Archived historical note.

The current production architecture is not CRDT-authoritative. The active design is:

- server-authoritative ordering in `apps/bilig`
- Zero as the narrow relational sync plane
- `@bilig/core` as the owner of local replica bookkeeping needed for replay and snapshot restore

## Current source of truth

- [design.md](/Users/gregkonush/github.com/bilig/docs/design.md)
- [architecture.md](/Users/gregkonush/github.com/bilig/docs/architecture.md)
- [05-06-next-phase.md](/Users/gregkonush/github.com/bilig/docs/05-06-next-phase.md)
- [replica-state-ownership-cleanup-2026-04-07.md](/Users/gregkonush/github.com/bilig/docs/replica-state-ownership-cleanup-2026-04-07.md)

## Why this file remains

This file is kept only as historical context for earlier local-first replication language. It must not be read as the current runtime contract.
