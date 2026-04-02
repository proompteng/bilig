# Durable Multiplayer Replication RFC

## Status

Archived historical RFC. The active production design has already moved to the monolith + Zero architecture.

## Current source of truth

- [zero-bilig-production-implementation-plan-v2.md](/Users/gregkonush/github.com/bilig/docs/zero-bilig-production-implementation-plan-v2.md)
- [bilig_production_plan_2026-03-30.md](/Users/gregkonush/github.com/bilig/docs/bilig_production_plan_2026-03-30.md)
- [production-stability-remediation-2026-04-02.md](/Users/gregkonush/github.com/bilig/docs/production-stability-remediation-2026-04-02.md)

## Current production shape

- `apps/bilig` is the only backend runtime.
- `apps/web` is the only browser shell.
- Zero is the read-sync and mutation ingress plane.
- Postgres is the durable source of truth.
- Durable recovery uses workbook checkpoints plus ordered event replay.
- There is no standalone `apps/local-server` or `apps/sync-server` product topology.

## Why this file remains

This file is kept only as historical context for the earlier replication design discussion. It is not an implementation checklist and it must not be treated as the current runtime contract.
