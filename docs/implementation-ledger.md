# Implementation Ledger

This ledger maps the current production path to concrete proof points in the checked-in source.

## Closed foundation rows

| Row | Proof |
| --- | --- |
| Worker-first browser shell | `apps/web/src/WorkerWorkbookApp.tsx` |
| Zero-backed authoritative viewport bridge | `apps/web/src/zero/ZeroWorkbookBridge.ts` |
| Monolith backend runtime | `apps/bilig/src/index.ts` |
| Zero service in monolith | `apps/bilig/src/zero/service.ts` |
| Semantic Zero mutators | `apps/bilig/src/zero/server-mutators.ts` |
| Recalc worker | `apps/bilig/src/zero/recalc-worker.ts` |
| Relational Zero schema (repo-local) | `packages/zero-sync/src/schema.ts` |
| Additive local Postgres schema | `docker/postgres/02-v2-schema.sql` |
| Transport-neutral workbook op layer | `packages/workbook-domain/src/index.ts` |

## Open work that still matters

- keep reducing projection and render write amplification
- finish the remaining non-production canonical formula rows

## Removed or retired product surfaces

- standalone `apps/local-server`
- standalone `apps/sync-server`
- Redis as a required product runtime component
- placeholder-only monolith files with no imports or shipping behavior

## Release rule

No runtime or deployment surface is considered complete until:

1. the monolith path is the only supported product backend
2. image build/publish workflows target the same Docker runtime names used locally and in deployment
3. browser tests, typecheck, lint, unit tests, and deployment manifests agree on the same app topology
