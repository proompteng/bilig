# Backend Sync Service

## Current state

The canonical backend is `apps/bilig`.

What the backend owns today:

- `/healthz`
- `/v2/session`
- `/api/zero/v2/query`
- `/api/zero/v2/mutate`
- agent ingress on `/v2/agent/frames`
- the embedded recalc worker and authoritative Postgres materialization path
- an integrated local listener that remains available for harnesses and import/export workflows

What is no longer the product authority:

- the retired `apps/sync-server` package
- the retired CRDT-first browser sync topology

## Product role

`apps/bilig` is the single production backend runtime for the spreadsheet product.
It serves:

- auth/session boot
- Zero query and mutate surfaces
- workbook serialization and authoritative write ordering
- recalc job processing and `cell_eval` materialization
- agent APIs and operational endpoints

## Operational target

- Fastify monolith
- Postgres as source of truth
- Zero as sync/cache runtime
- no Redis dependency on the correctness path
- no separate product backend packages for local or sync authority

## Exit gate

- production traffic runs through `apps/bilig`
- Zero query/mutate endpoints are served by the monolith
- browser product flows do not depend on the removed `apps/sync-server` package
- deployment manifests and image workflows target the monolith image only
