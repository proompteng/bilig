# `bilig` Canonical Product Design

`bilig` is a server-authoritative multiplayer spreadsheet with a worker-first browser shell, a deterministic semantic core, and a WASM-accelerated formula engine.

## Current state

- `@bilig/core`, `@bilig/formula`, `@bilig/wasm-kernel`, `@bilig/workbook-domain`, `apps/web`, and `apps/bilig` are the active product surface.
- The browser shell is worker-first and consumes viewport patches.
- The production collaboration path is Zero-backed and Postgres-backed.
- The monolith serves the active session, Zero, agent, and recalc surfaces.
- The canonical registry in code still contains a small non-production tail; feature parity exists, but the remaining gap is polish, performance, and operational hardening rather than basic capability.

## Canonical target

- formula semantics target Excel-class parity for supported families
- browser rendering stays worker-first
- authoritative workbook state stays server-ordered and relationally materialized
- common edits feel instant through worker preview plus fast authoritative convergence
- deployment stays intentionally simple:
  - `bilig-app` single deployable monolith image
  - `bilig-zero`
  - Postgres

## Competitive baselines

The current implementation should outperform consumer spreadsheet products on the dimensions we control directly:

- lower visible edit latency through worker-side preview and narrow authoritative tile sync
- cleaner rebuild semantics through `workbook_event` plus warm snapshots
- stronger formula/runtime determinism than ad hoc browser-state spreadsheets
- better operational observability than opaque hosted spreadsheet products

## Repo boundary

- `bilig` owns product runtime, semantic engine, browser shell, and acceptance behavior
- `lab` owns deployment manifests, Argo CD rollout, and cluster operations

## Exit gate

- the monolith is the only supported backend runtime in the repo
- browser product flows use Zero plus Postgres-backed authoritative state
- product docs no longer describe the retired duplicate app topology as current
