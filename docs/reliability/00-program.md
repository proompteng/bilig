# Reliability Rewrite Program

## Goal

Rewrite the product control plane around `Effect` and `XState v5` so that lifecycle, failure handling, retries, resource scoping, contract decoding, and workflow state are explicit and deterministic across the web app, worker runtime, sync server, local server, and Zero ingress.

This program is a clean-break rewrite of the control plane. It is not a request to wrap existing ad hoc logic in new abstractions.

## Non-goals

- rewriting the pure spreadsheet compute kernels for style over structure
- adding actor-per-cell or machine-per-row models
- preserving legacy `v1` public routes indefinitely
- layering Effect and XState on top of duplicated runtime truth

## Acceptance Bar

- all new public control-plane routes are `v2`
- every public payload and error envelope is defined in `@bilig/contracts`
- every long-lived workflow has an explicit XState machine
- every side effect runs through Effect services
- no expected runtime failure crosses a boundary as an untyped `Error`
- reconnect, retry, timeout, and shutdown behavior are deterministic under test

## Implementation Waves

1. Foundation
   - exact dependency pins
   - `docs/reliability/*`
   - `@bilig/contracts`
   - `@bilig/runtime-kernel`
   - `@bilig/actors`
2. Contract cutover
   - `v2` session, document state, snapshot, websocket, agent, and Zero ingress contracts
   - contract round-trip tests
3. Service cutover
   - Effect-backed sync/local/http services
   - resource-scoped fetch, websocket, timers, logging, metrics, and config
4. Actor cutover
   - web bootstrap/session/document/connection machines
   - worker runtime and transport machines
   - document supervisor machines in sync/local servers
5. Deletion and hardening
   - remove old `v1` paths
   - remove duplicated guards and fallback behavior
   - full reliability, fuzz, and browser acceptance pass

## Current Status

### Completed in `main`

- foundation dependencies are pinned exactly:
  - `effect@3.21.0`
  - `@effect/platform@0.96.0`
  - `@effect/platform-node@0.106.0`
  - `xstate@5.30.0`
  - `@xstate/react@6.1.0`
- the `docs/reliability/*` design-doc set exists and is now the control-plane source of truth
- `@bilig/contracts`, `@bilig/runtime-kernel`, and `@bilig/actors` exist in the workspace
- the web bootstrap path is actor-driven through `@bilig/actors`
- the web workbook shell now runs through a dedicated worker-runtime actor that owns worker bootstrap, runtime refresh, cache invalidation, selection synchronization, and optional Zero bridge lifecycle
- `v2` session, document state, snapshot, websocket, agent, and Zero ingress routes exist in the web/local/sync path
- the sync server now serves Zero only through `/api/zero/v2/query` and `/api/zero/v2/mutate`; GitOps must point Zero at the `v2` ingress paths
- shared runtime-kernel adapters now own server edge config, error envelopes, session shaping, request base-url resolution, websocket normalization, and message-byte decoding
- the web workbook shell now follows a single-authority update rule:
  - viewport and selected-cell render state come only from Zero-backed projection
  - persisted workbook mutations complete through Zero/server, not through a worker-first local mutation path
  - the old worker-plus-Zero dual render path is no longer allowed in the live shell

### Remaining before the rewrite is complete

- replace sync/local document managers with Effect-backed services supervised by document actors
- move worker runtime, reconnect, and transport lifecycle into XState + Effect
- finish web-app orchestration cutover beyond bootstrap
- delete the remaining old manager-backed and raw transport control paths
- complete the reliability acceptance suite and failure drills from `04-cutover-and-acceptance.md`

## Repo Rules

- no rewrite PR lands without updating at least one `docs/reliability/*` file
- `@bilig/contracts` is the only source of truth for runtime payload shapes
- new product code may not read ambient env, clock, random, timer, fetch, or websocket APIs directly
- if a behavior is not described in this doc set, it is not part of the rewrite contract

## Web Runtime Hard-Cut Rule

The live workbook shell may not render from two local authorities at once.

Specifically:

- no worker viewport subscription may run in parallel with Zero viewport projection for the same sheet viewport
- no persisted workbook mutation may be applied worker-first and then sent to Zero as a second write path
- no startup flow may bootstrap visible workbook state from browser-persisted worker data and then merge that state with live Zero data

The allowed live web path is:

1. Zero queries materialize authoritative workbook rows.
2. `ZeroWorkbookBridge` projects those rows into the in-browser cache used by the grid.
3. UI mutations go through Zero mutators and are considered successful only after the server leg succeeds.
4. Server-side mutators update Postgres and Zero replication sends the authoritative row changes back to the browser.

The worker runtime, if present, is not a second authority for visible workbook state.
