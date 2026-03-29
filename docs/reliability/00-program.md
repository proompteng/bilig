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
- `v2` session, document state, snapshot, websocket, agent, and Zero ingress routes exist in the web/local/sync path
- shared runtime-kernel adapters now own server edge config, error envelopes, session shaping, request base-url resolution, websocket normalization, and message-byte decoding

### Remaining before the rewrite is complete

- replace sync/local document managers with Effect-backed services supervised by document actors
- move worker runtime, reconnect, and transport lifecycle into XState + Effect
- finish web-app orchestration cutover beyond bootstrap
- delete old `v1` entrypoints and all duplicated fallback control paths
- complete the reliability acceptance suite and failure drills from `04-cutover-and-acceptance.md`

## Repo Rules

- no rewrite PR lands without updating at least one `docs/reliability/*` file
- `@bilig/contracts` is the only source of truth for runtime payload shapes
- new product code may not read ambient env, clock, random, timer, fetch, or websocket APIs directly
- if a behavior is not described in this doc set, it is not part of the rewrite contract
