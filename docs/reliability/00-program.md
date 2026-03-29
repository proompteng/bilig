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

## Repo Rules

- no rewrite PR lands without updating at least one `docs/reliability/*` file
- `@bilig/contracts` is the only source of truth for runtime payload shapes
- new product code may not read ambient env, clock, random, timer, fetch, or websocket APIs directly
- if a behavior is not described in this doc set, it is not part of the rewrite contract
