# Cutover and Acceptance

## Cutover Policy

This rewrite is a clean-break migration of the control plane.

- new routes land as `v2`
- web, worker, sync, and local surfaces are updated together
- old `v1` routes are removed once the `v2` path is green
- duplicated fallback paths are deleted in the same cutover wave

## Required Deletions

- legacy `v1` session/documents/agent/ws control-plane routes
- route-local guard code duplicated outside `@bilig/contracts`
- ambient fetch/websocket/time/random access in application code
- ad hoc boot and reconnect flows driven only by React effects or mutable service state

## Test Gates

- unit tests for every contract schema
- unit tests for every actor state machine
- deterministic Effect tests for retry, timeout, and interruption behavior
- browser tests for actor-driven boot, reconnect, and degraded-state UX
- fuzz tests for contract decode, websocket ordering, snapshot assembly, and browser shortcut flows

## Reliability SLOs

- session bootstrap failure is explicit and user-visible
- reconnect transitions are bounded and observable
- resource teardown leaves no live subscriptions or orphan queues
- service failures are attributable by tagged error type

## Failure Drills

- websocket disconnect during sync
- snapshot fetch failure during boot
- malformed request payload at every `v2` ingress
- agent frame decode failure
- shutdown during inflight async work
