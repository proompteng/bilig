# Agent API

## Current state

- `@bilig/agent-api` defines the shared request/response/event model and stdio framing helpers.
- the remote sync server accepts agent ingress frames, but live worksheet execution is still incomplete.
- agent frames are still serialized through JSON payload bodies inside a binary envelope; this is an explicit interim state, not the target wire format.

## Canonical goal

Support both:

- local low-overhead stdio agent control
- remote authenticated network agent control

using the same typed request/response/event model.

## Core operations

- open and close worksheet sessions
- read cells and ranges
- write values and formulas
- clear/fill/copy/paste ranges
- inspect precedents and dependents
- subscribe to range changes
- import and export snapshots
- query metrics and traces
- execute batched worksheet mutations with idempotency keys

## Transports

- stdio: length-prefixed binary frames
- remote: binary request frames over HTTP and websocket

## Target state

- stdio and remote transports use the same typed binary request/response/event frames
- every documented agent operation executes against a live worksheet session
- remote agent requests are authenticated, tenant-scoped, and idempotent where required

## Exit gate

- stdio and remote agent conformance tests return identical results for the same worksheet operations
- remote ingress no longer returns placeholder `NOT_IMPLEMENTED` responses for canonical worksheet mutations
- the wire format used by agents is binary end to end, not JSON-inside-binary
