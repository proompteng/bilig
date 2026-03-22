# Agent API

## Current state

- `@bilig/agent-api` defines the shared request/response/event model and stdio framing helpers.
- `apps/local-server` now executes the canonical worksheet mutation requests against live local workbook sessions.
- the remote sync server accepts agent ingress frames, but live worksheet execution is still incomplete.
- agent frames are still serialized through JSON payload bodies inside a binary envelope; this is an explicit interim state, not the target wire format.
- range subscription and chat-stream agent events are still open work; the current local server tranche focuses on live worksheet request/response execution.

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

- stdio and remote transports use the same typed binary request/response/event schema
- every documented agent operation executes against a live worksheet session
- local agent chat messages fan into the same ordered workbook commit stream as UI and replay mutations
- remote agent requests are authenticated, tenant-scoped, and idempotent where required

## Exit gate

- stdio and remote agent conformance tests return identical results for the same worksheet operations
- the local app server executes canonical read/write worksheet operations against live workbook sessions
- remote ingress no longer returns placeholder `NOT_IMPLEMENTED` responses for canonical worksheet mutations
- the wire format used by agents is binary end to end, not JSON-inside-binary

## See also

- [typed-agent-protocol-rfc.md](/Users/gregkonush/github.com/bilig/docs/typed-agent-protocol-rfc.md)
- [authoritative-workbook-op-model-rfc.md](/Users/gregkonush/github.com/bilig/docs/authoritative-workbook-op-model-rfc.md)
