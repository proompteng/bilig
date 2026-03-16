# Agent API

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

## Current tranche

`@bilig/agent-api` now defines the shared frame contracts and stdio framing helpers. The sync server exposes a remote ingress for those frames, with live worksheet execution still to be wired in a follow-up tranche.
