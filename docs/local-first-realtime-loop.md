# Local-First Realtime Loop

## Current state

- the repo now has all three runtime building blocks: browser shell, local workbook authority, and remote sync server.
- `apps/local-server` is the more complete runtime today: it hosts live workbook sessions, owns an in-memory engine, and emits committed binary batch frames to browsers and agents.
- `apps/sync-server` already handles ordered sync frames, snapshots, cursors, and presence, but many worksheet-level agent calls are still reserved or `NOT_IMPLEMENTED` unless an executor is injected.
- the browser worker and local-server loop are not yet unified into the final default boot path.
- the browser still does not consume a dedicated derived viewport patch stream; the runtime is ahead of the UI transport model.

## Stream model

The long-term loop should separate three different traffic classes:

1. durable workbook ops
2. derived viewport/render patches
3. ephemeral presence and awareness events

The current binary batch stream is the beginning of the first layer. The next runtime phases should avoid collapsing all three concerns into one packet family.

## Target state

1. browser restores local snapshot and queue
2. browser initializes worker and WASM
3. browser connects to the local agent server over localhost websocket
4. browser catches up by cursor from the authoritative workbook op stream
5. worker derives viewport/render patches for the visible browser surface
6. local user and local agent mutations commit through one ordered workbook stream
7. committed mutations render immediately in the browser from derived patches rather than ad hoc in-process coupling
8. committed mutations relay upstream to the remote sync backend for durability and cross-device fanout
9. presence and awareness events travel separately from the durable workbook log

## Exit gate

- local edit and local agent loops both converge on the same authoritative workbook op stream
- the browser renders visible updates from worker-derived viewport patches
- cursor catch-up is proven after reconnect
- remote replay converges with the same local commit stream without semantic drift
- presence/cursor channels stay separate from durable workbook history
