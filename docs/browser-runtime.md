# Browser Runtime

## Current state

- persistence is extracted into `@bilig/storage-browser`
- worker transport exists and is used by the shipping browser shell
- `apps/web` boots a dedicated worker-backed runtime by default
- the worker owns engine boot, persistence restore, WASM lifecycle, sync connectivity, and viewport patch derivation
- the UI consumes worker-derived viewport patches and worker-sourced sync state
- the browser can target a local app server runtime and fall back to local-only worker execution
- current browser-runtime gaps:
  - `apps/web` no longer depends on legacy demo CSS or sources
  - the viewport patch payload is JSON inside a byte envelope rather than a dedicated typed binary codec
  - remote catch-up and durable multiplayer depend on unfinished sync-server work

## Canonical boot flow

1. restore workbook snapshot from browser persistence when enabled
2. restore replica snapshot and local state
3. initialize worker transport
4. initialize engine and WASM in the worker
5. connect the local app server when configured
6. replay local and remote deltas
7. surface sync state in the shell

## Responsibilities

- UI thread: paint, input, clipboard, overlays, formula bar, and shell composition
- worker: engine, formula execution, authoritative ops, persistence, sync, WASM, and viewport patch derivation

## Preconditions

The browser runtime should not become the place where workbook semantics are re-invented. The remaining browser work depends on:

- the authoritative workbook op model fully covering the remaining workbook mutation surface
- workbook metadata such as names, tables, structured references, and spill ownership being first-class for `names:defined-name-range`, `tables:table-total-row-sum`, and `structured-reference:table-column-ref`
- the sync service growing into a durable remote worksheet host rather than a thin ingress layer

## Target state

- `apps/web` remains worker-first by default
- the UI thread owns only shell and input concerns
- the worker owns engine, authoritative ops, WASM, persistence, and live sync
- the worker derives viewport and render patches for visible ranges instead of exposing raw engine state directly to the grid
- the browser defaults to the local app server loop when available and falls back cleanly to local-only worker execution when it is not
- `apps/web` no longer depends on deprecated demo source or styling
- viewport patch payloads and agent payloads use typed codecs instead of JSON-in-binary wrappers

## Exit gate

- the production app boots through a worker-backed engine
- the production app can reconnect to the local app server and catch up by cursor
- live edits, restore, and reconnect flows pass browser tests without in-process engine fallback
- visible workbook updates are painted from worker-derived viewport patches
- the browser shell surfaces sync state sourced from the worker runtime
- the browser shell does not depend on retired demo-shell source or styling

## See also

- [architecture.md](/Users/gregkonush/github.com/bilig/docs/architecture.md)
- [durable-multiplayer-replication-rfc.md](/Users/gregkonush/github.com/bilig/docs/durable-multiplayer-replication-rfc.md)
