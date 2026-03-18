# Browser Runtime

## Current state

- persistence is already extracted into `@bilig/storage-browser`
- worker transport exists and is tested in isolation
- `apps/web` exists as the shipping browser shell wrapper
- the shipping browser shell still runs the engine in-process today
- the browser can now target a local app server runtime, but the worker-first path is not the default execution path yet
- sync state now exists in `@bilig/core`, but the UI does not yet consume it through a worker-backed runtime

## Canonical boot flow

1. restore workbook snapshot from IndexedDB
2. restore replica snapshot and outbound queue
3. initialize worker transport
4. initialize WASM in the worker
5. connect the local app server or sync client
6. replay local and remote deltas
7. surface sync state in the shell

## Responsibilities

- UI thread: paint, input, clipboard, overlays, formula bar
- worker: engine, formula execution, CRDT, WASM, persistence, sync

## Target state

- `apps/web` is worker-first by default
- the UI thread owns only shell and input concerns
- the worker owns engine, CRDT, WASM, persistence, and live sync
- the browser defaults to the local app server loop when available and falls back to local-only worker execution when it is not
- offline restore, reconnect, and sync state are visible product behavior rather than internal helpers

## Exit gate

- the production app boots through a worker-backed engine
- the production app can reconnect to the local app server and catch up by cursor
- live edits, restore, and reconnect flows pass browser tests without falling back to in-process engine execution
- the browser shell surfaces sync state sourced from the worker runtime
