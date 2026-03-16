# Browser Runtime

## Current state

- persistence is already extracted into `@bilig/storage-browser`
- worker transport exists and is tested in isolation
- the shipping browser shell still runs the engine in-process today
- sync state now exists in `@bilig/core`, but the UI does not yet consume it through a worker-backed runtime

## Canonical boot flow

1. restore workbook snapshot from IndexedDB
2. restore replica snapshot and outbound queue
3. initialize worker transport
4. initialize WASM in the worker
5. connect the sync client
6. replay local and remote deltas
7. surface sync state in the shell

## Responsibilities

- UI thread: paint, input, clipboard, overlays, formula bar
- worker: engine, formula execution, CRDT, WASM, persistence, sync

## Target state

- `apps/web` is worker-first by default
- the UI thread owns only shell and input concerns
- the worker owns engine, CRDT, WASM, persistence, and live sync
- offline restore, reconnect, and sync state are visible product behavior rather than internal helpers

## Exit gate

- the production app boots through a worker-backed engine
- live edits, restore, and reconnect flows pass browser tests without falling back to in-process engine execution
- the browser shell surfaces sync state sourced from the worker runtime
