# Browser Runtime

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

## Current tranche

- `@bilig/storage-browser` now owns reusable browser durability helpers
- `@bilig/worker-transport` now provides a concrete request/subscription bridge
- the playground still needs to switch fully to worker-first execution in a follow-up tranche
