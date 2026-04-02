# Local-First Realtime Loop

## Status

Archived summary of the old cutover discussion. The current shipped loop is worker-first UX on top of a server-authoritative monolith + Zero data plane.

## Current loop

1. The browser worker mounts immediately from local persistence and cache state.
2. The browser consumes Zero-backed viewport patches through `ZeroWorkbookBridge`.
3. The monolith owns mutation application, recalculation, checkpointing, and relational materialization.
4. Zero replicates relational source and eval rows back to the browser.

## Current proof points

- [apps/web/src/WorkerWorkbookApp.tsx](/Users/gregkonush/github.com/bilig/apps/web/src/WorkerWorkbookApp.tsx)
- [apps/web/src/zero/ZeroWorkbookBridge.ts](/Users/gregkonush/github.com/bilig/apps/web/src/zero/ZeroWorkbookBridge.ts)
- [apps/bilig/src/zero/recalc-worker.ts](/Users/gregkonush/github.com/bilig/apps/bilig/src/zero/recalc-worker.ts)
- [apps/bilig/src/zero/store.ts](/Users/gregkonush/github.com/bilig/apps/bilig/src/zero/store.ts)

## No longer current

- standalone `apps/local-server`
- standalone `apps/sync-server`
- websocket browser-sync as the default product path
