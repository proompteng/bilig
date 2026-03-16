# `bilig` Canonical Product Design

`bilig` is a local-first spreadsheet system with a browser-native Excel-like shell, a deterministic TypeScript semantic core, an AssemblyScript/WASM compute kernel, CRDT-based realtime collaboration, and a standalone sync backend deployed through the `lab` ArgoCD repo.

## Canonical target

The product target is fixed to these requirements:

- local-first CRDT collaboration is the canonical state model
- the browser loads a WASM binary and runs the engine in a worker-first runtime
- formula semantics target Excel 365 built-in worksheet parity as of `2026-03-15`
- the browser shell must feel native and spreadsheet-first, not dashboard-like
- realtime sync uses a binary transport, not JSON, on the hot path
- both local stdio and remote network agent APIs are first-class products
- deployment is a standalone `bilig` Argo product app in `/Users/gregkonush/github.com/lab`

## Package and app map

- `@bilig/core`: semantic spreadsheet authority, calc graph, selections, snapshots, sync hooks
- `@bilig/formula`: Excel grammar, binding, optimization, lowered plans, coercion/error semantics
- `@bilig/wasm-kernel`: AssemblyScript hot-path evaluator
- `@bilig/crdt`: deterministic LWW batch convergence
- `@bilig/grid`: Excel-like browser shell on Glide
- `@bilig/renderer`: declarative workbook DSL
- `@bilig/binary-protocol`: canonical binary sync framing
- `@bilig/worker-transport`: UI thread <-> worker engine bridge
- `@bilig/agent-api`: stdio and remote agent frame contracts
- `@bilig/storage-browser`: IndexedDB-backed local durability
- `@bilig/storage-server`: durable log/snapshot ownership abstractions
- `@bilig/excel-fixtures`: checked-in Excel parity goldens
- `apps/playground`: browser product shell and integration harness
- `apps/sync-server`: realtime sync and remote API service

## Runtime split

- UI thread owns painting, input, formula bar, editor overlays, clipboard, and shell chrome.
- Browser worker owns the engine, CRDT state, WASM kernel, persistence, and sync connection.
- The backend owns binary websocket ingress, session ownership, durable append log, snapshots, and replay/catch-up.

## Production constraints

- JS evaluation remains the semantic oracle.
- WASM only accelerates profitable overlap sets and must preserve exact JS parity.
- Local and remote mutations use the same deterministic apply path.
- Browser state must recover from refresh, offline work, and reconnect without losing local edits.
- Server acks happen only after durable append.
- Every architecture promise must be enforced by tests, benchmarks, or an explicit acceptance matrix gate.

## Canonical docs

- [architecture.md](/Users/gregkonush/github.com/bilig/docs/architecture.md)
- [public-api.md](/Users/gregkonush/github.com/bilig/docs/public-api.md)
- [formula-language.md](/Users/gregkonush/github.com/bilig/docs/formula-language.md)
- [crdt-model.md](/Users/gregkonush/github.com/bilig/docs/crdt-model.md)
- [browser-runtime.md](/Users/gregkonush/github.com/bilig/docs/browser-runtime.md)
- [binary-protocol.md](/Users/gregkonush/github.com/bilig/docs/binary-protocol.md)
- [backend-sync-service.md](/Users/gregkonush/github.com/bilig/docs/backend-sync-service.md)
- [agent-api.md](/Users/gregkonush/github.com/bilig/docs/agent-api.md)
- [performance-budgets.md](/Users/gregkonush/github.com/bilig/docs/performance-budgets.md)
- [production-acceptance-matrix.md](/Users/gregkonush/github.com/bilig/docs/production-acceptance-matrix.md)

## Current tranche status

This repo now contains the first production-foundation tranche for the canonical design:

- canonical docs rewritten to the production target
- binary frame codec package added
- worker transport package added
- browser persistence extracted into a reusable package
- server-side storage abstractions added
- sync-server skeleton added with binary HTTP ingress for sync frames and remote agent frames

The system is not yet at full Excel parity or full realtime production readiness. Those remain active implementation tranches, but the repo no longer describes them as out-of-scope future ideas.
