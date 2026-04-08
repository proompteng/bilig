# Replica state ownership cleanup
## Date: 2026-04-07
## Scope: remove `@bilig/crdt` as a package boundary without regressing local replay semantics

## Why this document exists

`bilig` no longer runs a CRDT-first product architecture. The production path is:

- server-authoritative ordering in `apps/bilig`
- Zero as a narrow relational sync layer
- worker-first browser execution with durable local state

That architecture is real, but the repo still carries a stale package seam:

- `@bilig/crdt` still owns replica-state helpers and re-exports workbook op types
- `@bilig/core` imports replica helpers from `@bilig/crdt`
- app code still calls `shouldApplyBatch(session.engine.replica, batch)` directly

This is not just naming debt. It leaves the wrong ownership boundary in production code.

## Problem statement

The current shape mixes three different concerns:

1. `@bilig/workbook-domain` owns transport-neutral semantic workbook ops
2. `@bilig/core` owns the runtime engine that actually applies and persists those ops
3. `@bilig/crdt` still owns replica bookkeeping that is only meaningful because `@bilig/core` and local session code use it

That split was useful during earlier migration phases. It is no longer a good production architecture.

Today:

- Zero owns authoritative shared ordering, persistence, and relational projection
- the engine still owns local replay idempotence and replica snapshot import/export
- app code still knows too much about engine internals

The result is half-migrated design:

- the package name says "CRDT" even though the product is not CRDT-authoritative anymore
- app code can read and reason about `engine.replica`
- docs say `@bilig/crdt` is compatibility-only while core runtime behavior still depends on it

## Current repo-grounded evidence

- `packages/core/src/engine.ts` imports `createReplicaState`, `hydrateReplicaState`, `exportReplicaSnapshot`, and `shouldApplyBatch` from `@bilig/crdt`
- `packages/core/src/engine.ts` exposes `exportReplicaSnapshot()`, `importReplicaSnapshot()`, and `applyRemoteBatch()`
- `apps/bilig/src/workbook-runtime/local-workbook-session-manager.ts` imports `shouldApplyBatch` directly from `@bilig/crdt`
- `apps/bilig/src/workbook-runtime/runtime-manager.ts` restores engine replica snapshots on server session bootstrap
- `apps/web/src/worker-runtime.ts` persists and restores replica snapshots for worker-first local execution

This proves the underlying semantics are still needed. What is wrong is where they live and who is allowed to call them.

## Current public contract inventory

### Runtime data shape that already crosses boundaries

`EngineReplicaSnapshot` is already the persisted and cross-runtime contract. Today it contains:

- `replica`
- `entityVersions`
- `sheetDeleteVersions`

That shape is defined in `packages/core/src/engine.ts` and validated by `isEngineReplicaSnapshot` in `packages/core/src/guards.ts`.

It currently crosses these boundaries:

- worker local-store persistence in `apps/web/src/worker-runtime.ts` and `packages/storage-browser/src/workbook-local-store.ts`
- worker engine bootstrap via `apps/web/src/worker-runtime-support.ts`
- monolith local-session snapshot state in `apps/bilig/src/workbook-runtime/local-session-snapshot-store.ts`
- monolith authoritative restore and backfill in `apps/bilig/src/zero/store.ts`

This matters because the cleanup must preserve the serialized shape through the refactor. Package ownership may change; snapshot compatibility must not.

### Current engine API that callers rely on

Public methods already relied on across the repo:

- `exportReplicaSnapshot()`
- `importReplicaSnapshot(snapshot)`
- `applyRemoteBatch(batch)`

Current problem:

- `applyRemoteBatch(batch)` is not sufficient as the full caller API because app code still reaches into `engine.replica` and separately calls `shouldApplyBatch(...)`

### Current config and docs that still keep `@bilig/crdt` alive

- `packages/core/package.json`
- `apps/bilig/package.json`
- `apps/bilig/tsconfig.json`
- `apps/web/vite.config.ts`
- `packages/workbook-domain/README.md`
- `packages/crdt/README.md`
- `docs/crdt-model.md`

The cleanup is not complete until all of those stop treating `@bilig/crdt` as an active architecture seam.

## Goals

- delete `packages/crdt`
- make `@bilig/core` the sole owner of replica bookkeeping behavior
- remove all app-layer imports of replica helpers
- remove all app-layer reads of `engine.replica`
- preserve current local replay, duplicate-batch suppression, and snapshot restore behavior
- keep Zero responsible only for authoritative shared ordering and materialization

## Non-goals

- moving replica bookkeeping into Zero
- removing replica snapshots from worker or monolith restore flows
- changing workbook op semantics in `@bilig/workbook-domain`
- redesigning the server-authoritative event model
- changing browser/local-first persistence strategy

## Design principles

- semantic ownership matters more than package deletion
- app code should call engine APIs, not manipulate engine internals
- `workbook-domain` should stay transport-neutral and declarative
- `core` should own mutable runtime bookkeeping needed to apply ops safely
- Zero should not become a dumping ground for local runtime concerns

## Target ownership

### `@bilig/workbook-domain`

Owns:

- `EngineOp`
- `EngineOpBatch`
- other transport-neutral workbook operation shapes

Does not own:

- mutable replica state
- batch idempotence bookkeeping
- engine entity version maps

### `@bilig/core`

Owns:

- replica state structures
- op ordering helpers needed by the engine
- snapshot import/export of engine replica state
- duplicate-batch suppression
- remote batch application behavior

Exposes:

- `exportReplicaSnapshot()`
- `importReplicaSnapshot(snapshot)`
- `applyRemoteBatch(batch): boolean`

The return value from `applyRemoteBatch` must become the public "was this batch newly applied?" contract so app code no longer calls `shouldApplyBatch(...)` itself.

Does not expose:

- mutable `replica` state as a public field
- public helper APIs that allow app code to reimplement engine dedupe policy externally

### `apps/web` and `apps/bilig`

May:

- restore replica snapshots through engine APIs
- apply remote batches through engine APIs

May not:

- import replica helpers directly
- inspect or mutate internal replica state

### Zero / monolith authoritative path

Owns:

- authoritative event ordering
- durable shared storage
- relational projection and recalculation

Does not own:

- worker-local replay idempotence
- local session duplicate suppression logic inside the engine

## Target code shape

Introduce an internal core module, for example:

- `packages/core/src/replica-state.ts`

Move the remaining logic from `packages/crdt/src/index.ts` into core:

- replica state types
- replica snapshot helpers
- op order helpers
- batch dedupe helpers

Then tighten the engine API:

- keep replica state private to `SpreadsheetEngine`
- make `applyRemoteBatch(batch)` return `true` when applied and `false` when ignored as already known
- keep `exportReplicaSnapshot()` and `importReplicaSnapshot()` as the only supported snapshot boundary

### Current-to-target API table

| Concern | Current state | Target state |
| --- | --- | --- |
| replica bookkeeping owner | `@bilig/crdt` package | internal `@bilig/core` module |
| app-layer duplicate check | `shouldApplyBatch(session.engine.replica, batch)` | `session.engine.applyRemoteBatch(batch)` boolean result |
| engine internal replica access | `engine.replica` public field | replica state private to engine |
| snapshot import/export | engine methods | same engine methods, same serialized shape |
| transport-neutral op shapes | `@bilig/workbook-domain` | unchanged |
| authoritative shared ordering | monolith + Zero | unchanged |

App code should move from:

```ts
if (!shouldApplyBatch(session.engine.replica, appendFrame.batch)) {
  // duplicate
}
session.engine.applyRemoteBatch(appendFrame.batch);
```

to:

```ts
const applied = session.engine.applyRemoteBatch(appendFrame.batch);
if (!applied) {
  // duplicate
}
```

That is the actual architecture win. The package deletion is just the cleanup that follows.

## Compatibility contract

The cleanup must preserve the following compatibility rules:

### Serialized snapshot compatibility

- the shape of `EngineReplicaSnapshot` remains readable by existing persisted worker and monolith state
- the shape accepted by `isEngineReplicaSnapshot` remains stable during the refactor
- existing persisted `replica_snapshot` payloads in Postgres remain importable through `engine.importReplicaSnapshot(...)`

### Behavioral compatibility

- replaying the same remote batch twice must still be idempotent
- restoring an engine from snapshot + replica snapshot must still suppress already-known remote batches
- sheet tombstone behavior must remain unchanged when stale cell batches replay after delete
- worker restore before authoritative hydrate must still produce the same local-first behavior

### Explicitly allowed change

- app code is no longer allowed to inspect `engine.replica` directly

## Why Zero is not the replacement

Zero already replaced the old shared-authority role.
It did not replace local runtime bookkeeping.

Replica bookkeeping is still needed for:

- durable worker restore before full authoritative hydrate
- monolith local session replay
- duplicate suppression when a batch is seen more than once
- deterministic engine import/export of previously applied batch ids and entity versions

Those are engine concerns.
They must not be reinterpreted as Zero concerns just because the old package name is now misleading.

## Migration plan

### Phase 1: internalize behavior in core

- add `packages/core/src/replica-state.ts`
- move `ReplicaState`, `ReplicaSnapshot`, `OpOrder`, and related helpers into core
- update `packages/core/src/engine.ts` to import from the new internal module
- update core tests to prove snapshot round-trip and duplicate-batch suppression still hold

Validation gate:

- `packages/core/src/__tests__/engine.test.ts`
- `packages/core/src/__tests__/guards.test.ts`

### Phase 2: remove app knowledge of replica internals

- change `SpreadsheetEngine.applyRemoteBatch(batch)` to return a boolean
- update `apps/bilig/src/workbook-runtime/local-workbook-session-manager.ts` to rely on that return value
- confirm no app code reads `engine.replica` directly

Validation gate:

- targeted local session tests under `apps/bilig/src/workbook-runtime/*.test.ts`
- worker runtime restore tests under `apps/web/src/__tests__/worker-runtime.test.ts`

### Phase 3: remove the package seam

- remove `@bilig/crdt` dependencies from `packages/core/package.json` and `apps/bilig/package.json`
- remove `@bilig/crdt` aliases from `apps/bilig/tsconfig.json` and `apps/web/vite.config.ts`
- remove the `../../packages/crdt` project reference from `apps/bilig/tsconfig.json`
- update docs and READMEs that still describe `@bilig/crdt` as a compatibility layer
- delete `packages/crdt`

Cut list:

- `packages/core/package.json`
- `apps/bilig/package.json`
- `apps/bilig/tsconfig.json`
- `apps/web/vite.config.ts`
- `packages/workbook-domain/README.md`
- `packages/crdt/README.md`
- `docs/crdt-model.md`

### Phase 4: verify production behavior

- run `pnpm typecheck`
- run `pnpm test`
- run `pnpm coverage`
- run `pnpm bench:smoke`
- run the Forgejo-equivalent non-browser checks
- run targeted browser/runtime tests that exercise worker restore and remote batch replay

Targeted suites:

- `packages/core/src/__tests__/engine.test.ts`
- `packages/core/src/__tests__/guards.test.ts`
- `apps/web/src/__tests__/worker-runtime.test.ts`
- `apps/web/src/__tests__/runtime-session.test.ts`
- `apps/bilig/src/workbook-runtime/agent-routing.test.ts`
- `apps/bilig/src/workbook-runtime/document-presence-session-store.test.ts`
- `apps/bilig/src/workbook-runtime/document-session-manager.test.ts`
- `apps/bilig/src/workbook-runtime/document-supervisor-shared.test.ts`
- `apps/bilig/src/workbook-runtime/local-agent-session-store.test.ts`
- `apps/bilig/src/workbook-runtime/local-workbook-session-manager.test.ts`
- `apps/bilig/src/workbook-runtime/workbook-session-core.test.ts`
- `apps/bilig/src/workbook-runtime/workbook-session-shared.test.ts`

## Acceptance criteria

The cleanup is complete when all of the following are true:

- no source file imports `@bilig/crdt`
- no app code reads `engine.replica`
- `SpreadsheetEngine` is the only public owner of replica apply/dedupe behavior
- `packages/crdt` no longer exists
- persisted worker and monolith `EngineReplicaSnapshot` payloads still deserialize without migration
- worker runtime restore still passes
- monolith local session replay still passes
- duplicate remote batches are still acknowledged without double-application
- docs no longer describe `@bilig/crdt` as current architecture

## Risks

### Risk: changing `applyRemoteBatch` semantics breaks callers

Mitigation:

- add focused tests for duplicate batch handling before refactor
- update call sites in one change
- keep method name stable and only tighten its contract

### Risk: snapshot import/export regressions break restore

Mitigation:

- keep the serialized snapshot shape unchanged during the refactor
- reuse existing worker and runtime-manager tests

### Risk: accidental broadening of `core` public API

Mitigation:

- move helpers into internal modules first
- only expose engine methods that product code actually needs

## Out of scope but adjacent cleanup

This document does not cover these follow-on cleanup items, even though they are real:

- renaming or removing the legacy-named `replaceSnapshot` worker RPC
- deleting legacy JSON viewport-patch fallback if no old clients remain
- moving one-time startup backfills out of steady-state service initialization

Those are separate cleanup tracks. They should not block replica-state ownership cleanup.

## Relation to existing docs

- this document supersedes `docs/crdt-model.md` as the statement of current ownership direction
- it does not replace the broader product architecture docs in `docs/design.md` or `docs/architecture.md`
- it is compatible with the current production contract described in `docs/05-06-next-phase.md`
- it assumes the authoritative runtime contract already documented in `docs/authoritative-workbook-op-model-rfc.md`

## Recommended implementation order

1. add tests around engine duplicate-batch behavior and replica snapshot round-trip
2. internalize replica helpers into core
3. change app callers to use engine APIs only
4. make `engine.replica` private or otherwise unreachable from app code
5. remove package references and aliases
6. delete `packages/crdt`
7. update docs
8. run full validation

## Exit gate

This cleanup is done when the repo no longer has a separate `@bilig/crdt` layer and no production path relies on app-layer access to replica internals, while Zero remains the shared authoritative system and `@bilig/core` owns local replay semantics cleanly.
