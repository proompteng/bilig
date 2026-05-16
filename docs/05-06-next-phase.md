# 05-06 next phase

This document used to describe a browser-owned SQLite/OPFS workbook store. That direction is retired.

The browser must not maintain a second workbook database beside Zero. Browser durability and local cache ownership belong to Zero client persistence, and the worker runtime should remain an ephemeral projection layer that hydrates from Zero/server authoritative state.

## Current Direction

- Use Zero as the browser persistence and sync layer.
- Keep `apps/bilig` and Postgres as the authoritative mutation, event, and projection source.
- Keep the worker runtime focused on projection, viewport patches, optimistic in-memory mutation replay, undo/redo, and rendering support.
- Rehydrate worker state from authoritative snapshots/events instead of a browser SQLite database.
- Do not add another OPFS, SQLite/WASM, IndexedDB workbook store, outbox, tile database, or schema migration layer in `apps/web`.

## Runtime Path

```text
UI shell
  -> Zero client state and server authoritative routes
  -> runtime worker ephemeral projection
  -> viewport/tile patch publishers
  -> renderer
```

Accepted durability lives in the server/Zero path:

```text
semantic mutation
  -> apps/bilig Zero mutation route
  -> Postgres workbook event/projection tables
  -> Zero client sync
  -> worker authoritative hydrate/reconcile
```

## Explicit Non-goals

- No `@bilig/storage-browser` package.
- No browser SQLite/WASM workbook database.
- No OPFS workbook cache owned by Bilig.
- No local pending-op journal outside Zero/server authoritative sync.
- No raw SQLite page/file replication through Zero.

## Validation

Use these gates for this direction:

- `pnpm lint`
- `pnpm typecheck`
- `pnpm --filter @bilig/web build`
- `bun scripts/release-check.ts`
- `pnpm exec vitest run apps/web/src/__tests__/worker-runtime-reconnect.test.ts apps/web/src/__tests__/worker-runtime-authoritative-bootstrap.test.ts apps/web/src/__tests__/worker-runtime-mutation-journal.test.ts scripts/__tests__/reliability-scorecard.test.ts`
