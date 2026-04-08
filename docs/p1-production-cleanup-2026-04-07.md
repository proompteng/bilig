# P1 production cleanup plan
## Date: 2026-04-07
## Scope: remove boot-path migrations, centralize workspace resolution, and rename the authoritative snapshot worker API

## Status

Ready for implementation.

This document covers the remaining `P1` production debt found in the April 7 repo scan:

- heavy data migration work still running during Zero service startup
- duplicated workspace-package resolution config across TypeScript, Vite, and Vitest
- a legacy-named `replaceSnapshot` worker RPC still used by the shipped browser bootstrap path

These items are grouped together because they are all the same class of problem:

- compatibility or migration seams that are still sitting on the live product path
- behavior and naming that no longer match the architecture described in the repo docs
- repeated configuration that can drift and already has drifted in practice

## Why this document exists

`bilig` is now much closer to a production architecture than an experimental one:

- `apps/bilig` is the shipped runtime
- the browser is worker-first and Zero-backed
- the monolith owns authoritative event ordering and persistence
- package boundaries are increasingly being tightened around semantic ownership

But three high-impact seams are still wrong:

1. the Zero service still performs expensive migration/backfill work every time it starts
2. workspace package resolution is declared in multiple independent configs
3. the browser product path still uses a worker RPC named `replaceSnapshot` even though the code is performing authoritative hydrate/recovery, not debug snapshot replacement

None of these are correctness bugs every day. All three are production-shape problems:

- they make the live path slower or harder to reason about
- they make upgrades riskier
- they create drift between docs and code

## Repo-grounded evidence

### Startup migration work still runs on boot

- `apps/bilig/src/zero/service.ts` calls `backfillAuthoritativeCellEval(this.pool)` during `initialize()`
- `apps/bilig/src/zero/service.ts` also calls `dropLegacyZeroSyncSchemaObjects(this.pool)` in the same boot path
- `apps/bilig/src/zero/workbook-migration-store.ts` rebuilds workbook source projections and `cell_eval` rows by:
  - scanning for stale projection versions
  - scanning for stale render rows
  - scanning legacy `sheet_style_ranges` and `sheet_format_ranges`
  - reconstructing workbook engines from stored checkpoints

This is a real data migration/repair workflow sitting in the same path that starts the recalc worker.

### Workspace resolution is duplicated

Current resolution sources:

- root TypeScript paths in `tsconfig.json`
- app-level TypeScript paths in `apps/web/tsconfig.json`
- app-level TypeScript paths in `apps/bilig/tsconfig.json`
- Vite aliases in `apps/web/vite.config.ts`
- Vitest aliases in `vitest.config.ts`

At the same time, workspace packages still publish `dist` entrypoints in package manifests such as `packages/core/package.json`.

This means local tooling must repeatedly override package export behavior just to point tests and apps at source.

### `replaceSnapshot` is still on the product path

- `apps/web/src/runtime-session.ts` calls worker method `replaceSnapshot` during initial authoritative hydrate
- the same file special-cases `replaceSnapshot` in generic `invoke(...)` post-refresh logic
- `apps/web/src/worker-runtime.ts` still exposes `replaceSnapshot(...)`
- `docs/browser-runtime.md` says `replaceSnapshot` is legacy-only debug plumbing, not the product write path

The docs and the runtime disagree.

## Goals

- remove one-shot data migration/backfill work from the normal Zero service startup path
- make workspace package resolution derive from one source of truth
- remove `replaceSnapshot` from the shipped browser bootstrap path
- preserve current runtime behavior while tightening naming and ownership
- reduce future config drift and migration regressions

## Non-goals

- redesigning the server-authoritative event model
- changing the Zero schema or query model beyond what is required to extract migrations
- changing package publish entrypoints away from `dist`
- replacing the viewport-patch transport in this document
- removing the `/v2/documents/:documentId/snapshot/latest` endpoint in this document

## Design principles

- heavy compatibility work should not live in the normal product boot path
- migration execution must be explicit, observable, and idempotent
- workspace package identity should come from package metadata, not repeated manual alias blocks
- API names should describe the actual production semantics
- internal cross-repo compatibility shims should be deleted once the repo can migrate in one coordinated change

## Current-to-target summary

| Area | Current state | Target state |
| --- | --- | --- |
| Zero startup | boot runs schema ensure, data backfills, and legacy cleanup | boot runs only cheap startup work; data migrations run through an explicit runner |
| Migration contract | hidden inside service init | required migrations complete before app becomes ready |
| Workspace resolution | repeated manually across TS, Vite, and Vitest | generated once from workspace metadata and consumed everywhere |
| Worker snapshot RPCs | `replaceSnapshot(...)` and `rebaseToSnapshot(...)` are public worker methods | one public `installAuthoritativeSnapshot(...)` method with explicit mode semantics |
| Docs/code alignment | docs call `replaceSnapshot` legacy-only while product uses it | docs and code describe the same authoritative hydrate path |

## Workstream A: remove data migrations from Zero service startup

### Current problem

`EnabledZeroSyncService.initialize()` currently mixes four concerns:

1. schema creation / schema evolution
2. boot-time service initialization
3. one-time data migrations
4. legacy cleanup

Only the first two belong on every startup.

The current boot sequence runs heavy migration logic before the service becomes steady:

- ensure schema tables/columns exist
- backfill authoritative source projection and `cell_eval`
- backfill workbook changes
- drop legacy schema objects
- start recalc worker

That has three production problems:

- startup latency depends on historical data shape instead of current runtime needs
- every pod restart can pay migration detection costs
- rollout safety is muddied because repair logic and service boot logic are coupled

### Target state

Split the current behavior into three layers:

#### 1. boot path

Always runs on startup:

- schema/table/column/index ensure calls
- cheap invariant checks
- recalc worker start

Must be:

- fast
- idempotent
- proportional to current runtime setup, not historical workbook volume

#### 2. explicit data migration runner

Runs versioned one-time data migrations behind a migration ledger and advisory lock.

Owns:

- backfilling source projections
- rebuilding `cell_eval`
- legacy table retirement
- any future rewrite that scans existing workbook rows

#### 3. repair or admin tasks

Owns post-migration repair workflows that may be re-runnable by operators or test harnesses.

These should not be silently hidden inside normal service startup.

### File inventory

Primary implementation files:

- `apps/bilig/src/index.ts`
- `apps/bilig/src/zero/service.ts`
- `apps/bilig/src/zero/workbook-migration-store.ts`
- `apps/bilig/src/zero/workbook-change-store.ts`
- `apps/bilig/src/zero/store.ts`

Primary validation files:

- `apps/bilig/src/zero/__tests__/workbook-migration-store.test.ts`
- `apps/bilig/src/zero/__tests__/workbook-change-store.test.ts`
- `apps/bilig/src/http/sync-server.test.ts`

### Proposed design

#### Migration ledger

Add a dedicated ledger table, for example:

- `bilig_data_migration`

Suggested shape:

- `name TEXT PRIMARY KEY`
- `applied_at TIMESTAMPTZ NOT NULL`
- `code_version TEXT NOT NULL`
- `details JSONB NOT NULL DEFAULT '{}'::jsonb`

This is separate from schema ensure logic. It records data migrations, not SQL DDL state.

#### Migration runner module

Introduce a small orchestrator, for example:

- `apps/bilig/src/zero/data-migration-runner.ts`

Responsibilities:

- ensure the migration ledger table exists
- acquire a PostgreSQL advisory lock so only one runner executes a migration batch
- inspect the ledger
- run pending migrations in a fixed order
- record each migration once it succeeds

This runner becomes the only place allowed to invoke heavy historical backfills.

#### Migration classification

Split current startup work into named migrations:

1. `workbook-source-projection-v2-backfill`
   - current source: `backfillAuthoritativeCellEval(...)` logic in `workbook-migration-store.ts`
   - scans `workbooks.source_projection_version`
   - rebuilds authoritative source projection rows

2. `cell-eval-style-json-backfill`
   - current source: the stale `cell_eval` scan in the same function
   - rebuilds rows where `style_id` is present but `style_json` is missing

3. `legacy-zero-style-format-table-retirement`
   - current source: `dropLegacyZeroSyncSchemaObjects(...)`
   - only runs after backfills are confirmed complete

4. `workbook-change-backfill`
   - current source: `backfillWorkbookChanges(...)`
   - remains explicit instead of hidden in startup

These names should be committed in code and used in logs/tests.

#### Migration classes

Each migration must declare whether it is:

- `required`: the app may not become healthy until it is complete
- `cleanup`: correctness does not depend on it, but rollout automation should still enforce it before the compatibility window closes

Initial classification:

- `workbook-source-projection-v2-backfill`: `required`
- `cell-eval-style-json-backfill`: `required`
- `workbook-change-backfill`: `required`
- `legacy-zero-style-format-table-retirement`: `cleanup`

#### Service startup contract

`EnabledZeroSyncService.initialize()` should stop running heavy migrations directly.

New boot contract:

1. run schema ensure functions
2. ensure migration ledger exists
3. assert that all `required` migrations are complete unless boot-time auto-run is explicitly enabled
4. start the recalc worker

The service should not scan every workbook row or rebuild workbook projections on each startup.

Because `apps/bilig/src/index.ts` awaits `zeroSyncService.initialize()` before Fastify starts listening, a failed migration precondition means the process never serves `/healthz`. That is the correct behavior for pending `required` migrations.

#### Production runtime assumption

This repo should define the runtime assumption, not the Argo implementation:

- before a new `bilig` app version is marked ready, all pending `required` data migrations for that code version must have completed successfully against the target database

Per `docs/bilig-lab-contract.md`, this repo should not prescribe the exact Argo hook or job form. What belongs here is the contract:

- `bilig` provides the migration runner command and required-migration checks
- `lab` decides how the runner is invoked during rollout
- app startup fails fast when required migrations are pending and boot-time auto-run is disabled

#### Release-time execution model

Use an explicit migration execution mode:

- local dev and CI may opt into auto-running data migrations
- production must run the migration runner before the new app version becomes healthy

Suggested control surface:

- `bun scripts/run-zero-data-migrations.ts`
- `BILIG_RUN_DATA_MIGRATIONS_ON_BOOT=true` for dev/test only
- `BILIG_ALLOW_PENDING_CLEANUP_MIGRATIONS=true` only during the compatibility window where cleanup migrations remain optional

This keeps the production boot path clean while preserving convenience for local iteration.

#### Ownership cleanup

After extraction:

- `workbook-migration-store.ts` owns migration and replacement helpers
- `store.ts` owns live runtime persistence APIs
- `service.ts` wires boot concerns, not repair workflows

The current partial extraction already points this way; this change completes it.

### Rollout plan

1. Introduce migration ledger and runner.
2. Move current migration functions behind named migration records.
3. Add a script entrypoint and local/dev auto-run option.
4. Remove direct migration calls from `EnabledZeroSyncService.initialize()`.
5. Add startup assertions so production cannot silently skip `required` migrations.
6. Once migration ledger proves clean in deployed environments, remove dead startup compatibility code that no longer has pending rows to process.

### Risks

- A migration that used to happen implicitly on boot may now be skipped operationally.
- Legacy table retirement must not happen before data backfill is recorded complete.
- Multi-pod deploys must not run the same migration concurrently.

### Validation

- unit tests for migration runner ledger semantics and lock behavior
- migration-store tests for each named migration
- boot tests proving service startup does not invoke heavy migration scans on a clean database
- deploy-playbook documentation proving how migrations run before app readiness

Concrete validation commands and files:

- `pnpm exec vitest run apps/bilig/src/zero/__tests__/workbook-migration-store.test.ts`
- `pnpm exec vitest run apps/bilig/src/zero/__tests__/workbook-change-store.test.ts`
- `pnpm exec vitest run apps/bilig/src/http/sync-server.test.ts`
- `pnpm typecheck`

## Workstream B: centralize workspace package resolution

### Current problem

Workspace packages are identified by package manifests, but source resolution is still restated in multiple tool-specific formats.

Today:

- package manifests publish `dist`
- TypeScript local development uses `paths`
- Vite uses a hand-maintained alias object
- Vitest uses a generated alias list

This creates two operational failure modes:

1. drift between toolchains
2. accidental fallback to package `dist` entrypoints on fresh checkouts

The repo already hit this class of bug with missing aliases for workspace packages in tests.

### File inventory

Primary implementation files:

- `tsconfig.json`
- `apps/web/tsconfig.json`
- `apps/bilig/tsconfig.json`
- `apps/web/vite.config.ts`
- `vitest.config.ts`
- `packages/*/package.json`

Planned new files:

- `scripts/workspace-resolution.ts`
- `scripts/gen-workspace-resolution.ts`
- `workspace-resolution.generated.json`
- `tsconfig.workspace-paths.json`

### Target state

There should be one canonical source of truth for workspace package source resolution:

- package name
- source entrypoint

Everything else should derive from it.

### Ownership rule

- workspace package manifests own package identity
- the generator owns derived alias/path metadata
- TypeScript, Vite, and Vitest may consume generated metadata, but may not define independent package maps by hand

Adding or renaming a workspace package must be a manifest change plus generator rerun, not a multi-file alias scavenger hunt.

### Proposed design

#### Canonical generated artifact

Generate one checked-in data file:

- `workspace-resolution.generated.json`

Suggested shape:

```json
{
  "@bilig/core": {
    "packageDir": "packages/core",
    "sourceEntry": "packages/core/src/index.ts"
  }
}
```

That file is the canonical derived artifact for non-`tsconfig` consumers.

`tsconfig.workspace-paths.json` is a generated projection of the same data for TypeScript consumption, not an independent source of truth.

#### Source of truth

Use workspace package manifests plus conventional source entrypoints:

- `packages/*/package.json` for package name
- `packages/*/src/index.ts` for source entrypoint

This matches the current successful Vitest approach and the repo’s actual package structure.

#### Shared tooling module

Introduce a shared helper, for example:

- `scripts/workspace-resolution.ts`

It should expose:

- `listWorkspacePackages()`
- `createWorkspaceAliasMap()`
- `createViteAliasRecord()`
- `createVitestAliasEntries()`
- `createTsconfigPaths()`

The helper should be the only place that scans `packages/`.

#### Generated TypeScript config

TypeScript cannot execute helper code inside `tsconfig`, so generate a checked-in shared config from the canonical generated artifact:

- `tsconfig.workspace-paths.json`

Suggested contents:

- `compilerOptions.paths` derived from workspace manifests

Then change:

- root `tsconfig.json`
- `apps/web/tsconfig.json`
- `apps/bilig/tsconfig.json`

to extend a common base that includes the generated paths rather than each restating package mappings manually.

#### Vite and Vitest

Replace hand-maintained alias blocks with imports from the shared helper:

- `apps/web/vite.config.ts` should build aliases from `createViteAliasRecord()`
- `vitest.config.ts` should use the same source rather than its own private package scan

This keeps Vite and Vitest aligned automatically.

#### Guardrail

Add a generator/check command, for example:

- `bun scripts/gen-workspace-resolution.ts`
- `bun scripts/gen-workspace-resolution.ts --check`

Then wire a check into CI so a new workspace package or renamed entrypoint cannot land without updating the generated config.

### Explicit non-goal inside this workstream

Do not change package publish behavior in this cleanup.

Packages should continue exporting `dist` for build/publish consumers. The cleanup only changes how in-repo tooling resolves source during development, test, and app bundling.

### Rollout plan

1. Introduce shared workspace-resolution helper.
2. Generate `workspace-resolution.generated.json`.
3. Generate `tsconfig.workspace-paths.json` from the same data.
4. Convert root and app `tsconfig` files to extend the shared config.
5. Convert `apps/web/vite.config.ts` to use the helper.
6. Convert `vitest.config.ts` to use the same helper.
7. Add `--check` CI guard.
8. Delete any remaining duplicated alias maps.

### Risks

- TypeScript path inheritance can be easy to misconfigure if the generated file sits in the wrong location.
- Nonstandard entrypoints like `@bilig/formula/program-arena` still need an explicit exception path.

### Validation

- `pnpm typecheck`
- `pnpm test`
- `pnpm bench:smoke`
- a fresh-checkout simulation where no package `dist` output exists
- generator `--check` in CI

Concrete validation commands:

- `bun scripts/gen-workspace-resolution.ts --check`
- `pnpm typecheck`
- `pnpm test`
- `pnpm bench:smoke`

## Workstream C: rename and tighten the authoritative snapshot worker API

### Current problem

The worker runtime currently exposes two snapshot-oriented methods:

- `replaceSnapshot(snapshot, authoritativeRevision?)`
- `rebaseToSnapshot(snapshot, authoritativeRevision)`

But only one of those names is semantically honest.

`replaceSnapshot` is used during the normal product bootstrap path when the worker needs to install the latest authoritative snapshot after cold boot or hydrate miss. That is not debug plumbing and not arbitrary snapshot replacement.

The current name causes three problems:

- code reads as if the product path is still using an old snapshot-era debug API
- docs now disagree with the actual runtime
- `runtime-session.ts` has a stringly-typed special case for a method that should be part of a deliberate authoritative hydrate contract

### File inventory

Primary implementation files:

- `apps/web/src/runtime-session.ts`
- `apps/web/src/worker-runtime.ts`
- `docs/browser-runtime.md`

Primary validation files:

- `apps/web/src/__tests__/runtime-session.test.ts`
- `apps/web/src/__tests__/worker-runtime.test.ts`
- `apps/web/src/__tests__/worker-runtime-reconnect.test.ts`
- `apps/web/src/__tests__/runtime-machine.test.ts`

### Target state

Replace `replaceSnapshot` with a semantically correct authoritative snapshot install API.

The product path should read in code the same way it is described in docs:

- bootstrap authoritative hydrate
- reconcile/recover via authoritative snapshot install when event replay cannot catch up

### Current-to-target API table

| Current public method | Current use | Target state |
| --- | --- | --- |
| `replaceSnapshot(snapshot, authoritativeRevision?)` | bootstrap authoritative hydrate | delete from the public worker RPC surface |
| `rebaseToSnapshot(snapshot, authoritativeRevision)` | reconcile/recover when event replay cannot catch up | delete from the public worker RPC surface |
| `installAuthoritativeSnapshot({ snapshot, authoritativeRevision, mode })` | does not exist | sole public snapshot-install RPC |
| `applyAuthoritativeEvents(events, authoritativeRevision)` | authoritative event replay | unchanged |
| `getAuthoritativeRevision()` | revision introspection | unchanged |

### Proposed design

#### New worker API

Introduce a single named worker method:

- `installAuthoritativeSnapshot(...)`

Suggested input shape:

```ts
interface InstallAuthoritativeSnapshotInput {
  snapshot: WorkbookSnapshot;
  authoritativeRevision: number;
  mode: "bootstrap" | "reconcile";
}
```

This replaces the overloaded meaning currently split between:

- `replaceSnapshot(...)`
- `rebaseToSnapshot(...)`

#### Semantics

`mode: "bootstrap"` means:

- install the authoritative engine from the fetched snapshot
- set `authoritativeRevision` to `max(currentAuthoritativeRevision, input.authoritativeRevision)`
- rebuild the projection engine and persist state

`mode: "reconcile"` means:

- install the authoritative engine from the fetched snapshot
- set `authoritativeRevision` to the exact target revision
- replay pending local mutations and rebuild projection state

This preserves the current behavioral split while removing the bad API naming.

#### Runtime-session call sites

`apps/web/src/runtime-session.ts` should change from:

- `replaceSnapshot` during initial hydrate
- `rebaseToSnapshot` during recovery path

to:

- `installAuthoritativeSnapshot({ ..., mode: "bootstrap" })`
- `installAuthoritativeSnapshot({ ..., mode: "reconcile" })`

The generic invoke refresh logic should key off the new method name once, not a legacy alias.

#### Worker-runtime implementation

`apps/web/src/worker-runtime.ts` should keep one shared internal helper for:

- creating the authoritative engine from snapshot
- rebuilding the projection engine
- invalidating viewport caches
- persisting state
- broadcasting viewport patches

The public worker API should expose only the semantically named entrypoint.

If the helper implementation remains split internally, that is fine. The public RPC surface should not remain split for legacy naming reasons.

Preferred internal end state:

- a private helper such as `installAuthoritativeSnapshotInternal(...)`
- no public `replaceSnapshot`
- no public `rebaseToSnapshot`

#### Backward compatibility decision

This worker RPC is internal to the repo. There is no need for a long-lived compatibility alias.

Acceptable migration plan:

- change all repo call sites in one coordinated refactor
- update tests and docs in the same commit
- delete `replaceSnapshot` immediately

### Rollout plan

1. Introduce `installAuthoritativeSnapshot(...)` in worker runtime.
2. Migrate bootstrap and reconcile call sites in `runtime-session.ts`.
3. Fold `replaceSnapshot` and `rebaseToSnapshot` internals behind one shared helper.
4. Update tests to use the new API name and mode semantics.
5. Update docs so `browser-runtime.md` and related notes describe the real product path.
6. Delete `replaceSnapshot`.

### Risks

- The bootstrap and reconcile revision semantics are intentionally different today; a sloppy merge would break pending-mutation replay or revision monotonicity.
- Tests that assert phase transitions need to continue proving `reconciling` and `recovering` behavior.

### Validation

- worker-runtime tests for bootstrap hydrate and reconcile/rebase behavior
- runtime-session tests for `reconciling` and `recovering` phases
- browser-runtime docs updated to remove the false “legacy-only” statement

Concrete validation commands and files:

- `pnpm exec vitest run apps/web/src/__tests__/worker-runtime.test.ts`
- `pnpm exec vitest run apps/web/src/__tests__/worker-runtime-reconnect.test.ts`
- `pnpm exec vitest run apps/web/src/__tests__/runtime-session.test.ts`
- `pnpm exec vitest run apps/web/src/__tests__/runtime-machine.test.ts`
- `pnpm test`

## Execution order

Recommended order:

1. workstream B: workspace resolution centralization
2. workstream C: authoritative snapshot API rename
3. workstream A: boot-path migration extraction

Why this order:

- B is low-risk and removes a recurring source of fresh-checkout/tooling breakage
- C is medium-risk but internal-only and keeps naming aligned before more runtime work lands
- A has the highest operational risk and should land after the toolchain and runtime API are cleaner

## Phase gates

### Phase 1: workspace resolution

Must be true before moving on:

- generated resolution artifacts are checked in
- root/app TS configs stop duplicating workspace maps manually
- Vite and Vitest consume shared resolution logic
- fresh-checkout tests no longer depend on prebuilt package `dist`

### Phase 2: authoritative snapshot API

Must be true before moving on:

- `replaceSnapshot` no longer appears in runtime-session, worker-runtime, or docs
- bootstrap hydrate and reconcile flows still pass existing runtime tests
- docs describe the same method name and semantics used in code

### Phase 3: migration extraction

Must be true before sign-off:

- Zero startup no longer executes heavy data scans directly
- required migration preconditions are enforced before the app begins listening
- migration runner and ledger are covered by tests
- the runtime assumption is mirrored in `lab` companion docs without embedding Argo implementation details here

## Acceptance criteria

This cleanup is complete when all of the following are true:

- `EnabledZeroSyncService.initialize()` no longer directly runs heavy data backfills or legacy table drops
- data migrations execute through a named migration runner with a persistent ledger
- app startup fails fast when `required` migrations are pending and boot-time auto-run is disabled
- root/app TypeScript configs no longer manually restate workspace package maps independently
- `apps/web/vite.config.ts` and `vitest.config.ts` derive workspace aliases from shared tooling
- `workspace-resolution.generated.json` is the canonical generated package-resolution artifact
- there is a CI `--check` path for generated workspace resolution metadata
- `replaceSnapshot` no longer exists on the worker RPC surface
- `rebaseToSnapshot` is no longer part of the public worker RPC surface
- browser bootstrap and reconcile paths use a semantically named authoritative snapshot install API
- runtime docs describe the same hydrate path that the code actually executes

## Out of scope but adjacent

These are real cleanup candidates, but they are not covered by this document:

- deleting legacy JSON viewport patch fallback
- collapsing giant modules like `packages/core/src/engine.ts`
- simplifying the root build script chain
- cleaning remaining manifest drift such as the stale GitHub repository URL in `package.json`

Those should be handled as separate cleanup work once these `P1` production seams are removed.
