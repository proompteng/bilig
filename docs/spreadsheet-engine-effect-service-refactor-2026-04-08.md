# SpreadsheetEngine Effect Service Refactor
## Date: 2026-04-08
## Status: proposed

## Why this document exists

`packages/core/src/engine.ts` is still the largest concentration of runtime debt in the repo.

Today `SpreadsheetEngine` is both:

- the public workbook runtime API that the rest of the application needs
- the implementation home for too many unrelated concerns

That shape is now the main maintainability risk in core runtime work. The file is still almost 6k lines and owns:

- workbook mutation orchestration
- formula binding and dependency graphs
- JS/WASM recalculation scheduling
- undo/redo
- replica sync bookkeeping
- snapshot import/export
- tables, spills, filters, sorts, and pivots
- event emission and selection state

The role is necessary. The current centralization is not.

This document defines a production refactor plan that keeps `SpreadsheetEngine` as the stable public façade while moving concrete responsibilities into typed Effect-backed services and smaller pure modules.

It is intentionally aligned with:

- `docs/reliability/00-program.md`
- `docs/reliability/01-runtime-kernel.md`
- `docs/replica-state-ownership-cleanup-2026-04-07.md`

## Problem statement

The current engine shape mixes four kinds of code that should not live in one file:

1. Pure spreadsheet logic
   - formula dependency binding
   - structural reference rewriting
   - snapshot projection transforms
   - pivot materialization helpers

2. Mutable runtime state ownership
   - workbook state
   - formula tables
   - reverse dependency indexes
   - typed arena buffers
   - undo/redo stacks
   - sync replica version state

3. Control-plane workflows
   - local transaction execution
   - remote batch apply
   - recalc orchestration
   - pivot refresh reconciliation
   - snapshot import/export workflows

4. Side-effect boundaries
   - metrics collection
   - event emission
   - sync-client lifecycle
   - clock/random access for volatile recalc
   - resource initialization and cleanup

The result is predictable:

- changes have large blast radius
- bugs hide in orchestration paths, not only compute kernels
- test setup is expensive because the unit of ownership is too large
- the current file can only shrink by “extract helper” work unless we create a better ownership model

## Goals

- keep `SpreadsheetEngine` as the stable public API for apps and tests during the refactor
- move orchestration and side-effecting responsibilities into concrete Effect services
- keep hot-path cell compute and low-level transforms pure and non-Effect
- reduce `packages/core/src/engine.ts` to a thin façade and composition layer
- make failure types explicit at service boundaries
- create a refactor path that allows bug fixes to land incrementally without a flag-day rewrite
- end with a structure that is easy to test, profile, and reason about

## Non-goals

- rewriting formula evaluation internals into Effect for style
- wrapping per-cell or per-edge compute in Effect
- changing workbook semantic behavior as part of the refactor
- changing the public op model in `@bilig/workbook-domain`
- changing snapshot formats during the structural cleanup
- replacing `SpreadsheetEngine` with actor-per-cell or service-per-cell models

This follows the same rule as the reliability program: use Effect for lifecycle, failure, resource, and workflow boundaries, not for inner numeric loops.

## Refactor principles

### 1. `SpreadsheetEngine` remains the façade

The application should continue to call:

- `setCellValue`
- `setCellFormula`
- `recalculateNow`
- `applyRemoteBatch`
- `exportSnapshot`
- `importSnapshot`
- `undo`
- `redo`

The class survives. Its implementation becomes delegation instead of ownership.

### 2. Pure logic stays pure

The following categories should remain plain functions and plain modules:

- address/range rewriting
- dependency materialization
- snapshot shape transforms
- format/style tiling
- value normalization and display
- pivot computation kernels
- formula compile/evaluate helpers

Effect services should call these modules. They should not replace them.

### 3. Services own workflows, resources, and failures

Effect services should own:

- sequencing
- resource lifetime
- retries/timeouts where needed
- structured error algebra
- metrics/logging emission
- coordination across runtime submodules

### 4. State is explicit

The current class fields should be collapsed behind an explicit runtime state object rather than left as dozens of private fields on one class.

### 5. Every extraction must be bug-positive

No large refactor slice should be a pure file move. Each slice must do at least one of:

- add direct tests for the extracted unit
- tighten validation
- remove an implicit ambient dependency
- fix an already-known bug or fragile branch

## Target architecture

## 1. Runtime state

Introduce a single runtime state container, for example:

- `packages/core/src/engine/runtime-state.ts`

It should hold the mutable engine internals currently spread across the class:

- workbook
- string pool
- formula table
- dependency reverse indexes
- typed arenas and scratch buffers
- selection state
- undo/redo stacks
- replica version maps
- last metrics

Suggested shape:

```ts
export interface EngineRuntimeState {
  readonly workbook: WorkbookStore;
  readonly strings: StringPool;
  readonly formulas: FormulaTable<RuntimeFormula>;
  readonly ranges: RangeRegistry;
  readonly scheduler: RecalcScheduler;
  readonly wasm: WasmKernelFacade;
  readonly dependencyIndex: EngineDependencyIndex;
  readonly scratch: EngineScratchBuffers;
  readonly history: EngineHistoryState;
  readonly replica: EngineReplicaState;
  readonly selection: SelectionState;
  readonly metrics: RecalcMetrics;
}
```

`SpreadsheetEngine` should hold exactly one state reference instead of owning every moving part directly.

## 2. Service layout

Create an internal engine service tree, for example:

- `packages/core/src/engine/services/`

Recommended concrete services:

### `EngineMutationService`

Owns:

- local transaction execution
- inverse-op generation hookup
- changed-input tracking
- explicit changed-set collection
- structural mutation dispatch

Public boundary:

```ts
export interface EngineMutationService {
  readonly executeLocal: (
    tx: TransactionRecord,
  ) => Effect.Effect<ReadonlyArray<number>, EngineMutationError>;
  readonly executeRemote: (
    tx: TransactionRecord,
  ) => Effect.Effect<ReadonlyArray<number>, EngineMutationError>;
}
```

### `EngineRecalcService`

Owns:

- dirty root composition
- formula rebinding scheduling
- JS/WASM program sync
- topological recalc
- volatile context injection
- pivot refresh reconciliation trigger

Public boundary:

```ts
export interface EngineRecalcService {
  readonly recalculateNow: () => Effect.Effect<ReadonlyArray<number>, EngineRecalcError>;
  readonly recalculateDirty: (
    dirty: ReadonlyArray<DirtyRegion>,
  ) => Effect.Effect<ReadonlyArray<number>, EngineRecalcError>;
}
```

### `EngineSnapshotService`

Owns:

- workbook snapshot export
- workbook snapshot import
- replica snapshot export/import
- snapshot compatibility checks

Public boundary:

```ts
export interface EngineSnapshotService {
  readonly exportWorkbook: () => Effect.Effect<WorkbookSnapshot, EngineSnapshotError>;
  readonly importWorkbook: (
    snapshot: WorkbookSnapshot,
  ) => Effect.Effect<void, EngineSnapshotError>;
  readonly exportReplica: () => Effect.Effect<EngineReplicaSnapshot, EngineSnapshotError>;
  readonly importReplica: (
    snapshot: EngineReplicaSnapshot,
  ) => Effect.Effect<void, EngineSnapshotError>;
}
```

### `EngineReplicaSyncService`

Owns:

- remote batch dedupe
- replica version updates
- sync-client connect/disconnect lifecycle
- sync state transitions

Public boundary:

```ts
export interface EngineReplicaSyncService {
  readonly applyRemoteBatch: (
    batch: EngineOpBatch,
  ) => Effect.Effect<boolean, EngineSyncError>;
  readonly connectClient: (
    client: EngineSyncClient,
  ) => Effect.Effect<void, EngineSyncError>;
  readonly disconnectClient: () => Effect.Effect<void, never>;
}
```

### `EngineHistoryService`

Owns:

- undo stack
- redo stack
- transaction replay depth
- replay-safe mutation invocation

### `EngineSelectionService`

Owns:

- selection state updates
- selection subscriptions
- selection event emission

### `EngineEventService`

Owns:

- batch event emission
- targeted cell subscriptions
- listener lifecycle
- event payload shaping

This should be the only service that talks directly to `EngineEventBus`.

### `EnginePivotService`

Owns:

- pivot refresh invalidation checks
- pivot output ownership
- pivot materialization workflow
- pivot output cleanup during structural changes

### `EngineStructureService`

Owns workflow-level structural edits:

- row/column insert/delete/move orchestration
- sheet rename propagation
- table/filter/sort/spill/pivot structural rewrites

The pure rewrite helpers already extracted under `engine-structural-utils.ts`, `engine-range-utils.ts`, and related modules remain beneath this service.

## 3. Effect construction model

Use the same Effect primitives already established in `@bilig/runtime-kernel`:

- `Context.GenericTag`
- `Data.TaggedError`
- `Effect`
- `Layer`

Suggested pattern:

```ts
export class EngineMutationError extends Data.TaggedError("EngineMutationError")<{
  readonly message: string;
  readonly cause?: unknown;
}> {}

export interface EngineMutationService {
  readonly executeLocal: (
    tx: TransactionRecord,
  ) => Effect.Effect<ReadonlyArray<number>, EngineMutationError>;
}

export const EngineMutationService =
  Context.GenericTag<EngineMutationService>("@bilig/core/EngineMutationService");
```

Live wiring should happen in one internal composition module, for example:

- `packages/core/src/engine/live.ts`

`SpreadsheetEngine` should call those services through a small local runner rather than reimplementing service construction ad hoc.

## 4. Public façade shape

Target shape for `SpreadsheetEngine`:

- constructor allocates `EngineRuntimeState`
- constructor installs live service layer
- public methods delegate to services
- getters/selectors delegate to pure read modules

Pseudo-shape:

```ts
export class SpreadsheetEngine {
  private readonly state: EngineRuntimeState;
  private readonly runtime: EngineServiceRuntime;

  setCellValue(sheet: string, address: string, value: LiteralInput): CellValue {
    return this.run(this.runtime.mutation.setCellValue(sheet, address, value));
  }

  recalculateNow(): number[] {
    return this.run(this.runtime.recalc.recalculateNow());
  }

  exportSnapshot(): WorkbookSnapshot {
    return this.run(this.runtime.snapshot.exportWorkbook());
  }
}
```

The façade remains synchronous where the current API is synchronous. Internally, services may still use Effect for composition, typed failure normalization, and shared dependencies.

## Dependency rules

### Allowed

- façade -> services
- services -> runtime state
- services -> pure helper modules
- services -> runtime-kernel utilities where relevant
- services -> other services only through narrow interfaces

### Forbidden

- app code importing internal engine services directly
- services reaching into unrelated service private state
- generic ambient `Date.now`, `Math.random`, `fetch`, or websocket usage inside services
- reintroducing large multi-concern utility files as a replacement for `engine.ts`

## Error algebra

Expected engine orchestration failures should be tagged.

Suggested error families:

- `EngineMutationError`
- `EngineRecalcError`
- `EngineSnapshotError`
- `EngineSyncError`
- `EngineHistoryError`
- `EnginePivotError`
- `EngineStructureError`

Rules:

- validation failures at service boundaries become tagged errors or normalized workbook error values
- generic thrown `Error` is allowed only inside low-level adapters and must be normalized immediately
- the public `SpreadsheetEngine` façade may convert internal service failures into current API-compatible behavior where necessary, but the service layer itself must remain typed

## Bug classes to squash during the refactor

These are the likely production bugs hidden by the current centralization. Each wave should explicitly target at least one:

### 1. Required-argument and invalid-shape fallthrough

Pattern:

- orchestration code assumes optional op fields are present
- coercion paths produce undefined behavior or wrong workbook errors

Action:

- add direct tests for each newly extracted service boundary
- prefer explicit tagged failure or explicit workbook error values

### 2. Reverse dependency index drift

Pattern:

- formula/table/spill/defined-name reverse edges are not fully cleaned up during structural edits or snapshot restore

Action:

- add invariant tests that reverse indexes match forward bindings after:
  - snapshot import
  - rename sheet
  - row/column move
  - table delete
  - spill delete

### 3. Pivot output ownership leaks

Pattern:

- pivot output cells remain marked or owned after structural changes, failed materialization, or source deletion

Action:

- centralize pivot owner bookkeeping in `EnginePivotService`
- add tests for clear-on-delete, clear-on-rename, clear-on-source-invalid

### 4. Snapshot drift

Pattern:

- snapshot export/import does not round-trip styles, formats, filters, sorts, defined names, spills, or pivots exactly

Action:

- add focused snapshot round-trip suites by metadata family
- keep snapshot compatibility stable across waves

### 5. Dirty-region recalc misses

Pattern:

- differential or dirty recalc misses impacted formulas because changed-root assembly and rebinding live in too many places

Action:

- centralize root composition in `EngineRecalcService`
- add regression tests for structural edits plus dirty-region recalc

### 6. Remote batch dedupe drift

Pattern:

- app-layer or service-layer logic reimplements dedupe policy differently

Action:

- keep `applyRemoteBatch()` as the single dedupe/apply boundary
- move all replica decisions behind `EngineReplicaSyncService`

## Module plan

The refactor should happen in waves, not as a single branch.

## Wave 0: Baseline and guardrails

Deliverables:

- direct tests for the service families to be extracted
- line-count baseline and hotspot inventory
- invariant tests for:
  - snapshot round-trip
  - pivot output cleanup
  - reverse dependency cleanup
  - undo/redo replay correctness
  - remote batch idempotence

Acceptance:

- no service extraction starts without direct tests
- every new service gets a dedicated `__tests__` file or suite

## Wave 1: Create façade-compatible runtime state and service shell

Deliverables:

- `EngineRuntimeState`
- service error types
- live layer composition
- thin runner used only inside `SpreadsheetEngine`

Do not extract major logic yet. First create the ownership skeleton.

Acceptance:

- `SpreadsheetEngine` public API unchanged
- no behavior changes
- state is held behind one runtime object

## Wave 2: Extract read-only and side-effect services

Target services:

- `EngineSelectionService`
- `EngineEventService`
- `EngineSnapshotService`
- `EngineHistoryService`

Why first:

- lower algorithmic risk
- high reduction in class noise
- easier direct test setup

Bug focus:

- snapshot round-trip drift
- selection event duplication
- undo/redo replay edge cases

## Wave 3: Extract sync and mutation orchestration

Target services:

- `EngineReplicaSyncService`
- `EngineMutationService`
- `EngineStructureService`

Why here:

- this is where app-facing runtime correctness concentrates
- most operational bugs happen at mutation boundaries, not pure compute

Bug focus:

- remote batch dedupe
- inverse-op generation correctness
- structural edit fallout on tables/spills/sorts/filters

## Wave 4: Extract recalculation and formula binding orchestration

Target services:

- `EngineRecalcService`
- `EngineFormulaBindingService` if needed as a separate service
- `EnginePivotService`

Why later:

- highest algorithmic risk
- most dependent on previous state/model cleanup

Bug focus:

- stale dependency edges
- dirty-root misses
- JS/WASM differential drift
- pivot refresh reconciliation

## Wave 5: Collapse the façade and delete dead seams

Deliverables:

- `packages/core/src/engine.ts` reduced to façade + composition only
- no remaining ad hoc side-effect ownership in the class
- service-local tests replace broad incidental coverage where possible
- docs and comments updated to match the new ownership model

Acceptance:

- `engine.ts` is orchestration-only
- no new file exceeds the same hotspot pattern we are removing

## File layout target

Suggested internal layout:

```text
packages/core/src/
  engine.ts
  engine/
    runtime-state.ts
    errors.ts
    live.ts
    services/
      mutation-service.ts
      recalc-service.ts
      snapshot-service.ts
      replica-sync-service.ts
      history-service.ts
      selection-service.ts
      event-service.ts
      pivot-service.ts
      structure-service.ts
    read/
      selectors.ts
      explain.ts
```

Existing extracted pure modules should stay where they are or move under `engine/` only if that improves clarity without causing churn.

## Testing plan

Every wave should run:

- `pnpm exec tsc -p packages/core/tsconfig.json --pretty false`
- focused `vitest` suites for the extracted module family
- targeted `oxlint`

At wave boundaries, run:

- `pnpm --filter @bilig/core test`
- relevant formula parity or differential tests when recalc code moves
- full repo `pnpm lint`

Add or expand:

- snapshot round-trip tests
- property-style structural transform tests
- pivot materialization regression tests
- remote batch idempotence tests
- JS/WASM differential smoke coverage

## Rollout strategy

This is an internal refactor. Prefer compatibility over flags.

Rules:

- keep `SpreadsheetEngine` public methods stable until the final cleanup wave
- land each service extraction in its own commit
- do not mix unrelated architecture shifts in the same slice
- if a slice grows beyond roughly 1000 lines of source diff, commit before the next extraction
- push only after the committed tree passes the intended verification set

## Production acceptance bar

The refactor is successful when:

- `packages/core/src/engine.ts` is below roughly 1000 lines and acts as a façade
- each engine service has a narrow, documented responsibility
- no service depends on ambient browser/node globals for expected runtime behavior
- snapshot import/export compatibility is preserved
- remote batch apply semantics are preserved
- undo/redo semantics are preserved
- pivot/table/filter/sort/spill behavior remains green
- `pnpm run ci` is green on the stabilized tree

## Immediate execution order

The highest-value next sequence is:

1. create `EngineRuntimeState`, `errors.ts`, and `live.ts`
2. extract `EngineSnapshotService`
3. extract `EngineHistoryService`
4. extract `EngineReplicaSyncService`
5. extract `EngineMutationService`
6. extract `EngineRecalcService`
7. extract `EnginePivotService`
8. collapse `SpreadsheetEngine` into façade-only form

## Final note

The point of this plan is not “make core use Effect everywhere.”

The point is:

- pure spreadsheet math stays fast and explicit
- engine workflows become typed and composable
- failures become attributable
- state ownership becomes manageable
- the class that the application depends on becomes small enough to trust again
