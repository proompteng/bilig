# High Severity Production Remediation Plan

## Date: 2026-04-11

## Status: proposed

## Purpose

This document turns the 2026-04-11 high-severity technical debt audit into an
execution-grade production plan.

Scope is intentionally limited to the audit's highest-risk findings:

- browser runtime split authority and worker-runtime centralization
- agent HTTP compatibility sprawl and in-memory session-state authority
- centralized workbook state and mutation/history orchestration in core

This plan is based on tracked `main` as of `5bbc2b0`.

It aligns with, but is stricter than:

- [next-iteration-production-plan-2026-04-10.md](/Users/gregkonush/github.com/bilig3/docs/next-iteration-production-plan-2026-04-10.md)
- [tech-debt-remediation-program-2026-04-09.md](/Users/gregkonush/github.com/bilig3/docs/tech-debt-remediation-program-2026-04-09.md)
- [browser-runtime.md](/Users/gregkonush/github.com/bilig3/docs/browser-runtime.md)
- [spreadsheet-engine-effect-service-refactor-2026-04-08.md](/Users/gregkonush/github.com/bilig3/docs/spreadsheet-engine-effect-service-refactor-2026-04-08.md)

## Non-negotiable rules

1. No long-lived compatibility surfaces survive completion.
2. No second authority layer is allowed to remain in the browser.
3. No public session-scoped agent API survives completion.
4. No in-memory mutable thread snapshot is allowed to be the durable source of truth.
5. No monolithic inverse-op switch or centralized workbook-state grab bag survives completion.
6. Every wave must delete the code it replaces before the wave is considered done.

Allowed transitional mechanics are narrow:

- additive storage migrations are allowed when physically necessary
- those migrations must not create a second runtime read/write path
- cutover commits must remove old readers and writers together

This is not a “preserve everything and wrap it” plan. It is a replacement and
deletion plan.

## Desired end state

### Browser runtime

- The browser persists exactly two durable local artifacts:
  - authoritative workbook base
  - pending mutation journal
- Projection state is derived, ephemeral, and rebuildable.
- Viewport patch publication is a read model over authoritative base plus
  pending journal, not an independent local authority.
- `WorkbookWorkerRuntime` becomes a thin façade over smaller services.

### Agent runtime

- `threadId` is the only durable public identifier for workbook chat.
- `/v2/documents/:documentId/chat/threads/*` is the only public agent route
  family.
- Durable thread state in storage is the only source of truth for timeline,
  context, pending bundle, workflow runs, and execution records.
- In-memory state is limited to ephemeral turn leases and stream subscribers.

### Core runtime

- `SpreadsheetEngine` remains the public façade.
- `WorkbookStore` no longer exists as a monolithic state owner.
- Workbook state is decomposed into focused repositories/modules.
- Mutation canonicalization, inverse-op construction, and history capture are
  explicit subsystems, not one giant service switch.

## Program structure

## Workstream A: Agent Durable Thread Cutover

Primary owner surface:

- [/Users/gregkonush/github.com/bilig3/apps/bilig/src/codex-app](/Users/gregkonush/github.com/bilig3/apps/bilig/src/codex-app)
- [/Users/gregkonush/github.com/bilig3/apps/bilig/src/http](/Users/gregkonush/github.com/bilig3/apps/bilig/src/http)
- [/Users/gregkonush/github.com/bilig3/apps/bilig/src/zero](/Users/gregkonush/github.com/bilig3/apps/bilig/src/zero)

Primary debt this removes:

- duplicate route families in [sync-server.ts](/Users/gregkonush/github.com/bilig3/apps/bilig/src/http/sync-server.ts)
- mutable in-memory session snapshots in [workbook-agent-service.ts](/Users/gregkonush/github.com/bilig3/apps/bilig/src/codex-app/workbook-agent-service.ts)

### Target architecture

- Introduce a durable thread domain boundary that owns:
  - thread metadata
  - timeline entries
  - thread context
  - pending bundle state
  - workflow runs
  - execution records
- Keep live Codex execution state outside that boundary:
  - active turn lease
  - stream subscriber registry
  - live transport binding
- Move all thread mutation into transactional repository methods with explicit
  optimistic concurrency, for example `threadVersion` or equivalent monotonic
  write token.

Recommended internal split:

- `WorkbookAgentThreadRepository`
- `WorkbookAgentTurnCoordinator`
- `WorkbookAgentBundleService`
- `WorkbookAgentWorkflowService`
- `WorkbookAgentStreamBroker`
- `WorkbookAgentCodexLeasePool`

### Execution waves

#### A1. Freeze the public contract around durable threads

Work:

- declare `/v2/documents/:documentId/chat/threads/*` as the only supported
  public route family
- move any browser caller still depending on `sessionId` to `threadId`
- make `sessionId` internal-only or remove it entirely from product payloads
- define exact durable-thread snapshot and stream contracts in
  `@bilig/contracts`

Exit bar:

- no browser product path depends on public session routes
- no browser product path stores a durable session identifier

#### A2. Replace mutable session snapshots with transactional thread updates

Work:

- replace `sessions`, `threadToSessionId`, and mutable `sessionState.snapshot`
  authority with repository-backed thread aggregate reads and writes
- require all mutations to flow through typed update methods that:
  - load durable thread state
  - validate ownership / rollout / review invariants
  - write the next durable snapshot atomically
  - emit a post-commit stream event
- keep in-memory maps only for live turn execution lease ownership and stream
  listeners

Exit bar:

- pending bundle, context, timeline, workflow runs, and execution records can
  all be reconstructed from durable storage alone
- monolith restart does not require rebuilding thread truth from memory

#### A3. Collapse the route surface and delete the legacy families

Work:

- replace route duplication in `sync-server.ts` with route modules centered on
  the thread contract
- delete:
  - `/agent/sessions/*`
  - `/agent/threads/*`
- keep only `/chat/threads/*`
- delete duplicated SSE handlers and replace them with one thread-stream
  controller

Exit bar:

- there is one route family
- there is one stream handler
- there is one controller path per operation

#### A4. Tighten stream semantics around durable state

Work:

- initial SSE snapshot must always come from durable thread storage
- live deltas must represent post-commit durable state changes or explicit live
  turn-lease signals
- connection / reconnection behavior must never synthesize missing timeline
  state from memory

Exit bar:

- refresh, reconnect, and monolith restart yield the same thread snapshot
- stale stream consumers never see state that durable storage cannot explain

### Explicit deletions

- duplicated route registrations in
  [sync-server.ts](/Users/gregkonush/github.com/bilig3/apps/bilig/src/http/sync-server.ts)
- public `sessionId`-centric browser flows
- mutable in-memory thread snapshot authority in
  [workbook-agent-service.ts](/Users/gregkonush/github.com/bilig3/apps/bilig/src/codex-app/workbook-agent-service.ts)
- tests that encode session-route parity as a product requirement

### Verification

- targeted route tests in `apps/bilig/src/http/*.test.ts`
- durable-thread store tests in `apps/bilig/src/zero/__tests__/*`
- thread restart/reconnect tests in
  `apps/bilig/src/codex-app/*.test.ts`
- `pnpm test:correctness:server`
- `pnpm lint`
- `pnpm typecheck`

### Final acceptance criteria

- the only public agent surface is thread-centric
- a thread survives refresh, reconnect, and monolith restart without fallback
  session hydration
- no accepted or pending agent state is lost when in-memory leases are evicted

## Workstream B: Browser Authority Unification

Primary owner surface:

- [/Users/gregkonush/github.com/bilig3/apps/web/src](/Users/gregkonush/github.com/bilig3/apps/web/src)
- [/Users/gregkonush/github.com/bilig3/packages/storage-browser/src](/Users/gregkonush/github.com/bilig3/packages/storage-browser/src)

Primary debt this removes:

- split authority in [worker-runtime.ts](/Users/gregkonush/github.com/bilig3/apps/web/src/worker-runtime.ts)
- heuristic cache authority in [projected-viewport-store.ts](/Users/gregkonush/github.com/bilig3/apps/web/src/projected-viewport-store.ts)

### Target architecture

- Durable local storage contains:
  - authoritative workbook base
  - pending mutation journal
- Projection overlay state is rebuilt from journal replay and not persisted as a
  second local truth.
- Viewport caches are subscription-scoped read models and have no mutation
  authority methods.
- `WorkbookWorkerRuntime` becomes composition only.

Recommended internal split:

- `WorkbookWorkerBootstrapService`
- `WorkbookWorkerAuthoritativeStore`
- `WorkbookWorkerMutationJournalService`
- `WorkbookWorkerProjectionService`
- `WorkbookWorkerViewportReadModel`
- `WorkbookWorkerPersistenceService`
- `WorkbookWorkerAgentPreviewService`

### Execution waves

#### B1. Simplify the persistence model to one durable workbook base plus journal

Work:

- stop persisting projection overlay tables as product state
- delete persisted `projection_overlay_*` authority once journal replay can
  rebuild local optimistic state deterministically
- keep authoritative tables and pending mutation journal only
- make bootstrap always derive projection state from:
  - authoritative base
  - pending journal

Exit bar:

- projection state can be reconstructed from persistent inputs without any
  persisted overlay table
- crash recovery does not depend on a second persisted viewport truth

#### B2. Convert viewport state into a derived read model

Work:

- replace `ProjectedViewportStore` mutation authority with a read-only
  projection cache API
- move optimistic mutation application, ack, rollback, and reconcile behavior
  into the projection or journal services
- remove correctness-bearing write methods from the viewport store

Exit bar:

- viewport caches are consumers of projection state, not owners of it
- grid reads one derived source for cell and axis state

#### B3. Delete branchy authority state from `WorkbookWorkerRuntime`

Work:

- remove broad runtime branching like:
  - `authoritativeStateSource`
  - `projectionMatchesLocalStore`
  - direct install/rebuild/persist orchestration in one class
- replace with explicit service boundaries for:
  - bootstrap
  - authoritative hydrate
  - authoritative event ingest
  - journal replay
  - persistence
  - viewport patch publication

Exit bar:

- each runtime boundary has a dedicated service with direct tests
- `WorkbookWorkerRuntime` is a façade and coordinator, not the implementation
  home for every lifecycle path

#### B4. Replace heuristic cell-cap cache policy with explicit residency rules

Work:

- delete `MAX_CACHED_CELLS_PER_SHEET = 6000`
- define cache residency in terms of:
  - active viewport subscriptions
  - tile generations
  - explicit memory budget policy module
- ensure cache policy cannot affect correctness, only memory footprint and
  recompute cost

Exit bar:

- cache eviction is predictable and measured
- correctness is independent of cache residency

### Explicit deletions

- `projection_overlay_*` persisted product state tables and write paths
- correctness-bearing mutation methods on
  [projected-viewport-store.ts](/Users/gregkonush/github.com/bilig3/apps/web/src/projected-viewport-store.ts)
- top-level heuristic cache caps that are not part of a dedicated policy module
- authority-branching fields in
  [worker-runtime.ts](/Users/gregkonush/github.com/bilig3/apps/web/src/worker-runtime.ts)
  that only exist to juggle overlapping truths

### Verification

- targeted runtime tests under `apps/web/src/__tests__/worker-runtime*.test.ts`
- targeted persistence tests under
  `apps/web/src/__tests__/worker-runtime-local-persistence.test.ts`
- targeted viewport tests under
  `apps/web/src/__tests__/projected-viewport-*.test.ts`
- reconnect and convergence browser tests
- `pnpm test:correctness:browser`
- `pnpm test:browser`
- `pnpm lint`
- `pnpm typecheck`

### Final acceptance criteria

- the browser has one durable local workbook authority and one journal
- authoritative reconcile and local replay always converge through the same
  derivation path
- refresh and crash recovery rebuild optimistic local state without persisted
  overlay truth

## Workstream C: Core Workbook State and Mutation/History Decomposition

Primary owner surface:

- [/Users/gregkonush/github.com/bilig3/packages/core/src](/Users/gregkonush/github.com/bilig3/packages/core/src)

Primary debt this removes:

- centralized ownership in [workbook-store.ts](/Users/gregkonush/github.com/bilig3/packages/core/src/workbook-store.ts)
- centralized inverse-op and history logic in
  [mutation-service.ts](/Users/gregkonush/github.com/bilig3/packages/core/src/engine/services/mutation-service.ts)

### Target architecture

- `SpreadsheetEngine` remains the public façade.
- Workbook state is decomposed behind explicit modules, for example:
  - `WorkbookSheetCatalog`
  - `WorkbookCellRepository`
  - `WorkbookStyleRepository`
  - `WorkbookNumberFormatRepository`
  - `WorkbookAxisRepository`
  - `WorkbookMetadataRepository`
  - `WorkbookStateAggregate`
- Mutation orchestration is decomposed into explicit services, for example:
  - `TransactionCanonicalizer`
  - `InverseOpRegistry`
  - `HistoryCaptureService`
  - `TransactionExecutor`
  - `FastCellMutationHistory`

### Execution waves

#### C1. Split workbook state by concern, not helper extraction

Work:

- move style/format interning out of the sheet/cell state owner
- move axis-entry and axis-metadata ownership out of the general workbook state
  owner
- move metadata families behind their own repository surface
- keep one aggregate for coordination, but stop storing unrelated concerns in a
  single monolithic class

Exit bar:

- focused modules own focused data
- tests can target cell, axis, style, and metadata concerns independently

#### C2. Replace giant inverse-op branching with registered op-family builders

Work:

- split the inverse-op switch into op-family modules:
  - workbook + sheet structure
  - cell value/formula
  - style + format
  - row/column structure
  - metadata
  - tables/spills/pivots
- register inverse builders through a typed registry instead of one monolithic
  switch
- keep the fast simple-cell history path, but make it one implementation behind
  a history-capture boundary

Exit bar:

- no giant inverse-op switch remains
- new op families can add undo support without editing one central god file

#### C3. Separate transaction execution from canonicalization and history capture

Work:

- isolate:
  - op canonicalization
  - undo capture
  - redo invalidation
  - transaction replay depth handling
  - batch execution
- make `mutation-service.ts` either a thin composition file or delete it
  entirely if it no longer adds value

Exit bar:

- execution order and history capture are explicit, typed boundaries
- undo/redo behavior is directly testable without broad engine setup

#### C4. Delete the monolithic `WorkbookStore`

Work:

- remove the current `WorkbookStore` class after call sites are migrated to the
  decomposed state modules or aggregate
- update internal imports, helpers, and tests to the new boundaries

Exit bar:

- there is no large internal state-owner file equivalent to `workbook-store.ts`
- engine consumers depend on stable façade or focused state modules, not a grab
  bag

### Explicit deletions

- monolithic [workbook-store.ts](/Users/gregkonush/github.com/bilig3/packages/core/src/workbook-store.ts)
- giant inverse-op branching in
  [mutation-service.ts](/Users/gregkonush/github.com/bilig3/packages/core/src/engine/services/mutation-service.ts)
- any helper that exists only to smuggle unrelated concerns through the old
  central store

### Verification

- focused tests for each extracted repository/service
- existing correctness suites:
  - `pnpm test:correctness:core`
  - `pnpm test:correctness:formula`
- undo/redo reversibility and snapshot parity property tests
- `pnpm lint`
- `pnpm typecheck`

### Final acceptance criteria

- workbook state concerns are isolated by ownership
- undo/redo correctness is enforced through registries and direct tests
- adding a new operation family no longer requires central switch editing across
  unrelated behavior

## Program ordering

Recommended merge order:

1. Workstream A contract freeze and route cutover
2. Workstream B persistence-model simplification and worker-runtime split
3. Workstream C workbook-state and mutation-history decomposition

Rationale:

- Workstream A removes duplicated product API surface first.
- Workstream B then simplifies browser state ownership against the now-canonical
  thread-centric product model.
- Workstream C is the deepest semantic refactor and should land after product
  surface simplification reduces concurrent change pressure.

Parallelism is still possible:

- Workstream C design and test preparation can start in parallel with A
- Workstream B schema and bootstrap planning can start in parallel with A
- merge sequencing should still follow the order above

## Definition of done

This program is done only when all of the following are true:

- no public `/agent/sessions/*` routes remain
- no public `/agent/threads/*` routes remain
- `threadId` is the only durable chat identifier in product code
- browser local persistence stores authoritative base and pending journal only
- no persisted projection overlay truth remains
- `WorkbookWorkerRuntime` is no longer the implementation home for all browser
  runtime lifecycles
- the monolithic `WorkbookStore` is gone
- the monolithic mutation inverse-op switch is gone
- correctness, browser, and CI gates pass on the replacement architecture

## Final release gate

Before declaring this complete, run the full repository bar:

```sh
pnpm run ci
```

Completion is not “the new code exists.” Completion is:

- the old code is deleted
- the new architecture is the only path
- correctness, durability, and multiplayer behavior are enforced by tests
