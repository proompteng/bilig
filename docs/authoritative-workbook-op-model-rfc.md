# Authoritative Workbook Op Model RFC

## Problem

`@bilig/crdt` carries ops for workbook metadata, row and column structure, freeze panes, filters, sorts, tables, spills, pivots, calc settings, and volatile context.

Current missing op families:

- some higher-level engine helpers normalize into many cell-level ops rather than first-class replicated range ops
- `renameSheet`
- `reorderSheets`
- explicit structured-reference binding ops
- explicit spill-owner ops

## Goals

- define an exhaustive authoritative workbook op family
- keep every stateful engine mutation behind a transaction executor
- make snapshots acceleration artifacts rather than semantic truth
- make undo/redo, replication, and local mutation paths consume the same transaction model
- give protocol, storage, browser runtime, and service layers one canonical mutation language

## Non-goals

- rewrite the engine around a new storage layout
- move all current and future behavior into one PR
- treat React descriptors or viewport patches as the authoritative workbook model

## Current gap

The existing replicated model in `packages/crdt/src/index.ts` covers:

- workbook metadata and calc settings
- sheet create and delete
- row and column insert, delete, move, and metadata update
- freeze panes, filters, and sorts
- cell value, formula, format, and clear
- defined names
- tables
- spill ranges
- pivot upsert and delete

The remaining local-engine-vs-replicated gap is:

- rename and reorder sheet operations
- first-class range mutation ops instead of only expanded cell-level replication
- explicit structured-reference binding surfaces
- explicit spill-owner and spill-blocking semantics
- workbook view or structural semantics that live as engine inference instead of authoritative state

## Design principles

1. one mutation language for local, replicated, undoable, and persisted workbook changes
2. local convenience methods become wrappers around transactions, not parallel mutation systems
3. snapshots, caches, and render patches are derived artifacts
4. op completeness matters more than shipping every consumer at once
5. JS semantics remain the oracle even as the mutation model grows

## Proposed model

### Terminology

- `WorkbookOp`: one logical workbook mutation
- `WorkbookTxn`: an ordered, atomic group of workbook ops
- `TxnExecutor`: the only stateful mutation ingress into the engine model
- `TxnCursor`: durable ordering cursor for replay and catch-up

### Op families

#### Workbook-level ops

- `upsertWorkbook`
- `setWorkbookMetadata`
- `setCalculationSettings`
- `setVolatileContext`

#### Sheet lifecycle ops

- `upsertSheet`
- `renameSheet`
- `deleteSheet`
- `reorderSheets`

#### Row and column structure ops

- `insertRows`
- `deleteRows`
- `moveRows`
- `updateRowMetadata`
- `insertColumns`
- `deleteColumns`
- `moveColumns`
- `updateColumnMetadata`

#### Cell content ops

- `setCellValue`
- `setCellFormula`
- `setCellFormat`
- `clearCell`
- first-class range ops where replication benefits from them instead of expanded cell-level batches

#### Workbook metadata ops

- `upsertDefinedName`
- `deleteDefinedName`
- `upsertTable`
- `deleteTable`
- explicit structured-reference binding ops where needed
- explicit spill-owner ops where needed
- `upsertPivotTable`
- `deletePivotTable`

#### UI-adjacent but durable worksheet semantics

- `setFreezePane`
- `clearFreezePane`
- `setFilter`
- `clearFilter`
- `setSort`
- `clearSort`

The rule is not that every one of these must be implemented at once. The rule is that the authoritative model must have a place for them so local-only mutation paths stop multiplying.

## Transaction executor boundary

`@bilig/core` exposes a transaction executor path and transaction-backed undo/redo. Remaining migration on this boundary is:

- validate input
- normalize into authoritative ops
- apply ops to workbook model and metadata store
- emit dependency invalidation and recalc scheduling effects
- append to undo/redo history
- publish outbound replication batches
- notify subscribers

## Snapshot and replication implications

- workbook snapshots remain important for restore speed
- replica snapshots remain important for replay trimming
- neither snapshot type should be treated as the semantic source of truth
- replication, catch-up, and local restore should all rebuild from the same transaction language plus snapshots as accelerators

## Package responsibilities

- `@bilig/protocol`
  - stable shared data types referenced by workbook ops
- `@bilig/crdt`
  - authoritative transaction ordering, dedupe, compaction, and replica clocks
- `@bilig/core`
  - transaction executor, invariant enforcement, graph invalidation, and history integration
- `@bilig/binary-protocol`
  - transport encoding for transactions and cursors
- `@bilig/storage-browser` and `@bilig/storage-server`
  - persistence of snapshots, transactions, and cursors
- `apps/local-server` and `apps/sync-server`
  - ordered transaction ingress, replay, catch-up, and fanout

## Remaining migration order

1. add the missing exhaustive-op families such as rename, reorder, and explicit metadata-binding ops
2. decide where first-class range ops are preferable to expanded cell-level replication
3. move remaining inferred spill and structured-reference semantics onto explicit authoritative state where needed
4. align binary protocol and persistence around the final transaction surface
5. make browser, local-server, and sync-server consume the same completed authoritative model

## Exit gate

- no stateful workbook mutation remains local-only by default
- defined names, spill ownership, and pivot lifecycle are authoritative ops
- undo/redo operates on transaction semantics rather than bespoke local branches
- snapshots are acceleration artifacts rather than the semantic source of truth
- browser, local-server, and sync-server all consume the same authoritative transaction model
