# Authoritative Workbook Op Model RFC

## Problem

The current repo has a stronger local engine surface than replicated mutation surface. `@bilig/core` can already represent workbook metadata and local engine behavior that the current `@bilig/crdt` op family does not yet express as first-class replicated state. That mismatch is now the highest-leverage architecture seam in the system.

## Goals

- define an exhaustive authoritative workbook op family
- move every stateful engine mutation behind a transaction executor
- make snapshots acceleration artifacts rather than semantic truth
- make undo/redo, replication, and local mutation paths consume the same transaction model
- give protocol, storage, browser runtime, and service layers one canonical mutation language

## Non-goals

- rewrite the engine around a new storage layout
- move all current and future behavior into one PR
- treat React descriptors or viewport patches as the authoritative workbook model

## Current gap

The existing replicated model in `packages/crdt/src/index.ts` covers:

- workbook upsert
- sheet upsert/delete
- cell value/formula/format
- clear cell
- pivot upsert

The local engine and workbook store already have broader stateful behavior around:

- defined names
- spill ownership and spill cleanup
- pivot deletion and pivot output lifecycle
- structural metadata that should not remain app-local forever
- future tables, structured references, row/column semantics, and workbook settings

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

### Proposed op families

#### Workbook-level ops

- `upsertWorkbook`
- `setWorkbookSetting`
- `setCalculationSetting`
- `setWorkbookViewSetting`

#### Sheet lifecycle ops

- `upsertSheet`
- `renameSheet`
- `deleteSheet`
- `reorderSheets`
- `setSheetView`

#### Row and column structure ops

- `insertRows`
- `deleteRows`
- `moveRows`
- `resizeRows`
- `hideRows`
- `insertColumns`
- `deleteColumns`
- `moveColumns`
- `resizeColumns`
- `hideColumns`

#### Cell content ops

- `setCellValue`
- `setCellFormula`
- `setCellFormat`
- `clearCell`
- `setRangeValues`
- `setRangeFormulas`
- `clearRange`
- `fillRange`
- `copyRange`
- `pasteRange`

#### Workbook metadata ops

- `setDefinedName`
- `deleteDefinedName`
- `upsertTable`
- `deleteTable`
- `setStructuredReferenceBinding`
- `setSpillOwner`
- `clearSpillOwner`
- `upsertPivotTable`
- `deletePivotTable`

#### UI-adjacent but durable worksheet semantics

- `setFreezePane`
- `setFilter`
- `clearFilter`
- `setSort`

The rule is not that all of these must be implemented immediately. The rule is that the authoritative model must have a place for them so local-only mutation paths stop multiplying.

## Transaction executor boundary

`@bilig/core` should expose one transaction executor boundary:

- validate input
- normalize into authoritative ops
- apply ops to workbook model and metadata store
- emit dependency invalidation and recalc scheduling effects
- append to undo/redo history
- publish outbound replication batches
- notify subscribers

All current direct engine mutators should migrate toward wrappers that call the executor. The executor becomes the place where invariants live.

## Undo/redo model

Undo/redo should stop being a special local-only stack of bespoke state mutations. It should become transaction-log based:

- each committed transaction is recorded with enough metadata to invert or replay
- undo produces an inverse transaction or replay step, not an ad hoc local branch
- redo replays the original transaction sequence

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
  - conceptually closer to an oplog/replication-model package than a broad CRDT abstraction
- `@bilig/core`
  - transaction executor, invariant enforcement, graph invalidation, history integration
- `@bilig/binary-protocol`
  - transport encoding for transactions and cursors
- `@bilig/storage-browser` and `@bilig/storage-server`
  - persistence of snapshots, transactions, and cursors
- `apps/local-server` and `apps/sync-server`
  - ordered transaction ingress, replay, catch-up, and fanout

## Migration order

1. define the expanded op family and transaction envelope types
2. add transaction executor scaffolding in `@bilig/core`
3. migrate defined names and pivot delete lifecycle first, because they already expose the seam clearly
4. migrate spill ownership into authoritative ops
5. add row/column structural ops and stable identities
6. move history to transaction-log semantics
7. update transport and storage paths to treat transactions as the canonical unit

## Suggested PR breakdown

1. shared op-family expansion in `@bilig/crdt` and protocol-facing types
2. transaction executor introduction in `@bilig/core`
3. defined-name replication and transaction migration
4. pivot delete and spill lifecycle transaction migration
5. row/column structural op introduction
6. undo/redo refactor onto transaction history
7. binary protocol and persistence alignment
8. browser/local/sync runtime adoption pass

## Exit gate

- no stateful workbook mutation remains local-only by default
- defined names, spill ownership, and pivot deletion are authoritative ops
- undo/redo operates on transaction semantics rather than bespoke local branches
- snapshots are acceleration artifacts rather than the semantic source of truth
- browser, local-server, and sync-server all consume the same authoritative transaction model
