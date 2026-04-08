# Workbook Metadata Runtime RFC

## Current state

Workbook metadata is now carried by the transaction-based workbook engine and projected into the Zero-backed relational model owned by the monolith.

Closed runtime facts:

- `@bilig/core` persists and exposes defined names, tables, pivots, spills, row and column metadata, calc settings, freeze panes, filters, and sorts.
- `@bilig/workbook-domain` owns the transport-neutral workbook operation language.
- `@bilig/core` owns the replica-state helpers needed for local replay and snapshot restore.
- `apps/bilig/src/zero/*` materializes authoritative workbook metadata into Postgres-backed rows for the browser path.

## Goals

- keep metadata first-class in the workbook runtime
- keep metadata serialized through one canonical transaction language
- keep Zero syncing relational metadata rows, not snapshot blobs
- keep metadata available to the binder, evaluator, render materializer, and browser bridge without re-inventing semantics in the UI

## Ownership by package

- `@bilig/core`
  - authoritative runtime storage
  - structural rewrite rules
  - snapshot import/export
- `@bilig/formula`
  - metadata-aware binding and evaluation
  - names, structured references, and spill semantics
- `@bilig/protocol`
  - snapshot and transport-facing metadata types
- `@bilig/workbook-domain`
  - transport-neutral workbook ops and transactions
- `apps/bilig/src/zero`
  - relational source/eval materialization for the Zero product path

## Product model

Metadata must participate in:

- binding
- dependency analysis
- structural rewrites
- render materialization
- snapshot rebuilds
- replay from `workbook_event`

Metadata must not depend on browser-only reconstruction rules.

## Runtime implications

- defined names, tables, pivots, spills, row and column metadata, filters, and sorts are durable workbook facts
- metadata mutations must be serializable per workbook and replayable through `workbook_event`
- render materialization must write only the derived rows needed for the viewport path
- snapshots remain warm-start artifacts rather than the synced product model

## Exit gate

- workbook metadata remains a first-class runtime concept in the engine and relational model
- metadata participates in binder/runtime behavior and rebuilds from events plus snapshots
- no production browser behavior depends on reconstructing metadata ad hoc in the client
