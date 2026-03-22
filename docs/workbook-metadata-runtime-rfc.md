# Workbook Metadata Runtime RFC

## Relationship to existing docs

`docs/workbook-metadata-model.md` remains the concise contract. This RFC expands it into a runtime and migration design with ownership, interfaces, and rollout order.

## Problem

Full spreadsheet parity is now blocked less by scalar builtin coverage and more by metadata and shape semantics. Names, tables, structured references, spills, pivots, and row/column structure need to become first-class runtime state that participates in binding, dependency analysis, snapshots, replication, and rendering.

## Goals

- define the runtime metadata domains the workbook model must own
- describe how metadata participates in binding, execution, snapshots, and replication
- make metadata changes authoritative transactions rather than side effects
- provide package boundaries and migration order

## Non-goals

- define every future Excel feature now
- make metadata a renderer concern
- push full metadata lowering into WASM before JS semantics are closed

## Metadata domains

### Workbook-scoped metadata

- defined names
- calculation settings
- volatile epoch and recalc context
- workbook view defaults where they affect durable semantics

### Sheet-scoped metadata

- stable sheet identity
- sheet order
- freeze panes, filters, sorts, and sheet-level view semantics where durable

### Table metadata

- table identity
- source bounds
- header configuration
- total rows
- column identities
- structured reference naming and resolution surfaces

### Dynamic-array metadata

- spill owner
- spill bounds
- blocked targets
- shape invalidation rules

### Pivot metadata

- pivot identity
- source range or source binding
- grouping and value definitions
- materialized output bounds
- delete and rebuild lifecycle

### Row and column structural metadata

- stable row and column identities independent of display index
- insert/delete/move semantics
- resize and hide state
- downstream reference rewrite rules

## Proposed runtime model

The workbook model should expose metadata as normalized, indexed state rather than a loose collection of ad hoc maps. A practical target shape is:

- workbook metadata root
- sheet records with stable IDs
- table registry
- defined-name registry
- spill registry
- pivot registry
- row/column structural registries

The exact storage form can stay optimized for the engine, but the semantic model should be explicit and testable.

## Ownership by package

- `@bilig/core`
  - authoritative runtime storage and transaction application
  - structural rewrite rules
  - snapshot import/export shape
- `@bilig/formula`
  - metadata-aware binding and evaluation rules
  - names, structured refs, and array semantics
- `@bilig/wasm-kernel`
  - lowered metadata only after JS semantics close
- `@bilig/protocol`
  - shared snapshot and transport-facing metadata types
- `@bilig/crdt`
  - authoritative mutation ordering for metadata ops

## Dependency consequences

Metadata must participate in dependency analysis, not sit beside it:

- defined names can point at values or references
- tables and structured references affect binding identity and downstream invalidation
- spill ownership affects visible result shape and error surfaces
- row/column structure affects reference rewriting and dependency expansion
- pivots must invalidate materialized regions and downstream consumers predictably

## Snapshot and replication consequences

- metadata must persist with workbook snapshots
- metadata must be expressible as authoritative workbook transactions
- restore and catch-up should rebuild the same metadata semantics as live mutation paths
- remote replay must not depend on app-local metadata reconstruction rules

## Suggested interface slices

### First tranche

- `DefinedNameRecord`
- `SpillRecord`
- `PivotRecord`
- authoritative metadata transaction ops for set/delete flows

### Second tranche

- `TableRecord`
- `StructuredReferenceBinding`
- row/column structural records and stable identities

### Third tranche

- calculation settings and more advanced workbook-level semantics
- filter/sort/freeze view structures where they affect durable behavior

## Migration order

1. move defined names onto authoritative transaction paths
2. formalize spill ownership and blocked-range semantics
3. complete pivot lifecycle including delete and output bounds
4. add table and structured-reference runtime model
5. add row/column structural identities and rewrite rules
6. add remaining workbook-level settings needed for parity families

## Suggested PR breakdown

1. metadata type normalization and snapshot shape cleanup
2. defined-name runtime and authoritative transaction migration
3. spill ownership runtime and invalidation rules
4. pivot lifecycle normalization
5. table registry and structured-reference binder integration
6. row/column structure runtime and rewrite rules
7. WASM metadata lowering follow-up for closed semantic families

## Exit gate

- names, tables, structured refs, spills, pivots, and row/column structure all exist as first-class runtime model concepts
- metadata participates in binding, dependency analysis, and snapshot/replication paths
- no major parity family remains blocked by missing metadata shape alone
- JS remains the semantic oracle and WASM consumes lowered metadata only after parity is proven
