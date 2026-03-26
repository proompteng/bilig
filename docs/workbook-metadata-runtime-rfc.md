# Workbook Metadata Runtime RFC

## Relationship to existing docs

`docs/workbook-metadata-model.md` remains the concise contract. This RFC expands it into a runtime and migration design with ownership, interfaces, and rollout order.

## Current state

Workbook metadata is first-class in these checked-in areas:

- `@bilig/core` persists and exposes defined names, tables, spills, pivots, row and column metadata, calculation settings, volatile context, freeze panes, filters, and sorts
- snapshots carry workbook metadata
- formula compilation and execution consult metadata for names, structured references, and spill references where implemented
- replication carries broad metadata mutations through `@bilig/crdt`

The remaining metadata-dependent canonical rows are:

- `names:defined-name-range`
- `tables:table-total-row-sum`
- `structured-reference:table-column-ref`
- explicit spill-owner and binding surfaces inferred rather than modeled directly

## Goals

- define the runtime metadata domains the workbook model must own
- describe how metadata participates in binding, execution, snapshots, and replication
- make metadata changes authoritative transactions rather than side effects
- provide package boundaries and remaining migration order

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
- insert, delete, and move semantics
- resize and hide state
- downstream reference rewrite rules

## Ownership by package

- `@bilig/core`
  - authoritative runtime storage and transaction application
  - structural rewrite rules
  - snapshot import and export shape
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
- row and column structure affects reference rewriting and dependency expansion
- pivots must invalidate materialized regions and downstream consumers predictably

## Snapshot and replication consequences

- metadata must persist with workbook snapshots
- metadata must be expressible as authoritative workbook transactions
- restore and catch-up should rebuild the same metadata semantics as live mutation paths
- remote replay must not depend on app-local metadata reconstruction rules

## Remaining migration order

1. finish reference-valued defined-name semantics
2. finish table registry semantics required for canonical formulas
3. finish structured-reference binder and production execution integration
4. decide whether explicit spill-owner and spill-blocking ops are required beyond the current spill-range model
5. lower the remaining closed metadata semantics into WASM only after JS parity is proven

## Exit gate

- names, tables, structured refs, spills, pivots, and row and column structure all exist as first-class runtime model concepts
- metadata participates in binding, dependency analysis, and snapshot and replication paths
- no major parity family remains blocked by missing metadata shape alone
- JS remains the semantic oracle and WASM consumes lowered metadata only after parity is proven
