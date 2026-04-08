# @bilig/workbook-domain

Transport-neutral workbook mutation language and batch types for bilig.

Use this package for workbook semantic types that should not depend on a transport,
runtime, or replica-state implementation, including:

- `WorkbookOp`
- `WorkbookTxn`
- `EngineOp`
- `EngineOpBatch`

Replica-state bookkeeping lives inside `@bilig/core`.
