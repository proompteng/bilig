# @bilig/crdt

Replica-state and batch ordering helpers for bilig workbook synchronization.

Workbook semantic op types such as `WorkbookOp`, `WorkbookTxn`, and `EngineOpBatch`
now live in `@bilig/workbook-domain`.

This package still re-exports those types for compatibility, but new code should
import them from `@bilig/workbook-domain`.

## Install

```bash
npm install @bilig/crdt
```

## Package entrypoints

- ESM: `./dist/index.js`
- Types: `./dist/index.d.ts`

## Migration

- Keep using `@bilig/crdt` for replica-state helpers such as `createReplicaState`,
  `createBatch`, `shouldApplyBatch`, `compareBatches`, and `compactLog`.
- Import transport-neutral workbook semantic types from `@bilig/workbook-domain`.

This package is part of the [bilig](https://github.com/proompteng/bilig) monorepo.
