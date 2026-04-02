# Production Zero mutation memory remediation - 2026-04-02

## Problem

`bilig-app` was crashing in production with V8 heap OOM during normal workbook bootstrap and edit traffic. The crash path was on the server-side Zero mutation route, not on ingress.

## Root cause

The hot mutation path exported a full workbook snapshot on every mutation in `apps/bilig/src/zero/server-mutators.ts` before persisting the authoritative source projection. For non-trivial workbooks that created an unnecessary deep clone on the request path.

At the same time, `WorkbookRuntimeManager` retained full snapshots in the in-memory runtime cache even after the engine had already imported the workbook, so the request-serving process carried duplicated workbook state.

## Fix

1. Removed `engine.exportSnapshot()` from the hot server mutation path.
2. Reworked source-projection persistence to materialize next-state rows directly from the live `SpreadsheetEngine`.
3. Reworked `WorkbookRuntimeManager` to cache engine plus authoritative source projection instead of engine plus workbook snapshot.
4. Kept full snapshot export only on the recalc checkpoint path, where durable checkpointing is required.
5. Added regression coverage for engine-backed source projection generation and runtime cache mutation commits.

## Files changed

- `apps/bilig/src/zero/server-mutators.ts`
- `apps/bilig/src/zero/store.ts`
- `apps/bilig/src/zero/projection.ts`
- `apps/bilig/src/workbook-runtime/runtime-manager.ts`
- `apps/bilig/src/zero/__tests__/projection.test.ts`
- `apps/bilig/src/zero/__tests__/runtime-manager.test.ts`

## Verification

Local verification for this change set:

- `pnpm exec vitest run apps/bilig/src/zero/__tests__/projection.test.ts apps/bilig/src/zero/__tests__/runtime-manager.test.ts`
- `pnpm typecheck`
- `pnpm run ci`

Production verification after push:

- confirm Forgejo built and pushed the new `bilig-app` image
- confirm Argo promoted the new image automatically
- verify `kubectl -n bilig get pods` shows fresh pods with no restarts
- verify `/healthz`, `/v2/session`, and `/` on `https://bilig.proompteng.ai`
- run a live Playwright smoke against the public hostname
