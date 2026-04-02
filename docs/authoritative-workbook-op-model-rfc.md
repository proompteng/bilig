# Authoritative Workbook Operation Model RFC

## Status

Archived historical RFC. The authoritative workbook model is now implemented inside the monolith runtime.

## Current production contract

- `apps/bilig` owns authoritative mutation ordering and persistence.
- Zero-backed relational source and eval rows are the browser-facing sync surface.
- Workbook checkpoints are warm-start and recovery artifacts, not the hot-path sync model.
- The browser renders viewport patches through `ZeroWorkbookBridge` and the worker cache.

## Active proof points

- [apps/bilig/src/zero/server-mutators.ts](/Users/gregkonush/github.com/bilig/apps/bilig/src/zero/server-mutators.ts)
- [apps/bilig/src/zero/store.ts](/Users/gregkonush/github.com/bilig/apps/bilig/src/zero/store.ts)
- [apps/web/src/zero/ZeroWorkbookBridge.ts](/Users/gregkonush/github.com/bilig/apps/web/src/zero/ZeroWorkbookBridge.ts)

## Historical note

References in the earlier draft to separate `local-server` and `sync-server` authorities are obsolete.
