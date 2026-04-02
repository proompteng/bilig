# Typed Agent Protocol RFC

## Status

Archived historical RFC.

## Current reality

- The typed agent frame contract remains active.
- The production ingress for that contract is the monolith runtime in `apps/bilig`.
- The monolith executes worksheet operations in-process.
- The browser product path does not depend on the retired websocket browser-sync transport.

## Current proof points

- [apps/bilig/src/http/sync-server.ts](/Users/gregkonush/github.com/bilig/apps/bilig/src/http/sync-server.ts)
- [apps/bilig/src/workbook-runtime/document-session-manager.ts](/Users/gregkonush/github.com/bilig/apps/bilig/src/workbook-runtime/document-session-manager.ts)
- [packages/agent-api/src/index.ts](/Users/gregkonush/github.com/bilig/packages/agent-api/src/index.ts)

## Historical note

The earlier split between `apps/local-server` and `apps/sync-server` is obsolete and kept only for context in old discussions.
