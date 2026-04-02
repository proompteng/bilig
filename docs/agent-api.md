# Agent API

## Status

Current summary for the monolith runtime.

## Current behavior

- Agent ingress is served by `apps/bilig`.
- Remote agent calls enter through `/v2/agent/frames` on the monolith.
- Worksheet execution is handled in-process by the embedded worksheet host, not by a second local HTTP server.
- Workbook import returns browser URLs rooted at the monolith public base URL.

## Proof points

- [apps/bilig/src/http/sync-server.ts](/Users/gregkonush/github.com/bilig/apps/bilig/src/http/sync-server.ts)
- [apps/bilig/src/workbook-runtime/document-session-manager.ts](/Users/gregkonush/github.com/bilig/apps/bilig/src/workbook-runtime/document-session-manager.ts)
- [apps/bilig/src/workbook-runtime/local-document-supervisor.ts](/Users/gregkonush/github.com/bilig/apps/bilig/src/workbook-runtime/local-document-supervisor.ts)
- [apps/bilig/src/workbook-runtime/worksheet-executor.ts](/Users/gregkonush/github.com/bilig/apps/bilig/src/workbook-runtime/worksheet-executor.ts)

## Historical note

Any earlier references to a standalone `apps/local-server` execution surface are no longer current.
