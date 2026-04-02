# Local Agent Server

## Status

Retired as a standalone HTTP surface.

## Current model

- There is no separate local agent server process in the product runtime.
- Worksheet execution now happens in-process inside `apps/bilig`.
- Agent ingress still enters through `/v2/agent/frames` on the monolith.

## Proof points

- [apps/bilig/src/index.ts](/Users/gregkonush/github.com/bilig/apps/bilig/src/index.ts)
- [apps/bilig/src/workbook-runtime/local-document-supervisor.ts](/Users/gregkonush/github.com/bilig/apps/bilig/src/workbook-runtime/local-document-supervisor.ts)
- [apps/bilig/src/workbook-runtime/worksheet-executor.ts](/Users/gregkonush/github.com/bilig/apps/bilig/src/workbook-runtime/worksheet-executor.ts)
