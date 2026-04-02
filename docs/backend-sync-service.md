# Backend Sync Service

## Status

Current summary for the monolith backend.

## Current backend runtime

- The only supported backend runtime is `apps/bilig`.
- It serves the browser shell, session bootstrap, agent ingress, Zero query/mutate endpoints, and the authoritative workbook runtime.
- The product path does not depend on the retired `apps/sync-server` package.

## Current proof points

- [apps/bilig/src/index.ts](/Users/gregkonush/github.com/bilig/apps/bilig/src/index.ts)
- [apps/bilig/src/http/sync-server.ts](/Users/gregkonush/github.com/bilig/apps/bilig/src/http/sync-server.ts)
- [apps/bilig/src/zero/service.ts](/Users/gregkonush/github.com/bilig/apps/bilig/src/zero/service.ts)
