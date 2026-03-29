# Actor Topology

## Web

- `appBootstrap`
- `session`
- `document`
- `zeroConnection`
- `workerRuntime`
- `selection`
- `editor`
- `clipboard`
- `ribbon`
- `dialogs`

The web app subscribes via `useSelector` only to narrow slices of these actors. React components are presentational and dispatch events; they do not own authoritative workflow state.

### Implemented in `main`

- `appBootstrap`
- `workerRuntime`

`workerRuntime` now owns worker creation, bootstrap, runtime-state refresh, selected-cell refresh, cache-driven invalidation, selection forwarding, and optional Zero bridge subscriptions. `WorkerWorkbookApp` renders the shell and dispatches actor events instead of orchestrating the worker lifecycle directly.

## Worker

- `workerBootstrap`
- `engineLifecycle`
- `snapshotHydration`
- `browserPersistence`
- `syncTransport`
- `viewportProjection`
- `recovery`

The worker runtime owns bootstrap, hydration, reconnect, and recovery states. Spreadsheet compute remains in pure engine structures.

## Sync / Local Servers

- `documentSupervisor`
- `presence`
- `batchLog`
- `snapshotAssembly`
- `browserSubscribers`
- `agentSessions`
- `importExport`
- `recalc`
- `leaseOwnership`

One document supervisor actor owns one document’s long-lived lifecycle.

## Zero

- `zeroQueryIngress`
- `zeroMutatorIngress`

These actors validate, decode, route, and supervise request execution. They do not own spreadsheet semantics.
