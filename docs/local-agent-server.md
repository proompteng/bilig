# Local Agent Server

## Current state

The standalone `apps/local-server` package has been retired.

Local worksheet and agent-oriented behavior now lives inside `apps/bilig` modules:

- `src/http/local-server.ts`
- `src/workbook-runtime/local-*`
- `src/agent/*`
- `src/import-export/*`

## Product rule

There is one backend runtime.

Local-only workflows may still use the monolith's local listener for:

- worksheet import/export
- local harnesses
- agent-driven workbook operations
- compatibility/debug workflows

But local browser product behavior is not a separate application anymore.

## Exit gate

- no standalone local-server package is required for development, CI, or deployment
- local agent and workbook behavior is hosted by monolith modules only
- documentation points engineers to `apps/bilig`, not a retired parallel app
