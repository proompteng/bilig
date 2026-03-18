# Implementation Ledger

This ledger maps acceptance rows to concrete checks so the docs, code, and release gates stay aligned.

## Current tranche coverage

| Acceptance row | Current proof |
| --- | --- |
| documented `@bilig/core` APIs exist in code | `packages/core/src/engine.ts`, `pnpm typecheck` |
| range mutation helpers are tested | `packages/core/src/__tests__/engine.test.ts` |
| undo/redo is tested | `packages/core/src/__tests__/engine.test.ts` |
| selection state and sync state are tested | `packages/core/src/__tests__/engine.test.ts`, `packages/core/src/__tests__/selectors.test.ts` |
| worker transport package exists and is tested | `packages/worker-transport/src/__tests__/worker-transport.test.ts` |
| binary protocol package exists and is tested | `packages/binary-protocol/src/__tests__/binary-protocol.test.ts` |
| sync-server app exists and is tested | `apps/sync-server/src/__tests__/sync-server.test.ts` |
| local-server app exists and is tested | `apps/local-server/src/__tests__/local-server.test.ts` |
| shipping browser app wrapper exists | `apps/web`, `pnpm --filter @bilig/web build` |

## Open rows

These rows are still open and should not be treated as complete:

- worker-first browser runtime is active
- offline restore and reconnect are proven under the worker runtime
- chat-driven local agent orchestration is wired to the local app server
- checked-in Excel parity corpus exists beyond the seed suite
- JS oracle matches the Excel parity corpus across the full target surface
- WASM overlap matches JS on every claimed kernel
- durable append-before-ack backend is proven
- cursor catch-up and snapshot restore are proven against durable storage
- remote agent execution is backed by live worksheet sessions
- performance budgets are green for the full canonical product, not just the current foundation slice

## Release rule

No acceptance row is closed until:

1. the code exists,
2. the docs describe the current and target state accurately,
3. an automated check or explicit production proof is linked to that row.
