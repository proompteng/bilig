# Browser Runtime

## Current state

- `apps/web` is the only supported browser shell.
- The browser boots through a dedicated worker runtime and consumes viewport patches instead of raw workbook state.
- The product path is Zero-backed:
  - workbook/session boot comes from `/v2/session`
  - authoritative workbook reads come from Zero tile queries
  - writes go through semantic Zero mutators
- The worker still owns parsing, optimistic previews, formatting helpers, clipboard behavior, and patch encoding.
- The browser no longer depends on the retired `apps/local-server` or `apps/sync-server` packages.

## Canonical boot flow

1. Load `runtime-config.json`.
2. Load `/v2/session` from the monolith.
3. Mount `ZeroProvider` with the session-derived `userID` and auth token.
4. Bootstrap the worker runtime.
5. Start `ZeroWorkbookBridge` subscriptions for workbook bootstrap, tile reads, selected-cell source, and axis metadata.
6. Project authoritative Zero rows back into the existing viewport-patch contract.
7. Layer optimistic worker previews over authoritative data until server state converges.

## Responsibilities

- UI thread:
  - rendering
  - input
  - formula bar
  - shell composition
  - connection-state presentation
- Worker runtime:
  - local parse / preview / formatting helpers
  - patch encoding
  - clipboard translation
  - selected-cell editing state
- Zero bridge:
  - tile subscriptions
  - authoritative workbook projection
  - patch emission into the grid cache

## Product rules

- The browser is server-authoritative in production mode.
- Disconnected or auth-failed Zero states must become read-only.
- `replaceSnapshot` is legacy-only debug plumbing, not the product write path.
- The grid contract remains viewport-patch based until a dedicated binary patch codec is introduced deliberately.

## Exit gate

- the production shell renders through Zero-backed viewport patches
- the shell uses real session-derived identity instead of hard-coded anonymous IDs
- common edits commit through Zero mutators and converge back through authoritative `cell_eval`
- reconnect, auth failure, and disconnected states are visible in the shell and disable unsafe writes

## See also

- [architecture.md](/Users/gregkonush/github.com/bilig/docs/architecture.md)
- [zero-bilig-production-implementation-plan-v2.md](/Users/gregkonush/github.com/bilig/docs/zero-bilig-production-implementation-plan-v2.md)
