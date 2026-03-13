# Testing and Benchmarks

- Vitest covers protocol, formula parsing/evaluation, CRDT ordering, engine behavior, WASM parity, and playground reconciler behavior.
- Playwright is reserved for browser smoke coverage in CI.
- `packages/benchmarks` emits JSON benchmark payloads for edit and load scenarios.
- `scripts/perf-smoke.mjs` enforces a lightweight smoke threshold for CI.

## CI policy

- Forgejo CI lives in `.forgejo/workflows/forgejo-ci.yml` and is the canonical pipeline for the private origin.
- The workflow targets Forgejo runners with the `docker` label and runs inside a pinned `node:24.14.0-bookworm` container so the toolchain is explicit.
- CI is strict by design:
  - `pnpm install --frozen-lockfile`
  - `pnpm run ci`
  - `git diff --exit-code` after the full pipeline
- The private origin keeps Forgejo-native workflow definitions only so Forgejo does not have to resolve duplicate workflow paths.
