# Testing and Benchmarks

- Vitest covers protocol, formula parsing/evaluation, CRDT ordering, engine behavior, WASM parity, and playground reconciler behavior.
- Playwright drives a browser smoke test against the built Vite playground in `e2e/tests/`.
- `packages/benchmarks` emits JSON benchmark payloads for edit and load scenarios.
- `scripts/perf-smoke.mjs` enforces a lightweight smoke threshold for CI and asserts that the run actually hits the WASM fast path.

## CI policy

- Forgejo CI lives in `.forgejo/workflows/forgejo-ci.yml` and is the canonical pipeline for the private origin.
- The workflow targets Forgejo runners with the `docker` label and runs inside a pinned `node:24.14.0-bookworm` container so the toolchain is explicit.
- CI is strict by design:
  - `pnpm install --frozen-lockfile`
  - `pnpm run ci`
  - `git diff --exit-code` after the full pipeline
- Local pre-push verification should use `pnpm run ci:strict`, which mirrors the Forgejo cleanliness gate in one command.
- Browser smoke runs only after the playground build and installs Chromium explicitly so Forgejo runners do not rely on ambient browser state.
- The private origin keeps Forgejo-native workflow definitions only so Forgejo does not have to resolve duplicate workflow paths.
