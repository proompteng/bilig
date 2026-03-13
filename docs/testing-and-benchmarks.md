# Testing and Benchmarks

- Vitest covers protocol, formula parsing/evaluation, CRDT ordering, engine behavior, WASM parity, and playground reconciler behavior.
- Playwright drives a browser smoke test against the built Vite playground in `e2e/tests/`.
- `packages/benchmarks` emits JSON benchmark payloads for:
  - load scenarios at 10k, 50k, and 100k materialized cells through snapshot import
  - downstream edit scenarios at 100, 1k, and 10k dependent formulas
  - renderer commit-style scenarios at 1k and 10k declared cells
- `scripts/perf-smoke.mjs` enforces a lightweight CI threshold against a 1k-downstream edit and asserts that the run actually dirties the expected formulas and hits the WASM fast path.
- `scripts/bench-contracts.mjs` enforces the documented performance contracts in CI:
  - 100k materialized-cell load under 1.5s
  - 10k-downstream edit under 120ms end to end
  - 10k-downstream recalc under 50ms
  - 10k-cell render commit under 50ms
  - Forgejo CI applies a bounded runner-variance tolerance so local runs stay strict while remote `main` remains stable under normal host variance.
- `scripts/release-check.mjs` enforces artifact budgets after the playground build:
  - largest built app JS asset must stay under 350KB gzip
  - largest built WASM asset must stay under 250KB gzip

## CI policy

- Forgejo CI lives in `.forgejo/workflows/forgejo-ci.yml` and is the canonical pipeline for the private origin.
- The workflow targets the Forgejo runner label `bilig-ci`.
- The runner contract is explicit in the workflow:
  - host Node must be `>=24.14.0`
  - Corepack must activate `pnpm 10.32.1`
  - Chromium is installed during the run for browser smoke instead of relying on ambient runner state
- CI is strict by design:
  - `pnpm install --frozen-lockfile`
  - `pnpm run ci`
  - `git diff --exit-code` after the full pipeline
- `pnpm run ci` now includes:
  - WASM build
  - typecheck
  - unit/integration tests
  - perf smoke
  - benchmark contract checks
  - browser smoke
  - artifact budget enforcement
- Local pre-push verification should use `pnpm run ci:strict`, which mirrors the Forgejo cleanliness gate in one command.
- Browser smoke runs only after the playground build and installs Chromium explicitly so Forgejo runners do not rely on ambient browser state.
- The private origin keeps Forgejo-native workflow definitions only so Forgejo does not have to resolve duplicate workflow paths.
