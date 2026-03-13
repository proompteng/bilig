# bilig

`bilig` is a local-first spreadsheet engine monorepo with a custom workbook reconciler, a React/Vite playground, a framework-agnostic core engine, CRDT-ready mutation pipelines, and an AssemblyScript/WASM numeric fast path.

## Workspace layout

- `apps/playground`: Vite 8 React app, custom reconciler, and playground UI
- `packages/protocol`: shared enums, opcodes, constants, and types
- `packages/formula`: A1 addressing, lexer, parser, binder, compiler, JS evaluator
- `packages/core`: spreadsheet engine, storage, scheduler, snapshots, selectors, WASM facade
- `packages/crdt`: replica clocks, op batches, merge rules, log compaction
- `packages/wasm-kernel`: AssemblyScript VM and numeric kernels
- `packages/benchmarks`: benchmark harness
- `docs`: architecture, API, reconciler layering, CRDT model, formula language

## Quickstart

```bash
pnpm install
pnpm wasm:build
pnpm typecheck
pnpm test
pnpm dev
```

## Commands

```bash
pnpm dev
pnpm build
pnpm typecheck
pnpm test
pnpm bench
pnpm release:check
pnpm run ci
pnpm run ci:strict
```

## Notes

- React is used only in `apps/playground`.
- The spreadsheet engine remains usable without React.
- The custom reconciler lives under `apps/playground/src/reconciler`.
- The WASM kernel is a custom AssemblyScript fast path, not an embedded proprietary spreadsheet runtime.
- The playground includes a scroll-windowed sheet surface, sheet tabs, keyboard cell navigation, dependency inspection, and recalc metrics.
- The demo workbook now exercises JS row/column range formulas and a WASM-backed branch formula in the visible UI so browser smoke covers both paths.
- The cell inspector now exposes formula mode, topo rank, versioning, and dependency edges from the core engine.
- The playground also demonstrates local-first replica mirroring through the engine’s outbound and inbound batch APIs.
- The playground persists workbook and replica snapshots in local storage so the demo survives reloads as a local-first app surface.
- The imperative engine now includes a single-sheet CSV bridge for import/export without pulling React into shared packages.
- `pnpm release:check` enforces the documented production budgets for the built app JS and bundled WASM asset.

## CI

- Forgejo Actions is the primary CI surface for this repo via `.forgejo/workflows/forgejo-ci.yml`.
- The workflow is strict: frozen lockfile install, full `pnpm run ci`, artifact budget checks, browser smoke, and a tracked-file cleanliness check.
- `pnpm run ci:strict` mirrors the remote cleanliness gate locally before a direct push to `main`.
- Forgejo runners must expose the `docker` label because the workflow uses a Node 24 container job instead of assuming a host toolchain.
