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
pnpm run ci
```

## Notes

- React is used only in `apps/playground`.
- The spreadsheet engine remains usable without React.
- The custom reconciler lives under `apps/playground/src/reconciler`.
- The WASM kernel is a custom AssemblyScript fast path, not an embedded proprietary spreadsheet runtime.
- The playground includes a scroll-windowed sheet surface, sheet tabs, keyboard cell navigation, dependency inspection, and recalc metrics.
- The playground also demonstrates local-first replica mirroring through the engine’s outbound and inbound batch APIs.

## CI

- Forgejo Actions is the primary CI surface for this repo via `.forgejo/workflows/forgejo-ci.yml`.
- The workflow is strict: frozen lockfile install, full `pnpm run ci`, and a tracked-file cleanliness check.
- Forgejo runners must expose the `docker` label because the workflow uses a Node 24 container job instead of assuming a host toolchain.
