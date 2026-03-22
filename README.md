# bilig

`bilig` is a local-first spreadsheet engine monorepo with a package-based custom workbook reconciler, a React/Vite playground shell, a framework-agnostic core engine, replication-ready mutation pipelines, and an AssemblyScript/WASM numeric fast path.

It already has the foundations of a serious spreadsheet/runtime stack: a real engine, a real local session loop, a real binary sync layer, a real reconciler, and a reasonably mature grid shell. The biggest remaining gap is not basic spreadsheet arithmetic; it is the seam between what the local engine can represent and what the authoritative replicated model can express.

## Workspace layout

- `apps/playground`: Vite 8 React app shell that composes the packages
- `packages/protocol`: shared enums, opcodes, constants, and types
- `packages/formula`: A1 addressing, lexer, parser, binder, compiler, JS evaluator
- `packages/core`: spreadsheet engine, storage, scheduler, snapshots, selectors, WASM facade
- `packages/crdt`: replica clocks, op batches, merge rules, log compaction
- `packages/renderer`: custom workbook reconciler and workbook DSL
- `packages/grid`: reusable React spreadsheet UI, hooks, selection, metrics, and inspectors
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
pnpm protocol:generate
pnpm build
pnpm typecheck
pnpm test
pnpm bench
pnpm bench:contracts
pnpm release:check
pnpm run ci
```

## Notes

- Reusable React code now lives in `packages/renderer` and `packages/grid`; `apps/playground` is a thin shell.
- The spreadsheet engine remains usable without React.
- The custom reconciler lives under `packages/renderer`.
- The public cell model supports `format` as a persisted attribute alongside `addr`, `value`, and `formula`.
- The WASM kernel is a custom AssemblyScript fast path, not an embedded proprietary spreadsheet runtime.
- The TS protocol enums/opcodes and AssemblyScript protocol mirror are generated from `scripts/gen-protocol.ts` so JS/WASM ABI drift fails fast in CI.
- The playground includes a scroll-windowed sheet surface, sheet tabs, keyboard cell navigation, dependency inspection, and recalc metrics.
- The playground operator surface now spans a 100k-row by 256-column virtualized window while keeping the engine hard limits at 1,048,576 rows by 16,384 columns.
- The demo workbook now exercises JS row/column range formulas and a WASM-backed branch formula in the visible UI so browser smoke covers both paths.
- The cell inspector now exposes formula mode, topo rank, versioning, and dependency edges from the core engine.
- The playground also demonstrates local-first replica mirroring through the engine’s outbound and inbound batch APIs.
- The playground persists workbook and replica snapshots in local storage so the demo survives reloads as a local-first app surface.
- The playground relay queue now persists paused replica traffic across reloads, so offline-style catch-up is visible instead of being memory-only.
- The paused relay queue is compacted with the CRDT entity-order rules, so repeated offline edits do not grow an unbounded replay backlog for the same cell or sheet entity.
- The imperative engine now includes a single-sheet CSV bridge for import/export without pulling React into shared packages.
- CI now enforces performance contracts for 100k snapshot load, 10k-downstream edits, and 10k-cell render commits instead of relying on a loose smoke check alone.
- `pnpm release:check` enforces the documented production budgets for the built app JS and bundled WASM asset.
- The next highest-leverage architecture work is to make the authoritative workbook op model exhaustive enough to match the local engine surface, then build worker-first runtime, durable multiplayer, and typed binary agent work on top of that seam.

## CI

- Forgejo Actions is the primary CI surface for this repo via `.forgejo/workflows/forgejo-ci.yml`.
- GitHub Actions mirrors the verification contract in `.github/workflows/ci.yml`.
- The workflow is strict: frozen lockfile install, full `pnpm run ci`, artifact budget checks, browser smoke, and a tracked-file cleanliness check.
- Forgejo runners must expose the `bilig-ci` label, provide Node `24.x`, and allow Corepack to activate `pnpm 10.32.1` during the job.
- GitHub Actions runs the same repository contract on Node 22 and Node 24.11.1 so compatibility drift is visible before release.
