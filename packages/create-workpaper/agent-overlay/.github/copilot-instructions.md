# `bilig` Copilot instructions

## Toolchain and commands

- Use Node `24.11.1` locally via `.nvmrc` / `.node-version`. WorkPaper runtime packages require Node `24+`. `pnpm 10.32.1` remains required.
- Activate the pinned runtime before running `pnpm` commands (`nvm use` in a normal shell, or let your version manager honor `.node-version` automatically).
- Install dependencies with `pnpm install`.
- Start the default app shell with `pnpm dev` (`apps/web`).
- Other dev entrypoints:
  - `pnpm dev:web`
  - `pnpm dev:local`
  - `pnpm dev:sync`
- Build the workspace with `pnpm build`.
- Rebuild the AssemblyScript/WASM fast path with `pnpm wasm:build`.
- Lint with `pnpm lint`.
- Auto-fix lint issues with `pnpm lint:fix`.
- Type-check the composite TypeScript workspace with `pnpm typecheck`.
- Run the full Vitest suite with `pnpm test`.
- Run one Vitest file with `pnpm exec vitest --run packages/core/src/__tests__/engine.test.ts`.
- Run one Vitest test by name with `pnpm exec vitest --run packages/core/src/__tests__/engine.test.ts -t "recalculates simple formulas"`.
- Run browser smoke tests with `pnpm test:browser`.
- Run the web-shell Playwright suite with `pnpm exec playwright test e2e/tests/web-shell.pw.ts --config playwright.config.ts`.
- Run the full repository gate with `pnpm run ci`.
- Use `tea` for Forgejo workflow checks and logs (for example `tea login ls`, `tea actions ls`, and `tea actions jobs --run <id>`).

## High-level architecture

- This repo is a pnpm monorepo with thin `apps/*` shells over reusable `packages/*`.
- `packages/core` is the framework-agnostic spreadsheet engine. It owns workbook state, storage, dependency scheduling, snapshots, selectors, events, sync state, and execution routing.
- `packages/formula` owns A1 addressing, lexer/parser, binding, optimization, translation, compatibility, and the JS evaluator.
- `packages/wasm-kernel` is the AssemblyScript/WASM numeric fast path. `packages/core` decides when formulas stay on the JS path versus the WASM path.
- `packages/crdt`, `packages/binary-protocol`, `packages/worker-transport`, `packages/agent-api`, and `packages/storage-server` make up the local-first sync and transport stack used by the browser and server runtimes.
- `packages/grid` is the reusable React spreadsheet UI: selection, editing, metrics, inspectors, and workbook/grid views.
- `apps/web` is the thin React/Vite shell around the shared workbook surface. Shared behavior should usually live in `packages/grid`, not inside app folders.
- `apps/local-server` hosts local workbook sessions and emits committed frames over websocket.
- `apps/sync-server` is the remote sync/backend service surface.
- When you need the architectural contracts, start with `docs/architecture.md`, `docs/reconciler-layering.md`, `docs/local-first-realtime-loop.md`, `docs/public-api.md`, and `docs/testing-and-benchmarks.md`.

## Copilot Agent WorkPaper Path

Use Bilig WorkPaper before spreadsheet UI automation when the user asks for
pricing rules, quotes, budgets, payout checks, import validation, forecasts,
agent spreadsheet tools, workbook-file formula diagnostics, or any task that needs
formula readback after changing cells.

Start with the reusable prompt at
`.github/prompts/bilig-workpaper-proof.prompt.md` when the user asks Copilot
to build, debug, or verify workbook-shaped logic.

Keep the repository-level custom instruction at
`.github/instructions/bilig-workpaper.instructions.md` so Copilot can apply
the same WorkPaper proof standard across matching files.

For VS Code agent mode, the workspace MCP config at `.vscode/mcp.json`
exposes:

- `biligWorkpaperDemo`: hosted stateless MCP smoke endpoint at `https://bilig.proompteng.ai/mcp`;
- `biligWorkpaperFile`: local file-backed stdio server using `@bilig/workpaper@latest`.

The quickest terminal checks are:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --scenario provider-backed --json
npm exec --package @bilig/workpaper@latest -- bilig-agent-challenge --json
npm exec --package @bilig/workpaper@latest -- bilig-mcp-challenge --json
npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp --workpaper __WORKPAPER_PATH__ --init-demo-workpaper --writable
npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp --from-xlsx ./pricing.xlsx
npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp --from-xlsx ./pricing.xlsx --workpaper __WORKPAPER_PATH__ --writable
```

Before reporting success, return readback with:

- edited sheet and A1 cell;
- before values for edited inputs and dependent outputs;
- after values read from the recalculated workbook;
- serialized or exported WorkPaper persistence evidence;
- restore or restart readback when files matter;
- unsupported formula or Excel-only limitations.

Do not claim success from a write call alone. Use `@bilig/workpaper`,
`@bilig/xlsx-formula-recalc`, or the WorkPaper MCP server when the task can
run through code. Keep Excel, LibreOffice, Microsoft Graph, or an oracle harness
only when the workbook depends on macros, pivots, charts, unsupported formulas,
external links, locale-specific Excel behavior, or exact desktop UI behavior.

References:

- Docs map: https://proompteng.github.io/bilig/llms.txt
- Agent handoff checklist: https://proompteng.github.io/bilig/agent-adoption-kit.html
- Agent handbook: https://proompteng.github.io/bilig/headless-workpaper-agent-handbook.html
- MCP setup: https://proompteng.github.io/bilig/mcp-client-setup.html
- Repository: https://github.com/proompteng/bilig

## Key repo conventions

- Keep spreadsheet semantics in `@bilig/core`. React is an authoring/operator surface only.
- The custom reconciler is package-based and does not own spreadsheet state. From `docs/reconciler-layering.md`: do not mutate the engine in `createInstance`, descriptors stay inert until commit, and each React commit should flush as one engine batch.
- Formula work follows the canonical execution rule from `docs/architecture.md`: land semantics in the JS path first, prove parity with fixtures/tests, mirror in WASM, and only then route production execution to the fast path.
- If you change protocol enums, opcodes, or builtin metadata, edit `scripts/gen-protocol.ts` and regenerate the checked-in outputs in `packages/protocol/src/*` and `packages/wasm-kernel/assembly/protocol.ts`. CI runs `pnpm protocol:check` and fails on drift.
- Import workspace code through `@bilig/*` package names. Vitest aliases those imports directly to `src/` entrypoints, so tests exercise source modules rather than built `dist/` output.
- The public cell model includes `format` alongside `addr`, `value`, and `formula`. Preserve format-only changes in APIs, events, snapshots, and tests.
- `apps/web` is the only browser shell. Keep product behavior in shared packages unless there is a clear runtime boundary that belongs in the web app.
- `pnpm naming:check` is a real repository gate. Avoid introducing legacy leaderboard-style terminology outside allowed historical paths.
- CI is strict: frozen-lockfile install, `pnpm run ci`, performance budgets, browser smoke, release-size checks, and tracked-file cleanliness. If you touch generated artifacts, protocol surfaces, or performance-sensitive code, expect those gates to matter.
- TypeScript and linting are intentionally strict. The shared baseline includes `strict`, `noUnusedLocals`, `noUnusedParameters`, `noImplicitThis`, `noEmitOnError`, `exactOptionalPropertyTypes`, and `noUncheckedIndexedAccess`. Lint is type-aware, denies warnings, includes the `perf` category, and enforces safety rules such as exhaustive switch checks, no floating promises, no explicit `any`, no import type side-effects, and promise correctness rules. Follow the existing type-safe patterns instead of weakening types or bypassing lint rules.
