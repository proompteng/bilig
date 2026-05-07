# bilig

[![CI](https://github.com/proompteng/bilig/actions/workflows/ci.yml/badge.svg)](https://github.com/proompteng/bilig/actions/workflows/ci.yml)
[![GitHub Repo stars](https://img.shields.io/github/stars/proompteng/bilig?style=social)](https://github.com/proompteng/bilig/stargazers)
[![npm: @bilig/headless](https://img.shields.io/npm/v/@bilig/headless?label=%40bilig%2Fheadless)](https://www.npmjs.com/package/@bilig/headless)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

**bilig is a local-first spreadsheet engine and runtime for browser workbooks,
agent workflows, and collaborative sync.**

Project site: <https://proompteng.github.io/bilig/>

If the WorkPaper package is relevant to your agent or Node workflow, star the
repo as a bookmark: <https://github.com/proompteng/bilig/stargazers>

If you want to try a small contribution first, start with the public
[`starter issues`](docs/starter-issues.md) list.

It is not a table widget. The repo contains a real workbook engine, formula
parser/compiler, React workbook reconciler, reusable grid shell, binary sync
protocol, agent API, browser/server persistence layers, and a conservative
AssemblyScript/WASM fast path for formula families that have proven parity.

The long-term target is a spreadsheet platform that can be edited by people or
agents, restored locally, synchronized through ordered mutation streams, and
benchmarked against serious spreadsheet-engine workloads.

## Why Watch This Repo

- **Spreadsheet engine, not just UI**: workbook mutation, formulas, snapshots,
  history, selections, dependency inspection, replica hooks, and import/export
  live below the React shell.
- **Local-first by design**: browser sessions restore from local state, preserve
  replica snapshots, and keep outbound edits as replayable mutation batches.
- **Agent-addressable workbooks**: the engine exposes stable request, response,
  event, and subscription shapes so agents can operate on spreadsheet state
  without screen scraping.
- **Performance tied to proof**: formula acceleration and WorkPaper benchmark
  work are backed by parity fixtures, differential checks, counters, and CI
  gates instead of benchmark-only rewrites.
- **Reusable package boundaries**: formula, core, grid, renderer, transport,
  protocol, storage, benchmark, and runtime concerns are split into packages.

## What Works Today

- Create, mutate, snapshot, restore, undo, redo, and inspect workbooks through
  `@bilig/core`.
- Parse, bind, compile, and evaluate spreadsheet formulas through
  `@bilig/formula`, with fixture-driven parity checks.
- Render and navigate a virtualized browser spreadsheet shell through
  `apps/web` and `@bilig/grid`.
- Author deterministic workbooks with React components through
  `@bilig/renderer`.
- Exercise the product runtime through the `apps/bilig` monolith, which serves
  the built web shell and backend APIs.
- Run WorkPaper and browser performance contracts from `packages/benchmarks`,
  `scripts/`, and `e2e/tests`.
- Build the AssemblyScript WASM kernel with `pnpm wasm:build`.

## Current Status

bilig is early, serious infrastructure. The architecture is broad and the
correctness bar is intentionally high, but it is not a finished Excel clone.

Known open areas include:

- full Excel formula parity
- defined names, tables, structured references, and deeper dynamic-array support
- worker-first browser runtime as the default boot path
- final durable multiplayer sync backend
- typed binary agent frames end to end
- more public package release hardening

## Headless WorkPaper In Five Minutes

Start here when you want to use the spreadsheet engine from Codex, Claude Code,
a service, or a Node script without opening the browser UI.

`@bilig/headless` is production-ready for applications that call the documented
WorkPaper API directly. The package README is the contract for install, API
usage, persistence, validation, supported scope, and agent workflow:
[packages/headless/README.md](packages/headless/README.md).

Install from npm:

```bash
pnpm add @bilig/headless
```

For a runnable external-consumer example, start with
[examples/headless-workpaper](examples/headless-workpaper). The repository smoke
test executes that same example against packed local runtime packages with
`pnpm workpaper:smoke:external`.

Minimal example:

```ts
import { WorkPaper, type WorkPaperCellAddress } from '@bilig/headless'

const workbook = WorkPaper.buildFromSheets(
  {
    Sheet1: [
      [10, 20, '=A1+B1'],
      [7, '=A2*3', null],
    ],
  },
  { maxRows: 1_000, maxColumns: 100, useColumnIndex: true },
)

const sheet = workbook.getSheetId('Sheet1')
if (sheet === undefined) {
  throw new Error('Sheet1 was not created')
}

const at = (row: number, col: number): WorkPaperCellAddress => ({ sheet, row, col })

workbook.setCellContents(at(1, 2), '=A2+B2')
console.log(workbook.getCellValue(at(1, 2)))
```

Rules for agents:

- Use public package exports from `@bilig/headless`; do not reach into `src/` or
  `dist/` unless the task is to change the package itself.
- Addresses are zero-based `{ sheet, row, col }`; resolve sheet ids with
  `getSheetId()`.
- Use `exportWorkPaperDocument()` and `createWorkPaperFromDocument()` for
  persistence round trips.
- Add tests before changing config rebuilds, range bounds, formulas,
  persistence, or structural edits.
- Run focused headless tests first, then `pnpm publish:runtime:check`,
  `pnpm workpaper:bench:competitive:check`, and `pnpm run ci` before publishing
  or claiming production readiness.

For a concise public evaluation path, share
[`docs/public-adoption-kit.md`](docs/public-adoption-kit.md). It includes
positioning, npm-only evaluation commands, proof links, shareable copy, and
guardrails for honest benchmark claims.

For the first technical adoption article, see
[`docs/why-agents-need-workbook-apis.md`](docs/why-agents-need-workbook-apis.md).
It explains why agents should operate on workbook APIs instead of spreadsheet
screenshots.

For the researched public-growth plan toward `1000` legitimate GitHub stars,
see [`docs/github-stars-growth-plan.md`](docs/github-stars-growth-plan.md).

## Quickstart

Use Node `24+`, Bun, and `pnpm@10.32.1`.

```bash
pnpm install
pnpm wasm:build
pnpm typecheck
pnpm test
pnpm dev
```

The default dev command runs the local web shell and monolith together.

Useful alternatives:

```bash
pnpm dev:web
pnpm dev:web-local
pnpm dev:sync
```

## Local Docker Compose

```bash
docker compose up --build
```

This brings up:

- `http://localhost:3000` for the monolith web shell with `/v2`,
  `/api/zero/v2`, and `/zero`
- `http://localhost:4321/healthz` for the monolith app runtime
- `http://localhost:4848/keepalive` for Zero cache
- `postgresql://bilig:bilig@localhost:5432/bilig` for Postgres

To reset local state:

```bash
docker compose down -v
```

## Package Map

| Path                        | Role                                                                          |
| --------------------------- | ----------------------------------------------------------------------------- |
| `apps/web`                  | Vite/React browser source compiled into the monolith                          |
| `apps/bilig`                | Fullstack monolith runtime, API surface, and static asset server              |
| `packages/protocol`         | Shared enums, opcodes, constants, and protocol types                          |
| `packages/formula`          | A1 addressing, lexer, parser, binder, compiler, JS evaluator                  |
| `packages/core`             | Workbook engine, scheduler, snapshots, selectors, sync ownership, WASM facade |
| `packages/headless`         | Headless WorkPaper runtime surfaces                                           |
| `packages/zero-sync`        | Zero schema, workbook queries, projection, and event payload helpers          |
| `packages/binary-protocol`  | Wire format for sync frames                                                   |
| `packages/agent-api`        | Agent request, response, event, and framing model                             |
| `packages/worker-transport` | Engine host/client bridge for worker execution                                |
| `packages/renderer`         | Custom workbook reconciler and workbook DSL                                   |
| `packages/grid`             | Reusable React spreadsheet UI components and hooks                            |
| `packages/wasm-kernel`      | AssemblyScript/WASM compute fast path                                         |
| `packages/storage-browser`  | Browser-side persistence                                                      |
| `packages/storage-server`   | Server-side storage integration points                                        |
| `packages/excel-fixtures`   | Checked-in formula parity fixtures                                            |
| `packages/benchmarks`       | Benchmark harness and performance contracts                                   |

## Verification

The repo has a strict local preflight. For small changes, run the narrowest
targeted command first; before publishing, use the full gate.

```bash
pnpm lint
pnpm typecheck
pnpm test
pnpm test:browser
pnpm bench:smoke
pnpm run ci
```

Generated sources are checked in and enforced:

```bash
pnpm protocol:check
pnpm formula-inventory:check
pnpm workspace-resolution:check
pnpm workpaper:bench:competitive:check
```

## Performance Work

The WorkPaper track is the repo's performance-leadership program. It compares
bilig's spreadsheet runtime against HyperFormula-style workloads and keeps the
important claims tied to benchmark artifacts, counters, and docs.

Current public evidence:

- `packages/benchmarks/baselines/workpaper-vs-hyperformula.json`, generated at
  `2026-05-06T14:54:57.091Z`, records WorkPaper `46/46` mean wins on
  scorecard-eligible comparable workloads: `38/38` public and `8/8` holdout.
- `docs/headless-workpaper-benchmark-evidence.md` explains what is measured,
  what is excluded, and why this is a mean-win claim rather than a blanket p95
  guarantee.

Start here:

- `docs/workpaper-engine-leadership-program.md`
- `docs/headless-workpaper-benchmark-evidence.md`
- `docs/workpaper-oracle-sota-performance-design-2026-04-21.md`
- `docs/workpaper-oracle-validated-performance-design-2026-04-26.md`
- `docs/workpaper-oracle-benchmark-expansion-performance-plan-2026-04-28.md`

Run the competitive benchmark with:

```bash
pnpm bench:workpaper:competitive
```

## Architecture Docs

Good entry points:

- `docs/architecture.md`
- `docs/public-api.md`
- `docs/formula-language.md`
- `docs/agent-api.md`
- `docs/local-first-realtime-loop.md`
- `docs/binary-protocol.md`
- `docs/wasm-runtime-contract.md`
- `docs/testing-and-benchmarks.md`

## Contributing

Read [CONTRIBUTING.md](CONTRIBUTING.md) before opening a PR. The highest-value
contributions are usually:

- formula parity fixtures and semantic tests
- WorkPaper benchmark scenarios with clear expected behavior
- focused engine correctness fixes
- grid accessibility and keyboard-behavior improvements
- docs that turn existing architecture notes into runnable examples

The shortest public on-ramp is the
[`starter issues`](docs/starter-issues.md) queue. Current starter issues are
scoped around WorkPaper recipes, benchmark
walkthroughs, and agent/tooling docs.

Please keep changes small, tested, and tied to the package that owns the
behavior.

## CI

Forgejo Actions is the primary CI surface for this repo via
`.forgejo/workflows/forgejo-ci.yml`. GitHub Actions mirrors the verification
contract in `.github/workflows/ci.yml`.

The strict gate includes frozen lockfile install, full `pnpm run ci`, artifact
budget checks, browser smoke, and tracked-file cleanliness checks.

## License

MIT.
