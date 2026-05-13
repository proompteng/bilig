# bilig

[![CI](https://github.com/proompteng/bilig/actions/workflows/ci.yml/badge.svg)](https://github.com/proompteng/bilig/actions/workflows/ci.yml)
[![GitHub Repo stars](https://img.shields.io/github/stars/proompteng/bilig?style=social)](https://github.com/proompteng/bilig/stargazers)
[![npm: @bilig/headless](https://img.shields.io/npm/v/@bilig/headless?label=%40bilig%2Fheadless)](https://www.npmjs.com/package/@bilig/headless)
[![npm weekly downloads](https://img.shields.io/npm/dw/@bilig/headless?label=npm%20downloads)](https://www.npmjs.com/package/@bilig/headless)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

<p align="center">
  <img src="docs/assets/github-social-preview.png" alt="bilig headless spreadsheet engine for Node.js programs" />
</p>

**bilig runs spreadsheet formulas from Node.js.** Build workbooks from arrays or
JSON records, edit cells, recalculate formulas, and save the document without
opening a spreadsheet UI.

Project site: <https://proompteng.github.io/bilig/>

## Current Public Proof

- Live growth snapshot:
  <https://proompteng.github.io/bilig/community-growth-snapshot.html>
- Latest checked-in snapshot: `24` GitHub stars, `13,427` npm downloads in the
  last week, `32` open `good first issue` tickets, `7` GitHub Discussions, and
  `393` recent repository views.
- Benchmark evidence:
  [`46/46` comparable WorkPaper mean wins](docs/what-workpaper-benchmark-proves.md),
  with the p95 caveat documented instead of hidden.

If the 90-second check below saves you a workbook-automation spike, star or
bookmark the repo after you see the verification output:
<https://github.com/proompteng/bilig/stargazers>.

## Choose Your Path

- **Evaluate in 90 seconds**: run the npm-only
  [`@bilig/headless` quickstart](#try-biligheadless-in-90-seconds).
- **Decide if it fits your backend**: read the
  [Node.js spreadsheet formula engine guide](docs/node-spreadsheet-formula-engine.md).
- **Compare spreadsheet engines**: read the
  [SheetJS and ExcelJS comparison](docs/sheetjs-exceljs-alternative-formula-workbook-api.md),
  [HyperFormula comparison](docs/hyperformula-alternative-headless-workpaper.md),
  [engine comparison](docs/headless-spreadsheet-engine-comparison.md), and
  [benchmark explainer](docs/what-workpaper-benchmark-proves.md).
- **Build a Node workflow**: start from the
  [five runnable workbook automation examples](docs/workbook-automation-examples-node.md),
  [runnable WorkPaper example](examples/headless-workpaper),
  [JSON records input example](examples/headless-workpaper#json-records-input),
  [invoice totals example](examples/headless-workpaper#invoice-totals),
  [budget variance example](examples/headless-workpaper#budget-variance-alerts),
  [fulfillment capacity example](examples/headless-workpaper#fulfillment-capacity-plan),
  [quote approval example](examples/headless-workpaper#quote-approval-threshold),
  [subscription MRR example](examples/headless-workpaper#subscription-mrr-forecast),
  [serverless API route example](examples/serverless-workpaper-api),
  [Node service recipe](docs/node-service-workpaper-recipe.md), or
  [serverless route walkthrough](docs/serverless-workpaper-api-route.md).
- **Wire a coding-agent tool**: use the
  [Vercel AI SDK / LangChain spreadsheet tool guide](docs/vercel-ai-sdk-langchain-spreadsheet-tool.md),
  [MCP spreadsheet tool server guide](docs/mcp-workpaper-tool-server.md),
  [agent tool-calling recipe](docs/agent-workpaper-tool-calling-recipe.md), or
  run the
  [Vercel AI SDK / LangChain adapter example](examples/headless-workpaper#agent-framework-adapters)
  and [MCP tool server example](examples/headless-workpaper#mcp-tool-server-shape).
- **Contribute a small patch**: pick a scoped
  [`good first issue`](docs/starter-issues.md).
- **Ask a question or share a workflow**: use
  [GitHub Discussions](https://github.com/proompteng/bilig/discussions) for
  Q&A, ideas, and show-and-tell posts.
- **Report an Excel compatibility gap**: use the issue templates and link the
  smallest workbook, formula, or fixture that reproduces the mismatch.
- **Follow the project**: star the repo as a bookmark:
  <https://github.com/proompteng/bilig/stargazers>.

Contributor and security docs:
[`CONTRIBUTING.md`](CONTRIBUTING.md) and [`SECURITY.md`](SECURITY.md).

## Try `@bilig/headless` in 90 seconds

The fastest evaluation path uses the published npm package only. It builds a
formula-backed workbook, applies an edit, persists the document, restores it,
and throws if formula readback does not survive the round trip.

```bash
mkdir bilig-headless-eval
cd bilig-headless-eval
npm init -y
npm pkg set type=module
npm install @bilig/headless
```

Create `eval.mjs`:

```js
import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from '@bilig/headless'

const workbook = WorkPaper.buildFromSheets({
  Revenue: [
    ['Region', 'Customers', 'ARPA', 'Revenue'],
    ['West', 20, 1200, '=B2*C2'],
    ['East', 30, 250, '=B3*C3'],
    ['Central', 18, 300, '=B4*C4'],
  ],
  Summary: [
    ['Metric', 'Value'],
    ['Total revenue', '=SUM(Revenue!D2:D4)'],
  ],
})

const numberValue = (cell) => {
  if (typeof cell === 'object' && cell !== null && typeof cell.value === 'number') {
    return cell.value
  }
  throw new Error(`expected numeric cell value, got ${JSON.stringify(cell)}`)
}

const revenue = workbook.getSheetId('Revenue')
const summary = workbook.getSheetId('Summary')
if (revenue === undefined || summary === undefined) {
  throw new Error('workbook sheets were not created')
}

const before = numberValue(workbook.getCellValue({ sheet: summary, row: 1, col: 1 }))
workbook.setCellContents({ sheet: revenue, row: 1, col: 1 }, 32)

const saved = serializeWorkPaperDocument(exportWorkPaperDocument(workbook, { includeConfig: true }))
const restored = createWorkPaperFromDocument(parseWorkPaperDocument(saved))
const restoredSummary = restored.getSheetId('Summary')
if (restoredSummary === undefined) {
  throw new Error('summary sheet was not restored')
}

const after = numberValue(restored.getCellValue({ sheet: restoredSummary, row: 1, col: 1 }))
const verified = before === 36900 && after === 51300 && saved.length > 0
if (!verified) {
  throw new Error(`unexpected formula readback: ${JSON.stringify({ before, after, bytes: saved.length })}`)
}

console.log({ before, after, sheets: restored.getSheetNames(), bytes: saved.length, verified })
```

Run it:

```bash
node eval.mjs
```

Expected output:

```json
{
  "before": 36900,
  "after": 51300,
  "sheets": ["Revenue", "Summary"],
  "bytes": 1064,
  "verified": true
}
```

The maintained repository example adds agent-style writeback verification:

```bash
git clone https://github.com/proompteng/bilig.git
cd bilig/examples/headless-workpaper
npm install
npm start
npm run agent:tool-call
npm run agent:framework-adapters
npm run agent:verify
```

Expected proof from `npm run agent:tool-call` includes:

```json
{
  "toolCall": { "toolName": "setInputCell" },
  "toolResult": {
    "editedCell": "Inputs!B3",
    "before": { "expectedArr": 60000, "targetGap": -34000 },
    "after": { "expectedArr": 96000, "targetGap": 5600 },
    "verified": {
      "formulasPersisted": true,
      "restoredMatchesAfter": true,
      "expectedArrImproved": true,
      "targetGapClosed": true
    }
  }
}
```

Expected proof from `npm run agent:verify` includes:

```json
{
  "after": {
    "customers": 65,
    "grossMrr": 15600,
    "expansionMrr": 18720,
    "annualizedArr": 224640,
    "arrTargetDelta": 74640
  },
  "verified": {
    "formulasUnchanged": true,
    "formulasPersisted": true,
    "restoredMatchesAfter": true
  }
}
```

The serverless route example gives HTTP and agent-tool evaluators a runnable
JSON boundary:

```bash
git clone https://github.com/proompteng/bilig.git
cd bilig/examples/serverless-workpaper-api
npm install
npm run smoke
```

Expected proof from `npm run smoke` includes:

```json
{
  "edit": {
    "records": 4,
    "after": {
      "totalRevenue": 48600,
      "westCustomers": 20,
      "largestDeal": 24000
    },
    "checks": {
      "totalRevenueChanged": true,
      "formulasPersisted": true
    }
  },
  "verified": true
}
```

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
For the shortest public method map, start with the
[WorkPaper read/write cheat sheet](packages/headless/README.md#workpaper-readwrite-cheat-sheet).

Install from npm:

```bash
pnpm add @bilig/headless
```

Try the package without cloning the monorepo:

```bash
mkdir bilig-headless-eval
cd bilig-headless-eval
npm init -y
npm pkg set type=module
npm install @bilig/headless
```

Create `eval.mjs` with the quickstart below, then run `node eval.mjs`. The
example builds a formula-backed workbook, edits source data, serializes the
document, restores it, and verifies that the recalculated value survives the
round trip.

For a runnable external-consumer example, start with
[examples/headless-workpaper](examples/headless-workpaper). The repository smoke
test executes that same example against packed local runtime packages with
`pnpm workpaper:smoke:external`.

For backend adoption, see
[`docs/node-service-workpaper-recipe.md`](docs/node-service-workpaper-recipe.md).
It shows a minimal Node service boundary that reads computed summaries, applies
one controlled edit, and persists the WorkPaper document.

For tabular service payloads, see
[`docs/csv-shaped-workpaper-input-recipe.md`](docs/csv-shaped-workpaper-input-recipe.md).
It normalizes a small CSV-shaped fixture into a WorkPaper workbook and reads
formula-backed summary values.

For JSON service/API payloads, the runnable example includes
[`npm run json-records`](examples/headless-workpaper#json-records-input). It
maps an array of opportunity records into `WorkPaper.buildFromSheets()`, adds
formula-backed summary cells, and validates exact computed output before
printing JSON.

For billing-style service payloads, the runnable example includes
[`npm run invoice-totals`](examples/headless-workpaper#invoice-totals). It
calculates line-item totals, subtotal, tax, and grand total formulas, then
validates exact computed and serialized formula readback.

For reporting and finance automation, the runnable example includes
[`npm run budget-variance`](examples/headless-workpaper#budget-variance-alerts).
It compares budget and actual rows, calculates dollar and percent variance, and
flags rows that need review with a formula-backed alert.

For subscription revenue forecasting, the runnable example includes
[`npm run subscription-mrr`](examples/headless-workpaper#subscription-mrr-forecast).
It models starting customers, churn, expansion, and new customers, then prints
starting MRR, ending MRR, net expansion MRR, and verified formula readback.

For sales-ops quote workflows, the runnable example includes
[`npm run quote-approval`](examples/headless-workpaper#quote-approval-threshold).
It calculates list total, discount amount, quote total, max line discount, and
an approval flag from formula-backed quote rows.

For operations planning, the runnable example includes
[`npm run fulfillment-capacity`](examples/headless-workpaper#fulfillment-capacity-plan).
It compares forecast order volume with available labor hours, calculates
required hours, capacity gap, short days, and a formula-backed status.

For formula errors, see
[`docs/unsupported-formula-troubleshooting-recipe.md`](docs/unsupported-formula-troubleshooting-recipe.md).
It shows how to read `#VALUE!`/`#NAME?` display text together with structured
diagnostics so services and agents can reject or normalize unsupported inputs.

That example also includes `npm run agent:verify`, a small agent writeback demo
that records the exact assumption cells changed, verifies dependent formula
readback, persists the workbook, restores it, and proves the formulas and values
survived the round trip.

For a tool-calling shape closer to agent SDKs, run `npm run agent:tool-call`.
It returns a compact tool call, before/after computed values, formula
contracts, persistence proof, and round-trip verification.

For Vercel AI SDK and LangChain-shaped wrappers, run
`npm run agent:framework-adapters`. The example keeps the same validated
WorkPaper read/write functions and exposes thin framework adapter shapes
without adding either framework as a dependency.

For an MCP-style shape, run `npm run agent:mcp-tools`. It returns a
dependency-free `tools/list` response, a `tools/call` read, and a verified
input edit with structured computed readback.

Quickstart:

```js
import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from '@bilig/headless'

const workbook = WorkPaper.buildFromSheets(
  {
    Revenue: [
      ['Region', 'Customers', 'ARPA', 'Revenue'],
      ['West', 20, 1200, '=B2*C2'],
      ['East', 30, 250, '=B3*C3'],
      ['Central', 18, 300, '=B4*C4'],
    ],
    Summary: [
      ['Metric', 'Value'],
      ['Total revenue', '=SUM(Revenue!D2:D4)'],
    ],
  },
  { maxRows: 1_000, maxColumns: 100, useColumnIndex: true },
)

const revenue = workbook.getSheetId('Revenue')
const summary = workbook.getSheetId('Summary')
if (revenue === undefined || summary === undefined) {
  throw new Error('Workbook sheets were not created')
}

const at = (row, col) => ({
  sheet: summary,
  row,
  col,
})

const before = workbook.getCellValue(at(1, 1))
workbook.setCellContents({ sheet: revenue, row: 1, col: 1 }, 32)

const saved = serializeWorkPaperDocument(exportWorkPaperDocument(workbook, { includeConfig: true }))
const restored = createWorkPaperFromDocument(parseWorkPaperDocument(saved))
const restoredSummary = restored.getSheetId('Summary')
if (restoredSummary === undefined) {
  throw new Error('Summary sheet was not restored')
}

const after = restored.getCellValue({
  sheet: restoredSummary,
  row: 1,
  col: 1,
})

console.log({ before, after, sheets: restored.getSheetNames(), bytes: saved.length })
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

For workflow feedback from people building Node services or agent tools, use
the GitHub discussion:
<https://github.com/proompteng/bilig/discussions/157>.

For the five runnable service-workflow examples as one shareable thread, see:
<https://github.com/proompteng/bilig/discussions/213>.

For the search-friendly version with commands and current output snippets, see
[`docs/workbook-automation-examples-node.md`](docs/workbook-automation-examples-node.md).

For the first technical adoption article, see
[`docs/why-agents-need-workbook-apis.md`](docs/why-agents-need-workbook-apis.md).
It explains why agents should operate on workbook APIs instead of spreadsheet
screenshots.

For a concrete framework-neutral agent tool loop, see
[`docs/agent-workpaper-tool-calling-recipe.md`](docs/agent-workpaper-tool-calling-recipe.md).
It wraps WorkPaper reads, validated writes, computed before/after checks, and
persistence into a small tool surface.

For the persistence-focused follow-up article and runnable example, see
[`docs/persisting-formula-backed-workpaper-documents-in-node.md`](docs/persisting-formula-backed-workpaper-documents-in-node.md)
and
[`examples/headless-workpaper/persistence-roundtrip.mjs`](examples/headless-workpaper/persistence-roundtrip.mjs).

For the benchmark-focused explainer, see
[`docs/what-workpaper-benchmark-proves.md`](docs/what-workpaper-benchmark-proves.md).
It states the `46/46` mean-win claim and the known p95 caveat without turning
the benchmark into a blanket performance claim.

For a local benchmark command walkthrough, see
[`docs/local-workpaper-benchmark-walkthrough.md`](docs/local-workpaper-benchmark-walkthrough.md).
It shows how to verify the checked-in artifact, run a reduced local smoke
benchmark, and compare benchmark diffs.

For a concise HyperFormula comparison and evaluation path, see
[`docs/hyperformula-alternative-headless-workpaper.md`](docs/hyperformula-alternative-headless-workpaper.md).

For a broader headless spreadsheet-engine comparison across `@bilig/headless`,
HyperFormula, IronCalc, ExcelJS, Formula.js, Hucre, Formualizer, and
JSpreadsheet Formula Pro, see
[`docs/headless-spreadsheet-engine-comparison.md`](docs/headless-spreadsheet-engine-comparison.md).

For a runnable revenue-model walkthrough, see
[`docs/building-a-revenue-model-with-headless-workpaper.md`](docs/building-a-revenue-model-with-headless-workpaper.md)
and
[`examples/headless-workpaper/revenue-scenarios.mjs`](examples/headless-workpaper/revenue-scenarios.mjs).

For the current Excel-compatibility boundaries, see
[`docs/where-bilig-is-not-excel-compatible-yet.md`](docs/where-bilig-is-not-excel-compatible-yet.md).
It names the macro, formula, XLSX corpus, and UI-claim gaps without treating
the project as a complete Excel clone.

For a short guide to interpreting XLSX cached-result corpus reports, see
[`docs/xlsx-corpus-verifier-walkthrough.md`](docs/xlsx-corpus-verifier-walkthrough.md).

For a formula-edge fixture walkthrough covering the exact-match `XLOOKUP` path,
see
[`docs/formula-edge-xlookup-exact-fixture.md`](docs/formula-edge-xlookup-exact-fixture.md).

For a formula-edge fixture walkthrough covering paired-criteria `SUMIFS`, see
[`docs/formula-edge-sumifs-paired-criteria-fixture.md`](docs/formula-edge-sumifs-paired-criteria-fixture.md).

For a formula-edge fixture walkthrough covering grouped dynamic-array `GROUPBY`
output, see
[`docs/formula-edge-groupby-spill-fixture.md`](docs/formula-edge-groupby-spill-fixture.md).

For the published DEV article, see
<https://dev.to/gregkonush/why-agents-need-workbook-apis-instead-of-spreadsheet-screenshots-3d61>.
The source mirror with front matter is
[`docs/dev-to-workbook-apis-post.md`](docs/dev-to-workbook-apis-post.md).

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

<p align="center">
  <img src="docs/assets/workpaper-benchmark-card.png" alt="WorkPaper benchmark evidence card with 46 out of 46 comparable mean wins and the visible p95 caveat" />
</p>

Current public evidence:

- `packages/benchmarks/baselines/workpaper-vs-hyperformula.json`, generated at
  `2026-05-08T15:00:27.603Z`, records WorkPaper `46/46` mean wins on
  scorecard-eligible comparable workloads: `38/38` public and `8/8` holdout.
- `docs/assets/workpaper-benchmark-card.png` is the shareable chart for the
  current scorecard. It is generated from the checked-in artifact with
  `pnpm docs:benchmark-card:generate`.
- `docs/headless-workpaper-benchmark-evidence.md` explains what is measured,
  what is excluded, and why this is a mean-win claim rather than a blanket p95
  guarantee.

Start here:

- `docs/workpaper-engine-leadership-program.md`
- `docs/headless-workpaper-benchmark-evidence.md`
- `docs/what-workpaper-benchmark-proves.md`
- `docs/local-workpaper-benchmark-walkthrough.md`
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

Read [CONTRIBUTING.md](CONTRIBUTING.md) before opening a PR. If this is your
first patch, start with the
[new contributor guide](docs/new-contributor-guide.md) and then claim a scoped
starter issue. The highest-value contributions are usually:

- formula parity fixtures and semantic tests
- WorkPaper benchmark scenarios with clear expected behavior
- focused engine correctness fixes
- grid accessibility and keyboard-behavior improvements
- docs that turn existing architecture notes into runnable examples

The shortest public on-ramp is the
[`starter issues`](docs/starter-issues.md) queue. Current starter issues are
scoped around small runnable examples with explicit acceptance commands, so a
first contribution can improve the public WorkPaper evaluation path without
understanding the whole engine.
If this is your first contribution to `bilig`, start with the
[`first-timers-only`](https://github.com/proompteng/bilig/issues?q=is%3Aissue%20state%3Aopen%20label%3Afirst-timers-only)
filter.

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
