# bilig

[![CI](https://github.com/proompteng/bilig/actions/workflows/ci.yml/badge.svg)](https://github.com/proompteng/bilig/actions/workflows/ci.yml)
[![GitHub Repo stars](https://img.shields.io/github/stars/proompteng/bilig?style=social)](https://github.com/proompteng/bilig/stargazers)
[![npm: @bilig/headless](https://img.shields.io/npm/v/@bilig/headless?label=%40bilig%2Fheadless)](https://www.npmjs.com/package/@bilig/headless)
[![npm weekly downloads](https://img.shields.io/npm/dw/@bilig/headless?label=npm%20downloads)](https://www.npmjs.com/package/@bilig/headless)
[![MCP server score](https://glama.ai/mcp/servers/proompteng/bilig/badges/score.svg)](https://glama.ai/mcp/servers/proompteng/bilig)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

<p align="center">
  <img src="docs/assets/bilig-hero-workbook-api.png" alt="bilig workbook grid showing an edited input cell and recalculated formula result" />
</p>

**bilig gives TypeScript code a workbook it can edit, calculate, and save.**

Use it when the clearest model is still a spreadsheet, but the runtime is a
Node service, queue worker, serverless route, test, or coding-agent tool. The
main package is [`@bilig/headless`](https://www.npmjs.com/package/@bilig/headless):
it gives you a `WorkPaper` object, not a browser grid.

The loop is small: create sheets, write cells, recalculate formulas, read the
value that came out, and persist the workbook as JSON. That is useful for
pricing rules, budget checks, payout models, import validation, and agent tools
that need real readback after an edit.

It is not trying to replace Excel's UI. It is the headless runtime you put
behind product code when formulas are the clearest way to explain the logic.

Project site: <https://proompteng.github.io/bilig/>

## Try It In 90 Seconds

This uses the published npm package. It builds a workbook, changes one input,
restores the saved document, and fails if the calculated result does not survive
the round trip.

```sh
mkdir bilig-headless-eval
cd bilig-headless-eval
npm init -y
npm pkg set type=module
npm install @bilig/headless
npm install -D tsx typescript @types/node
curl -fsSLo eval.ts https://proompteng.github.io/bilig/npm-eval.ts
npx tsx eval.ts
```

Expected output:

```json
{
  "before": 24000,
  "after": 38400,
  "afterRestore": 38400,
  "sheets": ["Inputs", "Summary"],
  "bytes": 1000,
  "verified": true
}
```

The TypeScript file is maintained in
[`examples/headless-workpaper/npm-eval.ts`](examples/headless-workpaper/npm-eval.ts).
The exact byte count can change between package versions; `verified: true` and
matching `after`/`afterRestore` values are the check.

## TypeScript API Shape

Most integrations are just this: build a workbook, write an input, read the
calculated value, and save the workbook state.

```ts
import {
  WorkPaper,
  exportWorkPaperDocument,
  serializeWorkPaperDocument,
} from '@bilig/headless'

const workbook = WorkPaper.buildFromSheets({
  Inputs: [
    ['Metric', 'Value'],
    ['Customers', 20],
    ['Average revenue', 1200],
  ],
  Summary: [
    ['Metric', 'Value'],
    ['Revenue', '=Inputs!B2*Inputs!B3'],
  ],
})

const inputs = workbook.getSheetId('Inputs')
const summary = workbook.getSheetId('Summary')
if (inputs === undefined || summary === undefined) {
  throw new Error('Workbook is missing required sheets')
}

workbook.setCellContents({ sheet: inputs, row: 1, col: 1 }, 32)

const revenue = workbook.getCellDisplayValue({ sheet: summary, row: 1, col: 1 })
const saved = serializeWorkPaperDocument(
  exportWorkPaperDocument(workbook, { includeConfig: true }),
)

console.log({ revenue, savedBytes: saved.length })
```

## When To Reach For It

Use `@bilig/headless` when:

- a Node service owns a workbook-shaped calculation;
- an agent needs tools such as `readRange` and `setInputCell`, with computed
  before/after values instead of screenshots;
- tests need deterministic spreadsheet state and formula readback;
- a workflow needs to save the edited workbook as JSON and restore it later.

Use something else when you need a visual spreadsheet grid, Office macros,
desktop Excel automation, or a one-off arithmetic helper. Do not treat embedded
XLSX cached formula values as truth; use the Excel oracle workflow when accuracy
matters.

## Start Here

Pick the path closest to what you are building.

- If you are evaluating the npm package:
  [90-second npm eval](#try-it-in-90-seconds) and
  [npm-only smoke test](docs/try-bilig-headless-in-node.md).
- If you are writing code against the API:
  [packages/headless/README.md](packages/headless/README.md) and the
  [WorkPaper read/write cheat sheet](packages/headless/README.md#workpaper-readwrite-cheat-sheet).
- If you are putting workbook logic in a service:
  [Node service recipe](docs/node-service-workpaper-recipe.md),
  [server-side spreadsheet automation](docs/server-side-spreadsheet-automation-node.md),
  [Node spreadsheet formula engine](docs/node-spreadsheet-formula-engine.md),
  [framework adapters](docs/node-framework-workpaper-adapters.md), and
  [workbook automation examples](docs/workbook-automation-examples-node.md).
  The runnable serverless boundary lives in
  [examples/serverless-workpaper-api](examples/serverless-workpaper-api);
  run `npm run next-route-handler`, `npm run next-server-action`,
  `npm run next-server-action-formdata`, `npm run framework-adapters`, and
  `npm run persistence-adapters` from that example.
- If an agent needs workbook tools:
  [agent tool-calling recipe](docs/agent-workpaper-tool-calling-recipe.md),
  [OpenAI Responses tools](docs/openai-responses-workpaper-tool-call.md),
  [AI SDK and LangChain adapters](docs/vercel-ai-sdk-langchain-spreadsheet-tool.md),
  [agent framework adapters](examples/headless-workpaper#agent-framework-adapters),
  [MCP server guide](docs/mcp-workpaper-tool-server.md),
  [MCP tool server shape](examples/headless-workpaper#mcp-tool-server-shape),
  [MCP directory status](docs/mcp-spreadsheet-server-directory.md),
  [MCP client setup](docs/mcp-client-setup.md), and
  [Claude Desktop MCPB bundle](docs/claude-desktop-mcpb-workpaper.md).
- If you are comparing libraries:
  [evaluate Excel formulas in Node.js](docs/evaluate-excel-formulas-in-node-typescript.md),
  [Google Sheets API boundary](docs/google-sheets-api-alternative-node-workpaper.md),
  [docs/javascript-spreadsheet-library-headless-node.md](docs/javascript-spreadsheet-library-headless-node.md),
  [docs/sheetjs-exceljs-alternative-formula-workbook-api.md](docs/sheetjs-exceljs-alternative-formula-workbook-api.md),
  [headless engine comparison](docs/headless-spreadsheet-engine-comparison.md), and
  [HyperFormula notes](docs/hyperformula-alternative-headless-workpaper.md).
- If you want a first contribution:
  [starter issues](docs/starter-issues.md),
  [first-timers-only issues](https://github.com/proompteng/bilig/issues?q=is%3Aissue%20state%3Aopen%20label%3Afirst-timers-only),
  [GitHub Discussions](https://github.com/proompteng/bilig/discussions), and
  [CONTRIBUTING.md](CONTRIBUTING.md).

## Examples You Can Run

The runnable examples are TypeScript files. Some source imports end in `.js`
because Node ESM resolves compiled package output that way; the files you edit
and run are still `.ts`.

From `examples/headless-workpaper`:

```sh
npm install
npm start
npm run json-records
npm run csv-shaped
npm run invoice-totals
npm run budget-variance
npm run fulfillment-capacity
npm run quote-approval
npm run subscription-mrr
npm run persistence
```

The most useful entry points:

- [JSON records input](examples/headless-workpaper#json-records-input)
- [CSV shaped input](examples/headless-workpaper#csv-shaped-input)
- [invoice totals](examples/headless-workpaper#invoice-totals)
- [budget variance alerts](examples/headless-workpaper#budget-variance-alerts)
- [fulfillment capacity plan](examples/headless-workpaper#fulfillment-capacity-plan)
- [quote approval threshold](examples/headless-workpaper#quote-approval-threshold)
- [subscription MRR forecast](examples/headless-workpaper#subscription-mrr-forecast)

For agent tools:

```sh
npm run agent:verify
npm run agent:tool-call
npm run agent:openai-responses
npm run agent:ai-sdk-generate-text
npm run agent:ai-sdk-stream-text
npm run agent:framework-adapters
npm run agent:mcp-tools
npm run agent:mcp-stdio
```

The AI SDK example uses
[`ai-sdk-generate-text-tool-smoke.ts`](examples/headless-workpaper/ai-sdk-generate-text-tool-smoke.ts).
The OpenAI Responses guide is
[`docs/openai-responses-workpaper-tool-call.md`](docs/openai-responses-workpaper-tool-call.md).
The agent framework guide is
[`docs/vercel-ai-sdk-langchain-spreadsheet-tool.md`](docs/vercel-ai-sdk-langchain-spreadsheet-tool.md).

The package also ships the MCP stdio binary:

```sh
npm exec --package @bilig/headless -- bilig-workpaper-mcp
```

It is published in the official MCP Registry as
`io.github.proompteng/bilig-workpaper`:
<https://registry.modelcontextprotocol.io/v0.1/servers?search=io.github.proompteng%2Fbilig-workpaper>.

## Proof You Can Reproduce

- The 90-second TypeScript check above edits one input, restores the saved JSON
  document, and verifies the dependent formula result.
- Run `pnpm workpaper:bench:competitive:check`. The checked-in artifact shows
  [`46/46` comparable WorkPaper mean wins](docs/what-workpaper-benchmark-proves.md)
  and names the slower p95 row: `lookup-approximate-duplicates` at `1.043x`.
- The benchmark card is generated from that artifact:
  [`docs/assets/workpaper-benchmark-card.png`](docs/assets/workpaper-benchmark-card.png).
- Read the [compatibility limits](docs/where-bilig-is-not-excel-compatible-yet.md)
  before importing real Excel workbooks.
- For XLSX accuracy audits, use the
  [Excel oracle harness](docs/xlsx-corpus-verifier-walkthrough.md#run-the-excel-oracle-harness).
  It separates import success, timeouts, stale cached formula values, and fresh
  Microsoft Excel recalculation results.
- Track public signals in the
  [growth snapshot](https://proompteng.github.io/bilig/community-growth-snapshot.html):
  stars, npm downloads, starter issues, Discussions, traffic, and clones.
- The WorkPaper MCP server is listed in the
  [official MCP Registry](https://registry.modelcontextprotocol.io/v0.1/servers?search=io.github.proompteng%2Fbilig-workpaper)
  and on [Glama](https://glama.ai/mcp/servers/proompteng/bilig). The
  [directory status page](docs/mcp-spreadsheet-server-directory.md) keeps the
  npm command and directory evidence in one place.
- Public feedback threads:
  [workflow questions](https://github.com/proompteng/bilig/discussions/157),
  [service examples](https://github.com/proompteng/bilig/discussions/213),
  [persistence adapters](https://github.com/proompteng/bilig/discussions/307),
  [JavaScript spreadsheet library guide](https://github.com/proompteng/bilig/discussions/308),
  [OpenAI Responses tool calls](https://github.com/proompteng/bilig/discussions/335),
  and [benchmark critique](https://github.com/proompteng/bilig/discussions/340).

If the 90-second check matches a problem you have, star or bookmark the repo:
<https://github.com/proompteng/bilig/stargazers>.

## XLSX Accuracy Policy

Cached formula values embedded in `.xlsx` files are cache diagnostics, not an
accuracy verdict. A Bilig correctness bug should only be claimed when the
expected value came from a fresh Excel recalculation oracle.

```sh
OUT=.cache/excel-oracle-evaluation
pnpm workpaper:xlsx-oracle -- prepare-oracle /path/to/xlsx-corpus "$OUT"
pnpm workpaper:xlsx-oracle -- evaluate-cache /path/to/xlsx-corpus "$OUT"
pnpm workpaper:xlsx-oracle -- evaluate-oracle /path/to/xlsx-corpus "$OUT/recalculated" "$OUT"
pnpm workpaper:xlsx-oracle -- summarize "$OUT"
```

`evaluate-cache` writes `cache-diagnostic.json` and stays non-authoritative.
`evaluate-oracle` writes `excel-oracle-report.json`, and `summarize` writes
`summary.md`. If Excel automation is unavailable, cells are classified as
`missing_excel_oracle` instead of being promoted to bugs.

## What Is In This Repo

- `packages/headless`: WorkPaper runtime and npm package.
- `packages/excel-import`: XLSX import/export boundary. Install both packages
  with `pnpm add @bilig/headless @bilig/excel-import` when you need file import
  and export.
- `packages/formula`: formula parser, binder, compiler, and evaluator.
- `packages/core`: workbook engine, snapshots, mutation flow, and scheduler.
- `packages/grid` and `apps/web`: browser spreadsheet shell.
- `apps/bilig`: fullstack monolith runtime, API surface, and static asset
  server.
- `packages/renderer`: React workbook renderer.
- `packages/protocol`, `packages/binary-protocol`, `packages/agent-api`, and
  `packages/worker-transport`: protocol and integration boundaries.
- `packages/wasm-kernel`: AssemblyScript/WASM numeric fast path.
- `packages/benchmarks`: benchmark harness and performance contracts.

For XLSX import/export from TypeScript:

```ts
import { WorkPaper } from '@bilig/headless'
import { exportXlsx, importXlsx } from '@bilig/excel-import'
```

Use `WorkPaper.buildFromSnapshot(imported.snapshot)` after import and
`workbook.exportSnapshot()` before `exportXlsx()`.

## Local Development

Use Node `24+`, Bun, and `pnpm@10.32.1`.

```sh
pnpm install
pnpm dev:web
pnpm dev:web-local
pnpm dev:sync
```

For a full local preflight:

```sh
pnpm lint
pnpm typecheck
pnpm test
pnpm test:browser
pnpm run ci
```

Generated sources and public evidence are checked:

```sh
pnpm protocol:check
pnpm formula-inventory:check
pnpm workspace-resolution:check
pnpm workpaper:bench:competitive:check
pnpm docs:discovery:check
```

## For Coding Agents

Start with the public package boundary unless the task is explicitly engine
work.

1. Read `packages/headless/README.md` before touching WorkPaper behavior.
2. Use public exports from `@bilig/headless`; do not reach into `src/` or
   `dist/` when writing consumer examples.
3. Keep examples TypeScript-first.
4. Do not call stale XLSX cached formula values an accuracy oracle.
5. Add focused tests before changing formulas, persistence, range bounds,
   config rebuilds, events, row/column moves, or sheet lifecycle.
6. Run the focused package tests first, then broaden to `pnpm run ci`.

## Contributing

Read [CONTRIBUTING.md](CONTRIBUTING.md) before opening a PR. If this is your
first patch, start with the
[new contributor guide](docs/new-contributor-guide.md) and then claim a scoped
starter issue.

Good first patches usually fit one of these shapes:

- formula fixtures with clear expected behavior;
- small WorkPaper examples that prove a real service or agent workflow;
- focused correctness fixes with regression tests;
- grid accessibility and keyboard-behavior improvements;
- docs that turn an existing architecture note into a runnable command.

The shortest public on-ramp is the
[`starter issues`](docs/starter-issues.md) queue. It keeps code/test picks,
example tasks, adapters, and focused docs work in one current list, with small
acceptance commands for first patches.

If this is your first contribution to `bilig`, use the
[`first-timers-only`](https://github.com/proompteng/bilig/issues?q=is%3Aissue%20state%3Aopen%20label%3Afirst-timers-only)
filter.

## CI

Forgejo Actions is the primary CI surface via
`.forgejo/workflows/forgejo-ci.yml`. GitHub Actions mirrors the verification
contract in `.github/workflows/ci.yml`.

The strict gate includes frozen lockfile install, full `pnpm run ci`, artifact
budget checks, browser smoke, and tracked-file cleanliness checks.

## License

MIT.
