# @bilig/headless

[![npm: @bilig/headless](https://img.shields.io/npm/v/@bilig/headless?label=%40bilig%2Fheadless)](https://www.npmjs.com/package/@bilig/headless)
[![npm weekly downloads](https://img.shields.io/npm/dw/@bilig/headless?label=npm%20downloads)](https://www.npmjs.com/package/@bilig/headless)
[![GitHub](https://img.shields.io/badge/GitHub-proompteng%2Fbilig-blue)](https://github.com/proompteng/bilig)
[![GitHub Repo stars](https://img.shields.io/github/stars/proompteng/bilig?style=social)](https://github.com/proompteng/bilig/stargazers)
[![MCP server score](https://glama.ai/mcp/servers/proompteng/bilig/badges/score.svg)](https://glama.ai/mcp/servers/proompteng/bilig)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](https://github.com/proompteng/bilig/blob/main/LICENSE)

`@bilig/headless` is a headless workbook runtime for TypeScript. It lets Node
services and agent tools build workbooks, write formulas, read calculated cells,
and save or restore the workbook as JSON.

Use it when spreadsheet logic belongs inside a service, test suite, local-first
app, queue worker, or coding-agent tool. The package gives code a workbook
boundary with stable sheet names, cell addresses, formulas, and persisted state.

It is not Excel desktop automation and it is not a visual grid. XLSX import and
export live in the repository import/export packages; this package executes the
validated WorkPaper model once data is in workbook form.

## Best Fit

| Use it for                                                             | Do not use it for                                           |
| ---------------------------------------------------------------------- | ----------------------------------------------------------- |
| Formula-backed calculations in Node services and serverless routes     | Manual spreadsheet editing                                  |
| Agent tools that write workbook inputs and return verified readback    | Browser grid rendering by itself                            |
| Persisting a workbook document after code changes cells or formulas    | Office macros, COM automation, or desktop Excel integration |
| Tests that need deterministic spreadsheet state instead of screenshots | One-off arithmetic where a workbook model adds no value     |

## Current Public Proof

- Live growth snapshot:
  <https://proompteng.github.io/bilig/community-growth-snapshot.html>
- Latest checked-in snapshot: `24` GitHub stars, `12` forks, `15,592` npm downloads in the
  last week, `23,240` npm downloads in the last 30 days, `88` open
  `good first issue` tickets, `10` GitHub Discussions, and `455` recent
  repository views.
- Benchmark evidence:
  [`46/46` comparable WorkPaper mean wins](https://github.com/proompteng/bilig/blob/main/docs/what-workpaper-benchmark-proves.md),
  with the p95 caveat documented instead of hidden.
- MCP discovery: listed in the
  [official MCP Registry](https://registry.modelcontextprotocol.io/v0.1/servers?search=io.github.proompteng%2Fbilig-workpaper)
  and on [Glama](https://glama.ai/mcp/servers/proompteng/bilig); the
  [directory status page](https://proompteng.github.io/bilig/mcp-spreadsheet-server-directory.html)
  tracks the npm command and pending directory reviews without claiming listings
  early.

If the sanity check below saves you a workbook automation spike, star the repo
so you can find it again:
<https://github.com/proompteng/bilig/stargazers>.

## Install

Requires Node `24+` and ESM imports.

```sh
npm install @bilig/headless
```

For a copy-paste evaluation from an empty directory, use the
[npm-only smoke test](https://proompteng.github.io/bilig/try-bilig-headless-in-node.html)
or the
[TypeScript guide for evaluating Excel formulas in Node.js](https://proompteng.github.io/bilig/evaluate-excel-formulas-in-node-typescript.html).

## Clean npm Sanity Check

Run this from an empty directory when you want to verify the published package
before cloning the repository. It builds a workbook, changes an input, saves the
document, restores it, and checks that the dependent formula still reads back
correctly.

```sh
mkdir bilig-headless-sanity
cd bilig-headless-sanity
npm init -y
npm pkg set type=module
npm install @bilig/headless
npm install -D tsx typescript @types/node
cat > sanity.ts <<'EOF'
import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from '@bilig/headless';

type NumericCell = {
  value: number;
};

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
});

const numberValue = (cell: unknown): number => {
  if (typeof cell === 'object' && cell !== null && typeof (cell as NumericCell).value === 'number') {
    return (cell as NumericCell).value;
  }
  throw new Error(`Expected numeric cell value, got ${JSON.stringify(cell)}`);
};

const revenue = workbook.getSheetId('Revenue');
const summary = workbook.getSheetId('Summary');
if (revenue === undefined || summary === undefined) {
  throw new Error('Workbook sheets were not created');
}

const before = numberValue(workbook.getCellValue({ sheet: summary, row: 1, col: 1 }));
workbook.setCellContents({ sheet: revenue, row: 1, col: 1 }, 32);

const saved = serializeWorkPaperDocument(exportWorkPaperDocument(workbook, { includeConfig: true }));
const restored = createWorkPaperFromDocument(parseWorkPaperDocument(saved));
const restoredSummary = restored.getSheetId('Summary');
if (restoredSummary === undefined) {
  throw new Error('Summary sheet was not restored');
}

const after = numberValue(restored.getCellValue({ sheet: restoredSummary, row: 1, col: 1 }));
const verified = before === 36900 && after === 51300 && saved.length > 0;
if (!verified) {
  throw new Error(`Unexpected formula readback: ${JSON.stringify({ before, after, bytes: saved.length })}`);
}

console.log({ before, after, sheets: restored.getSheetNames(), bytes: saved.length, verified });
EOF
npx tsx sanity.ts
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

Inside this monorepo:

```sh
pnpm install
pnpm --filter @bilig/headless build
```

## Start Here

| Job                          | Start with                                                                                                                                                                                                                                                                                                                                                                                                                                                           |
| ---------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Prove the package from npm   | [clean npm sanity check](#clean-npm-sanity-check)                                                                                                                                                                                                                                                                                                                                                                                                                    |
| Decide if the backend fits   | [evaluate Excel formulas in Node.js](https://proompteng.github.io/bilig/evaluate-excel-formulas-in-node-typescript.html), [Node.js spreadsheet formula engine guide](https://proompteng.github.io/bilig/node-spreadsheet-formula-engine.html), and [compatibility boundaries](https://github.com/proompteng/bilig/blob/main/docs/where-bilig-is-not-excel-compatible-yet.md)                                                                                         |
| Test formula edits and state | [quickstart](#quickstart) and [persistence round trip](https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#persistence-round-trip)                                                                                                                                                                                                                                                                                                             |
| Build from service data      | [`npm run json-records`](https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#json-records-input), [workbook-automation-examples-node guide](https://proompteng.github.io/bilig/workbook-automation-examples-node.html), [Node service recipe](https://github.com/proompteng/bilig/blob/main/docs/node-service-workpaper-recipe.md), [Express/Fastify/Hono adapters](https://proompteng.github.io/bilig/node-framework-workpaper-adapters.html) |
| Wrap agent tools             | [`npm run agent:tool-call`](https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#agent-tool-call-loop), [framework adapter guide](https://proompteng.github.io/bilig/vercel-ai-sdk-langchain-spreadsheet-tool.html)                                                                                                                                                                                                                             |
| Expose MCP tools             | [`npm run agent:mcp-tools`](https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#mcp-tool-server-shape), [MCP server guide](https://proompteng.github.io/bilig/mcp-workpaper-tool-server.html), [MCP client setup](https://proompteng.github.io/bilig/mcp-client-setup.html), [Claude Desktop MCPB bundle](https://proompteng.github.io/bilig/claude-desktop-mcpb-workpaper.html)                                                               |
| Compare engines              | [JavaScript spreadsheet library guide](https://github.com/proompteng/bilig/blob/main/docs/javascript-spreadsheet-library-headless-node.md), [headless engine comparison](https://github.com/proompteng/bilig/blob/main/docs/headless-spreadsheet-engine-comparison.md), [HyperFormula comparison](https://github.com/proompteng/bilig/blob/main/docs/hyperformula-alternative-headless-workpaper.md)                                                                 |

Run the packaged MCP stdio server with
`npm exec --package @bilig/headless -- bilig-workpaper-mcp`. The published MCP
Registry entry is
[`io.github.proompteng/bilig-workpaper`](https://registry.modelcontextprotocol.io/v0.1/servers?search=io.github.proompteng%2Fbilig-workpaper).
Use the
[MCP client setup guide](https://proompteng.github.io/bilig/mcp-client-setup.html)
for Claude, Cursor, VS Code, and Codex config snippets.
For Claude Desktop bundle installs, build
`build/mcpb/bilig-workpaper.mcpb` with `pnpm mcpb:workpaper:build`; the
[MCPB guide](https://proompteng.github.io/bilig/claude-desktop-mcpb-workpaper.html)
documents the manifest and verification prompt.

Star or bookmark the project: <https://github.com/proompteng/bilig>.

## Which Example Should I Run?

The full example catalog lives in
[`examples/headless-workpaper/README.md`](https://github.com/proompteng/bilig/blob/main/examples/headless-workpaper/README.md).

| Need                                 | Start with                         | Existing example                                                                                                                       |
| ------------------------------------ | ---------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------- |
| Evaluate formulas from API records   | `npm run json-records`             | [JSON records input](https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#json-records-input)                     |
| Load records already saved on disk   | `npm run json-file`                | [JSON file input](https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#json-file-input)                           |
| Persist and restore a workbook       | `npm run persistence`              | [Persistence round trip](https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#persistence-round-trip)             |
| Verify an agent writeback            | `npm run agent:verify`             | [Agent writeback verification](https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#agent-writeback-verification) |
| Wrap WorkPaper operations as tools   | `npm run agent:tool-call`          | [Agent tool call loop](https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#agent-tool-call-loop)                 |
| Adapt tools to agent frameworks      | `npm run agent:framework-adapters` | [Agent framework adapters](https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#agent-framework-adapters)         |
| Expose MCP-style tools               | `npm run agent:mcp-tools`          | [MCP tool server shape](https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#mcp-tool-server-shape)               |
| Flag budget variance rows            | `npm run budget-variance`          | [Budget variance alerts](https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#budget-variance-alerts)             |
| Check fulfillment capacity           | `npm run fulfillment-capacity`     | [Fulfillment capacity plan](https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#fulfillment-capacity-plan)       |
| Check quote approval threshold       | `npm run quote-approval`           | [Quote approval threshold](https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#quote-approval-threshold)         |
| Forecast subscription MRR            | `npm run subscription-mrr`         | [Subscription MRR forecast](https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#subscription-mrr-forecast)       |
| Return workbook results over HTTP    | `npm run http-json-summary`        | [HTTP JSON summary](https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#http-json-summary)                       |
| Calculate invoice totals             | `npm run invoice-totals`           | [Invoice totals](https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#invoice-totals)                             |
| Inspect restored workbook shape      | `npm run sheet-inspection`         | [Sheet inspection](https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#sheet-inspection)                         |
| Compare computed values and formulas | `npm run range-readback`           | [Range readback](https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#range-readback)                             |

## Production Status

Use this package in production for documented WorkPaper/headless API workflows:
programmatic workbook creation, formula evaluation, structural edits,
persistence round trips, service-side spreadsheet automation, and agent-driven
workbook operations.

Current release posture:

- Full local CI passed after the latest headless hardening work, including unit,
  contract, fuzz, browser, clean-diff, release-budget, runtime-publish, and
  WorkPaper competitive benchmark gates.
- The checked-in competitive artifact generated on `2026-05-08T15:00:27.603Z`
  shows `46/46` comparable WorkPaper mean wins against HyperFormula-style
  workloads: `38/38` public and `8/8` holdout.
- The shareable benchmark card is generated from that artifact:
  [`docs/assets/workpaper-benchmark-card.png`](https://github.com/proompteng/bilig/blob/main/docs/assets/workpaper-benchmark-card.png).
- The public benchmark evidence note explains the measured workload families,
  engine metadata, exclusions, and the current p95 nuance:
  [`docs/headless-workpaper-benchmark-evidence.md`](https://github.com/proompteng/bilig/blob/main/docs/headless-workpaper-benchmark-evidence.md).
- The shareable benchmark explainer states what the scorecard proves and what
  it does not:
  [`docs/what-workpaper-benchmark-proves.md`](https://github.com/proompteng/bilig/blob/main/docs/what-workpaper-benchmark-proves.md).
- Recently fixed and hardened P1 risks are covered by regression tests:
  - `updateConfig()` now applies `useColumnIndex` correctly when a rebuild-only
    config key changes in the same update.
  - `moveCells()` now respects `maxRows` and `maxColumns` and rejects target
    ranges outside configured bounds.
  - persisted WorkPaper documents reject sparse arrays, duplicate sheet names,
    non-JSON numeric values, invalid scoped named expressions, and invalid
    persisted config before restore.

Supported scope:

- The contract is the WorkPaper/headless API exported by this package.
- Excel-file ingestion belongs to import/export pipelines before data reaches
  `WorkPaper`; this package executes the validated WorkPaper workbook model.
  Use `WorkPaper.buildFromSnapshot()` for imported workbook snapshots so Excel
  defined names, tables, and translated formulas stay attached to the runtime
  model.
- XLSX cached-result parity investigations are covered by the repository
  verifier, not by the published package surface. Use
  `pnpm workpaper:xlsx-corpus:check -- <xlsx-file-or-directory>` for external
  corpora, and `pnpm workpaper:xlsx-corpus:fixtures:check` for the checked-in
  issue #8 reduction corpus. The verifier compares deterministic cached formula
  results and skips volatile or environment-dependent formulas such as `NOW()`
  and `CELL("filename")`; unsupported deterministic formulas remain visible as
  mismatches instead of being silently accepted.
- Custom function plugins and callback hooks are runtime registrations. Persist
  workbook data with the helpers below, then register custom behavior in
  application code before restore.

## XLSX Import And Export

XLSX ingestion and export are developed in this repository:

```sh
git clone https://github.com/proompteng/bilig.git
cd bilig
pnpm install
pnpm --filter @bilig/excel-import build
```

`@bilig/excel-import` lives in this monorepo, but its npm package name is still
being provisioned. Until that package is published on npm, use a repository
checkout for XLSX import/export work instead of adding `@bilig/excel-import` as
an external dependency.

Repository links:

- Website: <https://proompteng.github.io/bilig/>
- GitHub: <https://github.com/proompteng/bilig>
- Star or bookmark: <https://github.com/proompteng/bilig>
- Feedback discussion:
  <https://github.com/proompteng/bilig/discussions/157>
- Good first issues:
  <https://github.com/proompteng/bilig/blob/main/docs/starter-issues.md>
- First-timers-only issues:
  <https://github.com/proompteng/bilig/issues?q=is%3Aissue%20state%3Aopen%20label%3Afirst-timers-only>
- npm: <https://www.npmjs.com/package/@bilig/headless>
- runnable example:
  [`examples/headless-workpaper`](https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper)
- agent framework adapters example:
  [`npm run agent:framework-adapters`](https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#agent-framework-adapters)
- JSON records input example:
  [`npm run json-records`](https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#json-records-input)
- Formula diagnostics example:
  [`npm run formula-diagnostics`](https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#formula-diagnostics)
- Markdown report output example:
  [`npm run markdown-report`](https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#markdown-report-output)
- Snapshot diff example:
  [`npm run snapshot-diff`](https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#snapshot-diff)
- Range readback example:
  [`npm run range-readback`](https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#range-readback)
- Sheet inspection example:
  [`npm run sheet-inspection`](https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#sheet-inspection)
- serverless API route example:
  [`examples/serverless-workpaper-api`](https://github.com/proompteng/bilig/tree/main/examples/serverless-workpaper-api)
- Next.js App Router smoke:
  [`npm run next-route-handler`](https://github.com/proompteng/bilig/tree/main/examples/serverless-workpaper-api#nextjs-app-router-smoke)
- Express, Fastify, and Hono API adapters:
  [`npm run framework-adapters`](https://github.com/proompteng/bilig/tree/main/examples/serverless-workpaper-api#framework-adapters)
- Postgres JSONB, Redis/KV, and object-storage persistence adapters:
  [`npm run persistence-adapters`](https://github.com/proompteng/bilig/tree/main/examples/serverless-workpaper-api#persistence-adapters)
- persistence adapter feedback:
  <https://github.com/proompteng/bilig/discussions/307>
- JavaScript spreadsheet library guide feedback:
  <https://github.com/proompteng/bilig/discussions/308>
- server-side spreadsheet automation:
  [`docs/server-side-spreadsheet-automation-node.md`](https://github.com/proompteng/bilig/blob/main/docs/server-side-spreadsheet-automation-node.md)
- Node service recipe:
  [`docs/node-service-workpaper-recipe.md`](https://github.com/proompteng/bilig/blob/main/docs/node-service-workpaper-recipe.md)
- serverless API route recipe:
  [`docs/serverless-workpaper-api-route.md`](https://github.com/proompteng/bilig/blob/main/docs/serverless-workpaper-api-route.md)
- CSV-shaped input recipe:
  [`docs/csv-shaped-workpaper-input-recipe.md`](https://github.com/proompteng/bilig/blob/main/docs/csv-shaped-workpaper-input-recipe.md)
- unsupported formula troubleshooting:
  [`docs/unsupported-formula-troubleshooting-recipe.md`](https://github.com/proompteng/bilig/blob/main/docs/unsupported-formula-troubleshooting-recipe.md)
- agent tool-calling recipe:
  [`docs/agent-workpaper-tool-calling-recipe.md`](https://github.com/proompteng/bilig/blob/main/docs/agent-workpaper-tool-calling-recipe.md)
- revenue-model article:
  [`docs/building-a-revenue-model-with-headless-workpaper.md`](https://github.com/proompteng/bilig/blob/main/docs/building-a-revenue-model-with-headless-workpaper.md)
- compatibility boundaries:
  [`docs/where-bilig-is-not-excel-compatible-yet.md`](https://github.com/proompteng/bilig/blob/main/docs/where-bilig-is-not-excel-compatible-yet.md)
- XLSX corpus verifier walkthrough:
  [`docs/xlsx-corpus-verifier-walkthrough.md`](https://github.com/proompteng/bilig/blob/main/docs/xlsx-corpus-verifier-walkthrough.md)
- HyperFormula comparison:
  [`docs/hyperformula-alternative-headless-workpaper.md`](https://github.com/proompteng/bilig/blob/main/docs/hyperformula-alternative-headless-workpaper.md)
- SheetJS and ExcelJS comparison:
  [`docs/sheetjs-exceljs-alternative-formula-workbook-api.md`](https://github.com/proompteng/bilig/blob/main/docs/sheetjs-exceljs-alternative-formula-workbook-api.md)
- JavaScript spreadsheet library guide:
  [`docs/javascript-spreadsheet-library-headless-node.md`](https://github.com/proompteng/bilig/blob/main/docs/javascript-spreadsheet-library-headless-node.md)
- broader headless spreadsheet engine comparison:
  [`docs/headless-spreadsheet-engine-comparison.md`](https://github.com/proompteng/bilig/blob/main/docs/headless-spreadsheet-engine-comparison.md)
- local benchmark walkthrough:
  [`docs/local-workpaper-benchmark-walkthrough.md`](https://github.com/proompteng/bilig/blob/main/docs/local-workpaper-benchmark-walkthrough.md)
- XLOOKUP exact fixture walkthrough:
  [`docs/formula-edge-xlookup-exact-fixture.md`](https://github.com/proompteng/bilig/blob/main/docs/formula-edge-xlookup-exact-fixture.md)
- SUMIFS paired criteria fixture walkthrough:
  [`docs/formula-edge-sumifs-paired-criteria-fixture.md`](https://github.com/proompteng/bilig/blob/main/docs/formula-edge-sumifs-paired-criteria-fixture.md)
- GROUPBY spill fixture walkthrough:
  [`docs/formula-edge-groupby-spill-fixture.md`](https://github.com/proompteng/bilig/blob/main/docs/formula-edge-groupby-spill-fixture.md)
- published DEV article:
  <https://dev.to/gregkonush/why-agents-need-workbook-apis-instead-of-spreadsheet-screenshots-3d61>
- DEV article source:
  [`docs/dev-to-workbook-apis-post.md`](https://github.com/proompteng/bilig/blob/main/docs/dev-to-workbook-apis-post.md)

## Quickstart

This snippet is safe to paste into a one-file npm evaluation. It verifies formula
readback before and after a persisted restore instead of only printing a demo
value.

```ts
import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
  type WorkPaperCellAddress,
} from '@bilig/headless'

const workbook = WorkPaper.buildFromSheets(
  {
    Sheet1: [
      [10, 20, '=A1+B1'],
      [7, '=A2*3', null],
    ],
  },
  {
    maxRows: 1_000,
    maxColumns: 100,
    useColumnIndex: true,
  },
)

const numberValue = (cell: unknown): number => {
  if (typeof cell === 'object' && cell !== null && 'value' in cell) {
    const value = (cell as { value: unknown }).value
    if (typeof value === 'number') {
      return value
    }
  }
  throw new Error(`expected numeric cell value, got ${JSON.stringify(cell)}`)
}

const sheet = workbook.getSheetId('Sheet1')
if (sheet === undefined) {
  throw new Error('Sheet1 was not created')
}

const at = (row: number, col: number): WorkPaperCellAddress => ({
  sheet,
  row,
  col,
})

const initial = numberValue(workbook.getCellValue(at(0, 2)))
if (initial !== 30) {
  throw new Error(`unexpected initial formula value: ${String(initial)}`)
}

workbook.setCellContents(at(1, 2), '=A2+B2')
if (workbook.getCellFormula(at(1, 2)) !== '=A2+B2') {
  throw new Error('formula text was not recorded')
}

const document = exportWorkPaperDocument(workbook)
const json = serializeWorkPaperDocument(document)
const restored = createWorkPaperFromDocument(parseWorkPaperDocument(json))
const restoredSheet = restored.getSheetId('Sheet1')
if (restoredSheet === undefined) {
  throw new Error('Sheet1 was not restored')
}

const restoredValue = numberValue(restored.getCellValue({ sheet: restoredSheet, row: 1, col: 2 }))
const verified = restoredValue === 28 && restored.getSheetNames().includes('Sheet1')
if (!verified) {
  throw new Error(`unexpected restored formula value: ${String(restoredValue)}`)
}

console.log({ initial, restoredValue, bytes: json.length, verified })
```

## Runnable Example

The repo includes a small external-consumer project at
[`examples/headless-workpaper`](https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper). It builds a
revenue workbook, evaluates formulas, applies an agent-style edit, persists and
restores the workbook, and verifies the final result.

```sh
cd examples/headless-workpaper
npm install
npm start
npm run http-json-summary
npm run agent:tool-call
npm run agent:framework-adapters
```

For the agent-specific writeback proof, run:

```sh
npm run agent:verify
```

That demo records the exact assumption cells changed, verifies dependent formula
readback, persists the workbook, restores it, and checks that formulas and
outputs survived the round trip.

`npm run agent:tool-call` exposes a small SDK-neutral tool-call shape:
`readRange`, `setInputCell`, computed before/after values, formula contracts,
persistence bytes, and restored readback equality.

`npm run agent:framework-adapters` maps the same validated WorkPaper operations
into AI SDK, LangChain, Mastra, LlamaIndex.TS, LangGraph.js, CopilotKit, and
Cloudflare Agents tool shapes. The example verifies the same write and readback
contract for each wrapper without adding those agent frameworks as dependencies.

`npm run agent:mcp-tools` exposes the same operations through dependency-free
MCP-style `tools/list` and `tools/call` JSON-RPC responses with JSON Schema
inputs and structured formula readback.
`npm run agent:mcp-stdio` puts the same handlers behind newline-delimited
JSON-RPC over stdin/stdout. The `mcp` keyword points to these runnable local
adapters: no hosted service, no API key, and no agent framework dependency.

The npm package also ships the same demo server as a stdio binary:

```sh
npm exec --package @bilig/headless -- bilig-workpaper-mcp
```

Its package metadata includes `mcpName: io.github.proompteng/bilig-workpaper`
and `server.json`, and the server is published in the official MCP Registry:
<https://registry.modelcontextprotocol.io/v0.1/servers?search=io.github.proompteng%2Fbilig-workpaper>.

For copy-paste client configs, use the
[MCP client setup guide](https://proompteng.github.io/bilig/mcp-client-setup.html).

Before submitting the server to an MCP registry, verify this repo-specific
readiness checklist:

- `packages/headless/server.json` exists and describes the packaged stdio
  server.
- `packages/headless/package.json` exposes `bilig-workpaper-mcp` in `bin`.
- `packages/headless/package.json` includes
  `mcpName: io.github.proompteng/bilig-workpaper`.
- `pnpm publish:runtime:check` passes against the runtime packages.
- `pnpm workpaper:smoke:external` passes against packed local runtime packages.
- `pnpm mcpb:workpaper:build` creates `build/mcpb/bilig-workpaper.mcpb` from
  the published npm package when you want a Claude Desktop bundle.

Passing the checklist means the package metadata and smoke checks are ready for
registry submission; it does not mean this package version has already been
published.

For a framework-neutral recipe that wraps WorkPaper operations as agent-callable
tools, see
[`docs/agent-workpaper-tool-calling-recipe.md`](https://github.com/proompteng/bilig/blob/main/docs/agent-workpaper-tool-calling-recipe.md).
It covers validated sheet/address parsing, computed before/after readback, and
persistence after a successful edit.

Repository CI also runs the same example against packed local runtime packages
through `pnpm workpaper:smoke:external`.

For a minimal service boundary with no framework dependency, see
[`docs/node-service-workpaper-recipe.md`](https://github.com/proompteng/bilig/blob/main/docs/node-service-workpaper-recipe.md).
It shows a built-in Node HTTP route that reads a computed summary, applies one
controlled input edit, and persists the WorkPaper document.

For simple tabular service payloads, see
[`docs/csv-shaped-workpaper-input-recipe.md`](https://github.com/proompteng/bilig/blob/main/docs/csv-shaped-workpaper-input-recipe.md).
It normalizes a small CSV-shaped fixture into the `WorkPaper.buildFromSheets()`
array shape and reads formula-backed summaries.

For formula error handling, see
[`docs/unsupported-formula-troubleshooting-recipe.md`](https://github.com/proompteng/bilig/blob/main/docs/unsupported-formula-troubleshooting-recipe.md).
It shows how to pair `getCellDisplayValue()` with
`getCellFormulaDiagnostics()` so Node services and agent tools can return
actionable errors instead of silently accepting unsupported formula inputs.

For a focused persistence walkthrough, see
[`docs/persisting-formula-backed-workpaper-documents-in-node.md`](https://github.com/proompteng/bilig/blob/main/docs/persisting-formula-backed-workpaper-documents-in-node.md)
and run the example package:

```sh
cd examples/headless-workpaper
npm install
npm run persistence
```

For service storage, run the typed Postgres JSONB, Redis/KV, and object-storage
adapter smoke:

```sh
cd examples/serverless-workpaper-api
npm install
npm run persistence-adapters
```

## Core Concepts

- `WorkPaper` is the top-level workbook object. Create it with
  `WorkPaper.buildEmpty()`, `WorkPaper.buildFromArray()`,
  `WorkPaper.buildFromSheets()`, or `WorkPaper.buildFromSnapshot()` for
  importer-produced workbook snapshots.
- Addresses are zero-based `{ sheet, row, col }` objects. Use `getSheetId(name)`
  to resolve the numeric sheet id.
- A string beginning with `=` is a formula. Other strings are literal text.
  `null` clears a cell.
- `getCellValue()` returns a computed `CellValue`.
- `getCellDisplayValue()` returns the formatted user-facing text for a cell,
  including workbook error text such as `#VALUE!`.
- `getCellFormulaDiagnostics()` returns structured formula diagnostics for
  supported error families, including financial cash-flow/date validation for
  `XIRR()` and `XNPV()`.
- `getCellFormula()` returns the formula text when the cell is a formula.
- `getCellSerialized()` returns the persisted cell input shape.
- Mutation methods return WorkPaper change arrays. Empty arrays are valid when
  evaluation is suspended or a batch defers change publication.

## WorkPaper Read/Write Cheat Sheet

The common service and agent calls use the same zero-based address object:

```ts
const sheet = workbook.getSheetId('Sheet1')
if (sheet === undefined) throw new Error('Sheet1 was not created')
const at = (row: number, col: number) => ({ sheet, row, col })
```

| Operation               | Public API                                                     | Tiny snippet                                                                                          |
| ----------------------- | -------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------- |
| Create a workbook       | `WorkPaper.buildFromSheets()`                                  | `const workbook = WorkPaper.buildFromSheets({ Sheet1: [[1, '=A1*2']] })`                              |
| Set a value             | `workbook.setCellContents()`                                   | `workbook.setCellContents(at(0, 0), 42)`                                                              |
| Set a formula           | `workbook.setCellContents()`                                   | `workbook.setCellContents(at(0, 1), '=A1*2')`                                                         |
| Read a calculated value | `workbook.getCellValue()`                                      | `const value = workbook.getCellValue(at(0, 1))`                                                       |
| Read formula text       | `workbook.getCellFormula()`                                    | `const formula = workbook.getCellFormula(at(0, 1))`                                                   |
| Export/persist state    | `exportWorkPaperDocument()` and `serializeWorkPaperDocument()` | `const json = serializeWorkPaperDocument(exportWorkPaperDocument(workbook, { includeConfig: true }))` |
| Restore persisted state | `parseWorkPaperDocument()` and `createWorkPaperFromDocument()` | `const restored = createWorkPaperFromDocument(parseWorkPaperDocument(json))`                          |

Use `getCellDisplayValue()` for user-facing text, `getCellSerialized()` for the
stored input shape, and range helpers such as `getRangeValues()` when an API
response needs more than one cell.

## Common API Recipes

Create an empty workbook:

```ts
const workbook = WorkPaper.buildEmpty({
  maxRows: 10_000,
  maxColumns: 256,
})
```

Set values, formulas, and blanks:

```ts
workbook.setCellContents(at(0, 0), 42)
workbook.setCellContents(at(0, 1), 'label')
workbook.setCellContents(at(0, 2), '=A1*2')
workbook.setCellContents(at(0, 3), null)
```

Read ranges and whole sheets:

```ts
const range = { start: at(0, 0), end: at(9, 4) }

workbook.getRangeValues(range)
workbook.getRangeFormulas(range)
workbook.getRangeSerialized(range)
workbook.getSheetValues(sheet)
workbook.getSheetSerialized(sheet)
```

Inspect a formula error:

```ts
const value = workbook.getCellValue(at(7, 1))
const display = workbook.getCellDisplayValue(at(7, 1))
const diagnostics = workbook.getCellFormulaDiagnostics(at(7, 1))

console.log(value) // Raw CellValue protocol object
console.log(display) // "#VALUE!"
console.log(diagnostics[0]?.code) // e.g. "financial-unsupported-date-coercion"
```

Finance date inputs:

`XIRR(values, dates, [guess])` and `XNPV(rate, values, dates)` accept numeric
Excel serial dates in `dates`. Text labels and text date strings are not
coerced in headless formulas; they evaluate to `#VALUE!`. Use
`getCellFormulaDiagnostics()` to distinguish invalid date ranges, mismatched
range dimensions, invalid cash-flow values, missing positive or negative cash
flows, invalid rates/guesses, and solver non-convergence.

Batch related edits:

```ts
const changes = workbook.batch(() => {
  workbook.setCellContents(at(0, 0), 10)
  workbook.setCellContents(at(0, 1), '=A1*5')
})
```

Create and read named expressions:

```ts
import { WorkPaper, type WorkPaperCellAddress } from '@bilig/headless'

const workbook = WorkPaper.buildFromSheets({
  Plan: [
    ['Metric', 'Value'],
    ['Base revenue', 100],
    ['Target revenue', null],
  ],
})

const sheet = workbook.getSheetId('Plan')
if (sheet === undefined) {
  throw new Error('Plan sheet was not created')
}

const at = (row: number, col: number): WorkPaperCellAddress => ({
  sheet,
  row,
  col,
})

workbook.addNamedExpression('GrowthRate', 0.15)
workbook.setCellContents(at(2, 1), '=B2*(1+GrowthRate)')

console.log(workbook.getNamedExpressionValue('GrowthRate')) // CellValue for 0.15
console.log(workbook.getNamedExpressionFormula('GrowthRate')) // undefined for a scalar name
console.log(workbook.getCellValue(at(2, 1))) // CellValue for 115
console.log(workbook.getNamedExpressionValue('MissingName')) // undefined
```

Move cells inside configured bounds:

```ts
const source = { start: at(0, 0), end: at(0, 1) }

workbook.moveCells(source, at(2, 0))
```

`moveCells()` throws when the source or target falls outside `maxRows` or
`maxColumns`.

Update runtime config:

```ts
workbook.updateConfig({
  useColumnIndex: false,
  maxRows: 2_000,
})
```

`updateConfig()` preserves workbook state. Some config keys can trigger an
internal engine rebuild; callers should treat the returned workbook object as the
same public facade and read `getConfig()` after update when they need the active
configuration.

Persist and restore:

```ts
const saved = serializeWorkPaperDocument(exportWorkPaperDocument(workbook, { includeConfig: true }))

const restored = createWorkPaperFromDocument(parseWorkPaperDocument(saved))
```

Validate a persisted document before restore:

```ts
import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  isPersistedWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from '@bilig/headless'

const workbook = WorkPaper.buildFromSheets({
  Sheet1: [[10, '=A1*2']],
})

const document = exportWorkPaperDocument(workbook, { includeConfig: true })
const serialized = serializeWorkPaperDocument(document)

const parsed = parseWorkPaperDocument(serialized)
if (!isPersistedWorkPaperDocument(parsed)) {
  throw new Error('Persisted WorkPaper document failed validation')
}

const restored = createWorkPaperFromDocument(parsed)
```

`parseWorkPaperDocument()` validates JSON and throws on invalid payloads.
`isPersistedWorkPaperDocument()` is useful when a service already has an
unknown parsed object. Custom function implementations, callback hooks, and
other process state are not persisted; register those in application code before
restoring the workbook.

## Persistence Contract

`@bilig/headless` persists:

- ordered sheet content arrays
- sheet names
- named expressions
- the JSON-safe subset of `WorkPaperConfig` when `includeConfig` is enabled

It does not persist custom function implementations, callback hooks, or process
state. Register those in code before creating or restoring workbooks.

## Validation Commands

For a headless-only change, start with focused tests:

```sh
pnpm exec vitest run \
  packages/headless/src/__tests__/work-paper-runtime.test.ts \
  packages/headless/src/__tests__/work-paper-parity.test.ts \
  packages/headless/src/__tests__/persistence.test.ts \
  packages/headless/src/__tests__/persistence.fuzz.test.ts

pnpm --filter @bilig/headless build
```

Before publishing or claiming production readiness, run the full gates:

```sh
pnpm publish:runtime:check
pnpm workpaper:bench:competitive:check
pnpm run ci
```

For a newcomer-friendly benchmark command walkthrough, see
[`docs/local-workpaper-benchmark-walkthrough.md`](https://github.com/proompteng/bilig/blob/main/docs/local-workpaper-benchmark-walkthrough.md).
It explains the committed artifact check, a reduced local smoke run, and the
scorecard fields to compare in benchmark diffs.

Regenerate the competitive artifact only when intentionally updating benchmark
evidence:

```sh
pnpm workpaper:bench:competitive:generate
pnpm workpaper:bench:competitive:check
```

Do not change benchmark definitions, scoring, sampling, or workload sizes to hide
losses.

For XLSX compatibility investigations with cached formula results, run the
corpus verifier against real workbook files:

```sh
pnpm workpaper:xlsx-corpus:check -- /path/to/xlsx-corpus
```

The verifier reads `.xlsx`, `.xlsm`, and `.xls` files, builds a WorkPaper model
with `useColumnIndex: true`, and compares formula cells against cached workbook
results. It reports `totalFiles`, `failedTimeouts`, `comparableFormulaCells`,
`matchingFormulaCells`, `mismatchedFormulaCells`, `matchRate`, skipped formulas,
and actionable mismatch samples. Missing cached results and volatile or
environment-dependent formulas such as `NOW()` and `CELL()` are counted as
skipped instead of silently treated as parity evidence.

## Compatibility Notes

- The facade follows HyperFormula-style public workbook workflows, but it is not
  byte-for-byte compatible with HyperFormula.
- Public lookup helpers such as `getSheetId()`, `getSheetName()`,
  `simpleCellAddressFromString()`, `simpleCellRangeFromString()`, and
  named-expression reads return `undefined` on misses.
- `@bilig/headless` keeps `bilig` change arrays and also exposes richer detailed
  event payloads through `onDetailed()`, `onceDetailed()`, and `offDetailed()`.
- Stable compatibility adapters are available through `graph`, `rangeMapping`,
  `arrayMapping`, `sheetMapping`, `addressMapping`, `dependencyGraph`,
  `evaluator`, `columnSearch`, and `lazilyTransformingAstService`.
- Cached XLSX result parity is a testable corpus property, not a blanket
  package guarantee. Use `pnpm workpaper:xlsx-corpus:check -- <corpus>` for
  supplied workbook corpora, and keep unsupported or volatile workbook functions
  documented in the resulting report.

## For Coding Agents

Use this checklist when Codex, Claude Code, or another agent starts work here:

1. Read this README and the root `README.md` first.
2. Use public exports from `@bilig/headless`; do not import from `src/`,
   `dist/internal`, or `@bilig/core` unless the task explicitly requires engine
   integration work.
3. Use zero-based `{ sheet, row, col }` addresses and resolve sheet ids with
   `getSheetId()`.
4. Prefer `WorkPaper.buildFromSheets()` for hand-authored fixtures,
   `WorkPaper.buildFromSnapshot()` for importer-produced workbook snapshots, and
   `exportWorkPaperDocument()` / `createWorkPaperFromDocument()` for persistence
   round trips.
5. Add or tighten regression tests before changing behavior around config
   rebuilds, range bounds, formulas, persistence, events, row/column moves, or
   sheet lifecycle.
6. Run the focused headless tests before broader gates.
7. Preserve benchmark definitions and workload sizes. Performance improvements
   belong in production engine/headless code.
8. Document unsupported behavior honestly instead of implying full Excel
   compatibility.

## Public Entry Points

The package root exports:

- `WorkPaper`
- WorkPaper address, range, config, sheet, change, event, and adapter types
- WorkPaper error classes
- persistence helpers:
  - `exportWorkPaperDocument()`
  - `createWorkPaperFromDocument()`
  - `serializeWorkPaperDocument()`
  - `parseWorkPaperDocument()`
  - `isPersistedWorkPaperDocument()`
  - `pickPersistableWorkPaperConfig()`

## Versioning

`@bilig/headless` ships as part of the aligned bilig runtime package set. Treat
documented public exports as the supported surface, keep integration tests around
your own workbook corpus, and rerun the validation gates before upgrading in
production.
