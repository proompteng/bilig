# @bilig/headless

[![npm: @bilig/headless](https://img.shields.io/npm/v/@bilig/headless?label=%40bilig%2Fheadless)](https://www.npmjs.com/package/@bilig/headless)
[![npm weekly downloads](https://img.shields.io/npm/dw/@bilig/headless?label=npm%20downloads)](https://www.npmjs.com/package/@bilig/headless)
[![GitHub](https://img.shields.io/badge/GitHub-proompteng%2Fbilig-blue)](https://github.com/proompteng/bilig)
[![GitHub Repo stars](https://img.shields.io/github/stars/proompteng/bilig?style=social)](https://github.com/proompteng/bilig/stargazers)
[![MCP server score](https://glama.ai/mcp/servers/proompteng/bilig/badges/score.svg)](https://glama.ai/mcp/servers/proompteng/bilig)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](https://github.com/proompteng/bilig/blob/main/LICENSE)

`@bilig/headless` is a workbook runtime for Node.js services and agent tools.

Use it when the calculation is easiest to understand as sheets, cells, and
formulas, but the execution belongs in TypeScript. Your code owns a
`WorkPaper`: build sheets, write inputs, recalculate formulas, read the cell
value, and save the workbook as JSON.

Product code gets reviewable workbook-shaped logic without shipping a
spreadsheet UI. Coding agents get narrow tools such as `readRange` and
`setInputCell` instead of guessing state from screenshots.

This package is not a browser grid, desktop Excel automation, or a source of
truth for stale XLSX cached formula values. XLSX import/export is available from
the `@bilig/headless/xlsx` subpath for services that need workbook ingestion
around the same WorkPaper model.

## Install

Requires Node `24+` and ESM imports.

```sh
npm install @bilig/headless
```

<!-- headless-package-footprint:start -->

Current checked npm footprint for `@bilig/headless@0.16.0`:

- Pack dry run: `417 kB` tarball, `2.49 MB` unpacked, `420` package entries.
- Boundary: the main import is the WorkPaper formula/JSON runtime; XLSX
  import/export stays behind the `@bilig/headless/xlsx` subpath; MCP is the
  `bilig-workpaper-mcp` binary wrapper.
- Cold-start gate: Node imports the main entrypoint, builds a two-sheet
  WorkPaper, and reads `24000` under `1000 ms` without importing
  the XLSX subpath.
- Runtime: Node `>=24.0.0`; Node 22 support waits for release CI coverage.
<!-- headless-package-footprint:end -->

For a clean copy-paste run, use the
[Node quickstart](https://proompteng.github.io/bilig/try-bilig-headless-in-node.html).
If you are choosing between formula engines, read the
[TypeScript guide for evaluating Excel formulas in Node.js](https://proompteng.github.io/bilig/evaluate-excel-formulas-in-node-typescript.html)
and the
[Google Sheets API boundary](https://proompteng.github.io/bilig/google-sheets-api-alternative-node-workpaper.html).

## TypeScript API Shape

Most integrations are this loop: create a workbook, write an input, read the
calculated cell, and save the workbook state.

```ts
import { WorkPaper, exportWorkPaperDocument, serializeWorkPaperDocument } from '@bilig/headless'

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
const saved = serializeWorkPaperDocument(exportWorkPaperDocument(workbook, { includeConfig: true }))

console.log({ revenue, savedBytes: saved.length })
```

Addresses are zero-based `{ sheet, row, col }` objects. A formula is a string
that begins with `=`. Sheet ids are numeric and should be resolved with
`workbook.getSheetId(name)`.

## Clean npm Sanity Check

Run this before cloning the repository. It checks the published npm package by
building a workbook, changing an input, saving the document, restoring it, and
checking that the dependent formula still reads back correctly.

```sh
mkdir bilig-headless-sanity
cd bilig-headless-sanity
npm init -y
npm pkg set type=module
npm install @bilig/headless
npm install -D tsx typescript @types/node
curl -fsSLo sanity.ts https://proompteng.github.io/bilig/npm-eval.ts
npx tsx sanity.ts
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

The downloaded file is the maintained TypeScript example at
[`examples/headless-workpaper/npm-eval.ts`](https://github.com/proompteng/bilig/blob/main/examples/headless-workpaper/npm-eval.ts).
The exact byte count can change between package versions; `verified: true` and
matching `after`/`afterRestore` values are the check.

Inside this monorepo:

```sh
pnpm install
pnpm --filter @bilig/headless build
```

## When To Use It

Reach for `@bilig/headless` when:

- a service owns a workbook-shaped calculation and needs formula readback;
- an agent tool must prove the value after an edit;
- a queue worker or route needs deterministic spreadsheet state without a UI;
- tests need the same formula model that production code uses;
- a workbook document needs to round-trip as JSON after code changes it.

Use something else when you need:

- manual spreadsheet editing;
- a browser grid by itself;
- Office macros, COM automation, or desktop Excel integration;
- one-off arithmetic where a workbook model adds no value.

## Quickstart

The shortest local path is still TypeScript. Put the API shape above in a
`sanity.ts` file, run it with `tsx`, and expect the dependent formula to change
after the `setCellContents()` call. For a maintained file that already includes
restore verification, use the clean npm sanity check.

## WorkPaper Read/Write Cheat Sheet

The public surface is intentionally small:

- Create workbooks with `WorkPaper.buildEmpty()`, `WorkPaper.buildFromArray()`,
  `WorkPaper.buildFromSheets()`, or `WorkPaper.buildFromSnapshot()`.
- Edit values, formulas, and blanks with `workbook.setCellContents(address, value)`.
- Read computed values with `workbook.getCellValue(address)`.
- Read display text with `workbook.getCellDisplayValue(address)`.
- Read formula text with `workbook.getCellFormula(address)`.
- Read persisted cell input with `workbook.getCellSerialized(address)`.
- Read ranges with `getRangeValues()`, `getRangeFormulas()`, and
  `getRangeSerialized()`.
- Persist with `exportWorkPaperDocument()` and `serializeWorkPaperDocument()`.
- Restore with `parseWorkPaperDocument()` and `createWorkPaperFromDocument()`.

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
  { maxRows: 1_000, maxColumns: 100, useColumnIndex: true },
)

const sheet = workbook.getSheetId('Sheet1')
if (sheet === undefined) {
  throw new Error('Sheet1 was not created')
}

const at = (row: number, col: number): WorkPaperCellAddress => ({ sheet, row, col })

workbook.setCellContents(at(1, 2), '=A2+B2')

const saved = serializeWorkPaperDocument(exportWorkPaperDocument(workbook, { includeConfig: true }))
const restored = createWorkPaperFromDocument(parseWorkPaperDocument(saved))

console.log({
  formula: workbook.getCellFormula(at(1, 2)),
  display: workbook.getCellDisplayValue(at(1, 2)),
  sheets: restored.getSheetNames(),
})
```

For formula errors, pair `getCellDisplayValue()` with
`getCellFormulaDiagnostics()`. That lets a service return useful `#VALUE!` or
`#NAME?` diagnostics instead of silently accepting unsupported inputs.

## Runnable Examples

The example catalog lives in
[`examples/headless-workpaper`](https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper).
The examples are TypeScript files. Some imports end in `.js` because Node ESM
resolves compiled package output that way; the files you edit and run are still
`.ts`.

Start with the data shape closest to your app:

- `npm run json-records`:
  [JSON records input](https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#json-records-input)
- `npm run csv-shaped`:
  [CSV shaped input](https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#csv-shaped-input)
- `npm run invoice-totals`:
  [invoice totals](https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#invoice-totals)
- `npm run budget-variance`:
  [budget variance alerts](https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#budget-variance-alerts)
- `npm run fulfillment-capacity`:
  [fulfillment capacity plan](https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#fulfillment-capacity-plan)
- `npm run quote-approval`:
  [quote approval threshold](https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#quote-approval-threshold)
- `npm run subscription-mrr`:
  [subscription MRR forecast](https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#subscription-mrr-forecast)
- `npm run persistence`:
  [persistence round trip](https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#persistence-round-trip)
- `npm run range-readback`:
  [range readback](https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#range-readback)
- `npm run sheet-inspection`:
  [sheet inspection](https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#sheet-inspection)

Agent and tool-call examples:

- `npm run agent:verify` proves an agent writeback by checking the dependent
  formula, saved JSON, restored workbook, and formula text.
- `npm run agent:tool-call` exposes `readRange` and `setInputCell` style tool
  calls with computed before/after readback.
- `npm run agent:openai-responses` shows the
  [OpenAI Responses tool-call loop](https://github.com/proompteng/bilig/blob/main/docs/openai-responses-workpaper-tool-call.md).
- `npm run agent:ai-sdk-generate-text` uses the real Vercel AI SDK
  `generateText()` and `tool()` APIs; the runnable file is
  [`ai-sdk-generate-text-tool-smoke.ts`](https://github.com/proompteng/bilig/blob/main/examples/headless-workpaper/ai-sdk-generate-text-tool-smoke.ts).
- `npm run agent:ai-sdk-stream-text` covers the matching streamed tool-call
  path in
  [`ai-sdk-stream-text-tool-smoke.ts`](https://github.com/proompteng/bilig/blob/main/examples/headless-workpaper/ai-sdk-stream-text-tool-smoke.ts).
- `npm run agent:framework-adapters` maps the same validated WorkPaper
  operations into AI SDK, LangChain, Mastra, LlamaIndex.TS, LangGraph.js,
  CopilotKit, and Cloudflare Agents:
  <https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#agent-framework-adapters>.

MCP examples:

- `npm run agent:mcp-tools` returns dependency-free `tools/list` and
  `tools/call` JSON-RPC shapes:
  <https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#mcp-tool-server-shape>.
- `npm run agent:mcp-stdio` runs the same handlers over newline-delimited
  stdio.
- The package ships a stdio binary:

```sh
npm exec --package @bilig/headless -- bilig-workpaper-mcp
npm exec --package @bilig/headless -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --writable
```

Default mode starts the built-in demo workbook. File-backed mode loads a
persisted WorkPaper JSON document and exposes `list_sheets`, `read_range`,
`read_cell`, `set_cell_contents`, `get_cell_display_value`,
`export_workpaper_document`, and `validate_formula`; `--writable` persists
`set_cell_contents` edits back to the same file.

The package metadata includes
`mcpName: io.github.proompteng/bilig-workpaper`, and the server is listed in the
official MCP Registry:
<https://registry.modelcontextprotocol.io/v0.1/servers?search=io.github.proompteng%2Fbilig-workpaper>.

For setup details, use the
[MCP server guide](https://github.com/proompteng/bilig/blob/main/docs/mcp-workpaper-tool-server.md),
[MCP directory status](https://github.com/proompteng/bilig/blob/main/docs/mcp-spreadsheet-server-directory.md),
[MCP client setup](https://github.com/proompteng/bilig/blob/main/docs/mcp-client-setup.md),
and
[Claude Desktop MCPB guide](https://github.com/proompteng/bilig/blob/main/docs/claude-desktop-mcpb-workpaper.md).

## Service Routes

For HTTP and serverless examples, start with
[`examples/serverless-workpaper-api`](https://github.com/proompteng/bilig/tree/main/examples/serverless-workpaper-api).

```sh
cd examples/serverless-workpaper-api
npm install
npm run quote-approval-api
npm run next-route-handler
npm run next-server-action
npm run next-server-action-formdata
npm run framework-adapters
npm run persistence-adapters
```

Start with `npm run quote-approval-api` when you want the production-shaped
proof: input JSON writes `Inputs!B2:B6`, formulas recalculate, the WorkPaper
JSON is persisted, and a restored workbook returns the same approval decision.

Useful anchors:

- [quote approval API smoke](https://github.com/proompteng/bilig/tree/main/examples/serverless-workpaper-api#quote-approval-api-smoke)
- [Next.js App Router smoke](https://github.com/proompteng/bilig/tree/main/examples/serverless-workpaper-api#nextjs-app-router-smoke)
- [Next.js Server Action smoke](https://github.com/proompteng/bilig/tree/main/examples/serverless-workpaper-api#nextjs-server-action-smoke)
- [Next.js Server Action FormData smoke](https://github.com/proompteng/bilig/tree/main/examples/serverless-workpaper-api#nextjs-server-action-formdata-smoke)
- [framework adapters](https://github.com/proompteng/bilig/tree/main/examples/serverless-workpaper-api#framework-adapters)
- [persistence adapters](https://github.com/proompteng/bilig/tree/main/examples/serverless-workpaper-api#persistence-adapters)

The public framework guide is
<https://proompteng.github.io/bilig/node-framework-workpaper-adapters.html>.

## XLSX Import And Export

Use the `@bilig/headless/xlsx` subpath for XLSX import, WorkPaper calculation,
edits, and XLSX export from the same published npm package:

```sh
pnpm add @bilig/headless
```

```ts
import { readFileSync, writeFileSync } from 'node:fs'
import { WorkPaper } from '@bilig/headless'
import { exportXlsx, importXlsx } from '@bilig/headless/xlsx'

const imported = importXlsx(new Uint8Array(readFileSync('model.xlsx')), 'model.xlsx')
const workbook = WorkPaper.buildFromSnapshot(imported.snapshot, {
  evaluationTimeoutMs: 30_000,
  useColumnIndex: true,
})

const firstSheetName = imported.snapshot.sheets[0]?.name
const firstSheet = firstSheetName === undefined ? undefined : workbook.getSheetId(firstSheetName)
if (firstSheet === undefined) throw new Error('Workbook has no sheets')

workbook.setCellContents({ sheet: firstSheet, row: 1, col: 1 }, 150_000)
const displayValue = workbook.getCellDisplayValue({ sheet: firstSheet, row: 1, col: 1 })

writeFileSync('model-edited.xlsx', exportXlsx(workbook.exportSnapshot()))
workbook.dispose()

console.log({ displayValue })
```

`WorkPaper.buildFromSnapshot()` preserves imported XLSX metadata such as
defined names, tables, hidden sheets, and translated structured references. Use
`workbook.exportSnapshot()` with `exportXlsx()` when exporting a WorkPaper after
edits.

### External Workbook References

XLSX files can contain links to other workbooks. `@bilig/headless/xlsx`
preserves those package artifacts, but it does not open or recalculate linked
workbooks by itself.

The importer exposes linked-workbook state in structured metadata:

- `snapshot.workbook.metadata.externalWorkbookReferences`: linked workbook
  package paths, external targets, workbook names when available, and cached
  sheet names.
- `snapshot.workbook.metadata.unsupportedFormulaDependencies`: affected formula
  cells, original and imported formula text, linked workbook references, and
  whether cached formula or linked-cell values were used.

Use one of these policies:

- Resolve: provide ordinary local inputs or formulas after import, then
  recalculate with `WorkPaper`.
- Preserve stale: keep imported cached values and preserved external-link
  artifacts, but treat formula correctness as unaudited for those dependencies.
- Strict-fail: reject the import when either metadata field above is non-empty.

The real-workbook corpus scorecard reports external references as
`xlsx.externalLinks.workbookReferencesPreserved` and direct formula dependencies
as `xlsx.externalLinks.formulaDependenciesUnsupported`, with linked workbook,
affected formula, and cached-value counts.

## Accuracy Policy

Do not call a Bilig accuracy bug from stale XLSX cache data.

Embedded cached formula values are useful diagnostics, but they are not the
source of truth. For XLSX formula accuracy, prepare a fresh Microsoft Excel
oracle and evaluate against the recalculated copy:

```sh
OUT=.cache/excel-oracle-evaluation
pnpm workpaper:xlsx-oracle -- prepare-oracle /path/to/xlsx-corpus "$OUT"
pnpm workpaper:xlsx-oracle -- evaluate-cache /path/to/xlsx-corpus "$OUT"
pnpm workpaper:xlsx-oracle -- evaluate-oracle /path/to/xlsx-corpus "$OUT/recalculated" "$OUT"
pnpm workpaper:xlsx-oracle -- summarize "$OUT"
```

`evaluate-cache` writes `cache-diagnostic.json` and stays non-authoritative.
`evaluate-oracle` writes `excel-oracle-report.json`, and `summarize` writes
`summary.md`. If Excel automation is unavailable, the harness marks cells as
`missing_excel_oracle` instead of promoting cache mismatches to correctness
bugs.

For quick cache triage only:

```sh
pnpm workpaper:xlsx-corpus:check -- /path/to/xlsx-corpus
```

## Proof You Can Reproduce

- The clean TypeScript sanity check above edits one input, restores the saved
  JSON document, and verifies the dependent formula result.
- Auditing imported Excel files is a separate workflow. Cached formula values
  embedded in `.xlsx` files are useful for triage, but Bilig accuracy claims
  should be checked against a fresh Microsoft Excel recalculation.
- Run `pnpm workpaper:bench:competitive:check` from the repository. The
  checked-in artifact shows
  [`48/57` comparable WorkPaper mean wins](https://github.com/proompteng/bilig/blob/main/docs/what-workpaper-benchmark-proves.md)
  and names the worst p95 holdout: `structural-append-formula-rows` at `3.042x`.
- The shareable benchmark card is generated from the checked-in artifact:
  [`workpaper-benchmark-card.png`](https://github.com/proompteng/bilig/blob/main/docs/assets/workpaper-benchmark-card.png).
- Read the
  [compatibility limits](https://github.com/proompteng/bilig/blob/main/docs/where-bilig-is-not-excel-compatible-yet.md)
  before importing real Excel workbooks.
- For XLSX accuracy audits, use the
  [Excel oracle harness](https://github.com/proompteng/bilig/blob/main/docs/xlsx-corpus-verifier-walkthrough.md#run-the-excel-oracle-harness).
  It separates import success, timeouts, stale cached formula values, and fresh
  Microsoft Excel recalculation results.
- Open benchmark critique lives in
  [Discussion 340](https://github.com/proompteng/bilig/discussions/340).

If the sanity check matches a problem you have, star or bookmark the project:
<https://github.com/proompteng/bilig/stargazers>.

## Production Status

Use this package for documented WorkPaper workflows: programmatic workbook
creation, formula evaluation, structural edits, persistence round trips,
service-side spreadsheet automation, and agent-driven workbook operations.

Current release posture:

- The contract is the WorkPaper/headless API exported by this package.
- Excel-file ingestion belongs to import/export pipelines before data reaches
  `WorkPaper`.
- Use `WorkPaper.buildFromSnapshot()` for importer-produced workbook snapshots
  so Excel defined names, tables, and translated formulas stay attached to the
  runtime model.
- Custom function plugins and callback hooks are runtime registrations; persist
  workbook data, then register custom behavior in application code before
  restore.
- Recent hardening covered config rebuilds, move-range bounds, persisted
  document validation, and benchmark gates.

## Compatibility Notes

- The facade follows HyperFormula-style workbook workflows, but it is not
  byte-for-byte compatible with HyperFormula.
- Public lookup helpers such as `getSheetId()`, `getSheetName()`,
  `simpleCellAddressFromString()`, and named-expression reads return
  `undefined` on misses.
- `@bilig/headless` exposes `onDetailed()`, `onceDetailed()`, and
  `offDetailed()` for detailed event payloads.
- Stable compatibility adapters are available through `graph`, `rangeMapping`,
  `arrayMapping`, `sheetMapping`, `addressMapping`, `dependencyGraph`,
  `evaluator`, `columnSearch`, and `lazilyTransformingAstService`.
- Financial date formulas such as `XIRR()` and `XNPV()` accept numeric Excel
  serial dates. Text date strings are not coerced in headless formulas.

## Validation Commands

For a headless-only code change, start here:

```sh
pnpm exec vitest run \
  packages/headless/src/__tests__/work-paper-runtime.test.ts \
  packages/headless/src/__tests__/work-paper-parity.test.ts \
  packages/headless/src/__tests__/persistence.test.ts \
  packages/headless/src/__tests__/persistence.fuzz.test.ts

pnpm --filter @bilig/headless build
```

Before publishing or claiming production readiness:

```sh
pnpm publish:runtime:check
pnpm workpaper:bench:competitive:check
pnpm run ci
```

Regenerate the competitive benchmark artifact only when intentionally updating
benchmark evidence:

```sh
pnpm workpaper:bench:competitive:generate
pnpm workpaper:bench:competitive:check
```

Do not change benchmark definitions, scoring, sampling, or workload sizes to
hide losses.

## For Coding Agents

Start here when Codex, Claude Code, or another agent is modifying or consuming
this package:

1. Read this README and the root
   [`README.md`](https://github.com/proompteng/bilig/blob/main/README.md)
   first.
2. Use public exports from `@bilig/headless`; do not import from `src/`,
   `dist/internal`, or `@bilig/core` unless the task is package-internal engine
   work.
3. Use zero-based `{ sheet, row, col }` addresses and resolve sheet ids with
   `getSheetId()`.
4. Use `WorkPaper.buildFromSheets()` for hand-authored fixtures,
   `WorkPaper.buildFromSnapshot()` for importer-produced snapshots, and
   `exportWorkPaperDocument()` / `createWorkPaperFromDocument()` for persistence
   round trips.
5. Do not treat embedded XLSX cached formula values as an accuracy oracle.
6. Add or tighten regression tests before changing config rebuilds, range
   bounds, formulas, persistence, events, row/column moves, or sheet lifecycle.
7. Run focused headless tests before broader gates.
8. Preserve benchmark definitions and workload sizes.
9. Document unsupported behavior honestly instead of implying full Excel
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

## More Guides

When the sanity check passes, these are the next useful pages.

- Service workflows:
  [server-side spreadsheet automation](https://github.com/proompteng/bilig/blob/main/docs/server-side-spreadsheet-automation-node.md),
  [Node service recipe](https://github.com/proompteng/bilig/blob/main/docs/node-service-workpaper-recipe.md),
  [serverless API route recipe](https://github.com/proompteng/bilig/blob/main/docs/serverless-workpaper-api-route.md),
  [CSV-shaped input recipe](https://github.com/proompteng/bilig/blob/main/docs/csv-shaped-workpaper-input-recipe.md),
  [workbook automation examples](https://github.com/proompteng/bilig/blob/main/docs/workbook-automation-examples-node.md),
  and [framework adapters](https://github.com/proompteng/bilig/blob/main/docs/node-framework-workpaper-adapters.md).
- Agent and MCP workflows:
  [agent tool-calling recipe](https://github.com/proompteng/bilig/blob/main/docs/agent-workpaper-tool-calling-recipe.md),
  [OpenAI Responses guide](https://github.com/proompteng/bilig/blob/main/docs/openai-responses-workpaper-tool-call.md),
  [AI SDK, LangChain, and agent framework guide](https://github.com/proompteng/bilig/blob/main/docs/vercel-ai-sdk-langchain-spreadsheet-tool.md),
  [MCP server guide](https://github.com/proompteng/bilig/blob/main/docs/mcp-workpaper-tool-server.md),
  [MCP directory page](https://github.com/proompteng/bilig/blob/main/docs/mcp-spreadsheet-server-directory.md),
  [MCP client setup](https://github.com/proompteng/bilig/blob/main/docs/mcp-client-setup.md),
  and [Claude Desktop MCPB bundle](https://github.com/proompteng/bilig/blob/main/docs/claude-desktop-mcpb-workpaper.md).
- Choosing the stack:
  [screenshot automation boundary](https://github.com/proompteng/bilig/blob/main/docs/stop-driving-spreadsheets-with-screenshots.md),
  [Node spreadsheet formula engine](https://github.com/proompteng/bilig/blob/main/docs/node-spreadsheet-formula-engine.md),
  [Google Sheets API boundary](https://github.com/proompteng/bilig/blob/main/docs/google-sheets-api-alternative-node-workpaper.md),
  [docs/javascript-spreadsheet-library-headless-node.md](https://github.com/proompteng/bilig/blob/main/docs/javascript-spreadsheet-library-headless-node.md),
  [headless spreadsheet engine for Node services and agents](https://github.com/proompteng/bilig/blob/main/docs/headless-spreadsheet-engine-node-services-agents.md),
  [ExcelJS formula recalculation in Node.js](https://github.com/proompteng/bilig/blob/main/docs/exceljs-formula-recalculation-node.md),
  [docs/sheetjs-exceljs-alternative-formula-workbook-api.md](https://github.com/proompteng/bilig/blob/main/docs/sheetjs-exceljs-alternative-formula-workbook-api.md),
  [headless engine comparison](https://github.com/proompteng/bilig/blob/main/docs/headless-spreadsheet-engine-comparison.md),
  and [HyperFormula comparison](https://github.com/proompteng/bilig/blob/main/docs/hyperformula-alternative-headless-workpaper.md).
- Accuracy and compatibility:
  [compatibility boundaries](https://github.com/proompteng/bilig/blob/main/docs/where-bilig-is-not-excel-compatible-yet.md),
  [XLSX corpus verifier walkthrough](https://github.com/proompteng/bilig/blob/main/docs/xlsx-corpus-verifier-walkthrough.md),
  [local benchmark walkthrough](https://github.com/proompteng/bilig/blob/main/docs/local-workpaper-benchmark-walkthrough.md),
  and [benchmark proof note](https://github.com/proompteng/bilig/blob/main/docs/what-workpaper-benchmark-proves.md).
- Formula edge cases:
  [XLOOKUP exact fixture](https://github.com/proompteng/bilig/blob/main/docs/formula-edge-xlookup-exact-fixture.md),
  [SUMIFS paired criteria fixture](https://github.com/proompteng/bilig/blob/main/docs/formula-edge-sumifs-paired-criteria-fixture.md),
  and [GROUPBY spill fixture](https://github.com/proompteng/bilig/blob/main/docs/formula-edge-groupby-spill-fixture.md).

## Stay Connected

- Website: <https://proompteng.github.io/bilig/>
- GitHub: <https://github.com/proompteng/bilig>
- npm: <https://www.npmjs.com/package/@bilig/headless>
- Star or bookmark: <https://github.com/proompteng/bilig/stargazers>
- Ask a workflow question: <https://github.com/proompteng/bilig/discussions/157>
- Share service examples: <https://github.com/proompteng/bilig/discussions/213>
- Discuss persistence adapters:
  <https://github.com/proompteng/bilig/discussions/307>
- Discuss JavaScript spreadsheet library positioning:
  <https://github.com/proompteng/bilig/discussions/308>
- Discuss OpenAI Responses tool calls:
  <https://github.com/proompteng/bilig/discussions/335>
- Discuss benchmark fairness: <https://github.com/proompteng/bilig/discussions/340>
- Pick a scoped first patch:
  [starter issues](https://github.com/proompteng/bilig/blob/main/docs/starter-issues.md)
  or
  [first-timers-only issues](https://github.com/proompteng/bilig/issues?q=is%3Aissue%20state%3Aopen%20label%3Afirst-timers-only).

## Versioning

`@bilig/headless` ships as part of the aligned bilig runtime package set. Treat
documented public exports as the supported surface, keep integration tests around
your own workbook corpus, and rerun the validation gates before upgrading in
production.
