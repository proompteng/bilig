---
title: Spreadsheet APIs for AI agents
published: true
description: Why coding agents should edit workbook formulas through a Node.js WorkPaper API instead of spreadsheet screenshots, with a runnable @bilig/headless example.
tags: ai agents, spreadsheet, node, workpaper, typescript
canonical_url: https://proompteng.github.io/bilig/why-agents-need-workbook-apis.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Why Agents Need Workbook APIs, Not Spreadsheet Screenshots

Agents can click through a spreadsheet UI. That does not make the UI a good
runtime boundary.

Screenshots hide formulas, make structural edits ambiguous, and turn
verification into a visual guess. If the workflow depends on workbook state, the
agent needs an API that can read formulas, write cells, recalculate, and prove
what changed.

`@bilig/headless` is the public WorkPaper runtime from `bilig`. It is a
TypeScript package for Node services, coding agents, and local workbook
automation that need spreadsheet behavior without opening a browser grid.

## The Problem With Screen-Driven Spreadsheets

Spreadsheets are useful because they combine a document model, formulas,
structural editing, validation, and persistence. A grid UI is only one way to
operate that model.

When an agent drives the UI directly, several useful facts become hard to
trust:

- whether the displayed value came from a literal or a formula
- whether a formula was recalculated after an edit
- whether an inserted row moved the intended references
- whether hidden sheets, named expressions, or persisted state changed
- whether the agent verified the workbook or just saw plausible pixels

Screenshots are still useful for final human inspection. They should not be the
main contract between an agent and workbook logic.

## The Better Boundary

A workbook API gives the agent explicit operations and explicit readback:

- create sheets and cells from data
- write formulas as formulas
- evaluate values after changes
- apply structural edits through the same model used by the engine
- export and restore persisted workbook documents
- test the behavior without launching a browser

That shape fits backend jobs, coding-agent tools, and local-first workflows
better than asking a model to infer state from a rendered grid.

## Run The Maintained Eval First

The shortest trial starts from an empty Node project and uses the published npm
package:

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

The byte count can change between versions. The important check is that
`verified` is `true` and the restored workbook returns the same calculated
value as the edited workbook.

## What The API Looks Like

The public package is installable as a normal Node dependency:

```sh
npm install @bilig/headless
```

Build a workbook, change one input, read the recalculated value, and restore the
saved document:

```ts
import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
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
  throw new Error('Workbook did not create the expected sheets')
}

const before = workbook.getCellValue({ sheet: summary, row: 1, col: 1 })
workbook.setCellContents({ sheet: inputs, row: 1, col: 1 }, 32)
const after = workbook.getCellValue({ sheet: summary, row: 1, col: 1 })

const saved = serializeWorkPaperDocument(exportWorkPaperDocument(workbook, { includeConfig: true }))
const restored = createWorkPaperFromDocument(parseWorkPaperDocument(saved))
const restoredSummary = restored.getSheetId('Summary')
if (restoredSummary === undefined) {
  throw new Error('Restored workbook did not create the Summary sheet')
}

console.log({
  before,
  after,
  afterRestore: restored.getCellValue({ sheet: restoredSummary, row: 1, col: 1 }),
  sheets: restored.getSheetNames(),
})
```

The maintained external-consumer example is in
[`examples/headless-workpaper`](../examples/headless-workpaper).

## MCP In 30 Seconds

If the agent already supports MCP, skip the TypeScript wrapper and start the
published stdio server in file-backed mode:

```sh
npm exec --package @bilig/headless@0.40.31 -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable
```

That command creates `pricing.workpaper.json` only when it is missing, exposes
typed workbook tools, and persists `set_cell_contents` edits back to the same
WorkPaper JSON file.

Expected tools:

- `list_sheets`
- `read_range`
- `read_cell`
- `set_cell_contents`
- `get_cell_display_value`
- `export_workpaper_document`
- `validate_formula`

Ask the agent for this proof object after an edit:

```json
{
  "editedCell": "Inputs!B3",
  "before": {
    "Summary!B3": 60000
  },
  "after": {
    "Summary!B3": 96000
  },
  "checks": {
    "formulaReadbackChanged": true,
    "exportedWorkPaperDocument": true,
    "restoredMatchesAfter": true
  }
}
```

The important distinction is that the agent is not reporting “clicked” or
“updated.” It is reporting the exact edited cell, the recalculated dependent
value, and persistence evidence.

## What This Enables

For an agent tool, a WorkPaper API can expose a small set of reliable commands:

- `buildWorkbookFromSheets`
- `setCellContents`
- `getCellValue`
- `listSheets`
- `exportWorkbookDocument`
- `restoreWorkbookDocument`

Those commands produce deterministic outputs that can be tested, logged, and
replayed. The rendered spreadsheet can stay a human-facing view instead of the
agent's source of truth.

For a Node service, the same model supports formula-backed business logic
without bundling a spreadsheet application into the service path.

## What Bilig Does Not Claim

`bilig` is not a finished Excel clone. It does not claim full Excel formula
parity. It does not claim every benchmark p95 row is faster than HyperFormula.

The current public claim is narrower: `@bilig/headless` exposes a WorkPaper API
for programmatic workbook creation, formulas, structural operations,
persistence, and checked-in benchmark evidence. The evidence note records
WorkPaper `100/100` mean wins on scorecard-eligible comparable workloads against
HyperFormula-style workloads, with the p95 evidence documented separately.

Read the benchmark note here:
[`docs/headless-workpaper-benchmark-evidence.md`](headless-workpaper-benchmark-evidence.md).

## Try It

- GitHub: <https://github.com/proompteng/bilig>
- Website: <https://proompteng.github.io/bilig/>
- npm: <https://www.npmjs.com/package/@bilig/headless>
- Empty-directory eval:
  <https://proompteng.github.io/bilig/try-bilig-headless-in-node.html>
- Runnable example: [`examples/headless-workpaper`](../examples/headless-workpaper)

If this is relevant to an agent or Node workflow, star the repo as a bookmark:
<https://github.com/proompteng/bilig/stargazers>
