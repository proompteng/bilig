# Why Agents Need Workbook APIs, Not Spreadsheet Screenshots

AI agents can click through a spreadsheet UI, but screenshots are a weak
runtime boundary. They hide formulas, make structural edits ambiguous, and turn
verification into a visual guess. If a workflow depends on workbook state, the
agent should operate on a workbook API.

`@bilig/headless` is the small public wedge of `bilig`: a TypeScript WorkPaper
runtime for Node services, coding agents, and local workbook automation. It is
for cases where a program needs spreadsheet behavior without opening a browser
grid.

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

## A Small Example

The public package is installable as a normal Node dependency:

```sh
pnpm add @bilig/headless
```

Build a workbook, read a formula result, and persist the document:

```ts
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
  ],
  Summary: [
    ['Metric', 'Value'],
    ['Total revenue', '=SUM(Revenue!D2:D3)'],
  ],
})

const summary = workbook.getSheetId('Summary')
if (summary === undefined) {
  throw new Error('Summary sheet was not created')
}

const total = workbook.getCellValue({ sheet: summary, row: 1, col: 1 })
const saved = serializeWorkPaperDocument(exportWorkPaperDocument(workbook, { includeConfig: true }))
const restored = createWorkPaperFromDocument(parseWorkPaperDocument(saved))

console.log({
  total,
  sheets: restored.getSheetNames(),
})
```

The maintained external-consumer example is in
[`examples/headless-workpaper`](../examples/headless-workpaper).

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
WorkPaper `46/46` mean wins on scorecard-eligible comparable workloads against
HyperFormula-style workloads, with the p95 nuance documented separately.

Read the benchmark note here:
[`docs/headless-workpaper-benchmark-evidence.md`](headless-workpaper-benchmark-evidence.md).

## Try It

- GitHub: <https://github.com/proompteng/bilig>
- Website: <https://proompteng.github.io/bilig/>
- npm: <https://www.npmjs.com/package/@bilig/headless>
- Evaluation kit: [`docs/public-adoption-kit.md`](public-adoption-kit.md)
- Runnable example: [`examples/headless-workpaper`](../examples/headless-workpaper)

If this is relevant to an agent or Node workflow, star the repo as a bookmark:
<https://github.com/proompteng/bilig/stargazers>
