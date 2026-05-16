---
title: Why agents need workbook APIs instead of spreadsheet screenshots
published: true
description: A practical case for treating spreadsheet state as an API boundary for Node services and coding agents.
tags: typescript, node, opensource, ai
canonical_url: https://proompteng.github.io/bilig/dev-to-workbook-apis-post.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

Spreadsheets are still one of the most common ways teams encode business logic.
Pricing models, revenue plans, reconciliations, forecasts, data checks, and
ad hoc operational tools often start in a workbook because the shape is flexible
and everyone can inspect it.

That creates an awkward automation problem: the logic is structured like a
workbook, but most programmatic workflows either drive a browser grid or rewrite
the model in application code.

That is especially brittle for coding agents. A screenshot can show a grid, but
it cannot give the agent a stable contract for formulas, structural edits,
persistence, validation, or post-write readback.

I maintain an open-source TypeScript project called
[`bilig`](https://github.com/proompteng/bilig). The public package is
[`@bilig/headless`](https://www.npmjs.com/package/@bilig/headless), a
Node-facing workbook runtime for programmatic spreadsheet automation.

This post is the practical argument for the API boundary. It is not a claim that
the project is a finished Excel clone.

## Screenshots are a weak runtime boundary

A spreadsheet UI is useful for humans. It is not enough as the main contract for
an agent.

When an agent drives a grid by pixels, several facts become hard to verify:

- whether the displayed value came from a literal or a formula
- whether formulas recalculated after an edit
- whether an inserted row moved the intended references
- whether hidden sheets or named expressions changed
- whether the saved workbook still round-trips into the same state
- whether the agent actually verified the workbook or just saw plausible cells

Screenshots are still useful for final inspection. They should not be the only
evidence that a workflow did the right thing.

## The better boundary is workbook state

A workbook API gives an agent explicit operations and explicit readback:

- create sheets and ranges from data
- write formulas as formulas
- evaluate values after edits
- apply structural operations through the engine model
- serialize and restore workbook documents
- validate behavior without opening a browser

That shape fits backend jobs and agent tools better than asking a model to infer
state from a rendered grid.

The UI can stay a human-facing view. The API becomes the source of truth.

## Run the maintained eval first

The shortest trial starts from an empty Node project and uses the published npm
package:

```sh
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
`verified` is `true` and the restored workbook returns the same calculated value
as the edited workbook.

## What the API looks like

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

That is the basic pattern I want from agent-facing spreadsheet tools:

1. create or load workbook state
2. make the edit
3. recalculate through the workbook engine
4. read back exact values
5. persist and restore
6. verify the same behavior without a browser session

## What this gives an agent

An agent tool can expose a small set of reliable commands:

- `buildWorkbookFromSheets`
- `setCellContents`
- `getCellValue`
- `readRange`
- `listSheets`
- `exportWorkbookDocument`
- `restoreWorkbookDocument`

Those commands produce deterministic outputs that can be logged, tested, and
replayed.

The important shift is that the agent no longer has to treat the grid as the
database. It can operate on workbook state and use the rendered spreadsheet only
when a human needs to inspect the result.

## What bilig currently claims

The current public claim is intentionally narrow:

`@bilig/headless` exposes a WorkPaper API for programmatic workbook creation,
formula evaluation, structural operations, persistence, validation, and
readback.

The repo also includes checked-in benchmark evidence against
HyperFormula-style workloads. The current artifact records `46/57` mean wins on
scorecard-eligible comparable workloads, with a p95 caveat documented instead
of hidden.

benchmark evidence:
<https://github.com/proompteng/bilig/blob/main/docs/headless-workpaper-benchmark-evidence.md>

What it does not claim:

- not full Excel compatibility
- not a finished spreadsheet application
- not faster on every p95 row
- not complete xlsx import/export fidelity for every workbook in the wild

I think those caveats matter. Developer infrastructure gets more trust from
clear boundaries than from broad compatibility claims.

## Try it

published package:

```sh
npm install @bilig/headless
```

Empty-directory eval:

<https://proompteng.github.io/bilig/try-bilig-headless-in-node.html>

useful links:

- repo: <https://github.com/proompteng/bilig>
- npm: <https://www.npmjs.com/package/@bilig/headless>
- package readme:
  <https://github.com/proompteng/bilig/tree/main/packages/headless#readme>
- runnable example:
  <https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper>
- compatibility caveats:
  <https://github.com/proompteng/bilig/blob/main/docs/where-bilig-is-not-excel-compatible-yet.md>
- starter issues:
  <https://github.com/proompteng/bilig/blob/main/docs/starter-issues.md>

If this is relevant to a Node service, spreadsheet engine, or coding-agent
workflow you are building, the most useful feedback is concrete: API friction,
missing formula semantics, import/export expectations, or a real workbook case
that should become a fixture.
