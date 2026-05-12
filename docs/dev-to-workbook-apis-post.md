---
title: Why agents need workbook APIs instead of spreadsheet screenshots
published: true
description: A practical case for treating spreadsheet state as an API boundary for Node services and coding agents.
tags: typescript, node, opensource, ai
canonical_url: https://proompteng.github.io/bilig/dev-to-workbook-apis-post.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

spreadsheets are still one of the most common ways teams encode business logic.
pricing models, revenue plans, reconciliations, forecasts, data checks, and
ad hoc operational tools often start in a workbook because the shape is flexible
and everyone can inspect it.

that creates an awkward automation problem: the logic is structured like a
workbook, but most programmatic workflows either drive a browser grid or rewrite
the model in application code.

that is especially brittle for coding agents. a screenshot can show a grid, but
it cannot give the agent a stable contract for formulas, structural edits,
persistence, validation, or post-write readback.

i maintain an open-source typescript project called
[`bilig`](https://github.com/proompteng/bilig). the public package is
[`@bilig/headless`](https://www.npmjs.com/package/@bilig/headless), a
node-facing workbook runtime for programmatic spreadsheet automation.

this post is the practical argument for the API boundary. it is not a claim that
the project is a finished excel clone.

## screenshots are a weak runtime boundary

a spreadsheet UI is useful for humans. it is not enough as the main contract for
an agent.

when an agent drives a grid by pixels, several facts become hard to verify:

- whether the displayed value came from a literal or a formula
- whether formulas recalculated after an edit
- whether an inserted row moved the intended references
- whether hidden sheets or named expressions changed
- whether the saved workbook still round-trips into the same state
- whether the agent actually verified the workbook or just saw plausible cells

screenshots are still useful for final inspection. they should not be the only
evidence that a workflow did the right thing.

## the better boundary is workbook state

a workbook API gives an agent explicit operations and explicit readback:

- create sheets and ranges from data
- write formulas as formulas
- evaluate values after edits
- apply structural operations through the engine model
- serialize and restore workbook documents
- validate behavior without opening a browser

that shape fits backend jobs and agent tools better than asking a model to infer
state from a rendered grid.

the UI can stay a human-facing view. the API becomes the source of truth.

## a small node example

install the package:

```sh
npm install @bilig/headless
```

build a workbook, evaluate a formula, and round-trip the document:

```js
import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from "@bilig/headless";

const workbook = WorkPaper.buildFromSheets({
  Revenue: [
    ["Region", "Customers", "ARPA", "Revenue"],
    ["West", 20, 1200, "=B2*C2"],
    ["East", 30, 250, "=B3*C3"],
    ["Central", 18, 300, "=B4*C4"],
  ],
  Summary: [
    ["Metric", "Value"],
    ["Total revenue", "=SUM(Revenue!D2:D4)"],
    ["West customers", '=SUMIF(Revenue!A2:A4,"West",Revenue!B2:B4)'],
  ],
});

const summary = workbook.getSheetId("Summary");
if (summary === undefined) {
  throw new Error("Summary sheet was not created");
}

const total = workbook.getCellValue({ sheet: summary, row: 1, col: 1 });
const saved = serializeWorkPaperDocument(
  exportWorkPaperDocument(workbook, { includeConfig: true }),
);
const restored = createWorkPaperFromDocument(parseWorkPaperDocument(saved));

console.log({
  total,
  sheets: restored.getSheetNames(),
});
```

that is the basic pattern i want from agent-facing spreadsheet tools:

1. create or load workbook state
2. make the edit
3. recalculate through the workbook engine
4. read back exact values
5. persist and restore
6. verify the same behavior without a browser session

## what this gives an agent

an agent tool can expose a small set of reliable commands:

- `buildWorkbookFromSheets`
- `setCellContents`
- `getCellValue`
- `readRange`
- `listSheets`
- `exportWorkbookDocument`
- `restoreWorkbookDocument`

those commands produce deterministic outputs that can be logged, tested, and
replayed.

the important shift is that the agent no longer has to treat the grid as the
database. it can operate on workbook state and use the rendered spreadsheet only
when a human needs to inspect the result.

## what bilig currently claims

the current public claim is intentionally narrow:

`@bilig/headless` exposes a WorkPaper API for programmatic workbook creation,
formula evaluation, structural operations, persistence, validation, and
readback.

the repo also includes checked-in benchmark evidence against
hyperformula-style workloads. the current artifact records `46/46` mean wins on
scorecard-eligible comparable workloads, with a p95 caveat documented instead
of hidden.

benchmark evidence:
<https://github.com/proompteng/bilig/blob/main/docs/headless-workpaper-benchmark-evidence.md>

what it does not claim:

- not full excel compatibility
- not a finished spreadsheet application
- not faster on every p95 row
- not complete xlsx import/export fidelity for every workbook in the wild

i think those caveats matter. developer infrastructure gets more trust from
clear boundaries than from broad compatibility claims.

## try it

published package:

```sh
npm install @bilig/headless
```

maintained external example:

```sh
git clone https://github.com/proompteng/bilig.git
cd bilig/examples/headless-workpaper
npm install
npm start
```

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

if this is relevant to a node service, spreadsheet engine, or coding-agent
workflow you are building, the most useful feedback is concrete: api friction,
missing formula semantics, import/export expectations, or a real workbook case
that should become a fixture.
