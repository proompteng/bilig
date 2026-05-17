---
title: Why agents need workbook APIs instead of spreadsheet screenshots
published: true
description: A maintainer note about using workbook state, not screenshots, as the boundary for spreadsheet automation in Node.
tags: typescript, node, opensource, ai
canonical_url: https://proompteng.github.io/bilig/dev-to-workbook-apis-post.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

I keep running into the same spreadsheet automation failure.

The screenshot looks fine. The total changed. The agent says it edited the right
input. Then you ask the boring questions and the story falls apart:

- was that cell a literal or a formula?
- did dependent formulas recalculate?
- did the workbook save with the formulas intact?
- if I restore the saved state in CI, do I get the same answer?

That is the point where a screenshot stops being evidence.

I maintain [`bilig`](https://github.com/proompteng/bilig). The public package is
[`@bilig/headless`](https://www.npmjs.com/package/@bilig/headless). It is a
small TypeScript runtime for workbook-shaped business logic in Node services and
agent tools.

It is not an Excel clone. The promise is smaller: build or load a workbook,
write inputs, recalculate formulas, read the answer back, and save the state as
JSON.

## The bug

Spreadsheet UIs are good for people. They are a weak runtime boundary for code.

If an agent has to click cells and inspect pixels, it can easily produce a
plausible-looking result without proving the workbook state is right. The grid
does not tell you whether hidden formulas moved, whether a structural edit
retargeted references, or whether the saved document still round-trips.

For backend jobs and agent tools, the rendered spreadsheet should be inspection
only. The contract should be the workbook state.

## The boundary I want

A useful workbook API should let code do this without opening a browser:

1. build or load sheets
2. write typed values and formulas
3. recalculate after edits
4. read exact values back
5. export a workbook document
6. restore that document and check the same result again

That gives an agent a loggable operation instead of a screenshot and a shrug.

## Run The Small Proof

This starts from an empty directory and uses the published npm package:

```sh
mkdir bilig-workpaper-eval
cd bilig-workpaper-eval
npm init -y
npm pkg set type=module
npm install @bilig/headless
npm install -D tsx typescript @types/node
curl -fsSLo eval.ts https://proompteng.github.io/bilig/npm-eval.ts
npx tsx eval.ts
```

Expected shape:

```json
{
  "before": 24000,
  "after": 38400,
  "afterRestore": 38400,
  "sheets": ["Inputs", "Summary"],
  "bytes": 1000,
  "verified": true,
  "nextStep": "If this proof matches your service or agent workflow, star or bookmark Bilig: https://github.com/proompteng/bilig/stargazers"
}
```

The byte count can move between releases. The important part is
`"verified": true`: the script edits one input, reads the recalculated value,
saves WorkPaper JSON, restores it, and gets the same value again.

## What The API Looks Like

Here is the whole shape without the surrounding quickstart script:

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

That is the loop I want exposed to agents:

- make the edit
- read the formula output
- persist the workbook state
- restore it
- prove the value did not change

## Where it fits

This is useful when a workbook is really business logic:

- quote approval
- pricing and discount rules
- payout checks
- import validation
- finance sanity checks
- MCP or coding-agent tools that need read-after-write proof

It is a bad fit when the job is human collaboration, macros, chart fidelity, or
full desktop Excel behavior. For those jobs, use Excel, Google Sheets, or a
mature spreadsheet UI product.

It is also not automatically better than HyperFormula, ExcelJS, or SheetJS.
Those are the first tools to check when you mainly need a broad formula engine
or spreadsheet file reading/writing. Bilig is for the narrower case where the
Node code owns the workbook state and needs formula readback plus JSON
persistence.

## Current evidence

The benchmark claim is deliberately narrow. The checked artifact currently
records `76/100` mean-latency wins against HyperFormula-style comparable
workloads and `76/100` workloads winning both mean and p95. The p95 misses are
called out instead of hidden.

Benchmark note:
<https://github.com/proompteng/bilig/blob/main/docs/what-workpaper-benchmark-proves.md>

Compatibility limits:
<https://github.com/proompteng/bilig/blob/main/docs/where-bilig-is-not-excel-compatible-yet.md>

## Try it or reject it

The best feedback is a concrete rejection reason:

- a formula family that blocks you
- an XLSX case that has to round-trip
- a persistence shape that would be awkward in your service
- an API call you would not want to maintain
- a benchmark that would make the library easier to trust

Links:

- repo: <https://github.com/proompteng/bilig>
- npm: <https://www.npmjs.com/package/@bilig/headless>
- empty-directory eval:
  <https://proompteng.github.io/bilig/try-bilig-headless-in-node.html>
- runnable examples:
  <https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper>
- adoption blocker discussion:
  <https://github.com/proompteng/bilig/discussions/new?category=general>
