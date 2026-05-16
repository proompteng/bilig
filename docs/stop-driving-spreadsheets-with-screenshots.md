---
title: Stop driving spreadsheets with screenshots. Run formula workbooks in Node.
description: A practical argument for replacing spreadsheet screen automation with formula-backed WorkPaper APIs in backend services and agent tools.
canonical_url: https://proompteng.github.io/bilig/stop-driving-spreadsheets-with-screenshots.html
image: /assets/github-social-preview.png
---

# Stop driving spreadsheets with screenshots. Run formula workbooks in Node.

Research date: 2026-05-16.

Spreadsheets are still where teams keep a lot of operational logic: pricing
rules, revenue models, quote approvals, capacity plans, billing checks, and
import validation. The problem is not that spreadsheets exist. The problem is
that automation often treats the grid as pixels instead of as state.

If a backend job or coding agent has to click cells, infer formulas from a
rendered view, and trust a screenshot after the edit, the verification boundary
is weak. The automation can look right while still failing to prove the formula
actually recalculated, the workbook state persisted, or a later restore returns
the same computed value.

`@bilig/headless` is built around the smaller claim: keep spreadsheet-shaped
business logic in a workbook model, but run it through a TypeScript API in Node
services and agent tools.

## The Failure Mode

Screenshot-driven spreadsheet automation is brittle because the visible grid is
not the whole workbook contract.

A serious workflow usually needs to answer these questions:

- Which cells were changed?
- Which cells are formulas rather than literals?
- Did dependent formulas recalculate after the edit?
- Did the value survive export and restore?
- Can the same operation run in CI without a browser session?
- Can an agent return exact readback instead of a plausible visual answer?

Those are state questions. A screenshot can help a human inspect the final
shape, but it should not be the only proof that a calculation is correct.

## The API Boundary

A WorkPaper gives code explicit operations:

- build sheets from arrays or records
- write cells and formulas
- read calculated values and display values
- export a JSON workbook document
- restore that document and re-read the result
- expose the same operations as agent or MCP tools

That makes the workbook a reviewable calculation artifact instead of a browser
grid that an automation script has to push around.

## Run The Proof

This quickstart uses the published npm package, not source imports from the
repository.

```sh
mkdir bilig-workpaper-proof
cd bilig-workpaper-proof
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
  "verified": true
}
```

The exact serialized byte count can change between releases. The important
part is `verified: true`: the edited workbook and restored workbook agree on the
calculated output.

## Minimal TypeScript

```ts
import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from "@bilig/headless";

const workbook = WorkPaper.buildFromSheets({
  Inputs: [
    ["Metric", "Value"],
    ["Seats", 25],
    ["Price", 147],
  ],
  Summary: [
    ["Metric", "Value"],
    ["Total", "=Inputs!B2*Inputs!B3"],
  ],
});

const inputs = workbook.getSheetId("Inputs");
const summary = workbook.getSheetId("Summary");
if (inputs === undefined || summary === undefined) {
  throw new Error("Workbook did not create the expected sheets");
}

const before = readNumber(workbook.getCellValue({ sheet: summary, row: 1, col: 1 }));
workbook.setCellContents({ sheet: inputs, row: 1, col: 1 }, 40);
const after = readNumber(workbook.getCellValue({ sheet: summary, row: 1, col: 1 }));

const saved = serializeWorkPaperDocument(
  exportWorkPaperDocument(workbook, { includeConfig: true }),
);
const restored = createWorkPaperFromDocument(parseWorkPaperDocument(saved));
const restoredSummary = restored.getSheetId("Summary");
if (restoredSummary === undefined) {
  throw new Error("Restored workbook did not create the Summary sheet");
}

const afterRestore = readNumber(
  restored.getCellValue({
    sheet: restoredSummary,
    row: 1,
    col: 1,
  }),
);

console.log({
  before,
  after,
  afterRestore,
  verified: after === afterRestore,
});

function readNumber(cell: unknown): number {
  if (
    typeof cell === "object" &&
    cell !== null &&
    typeof (cell as { value: unknown }).value === "number"
  ) {
    return (cell as { value: number }).value;
  }
  if (typeof cell === "number") {
    return cell;
  }
  throw new Error(`Expected numeric cell value, got ${JSON.stringify(cell)}`);
}
```

Expected output:

```json
{
  "before": 3675,
  "after": 5880,
  "afterRestore": 5880,
  "verified": true
}
```

## Where This Fits

Use a WorkPaper when the code owns the workflow:

- quote approval endpoints
- pricing and discount rules
- finance checks and payout validation
- import validation that needs formula readback
- agent tools that must prove the value after editing inputs
- MCP servers that need workbook state without opening Excel or Sheets

Keep Google Sheets or Excel when the primary job is human collaboration,
desktop workbook authoring, or full XLSX compatibility. `@bilig/headless` is a
formula-backed runtime boundary, not a finished Excel clone.

## Useful Next Pages

- [Try `@bilig/headless` in Node](try-bilig-headless-in-node.md)
- [Node service WorkPaper recipe](node-service-workpaper-recipe.md)
- [ExcelJS formula recalculation in Node.js](exceljs-formula-recalculation-node.md)
- [Agent tool-calling recipe](agent-workpaper-tool-calling-recipe.md)
- [MCP spreadsheet tool server](mcp-workpaper-tool-server.md)
- [Headless spreadsheet engine comparison](headless-spreadsheet-engine-comparison.md)

If this solves a workflow you have, the most useful signal is a star on the
repository or a concrete issue with the workbook shape you need:

<https://github.com/proompteng/bilig>
