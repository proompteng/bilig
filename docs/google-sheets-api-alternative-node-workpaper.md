---
title: Google Sheets API alternative for local Node workbook execution
published: true
description: Decide when to use Google Sheets API and when a local @bilig/headless WorkPaper is a better fit for formula execution, verified readback, and JSON persistence in Node services.
tags: google sheets api, node, spreadsheet, formulas, workbook automation
canonical_url: https://proompteng.github.io/bilig/google-sheets-api-alternative-node-workpaper.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Google Sheets API alternative for local Node workbook execution

If a real spreadsheet with sharing, permissions, comments, and a URL is the
product, use Google Sheets API.

If the spreadsheet logic belongs inside a Node service, queue worker, CLI, or
agent tool, use `@bilig/headless`. Keep the workbook local, write inputs, read
calculated cells, and persist the WorkPaper document as JSON.

That is the boundary. `bilig` is not trying to replace Google Sheets. It is for
code that needs sheet-shaped business logic without turning a hosted
spreadsheet into the system of record.

## Quick decision

| You need | Start with |
| --- | --- |
| People editing the same hosted spreadsheet | Google Sheets |
| OAuth, spreadsheet IDs, A1 ranges, and Google Workspace permissions | Google Sheets API |
| A Node service that owns workbook state and formula execution | `@bilig/headless` |
| An agent tool that edits a cell and returns checked readback | `@bilig/headless` |
| An XLSX file for a person to open later | SheetJS, ExcelJS, or Excel automation |

Google describes the Sheets API as a REST interface for reading and modifying
spreadsheet data. Its values guide is built around the `spreadsheets.values`
resource, spreadsheet IDs, and ranges. That is the right shape when the
spreadsheet already lives in Google Sheets.

A WorkPaper is smaller. Your code owns the workbook object. It does not need
OAuth or a network round trip to recalculate a dependent cell.

## TypeScript smoke test

Run this from an empty Node project:

```sh
mkdir bilig-google-sheets-api-boundary
cd bilig-google-sheets-api-boundary
npm init -y
npm pkg set type=module
npm install @bilig/headless
npm install -D tsx typescript @types/node
cat > eval.ts <<'EOF'
import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from "@bilig/headless";

type NumericCell = {
  value: number;
};

function numberValue(cell: unknown, label: string): number {
  if (typeof cell === "object" && cell !== null && typeof (cell as NumericCell).value === "number") {
    return (cell as NumericCell).value;
  }

  throw new Error(`expected ${label} to be numeric, got ${JSON.stringify(cell)}`);
}

const workbook = WorkPaper.buildFromSheets({
  Inputs: [
    ["Metric", "Value"],
    ["Units", 300],
    ["Price", 19],
    ["Discount", 0.1],
  ],
  Summary: [
    ["Metric", "Value"],
    ["Net revenue", "=Inputs!B2*Inputs!B3*(1-Inputs!B4)"],
  ],
});

const inputs = workbook.getSheetId("Inputs");
const summary = workbook.getSheetId("Summary");
if (inputs === undefined || summary === undefined) {
  throw new Error("missing sheet");
}

const before = numberValue(workbook.getCellValue({ sheet: summary, row: 1, col: 1 }), "before revenue");
workbook.setCellContents({ sheet: inputs, row: 1, col: 1 }, 420);

const after = numberValue(workbook.getCellValue({ sheet: summary, row: 1, col: 1 }), "after revenue");
const saved = serializeWorkPaperDocument(exportWorkPaperDocument(workbook, { includeConfig: true }));
const restored = createWorkPaperFromDocument(parseWorkPaperDocument(saved));
const restoredSummary = restored.getSheetId("Summary");
if (restoredSummary === undefined) {
  throw new Error("missing restored Summary sheet");
}

const afterRestore = numberValue(
  restored.getCellValue({ sheet: restoredSummary, row: 1, col: 1 }),
  "restored revenue",
);

console.log(
  JSON.stringify(
    {
      before,
      after,
      afterRestore,
      verified: before === 5130 && after === 7182 && afterRestore === 7182,
    },
    null,
    2,
  ),
);

workbook.dispose();
restored.dispose();
EOF
npx tsx eval.ts
```

Expected output:

```json
{
  "before": 5130,
  "after": 7182,
  "afterRestore": 7182,
  "verified": true
}
```

That proves the local backend loop: build workbook state, change one input, read
the recalculated value, save JSON, restore the workbook, and read the same value
again.

## Use Google Sheets API when

- users already collaborate in a Google Sheet
- the spreadsheet URL is the durable artifact
- Google permissions are part of the product
- your app reads or writes ranges in an existing Google Sheet
- reviewers expect to inspect the result in Google Workspace

Do not move that workflow to `bilig` just to avoid an API call.

## Use a local WorkPaper when

- a route handler, job, CLI, or coding agent owns the flow
- the workbook is service state, not a shared document
- tests should run without network credentials
- the app needs checked before/after values after one input changes
- formula-backed state should sit beside the rest of the service data

The useful unit is a state transition your code can test: input changed,
formulas recalculated, output read, document saved.

## Related paths

- [Server-side spreadsheet automation in Node.js](server-side-spreadsheet-automation-node.md)
- [Node service WorkPaper recipe](node-service-workpaper-recipe.md)
- [Persist formula-backed WorkPaper documents in Node](persisting-formula-backed-workpaper-documents-in-node.md)
- [Agent spreadsheet tool-call loop](agent-spreadsheet-tool-call-loop.md)
- [MCP spreadsheet tool server](mcp-workpaper-tool-server.md)
- [Where bilig is not Excel-compatible yet](where-bilig-is-not-excel-compatible-yet.md)

If this boundary saves you a Google Sheets automation spike, star the repository
so the next backend developer can find it:
<https://github.com/proompteng/bilig/stargazers>.

## Sources

- Google Sheets API overview:
  <https://developers.google.com/workspace/sheets/api/guides/concepts>
- Google Sheets API read and write values guide:
  <https://developers.google.com/workspace/sheets/api/guides/values>
