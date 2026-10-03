---
title: Server-side spreadsheet automation in Node.js
published: true
description: 'Automate spreadsheet formulas inside Node services with @bilig/workpaper: edit inputs, read calculated cells, and persist WorkPaper JSON without a browser grid.'
tags: typescript, node, spreadsheet, automation, opensource
canonical_url: https://proompteng.github.io/bilig/server-side-spreadsheet-automation-node.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Server-side spreadsheet automation in Node.js

Server-side spreadsheet automation is useful when the spreadsheet is logic, not
the user interface. A Node service may need to price a quote, check a budget,
score an import, or run a forecast using formulas that already exist in a
workbook-shaped model.

Use `@bilig/workpaper` for that middle case: the service owns a workbook object,
changes narrow input cells, reads calculated outputs, and stores the WorkPaper
document as JSON for the next request or job.

## The service boundary

Keep the automation boundary small:

1. Load or create a WorkPaper.
2. Apply one business input.
3. Read the dependent result through the workbook API.
4. Serialize the WorkPaper document.
5. Return the value and the checks that prove the edit mattered.

That is a better backend contract than screen scraping a grid or duplicating
spreadsheet formulas in application code.

## TypeScript smoke test

Run this from an empty Node project:

```sh
mkdir bilig-server-automation-eval
cd bilig-server-automation-eval
npm init -y
npm pkg set type=module
npm install @bilig/workpaper
npm install -D tsx typescript @types/node
cat > eval.ts <<'EOF'
import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from '@bilig/workpaper'

type NumericCell = {
  value: number
}

function numberValue(cell: unknown, label: string): number {
  if (typeof cell === 'object' && cell !== null && typeof (cell as NumericCell).value === 'number') {
    return (cell as NumericCell).value
  }

  throw new Error(`expected ${label} to be numeric, got ${JSON.stringify(cell)}`)
}

const workbook = WorkPaper.buildFromSheets({
  Inputs: [
    ['Metric', 'Value'],
    ['Customers', 25],
    ['ARPA', 140],
    ['Discount', 0.05],
  ],
  Summary: [
    ['Metric', 'Value'],
    ['Net revenue', '=Inputs!B2*Inputs!B3*(1-Inputs!B4)'],
  ],
})

const inputs = workbook.getSheetId('Inputs')
const summary = workbook.getSheetId('Summary')
if (inputs === undefined || summary === undefined) {
  throw new Error('missing sheet')
}

const before = numberValue(workbook.getCellValue({ sheet: summary, row: 1, col: 1 }), 'before revenue')
workbook.setCellContents({ sheet: inputs, row: 1, col: 1 }, 42)

const after = numberValue(workbook.getCellValue({ sheet: summary, row: 1, col: 1 }), 'after revenue')
const saved = serializeWorkPaperDocument(exportWorkPaperDocument(workbook, { includeConfig: true }))
const restored = createWorkPaperFromDocument(parseWorkPaperDocument(saved))
const restoredSummary = restored.getSheetId('Summary')
if (restoredSummary === undefined) {
  throw new Error('missing restored Summary sheet')
}

const afterRestore = numberValue(restored.getCellValue({ sheet: restoredSummary, row: 1, col: 1 }), 'restored revenue')

console.log(
  JSON.stringify(
    {
      before,
      after,
      afterRestore,
      persistedBytes: saved.length,
      verified: before === 3325 && after === 5586 && afterRestore === 5586,
    },
    null,
    2,
  ),
)
EOF
npx tsx eval.ts
```

Expected output:

```json
{
  "before": 3325,
  "after": 5586,
  "afterRestore": 5586,
  "persistedBytes": 1024,
  "verified": true
}
```

The byte count can change between package versions. The important check is that
`verified` is `true` and the restored workbook returns the same computed value.

## What to automate

Good server-side fits:

| Workflow              | Service input               | WorkPaper output                          |
| --------------------- | --------------------------- | ----------------------------------------- |
| Quote approval        | discount, quantity, term    | total price and approval status           |
| Budget variance       | department actuals          | variance amount and review flag           |
| Subscription forecast | churn, expansion, new seats | ending MRR and ARR                        |
| Import validation     | normalized records          | totals, rejected rows, and formulas       |
| Agent tool call       | one changed assumption      | before/after values and persistence proof |

This is not an XLSX styling library and it is not full Excel automation. Use
ExcelJS or SheetJS when the main artifact is an Excel file. Use HyperFormula
when broad Excel-compatible formula coverage is the primary requirement. Use
`@bilig/workpaper` when a Node process needs a small workbook model it can edit,
verify, and save.

## Next paths

- [Run the npm-only smoke test](try-bilig-headless-in-node.md)
- [Use WorkPaper behind a Node service](node-service-workpaper-recipe.md)
- [Persist formula-backed WorkPaper documents in Node](persisting-formula-backed-workpaper-documents-in-node.md)
- [Try five runnable workbook automation examples](workbook-automation-examples-node.md)
- [Compare headless spreadsheet engines](headless-spreadsheet-engine-comparison.md)
- [Check current Excel compatibility gaps](where-bilig-is-not-excel-compatible-yet.md)

Repository and release notes:
<https://github.com/proompteng/bilig>.

If it almost matches but a gap blocks adoption, open an implementation gap discussion:
<https://github.com/proompteng/bilig/discussions/new?category=general>.
