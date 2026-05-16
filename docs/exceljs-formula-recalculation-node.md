---
title: ExcelJS formula recalculation in Node.js
published: true
description: Decide what to do when ExcelJS can write formula records but your Node.js service needs recalculated formula values after changing workbook inputs.
tags: typescript, node, exceljs, spreadsheet, formulas
canonical_url: https://proompteng.github.io/bilig/exceljs-formula-recalculation-node.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# ExcelJS formula recalculation in Node.js

This page is for the specific failure mode where an ExcelJS workflow can create
or update formula cells, but the Node.js process also needs the calculated
result after changing inputs.

That is a runtime problem, not just a file-writing problem.

## Short answer

Use ExcelJS when the product is an `.xlsx` file that Excel, LibreOffice, or
another spreadsheet application will open and calculate later.

Use a formula runtime when the service needs to write inputs and read the
recalculated values before returning a response. `@bilig/headless` is one
option for the narrower case where the workbook can live as a TypeScript
WorkPaper document with JSON persistence and verified readback.

## Why cached values are not enough

Spreadsheet files can contain formula text and cached results. A file library
can preserve those records, and some libraries let you supply the cached result
yourself. That does not mean the library recalculated the dependency graph
after a service changed an input cell.

The difference matters in production:

1. A user changes a discount, quantity, tax rate, or threshold.
2. The service writes the input into the workbook model.
3. Dependent formulas must update now.
4. The service must reject or persist based on the value it actually read back.

If step 3 happens later in Excel, the backend never owned the decision.

## Decision table

| Job                                                                        | Better starting point                       |
| -------------------------------------------------------------------------- | ------------------------------------------- |
| Generate an XLSX report with styles, sheets, tables, and formula strings   | ExcelJS                                     |
| Open a file later in Excel and let Excel calculate formulas                | ExcelJS                                     |
| Preserve formula records and cached values from an existing workbook       | ExcelJS or SheetJS-style tooling            |
| Recalculate workbook formulas inside a Node.js request, job, or agent tool | A formula runtime such as `@bilig/headless` |
| Persist formula-backed state as JSON and verify it after restore           | `@bilig/headless` WorkPaper                 |

## Minimal WorkPaper replacement for the recalculation step

Install the runtime in a scratch project:

```sh
mkdir exceljs-recalc-eval
cd exceljs-recalc-eval
npm init -y
npm pkg set type=module
npm install @bilig/headless
npm install -D tsx typescript @types/node
```

Create `recalculate.ts`:

```ts
import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from '@bilig/headless'

type NumericCell = {
  value: number
}

function readNumber(cell: unknown, label: string): number {
  if (typeof cell === 'object' && cell !== null && typeof (cell as NumericCell).value === 'number') {
    return (cell as NumericCell).value
  }

  throw new Error(`Expected ${label} to be numeric, got ${JSON.stringify(cell)}`)
}

const workbook = WorkPaper.buildFromSheets({
  Inputs: [
    ['Metric', 'Value'],
    ['Units', 100],
    ['Unit price', 49],
    ['Discount', 0.1],
  ],
  Quote: [
    ['Metric', 'Value'],
    ['Net total', '=Inputs!B2*Inputs!B3*(1-Inputs!B4)'],
  ],
})

const inputs = workbook.getSheetId('Inputs')
const quote = workbook.getSheetId('Quote')
if (inputs === undefined || quote === undefined) {
  throw new Error('Expected Inputs and Quote sheets')
}

const netTotalCell = { sheet: quote, row: 1, col: 1 }
const before = readNumber(workbook.getCellValue(netTotalCell), 'before')

workbook.setCellContents({ sheet: inputs, row: 3, col: 1 }, 0.25)
const after = readNumber(workbook.getCellValue(netTotalCell), 'after')

const serialized = serializeWorkPaperDocument(exportWorkPaperDocument(workbook, { includeConfig: true }))
const restored = createWorkPaperFromDocument(parseWorkPaperDocument(serialized))
const restoredQuote = restored.getSheetId('Quote')
if (restoredQuote === undefined) {
  throw new Error('Expected restored Quote sheet')
}

const afterRestore = readNumber(restored.getCellValue({ sheet: restoredQuote, row: 1, col: 1 }), 'after restore')

console.log({
  before,
  after,
  afterRestore,
  verified: before === 4410 && after === 3675 && afterRestore === after,
})
```

Run it:

```sh
npx tsx recalculate.ts
```

Expected output:

```json
{ "before": 4410, "after": 3675, "afterRestore": 3675, "verified": true }
```

That verifies the part ExcelJS is not designed to own: an input changed,
dependent formulas recalculated in the same Node.js process, and the persisted
document restored with the same computed result.

## How to combine ExcelJS and Bilig

The honest architecture is to keep file generation and formula runtime separate:

1. Use ExcelJS for `.xlsx` files, styling, worksheets, and reports.
2. Use `@bilig/headless` for the formula-backed business state your service
   must trust immediately.
3. Add compatibility tests at the boundary if you import or export XLSX files.

Do not mix these responsibilities silently. A cached value in a file is not the
same as recalculated business state.

## When not to use Bilig

Do not choose `@bilig/headless` only to generate styled XLSX files.

Do not choose it if a human can open the workbook in Excel before any business
decision depends on the calculated value.

Do not choose it when you need full Excel compatibility across every formula,
chart, pivot table, macro, or workbook artifact. Check the
[compatibility limits](where-bilig-is-not-excel-compatible-yet.md) first.

## Related proof

- [SheetJS and ExcelJS boundary](sheetjs-exceljs-alternative-formula-workbook-api.md)
- [ExcelJS shared formulas and Node.js recalculation](exceljs-shared-formula-recalculation-node.md)
- [Node spreadsheet formula engine](node-spreadsheet-formula-engine.md)
- [Headless spreadsheet engine for Node services and agents](headless-spreadsheet-engine-node-services-agents.md)
- [Persist formula-backed WorkPaper documents in Node](persisting-formula-backed-workpaper-documents-in-node.md)
- [90-second Node quickstart](try-bilig-headless-in-node.md)

If this saves you an ExcelJS recalculation workaround, star the repository so
the project is easier for the next backend developer to find:
<https://github.com/proompteng/bilig/stargazers>.
