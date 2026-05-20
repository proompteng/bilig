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
recalculated values before returning a response.
[`@bilig/exceljs-formula-recalc`](https://www.npmjs.com/package/@bilig/exceljs-formula-recalc)
is the narrow package for existing ExcelJS workflows. It writes the workbook to
bytes, recalculates through Bilig WorkPaper, reloads the workbook, and patches
the read cells with fresh formula results.

Use `@bilig/workpaper` or `@bilig/workpaper` directly when the workbook can live
as a TypeScript WorkPaper document with JSON persistence and verified readback.

## The Exact ExcelJS Failure

Searches like "ExcelJS formula result not updating", "updating formula result",
and "get computed value of Excel sheet cell in Node.js" usually describe the
same boundary:

1. ExcelJS can write the formula record.
2. ExcelJS can preserve a cached `result`.
3. ExcelJS does not recalculate the dependency graph after your service edits
   an input cell.

If the service needs the computed value in the same request or job, bridge the
ExcelJS workbook through `@bilig/exceljs-formula-recalc` and read the returned proof
values before sending the response.

## Why cached values are not enough

Spreadsheet files can contain formula text and cached results. A file library
can preserve those records, and some libraries let you supply the cached result
yourself. That does not mean the library recalculated the dependency graph
after a service changed an input cell.

A common ExcelJS trap looks like this:

```ts
workbook.calcProperties.fullCalcOnLoad = true
worksheet.getCell('A1').value = 15

console.log(worksheet.getCell('C1').value)
// { formula: 'A1+B1', result: 20 }
```

If `C1` originally cached `20`, setting `fullCalcOnLoad` does not make
`worksheet.getCell('C1').value` become `25` inside the same Node.js process.
That flag tells a spreadsheet application to recalculate when it opens the
file. It is a file instruction, not an in-process calculation engine.

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
| Existing ExcelJS workbook needs recalculated values inside Node            | `@bilig/exceljs-formula-recalc`                    |
| Raw XLSX bytes need recalculated values inside Node                        | `@bilig/xlsx-formula-recalc`                       |
| Recalculate workbook formulas inside a Node.js request, job, or agent tool | A formula runtime such as `@bilig/workpaper` |
| Persist formula-backed state as JSON and verify it after restore           | `@bilig/workpaper` WorkPaper                 |

## One-command proof

Run this before wiring it into an app:

```sh
npx --package @bilig/exceljs-formula-recalc exceljs-recalc --demo --json
```

The demo creates a small workbook, changes `Inputs!B2` and `Inputs!B3`,
recalculates `Summary!B2`, writes `bilig-formula-recalc-demo.xlsx`, and prints
`verified: true`.

Exact reproduction for the high-view Stack Overflow question:

```sh
git clone https://github.com/proompteng/bilig.git
cd bilig
npm --prefix examples/recalc-bridge-workflows install
npm --prefix examples/recalc-bridge-workflows run so:exceljs-44199441
```

That script mirrors
[Get computed value of Excel sheet cell in Node.js](https://stackoverflow.com/questions/44199441/get-computed-value-of-excel-sheet-cell-in-node-js):
`A1` changes from `1` to `3`, ExcelJS still has the stale formula `result = 3`,
then `@bilig/exceljs-formula-recalc` verifies and patches the formula result to `5`.

## Minimal ExcelJS bridge

Install the runtime in a scratch project:

```sh
mkdir exceljs-recalc-eval
cd exceljs-recalc-eval
npm init -y
npm pkg set type=module
npm install exceljs @bilig/exceljs-formula-recalc
npm install -D tsx typescript @types/node
```

Create `recalculate.ts`:

```ts
import ExcelJS from 'exceljs'
import { recalculateExceljsWorkbook } from '@bilig/exceljs-formula-recalc'

const workbook = new ExcelJS.Workbook()
const inputs = workbook.addWorksheet('Inputs')
inputs.getCell('A1').value = 'Metric'
inputs.getCell('B1').value = 'Value'
inputs.getCell('A2').value = 'Units'
inputs.getCell('B2').value = 100
inputs.getCell('A3').value = 'Unit price'
inputs.getCell('B3').value = 49
inputs.getCell('A4').value = 'Discount'
inputs.getCell('B4').value = 0.1

const quote = workbook.addWorksheet('Quote')
quote.getCell('A1').value = 'Metric'
quote.getCell('B1').value = 'Value'
quote.getCell('A2').value = 'Net total'
quote.getCell('B2').value = {
  formula: 'Inputs!B2*Inputs!B3*(1-Inputs!B4)',
  result: 4410,
}

const result = await recalculateExceljsWorkbook(workbook, {
  edits: [{ target: 'Inputs!B4', value: 0.25 }],
  reads: ['Quote!B2'],
})

const readback = readNumber(result.reads['Quote!B2'])
const recalculatedCell = workbook.getWorksheet('Quote')?.getCell('B2').value

console.log({
  readback,
  exceljsCell: recalculatedCell,
  verified: readback === 3675,
})

function readNumber(cell: unknown): number {
  if (typeof cell === 'object' && cell !== null && 'value' in cell && typeof cell.value === 'number') {
    return cell.value
  }
  throw new Error(`Expected numeric formula result, got ${JSON.stringify(cell)}`)
}
```

Run it:

```sh
npx tsx recalculate.ts
```

Expected output includes:

```json
{ "verified": true }
```

That verifies the part ExcelJS is not designed to own: an input changed and a
dependent formula recalculated in the same Node.js process.

## How to combine ExcelJS and Bilig

The honest architecture is to keep file generation and formula runtime separate:

1. Use ExcelJS for `.xlsx` files, styling, worksheets, and reports.
2. Use `@bilig/exceljs-formula-recalc` for an ExcelJS workbook that needs fresh formula
   results before the process returns.
3. Use `@bilig/workpaper` for formula-backed business state your service
   must trust immediately.
4. Add compatibility tests at the boundary if you import or export XLSX files.

Do not mix these responsibilities silently. A cached value in a file is not the
same as recalculated business state.

## When not to use Bilig

Do not choose `@bilig/workpaper` only to generate styled XLSX files.

Do not choose it if a human can open the workbook in Excel before any business
decision depends on the calculated value.

Do not choose it when you need full Excel compatibility across every formula,
chart, pivot table, macro, or workbook artifact. Check the
[compatibility limits](where-bilig-is-not-excel-compatible-yet.md) first.

## Related proof

- [SheetJS and ExcelJS boundary](sheetjs-exceljs-alternative-formula-workbook-api.md)
- [Excel file as a calculation engine in Node.js](excel-file-calculation-engine-node.md)
- [Microsoft Graph Excel recalculation in Node.js](microsoft-graph-excel-recalculation-node.md)
- [ExcelJS shared formulas and Node.js recalculation](exceljs-shared-formula-recalculation-node.md)
- [Node spreadsheet formula engine](node-spreadsheet-formula-engine.md)
- [Headless spreadsheet engine for Node services and agents](headless-spreadsheet-engine-node-services-agents.md)
- [Persist formula-backed WorkPaper documents in Node](persisting-formula-backed-workpaper-documents-in-node.md)
- [90-second Node quickstart](try-bilig-headless-in-node.md)

If this saves you an ExcelJS recalculation workaround, star the repository so
the project is easier for the next backend developer to find:
<https://github.com/proompteng/bilig/stargazers>.
