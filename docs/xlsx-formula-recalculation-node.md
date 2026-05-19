---
title: XLSX formula recalculation in Node.js
published: true
description: Recalculate formula-backed XLSX workbooks in a Node.js process after editing input cells, then export and verify the workbook round trip.
tags: typescript, node, xlsx, spreadsheet, formulas
canonical_url: https://proompteng.github.io/bilig/xlsx-formula-recalculation-node.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# XLSX formula recalculation in Node.js

This page is for the backend workflow where an `.xlsx` file is not just a
report artifact. The service imports a workbook, edits input cells, needs the
new formula values immediately, and still has to export an `.xlsx` at the end.

That is different from simply writing formula strings into a file and waiting
for Excel to calculate later.

## The production shape

A realistic server-side loop looks like this:

1. Load or generate a pricing, payout, quote, or validation workbook.
2. Pass the bytes to `xlsx-formula-recalc`.
3. Write request inputs into known cells.
4. Read the recalculated formula outputs before returning a response.
5. Export the edited workbook back to `.xlsx`.
6. Reimport the exported workbook in a test and verify formulas still produce
   the same values.

The last step matters. It catches the difference between "the in-memory model
looked right" and "the workbook artifact still works after the XLSX boundary."

## Install the narrow package

```sh
npm install xlsx-formula-recalc
```

For a one-off CLI run without adding it to a project:

```sh
npx --package xlsx-formula-recalc xlsx-recalc --demo --json
```

The demo creates a tiny workbook, edits input cells, recalculates `Summary!B2`,
writes `bilig-formula-recalc-demo.xlsx`, and prints `verified: true`.

For an existing workbook:

```sh
npx --package xlsx-formula-recalc xlsx-recalc pricing.xlsx \
  --set Inputs!B2=48 \
  --set Inputs!B3=1500 \
  --read Summary!B7 \
  --out pricing.recalculated.xlsx \
  --json
```

The CLI writes the updated workbook and prints the values read after
recalculation. Cell targets are sheet-qualified A1 references such as
`Inputs!B2` or `'Pricing Model'!F12`.

## Run the maintained example

If you want the shortest proof without cloning the repo, run the
[curlable XLSX recalculation proof](xlsx-recalculation-proof.md). It creates
source and edited `.xlsx` files in a blank folder and verifies the round trip.

From a clean clone:

```sh
git clone https://github.com/proompteng/bilig.git
cd bilig/examples/xlsx-recalculation-node
npm install
npm start
```

The example builds a pricing workbook, exports `pricing-model-source.xlsx`,
passes it through `xlsx-formula-recalc`, changes input cells, reads a
recalculated approval decision, exports `pricing-model-edited.xlsx`, and
reimports the edited workbook.

Expected output includes:

```json
{
  "before": {
    "decision": "review"
  },
  "after": {
    "decision": "approved"
  },
  "checks": {
    "decisionChanged": true,
    "exportedReimportMatchesAfter": true,
    "formulasSurvivedXlsxRoundTrip": true,
    "verified": true
  }
}
```

The source is intentionally small enough to read in one sitting:
[`examples/xlsx-recalculation-node/recalculate-xlsx.ts`](https://github.com/proompteng/bilig/blob/main/examples/xlsx-recalculation-node/recalculate-xlsx.ts).

## Minimal API boundary

Use the package API at the file boundary when the service needs a recalculated
answer immediately:

```ts
import { readFile, writeFile } from 'node:fs/promises'
import { recalculateXlsx } from 'xlsx-formula-recalc'

const result = recalculateXlsx(await readFile('pricing-model-source.xlsx'), {
  fileName: 'pricing-model-source.xlsx',
  edits: [
    { target: 'Inputs!B2', value: 48 },
    { target: 'Inputs!B3', value: 1250 },
  ],
  reads: ['Summary!B7'],
})

await writeFile('pricing-model-edited.xlsx', result.xlsx)
console.log({ decision: result.reads['Summary!B7'], warnings: result.warnings })
```

In production, keep a narrow adapter around this boundary. Your application
should know exactly which cells are inputs, which cells are outputs, and which
checks prove the workbook is still valid after export.

## When to keep using ExcelJS or SheetJS

Use ExcelJS or SheetJS first when the job is workbook-file manipulation:
styling, rows, sheets, images, tables, streaming writes, or broad spreadsheet
format interchange.

Use `xlsx-formula-recalc` when the Node process must own the recalculated answer
before it accepts, rejects, queues, or persists a workflow. Use
`@bilig/headless/xlsx` directly only when you need the lower-level WorkPaper
snapshot boundary.

Many services should combine the tools: use a file library for presentation
details and use a formula runtime for the auditable decision path.

## Trust checks before production

Before putting XLSX recalculation on a customer-critical path, add tests for:

- the exact input cells your service writes
- the exact output cells your service reads
- unsupported formulas and compatibility limits
- exported workbook reimport
- stale cached values in source files
- a golden workbook fixture that opens in Excel or another spreadsheet app

Bilig is not full Excel. The useful promise is narrower: a Node service can run
formula-backed workbook logic, prove readback after edits, and keep a checked
XLSX boundary around the parts it owns.

## Related proof

- [Curlable XLSX recalculation proof](xlsx-recalculation-proof.md)
- [Runnable XLSX recalculation example](https://github.com/proompteng/bilig/tree/main/examples/xlsx-recalculation-node)
- [Excel file as a calculation engine in Node.js](excel-file-calculation-engine-node.md)
- [Stale XLSX formula cache in Node.js](stale-xlsx-formula-cache-node.md)
- [ExcelJS formula recalculation in Node.js](exceljs-formula-recalculation-node.md)
- [ExcelJS shared formulas and Node.js recalculation](exceljs-shared-formula-recalculation-node.md)
- [SheetJS and ExcelJS boundary](sheetjs-exceljs-alternative-formula-workbook-api.md)
- [Stale XLSX cache and Excel oracle checks](xlsx-corpus-verifier-walkthrough.md)
- [Compatibility limits](where-bilig-is-not-excel-compatible-yet.md)

If this saves you from opening Excel in a backend job just to refresh formulas,
star the repository so the project is easier for the next Node developer to
find: <https://github.com/proompteng/bilig/stargazers>.
