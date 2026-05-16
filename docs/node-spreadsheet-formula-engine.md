---
title: Node.js spreadsheet formula engine for services
published: true
description: Use @bilig/headless when a Node.js service needs workbook formulas, computed readback, JSON persistence, and verified edits without a browser grid.
tags: typescript, node, spreadsheet, formulas, opensource
canonical_url: https://proompteng.github.io/bilig/node-spreadsheet-formula-engine.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Node.js spreadsheet formula engine for services

Use `@bilig/headless` when a Node.js service needs spreadsheet formulas as part
of its own runtime. The useful case is not "make a prettier spreadsheet." It is
"accept inputs, update workbook state, recalculate formulas, persist the
document, and return values that were actually read back from the engine."

That shows up in ordinary backend work:

- pricing rules that business users still think about as spreadsheet formulas
- invoice, quote, commission, or usage-billing checks inside an API route
- forecasts and variance reports in queue workers
- coding-agent tools that must prove a workbook edit happened
- tests that need a formula-backed fixture instead of a hand-coded math clone

If you only need to write an XLSX file for Excel to open later, use an XLSX
library. If you need maximum Excel-function coverage today, evaluate
HyperFormula first. If you need a small TypeScript WorkPaper object that a Node
process can mutate, verify, and save as JSON, `@bilig/headless` is the slice to
try.

## What the engine owns

`@bilig/headless` gives the service a WorkPaper object. A WorkPaper is a
programmatic workbook with sheets, cell addresses, formulas, computed values,
structural edits, and persistence helpers.

The important boundary is readback. Do not trust that a service "wrote the
formula" just because a string was assigned. Read the calculated cell value
through the WorkPaper API, serialize the document, restore it, and read the
same value again. That is the path the package is built around.

## Quick smoke test

Run this from an empty directory:

```sh
mkdir bilig-formula-engine-eval
cd bilig-formula-engine-eval
npm init -y
npm pkg set type=module
npm install @bilig/headless
npm install -D tsx typescript @types/node
cat > formula-engine-smoke.ts <<'EOF'
import { WorkPaper } from '@bilig/headless'

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
  Plan: [
    ['Metric', 'Value'],
    ['Customers', 80],
    ['Price', 49],
    ['Discount', 0.08],
  ],
  Summary: [
    ['Metric', 'Value'],
    ['Net revenue', '=Plan!B2*Plan!B3*(1-Plan!B4)'],
  ],
})

const sheet = workbook.getSheetId('Summary')
if (sheet === undefined) {
  throw new Error('Summary sheet was not created')
}

const netRevenue = readNumber(workbook.getCellValue({ sheet, row: 1, col: 1 }), 'net revenue')
if (netRevenue !== 3606.4) {
  throw new Error(`Unexpected formula readback: ${netRevenue}`)
}

console.log({ netRevenue, verified: true })
EOF
npx tsx formula-engine-smoke.ts
```

Expected output:

```json
{ "netRevenue": 3606.4, "verified": true }
```

That tiny check proves three things at once: multi-sheet workbook creation,
formula evaluation, and computed readback from Node.

## When it fits

Start here when the service needs a workbook model, not just isolated formula
functions.

| Need                                   | Why `@bilig/headless` helps                                                                          |
| -------------------------------------- | ---------------------------------------------------------------------------------------------------- |
| Put spreadsheet formulas behind an API | The service can build a WorkPaper, edit input cells, and return computed cells.                      |
| Let an agent edit a workbook safely    | The agent can report exact changed cells and post-write readback instead of only narrating intent.   |
| Persist formula-backed state           | Export the WorkPaper document, store JSON, restore it later, and verify formulas still calculate.    |
| Keep examples runnable                 | The repo includes invoice, budget variance, subscription MRR, quote approval, and capacity examples. |

## When to choose something else

Use a different tool when the job is outside this package's current shape:

- Use ExcelJS or SheetJS-style tooling when the main product is an XLSX file
  with styles, tables, images, or import/export fidelity.
- Use HyperFormula when the primary need is broad, mature Excel-compatible
  formula coverage inside a headless calculation engine.
- Use a visual spreadsheet grid when people need to edit cells directly in the
  browser.
- Do not choose `bilig` expecting a finished Excel clone. The compatibility
  boundaries are public because the project is still early.

## Evaluation path

If the smoke test matches your use case, run the maintained example next:

```sh
git clone https://github.com/proompteng/bilig.git
cd bilig/examples/headless-workpaper
npm install
npm start
npm run agent:verify
```

Then inspect the focused examples:

- [invoice totals](https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#invoice-totals)
- [budget variance alerts](https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#budget-variance-alerts)
- [subscription MRR forecast](https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#subscription-mrr-forecast)
- [quote approval threshold](https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#quote-approval-threshold)
- [fulfillment capacity plan](https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#fulfillment-capacity-plan)
- [serverless WorkPaper API route](https://github.com/proompteng/bilig/tree/main/examples/serverless-workpaper-api)

## Related docs

- [Five Node.js workbook automation examples](workbook-automation-examples-node.md)
- [WorkPaper Node service recipe](node-service-workpaper-recipe.md)
- [Excel file as a calculation engine in Node.js](excel-file-calculation-engine-node.md)
- [Persist formula-backed WorkPaper documents in Node](persisting-formula-backed-workpaper-documents-in-node.md)
- [Headless spreadsheet engine comparison for Node services and agents](headless-spreadsheet-engine-comparison.md)
- [JavaScript spreadsheet library for Node services](javascript-spreadsheet-library-headless-node.md)
- [SheetJS and ExcelJS alternative for formula-backed workbook APIs](sheetjs-exceljs-alternative-formula-workbook-api.md)
- [HyperFormula alternative for headless WorkPaper workflows](hyperformula-alternative-headless-workpaper.md)
- [Where bilig is not Excel-compatible yet](where-bilig-is-not-excel-compatible-yet.md)
- [Starter issues for first-time contributors](starter-issues.md)

If this package saves you a workbook automation spike, star the repository so
the project is easier for the next backend developer to find:
<https://github.com/proompteng/bilig/stargazers>.

If it almost matches but a gap blocks adoption, use the adoption blocker form:
<https://github.com/proompteng/bilig/discussions/new?category=general>.
