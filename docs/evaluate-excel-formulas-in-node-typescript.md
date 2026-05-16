---
title: Evaluate Excel formulas in Node.js with TypeScript
published: true
description: A TypeScript-first path for evaluating spreadsheet formulas in Node.js with @bilig/headless, verified readback, and JSON persistence.
tags: typescript, node, excel, spreadsheet, formulas
canonical_url: https://proompteng.github.io/bilig/evaluate-excel-formulas-in-node-typescript.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Evaluate Excel formulas in Node.js with TypeScript

Use `@bilig/headless` when a Node.js program needs a small workbook model it can
edit, recalculate, read, and persist. The useful case is not "open Excel on the
server." It is "run spreadsheet-shaped logic in a service and prove the result
that came back."

That usually means one of these jobs:

- an API route needs to price a quote from workbook formulas
- a worker needs to update forecast inputs and read the new total
- a test needs a formula-backed fixture instead of duplicated math
- a coding agent needs to write a cell and return verified readback
- a service needs to save workbook state as JSON and restore it later

If the main job is styling or writing an `.xlsx` file, start with an XLSX
library. If the main job is a long tail of arbitrary Excel edge cases with no
fixture-reduction step, evaluate a mature formula engine first. If the job is a
TypeScript workbook object with formulas, edits, readback, and persistence, try
the WorkPaper path below.

## Run a TypeScript formula smoke test

Start from an empty directory:

```sh
mkdir bilig-node-formulas
cd bilig-node-formulas
npm init -y
npm pkg set type=module
npm install @bilig/headless
npm install -D tsx typescript @types/node
```

Create `eval-node-formulas.ts`:

```ts
import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
  type WorkPaperCellAddress,
} from '@bilig/headless'

type NumericCell = {
  value: number
}

function readNumber(workbook: WorkPaper, address: WorkPaperCellAddress): number {
  const value = workbook.getCellValue(address)

  if (typeof value !== 'object' || value === null || typeof (value as NumericCell).value !== 'number') {
    throw new Error(`Expected numeric cell, got ${JSON.stringify(value)}`)
  }

  return (value as NumericCell).value
}

const workbook = WorkPaper.buildFromSheets({
  Quote: [
    ['Metric', 'Value'],
    ['Seats', 12],
    ['Price', 49],
    ['Discount', 0.1],
  ],
  Summary: [
    ['Metric', 'Value'],
    ['Monthly total', '=Quote!B2*Quote!B3*(1-Quote!B4)'],
  ],
})

const quote = workbook.getSheetId('Quote')
const summary = workbook.getSheetId('Summary')

if (quote === undefined || summary === undefined) {
  throw new Error('Expected Quote and Summary sheets')
}

const total: WorkPaperCellAddress = { sheet: summary, row: 1, col: 1 }
const before = readNumber(workbook, total)

workbook.setCellContents({ sheet: quote, row: 1, col: 1 }, 20)
const afterEdit = readNumber(workbook, total)

const serialized = serializeWorkPaperDocument(exportWorkPaperDocument(workbook, { includeConfig: true }))
const restored = createWorkPaperFromDocument(parseWorkPaperDocument(serialized))
const afterRestore = readNumber(restored, total)

if (before !== 529.2 || afterEdit !== 882 || afterRestore !== afterEdit) {
  throw new Error(JSON.stringify({ before, afterEdit, afterRestore }))
}

console.log({ before, afterEdit, afterRestore, verified: true })
```

Run it:

```sh
npx tsx eval-node-formulas.ts
```

Expected output:

```json
{ "before": 529.2, "afterEdit": 882, "afterRestore": 882, "verified": true }
```

That proves the Node process created a workbook, evaluated a cross-sheet
formula, edited an input cell, read the dependent result, serialized the
WorkPaper document, restored it, and read the same calculated value again.

## Where this is different from calling formula functions

A direct formula-function library is useful when code wants to call something
like `SUM()` as a JavaScript function. A workbook runtime is useful when the
formula depends on sheet state:

- the formula lives in a cell
- the formula refers to other cells or sheets
- edits should recalculate dependent cells
- the service needs to persist the workbook document
- an agent or test needs evidence that the write changed the expected output

That is the boundary `@bilig/headless` is built around.

## Where this is different from writing XLSX files

An XLSX library is the right first choice when the file is the product: reports,
styles, images, tables, and handoff to Excel or another spreadsheet app.

Use `@bilig/headless` when the service needs calculated workbook state before a
person opens any file. You can still export or import at the system boundary,
but the WorkPaper model is the part that calculates and verifies the values in
Node.

## Production checklist

Before using any headless spreadsheet engine in a service, check these items:

- Does the engine support the formulas your workbook actually uses?
- Can your service read the calculated value after an edit?
- Can you persist and restore the workbook state?
- Does the package expose useful errors for unsupported formulas?
- Can you keep a small fixture in CI that covers your real workflow?

For `@bilig/headless`, start with one TypeScript fixture and keep it boring:
build the workbook, edit one input, read one dependent output, persist, restore,
and assert the same value after restore.

## Next paths

- [Run the npm-only smoke test](try-bilig-headless-in-node.md)
- [Node.js spreadsheet formula engine guide](node-spreadsheet-formula-engine.md)
- [Node service WorkPaper recipe](node-service-workpaper-recipe.md)
- [Persist formula-backed WorkPaper documents in Node](persisting-formula-backed-workpaper-documents-in-node.md)
- [Five Node workbook automation examples](workbook-automation-examples-node.md)
- [Headless spreadsheet engine comparison](headless-spreadsheet-engine-comparison.md)
- [Where bilig is not Excel-compatible yet](where-bilig-is-not-excel-compatible-yet.md)

If this is the shape of formula automation you were looking for, star the repo
so the package is easier to find later:
<https://github.com/proompteng/bilig/stargazers>.

If it almost matches but a gap blocks adoption, use the adoption blocker form:
<https://github.com/proompteng/bilig/discussions/new?category=general>.
