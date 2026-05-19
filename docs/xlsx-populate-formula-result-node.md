---
title: xlsx-populate formula results in Node.js
published: true
description: How to handle calculated formula values when xlsx-populate writes formulas but a Node.js service needs fresh readback.
tags: typescript, node, xlsx-populate, xlsx, formulas, recalculation
canonical_url: https://proompteng.github.io/bilig/xlsx-populate-formula-result-node.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# xlsx-populate formula results in Node.js

`xlsx-populate` can write cells and formulas into XLSX files. It is not a
workbook calculation engine.

This page is for the practical `xlsx-populate` issue where a generated workbook
needs to keep the formula and also expose the new calculated value before
anyone opens the file in Excel. A public `xlsx-populate` issue asks exactly
that: setting a formula preserves the formula, setting a value replaces it, and
the result is not calculated until another spreadsheet engine opens the file.

That matters when a backend flow needs both:

- formula text in the generated workbook; and
- a fresh calculated value immediately after changing inputs.

Those are different responsibilities. A file writer can serialize the formula.
A calculator has to evaluate the dependency graph.

## The common trap

Setting a formula does not mean the cached formula result changed:

```ts
cell.formula('A1*10')
```

Some XLSX libraries can serialize a `{ formula, result }` pair, but that still
requires you to compute the result somewhere. Setting `fullCalcOnLoad` can ask
Excel or LibreOffice to recalculate later; it does not give a Node API route a
fresh value now.

If the backend is about to approve a quote, price an order, validate an import,
or enqueue a payout, the cached value inside the XLSX file is not a decision
source. It is only the last value some other spreadsheet engine happened to
write.

## Add a recalculation step after xlsx-populate

If the rest of the pipeline already uses `xlsx-populate`, keep it there. Treat
it as the authoring step, then hand the resulting XLSX bytes to
`xlsx-formula-recalc` before the service reads output cells.

```ts
import { readFile, writeFile } from 'node:fs/promises'
import { recalculateXlsx } from 'xlsx-formula-recalc'

// This can also be the Buffer returned by an xlsx-populate output call.
const source = await readFile('quote.xlsx')

const result = recalculateXlsx(source, {
  edits: [
    { target: 'Inputs!B2', value: 42 },
    { target: 'Inputs!B3', value: 1500 },
  ],
  reads: ['Summary!B7'],
})

console.log(result.reads['Summary!B7'])
await writeFile('quote.recalculated.xlsx', result.xlsx)
```

For a file-level repair step:

```sh
npx --package xlsx-formula-recalc xlsx-recalc quote.xlsx \
  --set Inputs!B2=42 \
  --set Inputs!B3=1500 \
  --read Summary!B7 \
  --out quote.recalculated.xlsx \
  --json
```

This is the narrowest bridge for the common public questions:

- "xlsx-populate wrote the formula but the value did not update"
- "I need formula plus calculated value in the generated workbook"
- "I need to read the calculated formula value in Node before returning"

It is still not an Excel clone. If the workbook uses unsupported functions,
macros, external links, or volatile formulas, keep a reduced fixture in tests
and compare the answer to Excel or LibreOffice before treating it as production
logic.

## TODAY() and cached date values

`TODAY()` is a volatile formula. `xlsx-populate` can read the cached serial date
stored in the file, but it does not execute `TODAY()` again.

If the cached value is enough, convert the Excel serial number explicitly:

```ts
function excelSerialDateToUtcDate(serial: number): Date {
  return new Date((serial - 25569) * 86400 * 1000)
}
```

That assumes the normal Excel 1900 date system. If the workbook uses the 1904
date system, or if the backend needs today's value at request time, recalculate
the workbook before reading the cell.

## The small rule

Keep `xlsx-populate` for file authoring. Move the calculation step to a runtime
that can answer immediately.

That gives the backend a plain contract:

1. write request values into known input cells;
2. recalculate in the same Node process;
3. read known output cells;
4. persist the workbook state;
5. export XLSX if a human still needs the file.

## Pick the boundary

Use `xlsx-populate` when the job is primarily XLSX file generation or editing.

Use Excel, LibreOffice, or Microsoft Graph when the result must match Excel and
the operational cost of a spreadsheet host is acceptable.

Use `@bilig/headless` when the workbook is service-owned business logic and the
backend needs:

- input writes through an API;
- formula recalculation in-process;
- output readback before returning a response;
- JSON persistence and restore proof;
- optional XLSX import/export at the edge.

## WorkPaper version of the flow

```ts
import { WorkPaper } from '@bilig/headless'

const workbook = new WorkPaper()
const sheet = workbook.addSheet('Quote')

workbook.setCellContents({ sheet, row: 0, col: 0 }, 42)
workbook.setCellContents({ sheet, row: 0, col: 1 }, '=A1*10')

const value = workbook.getCellDisplayValue({ sheet, row: 0, col: 1 })
const snapshot = workbook.exportSnapshot()
const restored = WorkPaper.buildFromSnapshot(snapshot)

try {
  const restoredSheet = restored.getSheetId('Quote')
  if (restoredSheet === undefined) {
    throw new Error('Missing Quote sheet after restore')
  }

  const restoredValue = restored.getCellDisplayValue({
    sheet: restoredSheet,
    row: 0,
    col: 1,
  })

  console.log({ value, restoredValue })
} finally {
  restored.dispose()
  workbook.dispose()
}
```

That shape avoids the "write a formula, then wait for Excel to populate the
cache" loop. The service owns the calculated state before it emits a response.

If the XLSX file is the source of truth, use the XLSX edge instead of creating a
blank WorkPaper:

```ts
import { readFile } from 'node:fs/promises'
import { WorkPaper } from '@bilig/headless'
import { importXlsx } from '@bilig/headless/xlsx'

const imported = importXlsx(await readFile('quote-template.xlsx'), 'quote-template.xlsx')
const workbook = WorkPaper.buildFromSnapshot(imported.snapshot)

const inputs = workbook.getSheetId('Inputs')
const outputs = workbook.getSheetId('Outputs')
if (inputs === undefined || outputs === undefined) {
  throw new Error('Expected Inputs and Outputs sheets')
}

workbook.setCellContents({ sheet: inputs, row: 1, col: 1 }, 42)
const approved = workbook.getCellDisplayValue({ sheet: outputs, row: 1, col: 1 })

console.log({ approved })
```

That is the part `xlsx-populate` does not try to own: the service changed an
input and read the dependent value back before returning.

## Test with a reduced workbook

If an existing `xlsx-populate` pipeline has a real formula case, reduce it to a
public workbook fixture and run:

```sh
curl -fsSLo formula-clinic-report.ts \
  https://proompteng.github.io/bilig/formula-clinic-report.ts
npx tsx formula-clinic-report.ts ./reduced.xlsx \
  --cells "Quote!B1"
```

The report prints package version, imported sheets, formula samples, requested
readback, and a fixture checklist. It runs locally and does not upload workbook
contents.

## Related

- [Fix stale XLSX formula values in Node.js](stale-xlsx-formula-cache-node.md)
- [ExcelJS formula recalculation in Node.js](exceljs-formula-recalculation-node.md)
- [XLSX formula recalculation in Node.js](xlsx-formula-recalculation-node.md)
- [Formula bug clinic](formula-bug-clinic.md)

Source issue:
<https://github.com/dtjohnson/xlsx-populate/issues/265>

Related public issues:

- <https://github.com/dtjohnson/xlsx-populate/issues/354>
- <https://github.com/dtjohnson/xlsx-populate/issues/275>
