---
title: Fix stale XLSX formula values in Node.js
published: true
description: Explain why Node XLSX readers return old cached formula values after input edits, and when to use Excel, Graph, xlsx-calc, HyperFormula, or @bilig/workpaper.
tags: typescript, node, excel, xlsx, formulas, recalculation
canonical_url: https://proompteng.github.io/bilig/stale-xlsx-formula-cache-node.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Fix stale XLSX formula values in Node.js

The failure is simple and easy to miss:

1. Your Node job opens an `.xlsx`.
2. It changes an input cell.
3. It reads a formula cell.
4. The formula cell still has yesterday's answer.
5. Somebody opens the file in Excel, hits save, and now the number is right.

That is not a write bug. An `.xlsx` can store both the formula text and the last
calculated value. Most Node file libraries can change the cell contents; they do
not run Excel's calculation engine.

## Short answer

If a person opens the file before the number matters, mark the workbook to
recalculate on open and treat cached formula values as stale.

If the backend makes a decision from the number, calculate before you read.

For exact Excel behavior, use Excel, LibreOffice, or Microsoft Graph. For a JS
runtime, test `xlsx-calc` or HyperFormula against your actual workbook. Use
`@bilig/workpaper` when the service can own the workbook state locally and needs
formula readback, JSON persistence, and restore tests.

If you already have an XLSX file and just need fresh values before returning,
start with the narrow file-level command:

```sh
npx --package @bilig/xlsx-formula-recalc xlsx-recalc quote.xlsx \
  --set Inputs!B2=42 \
  --read Summary!B7 \
  --out quote.recalculated.xlsx \
  --json
```

That is the first path to try for `xlsx-populate`, SheetJS, or template
generation jobs where the stale value is the blocker, not workbook ownership.

## What not to trust

Do not treat this as a fresh calculated value:

```ts
const workbook = XLSX.readFile('quote.xlsx')
const sheet = workbook.Sheets.Inputs

sheet.B2.v = 42_000
XLSX.writeFile(workbook, 'quote-edited.xlsx')

const reread = XLSX.readFile('quote-edited.xlsx')
console.log(reread.Sheets.Quote.B8.v)
```

That value can still be the cached result from before the edit. Recalc-on-open
helps the next human who opens the file. It does not help an API route that
needs the answer now.

## Local WorkPaper shape

```ts
import { WorkPaper } from '@bilig/workpaper'

const workbook = new WorkPaper()
const inputs = workbook.addSheet('Inputs')
const quote = workbook.addSheet('Quote')

workbook.setCellContents({ sheet: inputs, row: 1, col: 1 }, 40_000)
workbook.setCellContents({ sheet: inputs, row: 2, col: 1 }, 0.2)
workbook.setCellContents({ sheet: quote, row: 1, col: 1 }, '=Inputs!B2*Inputs!B3')

const before = workbook.getCellValue({ sheet: quote, row: 1, col: 1 })
workbook.setCellContents({ sheet: inputs, row: 1, col: 1 }, 42_000)
const after = workbook.getCellValue({ sheet: quote, row: 1, col: 1 })

const restored = WorkPaper.buildFromSnapshot(workbook.exportSnapshot())
const restoredQuote = restored.getSheetId('Quote')
if (restoredQuote === undefined) {
  throw new Error('Expected Quote sheet')
}

const afterRestore = restored.getCellValue({ sheet: restoredQuote, row: 1, col: 1 })
console.log({ before, after, afterRestore })
```

For a real service, keep the boundary boring:

- write only named input cells;
- read only named output cells;
- fixture-test unsupported formulas;
- persist the WorkPaper JSON used for the decision;
- export XLSX only when a human needs the file artifact.

## Pick the tool by the failure mode

| Situation                                            | First thing to try                   |
| ---------------------------------------------------- | ------------------------------------ |
| File is only a downloadable report                   | Recalculate on open                  |
| Backend must return the computed value               | Run a formula runtime before reading |
| Must match Excel Online                              | Microsoft Graph calculate API        |
| Must match desktop Excel or add-ins                  | Excel or LibreOffice automation      |
| SheetJS-shaped workbook with supported formulas      | `xlsx-calc`                          |
| Broad JS formula engine needed                       | HyperFormula                         |
| Service-owned formula state and restore proof needed | `@bilig/workpaper`                    |

## Related

- [Excel file as a Node calculation engine](excel-file-calculation-engine-node.md)
- [Microsoft Graph Excel recalculation in Node.js](microsoft-graph-excel-recalculation-node.md)
- [XLSX formula recalculation in Node.js](xlsx-formula-recalculation-node.md)
- [xlsx-calc alternative for Node workbook recalculation](xlsx-calc-alternative-node-workbook-recalculation.md)
- [ExcelJS formula recalculation in Node.js](exceljs-formula-recalculation-node.md)
- [Compatibility limits](where-bilig-is-not-excel-compatible-yet.md)
- [Formula bug clinic](formula-bug-clinic.md)
- [Formula clinic report script](formula-clinic-report.ts)

If this saved you from trusting a stale cached formula value, star or bookmark
the repo:
<https://github.com/proompteng/bilig/stargazers>.
