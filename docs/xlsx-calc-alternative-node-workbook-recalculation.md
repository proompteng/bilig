---
title: xlsx-calc alternative for Node workbook recalculation
published: true
description: When xlsx-calc is enough, when it is not, and how to test formula recalculation in a Node service without opening Excel.
tags: typescript, node, xlsx, spreadsheet, formulas
canonical_url: https://proompteng.github.io/bilig/xlsx-calc-alternative-node-workbook-recalculation.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# xlsx-calc alternative for Node workbook recalculation

You probably got here because you have an `.xlsx` file, you changed an input
cell in Node, and now the cached formula result is wrong.

`xlsx-calc` can be the right answer. If your workbook is already a SheetJS
object and the formulas you use are in its supported set, the API is simple:
edit cells, call `XLSX_CALC(workbook)`, read the values.

Use `@bilig/workpaper` when the spreadsheet is not just a file. The usual case
is a backend decision path: quote approval, payout checks, import validation,
or an agent tool that needs to write inputs, recalculate, read outputs, and
save a state it can test again later.

## Quick choice

| You need                                                         | Start with         |
| ---------------------------------------------------------------- | ------------------ |
| Recalculate a supported formula set on a SheetJS workbook object | `xlsx-calc`        |
| Read and write lots of spreadsheet file formats                  | SheetJS            |
| Build styled `.xlsx` files                                       | ExcelJS or SheetJS |
| Keep a formula workbook as service state                         | `@bilig/workpaper`  |
| Read recalculated outputs before accepting a request             | `@bilig/workpaper`  |
| Persist JSON state and still import or export XLSX at the edge   | `@bilig/workpaper`  |

If you only need to refresh an existing XLSX file before your service returns,
try the file-level package before migrating workbook state:

```sh
npx --package @bilig/xlsx-formula-recalc xlsx-recalc pricing.xlsx \
  --set Inputs!B2=48 \
  --read Summary!B7 \
  --out pricing.recalculated.xlsx \
  --json
```

That is the whole distinction. `xlsx-calc` is a calculator over a workbook
object. Bilig is a workbook runtime with import/export at the edges.

## Node service recalculation path

```ts
import { readFile, writeFile } from 'node:fs/promises'
import { WorkPaper } from '@bilig/workpaper'
import { exportXlsx, importXlsx } from '@bilig/workpaper/xlsx'

const source = await readFile('pricing-model.xlsx')
const imported = importXlsx(source, 'pricing-model.xlsx')
const workbook = WorkPaper.buildFromSnapshot(imported.snapshot)

const inputs = workbook.getSheetId('Inputs')
const summary = workbook.getSheetId('Summary')
if (inputs === undefined || summary === undefined) {
  throw new Error('Expected Inputs and Summary sheets')
}

workbook.setCellContents({ sheet: inputs, row: 1, col: 1 }, 48)
workbook.setCellContents({ sheet: inputs, row: 2, col: 1 }, 1250)

const decision = workbook.getCellValue({ sheet: summary, row: 6, col: 1 })
if (decision !== 'approved') {
  throw new Error(`Expected approved decision, got ${String(decision)}`)
}

const edited = exportXlsx(workbook.exportSnapshot())
await writeFile('pricing-model-edited.xlsx', edited)
```

The maintained example is small enough to inspect:

```sh
git clone https://github.com/proompteng/bilig.git
cd bilig/examples/xlsx-recalculation-node
npm install
npm start
```

It should end with checks like these:

```json
{
  "decisionChanged": true,
  "exportedReimportMatchesAfter": true,
  "formulasSurvivedXlsxRoundTrip": true,
  "verified": true
}
```

## Measured Lane

There is a checked-in `xlsx-calc` comparison, but it is deliberately narrow.
It covers four workbook-wide recalculation workloads: aggregate, exact-match
lookup, approximate lookup, and formula-chain recalculation.

Current artifact:
[`packages/benchmarks/baselines/workpaper-vs-xlsx-calc.json`](https://github.com/proompteng/bilig/blob/main/packages/benchmarks/baselines/workpaper-vs-xlsx-calc.json)

The artifact records:

- `xlsx-calc` version: `0.9.2`
- comparable workloads: `4`
- WorkPaper mean wins: `4/4`
- WorkPaper p95 wins: `4/4`
- WorkPaper mean-and-p95 wins: `4/4`
- coverage note: this is a limited SheetJS-style workbook-wide comparison, not
  a claim about every spreadsheet workload

Run the local check with:

```sh
pnpm workpaper:bench:xlsx-calc:check
```

That benchmark does not mean "Bilig replaces Excel." It only says that this
particular Node recalculation lane is measured, checked in, and easy to rerun.

## What I Would Test

For a production service, I would not ship this based on one happy-path example.
I would add tests for:

- exact input cells the service writes
- exact output cells the service reads
- exported XLSX reimport
- formulas preserved after the round trip
- unsupported formulas and compatibility limits
- fixtures that look like your real workbooks
- stale cached values in source XLSX files

If the service mostly formats files, keep using a file library. If the service
acts on the calculated result, put the cells behind a small adapter and test
that adapter like business logic.

## Sources and related pages

- `xlsx-calc` repository:
  <https://github.com/fabiooshiro/xlsx-calc>
- SheetJS formula docs:
  <https://docs.sheetjs.com/docs/csf/features/formulae/>
- [Excel file as a calculation engine in Node.js](excel-file-calculation-engine-node.md)
- [XLSX formula recalculation in Node.js](xlsx-formula-recalculation-node.md)
- [SheetJS and ExcelJS boundary](sheetjs-exceljs-alternative-formula-workbook-api.md)
- [Headless benchmark evidence](headless-workpaper-benchmark-evidence.md)
- [Compatibility limits](where-bilig-is-not-excel-compatible-yet.md)

If this saved you a spreadsheet-recalculation detour, star the repo so the next
Node developer can find it:
<https://github.com/proompteng/bilig/stargazers>.
