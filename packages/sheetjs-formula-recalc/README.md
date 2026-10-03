# @bilig/sheetjs-formula-recalc

Recalculate SheetJS / `xlsx` formula workbooks in Node.js without opening
Excel, LibreOffice, or a browser.

Use this package when the rest of your pipeline already uses SheetJS or the
`xlsx` package for workbook file I/O, but a backend job needs fresh formula
readback after changing inputs.

## If You Arrived From a SheetJS Formula Issue

SheetJS is good at reading and writing spreadsheet files. The common production
gap is different:

- `SheetJS formula result not updating`
- `xlsx formula value stale after edit`
- `js-xlsx recalculate formulas`
- `refresh formula cells in xlsx node`

Formula cells can carry cached results. When a Node process edits `Inputs!B2`,
the cached value in `Summary!B7` is not automatically recalculated inside that
process.

Use this package at the file boundary:

1. let SheetJS produce or update XLSX bytes;
2. call `recalculateSheetjsWorkbook(...)`;
3. read proof cells from `result.reads`;
4. write `result.xlsx` if the updated artifact is needed.

This package is a SheetJS-named bridge over `@bilig/xlsx-formula-recalc`, so teams
searching for a SheetJS answer can find the right boundary directly.

If the workflow might belong at the generic XLSX, ExcelJS, template, CI,
WorkPaper, or agent boundary instead, use the
[stale formula readback chooser](https://proompteng.github.io/bilig/stale-formula-readback-chooser.html)
before installing another bridge.

## Install

```sh
npm install @bilig/sheetjs-formula-recalc
```

## CLI

Run a self-contained proof first:

```sh
npx --package @bilig/sheetjs-formula-recalc sheetjs-recalc --demo --json
```

For a real workbook:

```sh
npx --package @bilig/sheetjs-formula-recalc sheetjs-recalc quote.xlsx \
  --set Inputs!B2=48 \
  --set Inputs!B3=1500 \
  --read Summary!B7 \
  --out quote.recalculated.xlsx \
  --json
```

The command writes the recalculated XLSX and prints the requested read cells.

## TypeScript

```ts
import { readFile, writeFile } from 'node:fs/promises'
import { recalculateSheetjsWorkbook } from '@bilig/sheetjs-formula-recalc'

const result = recalculateSheetjsWorkbook(await readFile('quote.xlsx'), {
  fileName: 'quote.xlsx',
  edits: [
    { target: 'Inputs!B2', value: 48 },
    { target: 'Inputs!B3', value: 1500 },
  ],
  reads: ['Summary!B7'],
})

await writeFile('quote.recalculated.xlsx', result.xlsx)

console.log({
  value: result.reads['Summary!B7'],
  warnings: result.warnings,
})
```

## Proof Against SheetJS, xlsx-populate, and ExcelJS

The repository includes a cross-library proof:

```sh
git clone https://github.com/proompteng/bilig.git
cd bilig
npm --prefix examples/recalc-bridge-workflows install
npm --prefix examples/recalc-bridge-workflows run smoke
```

It edits the same workbook through SheetJS/`xlsx`, `xlsx-populate`, and
ExcelJS, then verifies that Bilig refreshes the stale `48000` result to
`72000`.

## What This Is Not

This is not a full Excel clone and not a replacement for SheetJS file I/O. Keep
SheetJS where it is strongest: parsing, writing, and transforming workbook
files. Add this package only where the Node process must own recalculated
formula readback before accepting, rejecting, returning, or persisting a
workflow.

Review `result.warnings` and keep fixtures for unsupported functions, external
workbook links, macros, volatile functions, and customer-critical templates.

## After The Proof

If the demo or a reduced SheetJS workbook returns the fresh formula values you
needed, keep the repository and release feed nearby:
<https://github.com/proompteng/bilig>.

Watch releases if this is close to a production workflow:
<https://github.com/proompteng/bilig/subscription>.

If a SheetJS workbook still fails after a reduced test case, open the blocker
with the formula, input cells, output cells, and warning output:
<https://github.com/proompteng/bilig/discussions/new?category=general>.

## Related

- [SheetJS formula result not updating in Node.js](https://proompteng.github.io/bilig/sheetjs-formula-result-not-updating-node.html)
- [XLSX formula recalculation in Node.js](https://proompteng.github.io/bilig/xlsx-formula-recalculation-node.html)
- [SheetJS and ExcelJS boundary guide](https://proompteng.github.io/bilig/sheetjs-exceljs-alternative-formula-workbook-api.html)
- [`@bilig/xlsx-formula-recalc`](https://www.npmjs.com/package/xlsx-formula-recalc)
- [`@bilig/workpaper`](https://www.npmjs.com/package/@bilig/workpaper)
