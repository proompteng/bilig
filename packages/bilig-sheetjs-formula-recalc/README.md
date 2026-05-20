# @bilig/sheetjs-formula-recalc

Scoped SheetJS formula recalculation package for Node.js without Excel, LibreOffice, or browser automation.

Use this package when the rest of your pipeline already uses SheetJS or the
`xlsx` package for workbook file I/O, but a backend job needs fresh formula
readback after changing inputs.

The unscoped `sheetjs-formula-recalc` package remains published as a
compatibility and search alias.

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

This package is a SheetJS-named bridge over `@bilig/xlsx-formula-recalc`, so
teams searching for a SheetJS answer can find the right boundary directly.

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

Full docs: <https://proompteng.github.io/bilig/sheetjs-formula-result-not-updating-node.html>
