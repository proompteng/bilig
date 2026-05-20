---
title: SheetJS formula result not updating in Node.js
published: true
description: What to do when SheetJS or xlsx edits an input cell but a formula cell still returns the old cached value in Node.js.
tags: typescript, node, sheetjs, xlsx, formulas, recalculation
canonical_url: https://proompteng.github.io/bilig/sheetjs-formula-result-not-updating-node.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# SheetJS formula result not updating in Node.js

This page is for the exact SheetJS / `xlsx` failure mode where a Node service
loads or creates an `.xlsx`, changes an input cell, then reads a formula cell
and still sees the old cached result.

Short answer: keep SheetJS for file I/O, but add a recalculation step before
you trust formula readback.

## Why the value is stale

XLSX formula cells can carry both formula text and a cached value. SheetJS cell
objects expose formula text and cell values, but the Community Edition path is
file-centric. If your process changes `Inputs!B2`, the cached value in
`Summary!B7` does not become fresh just because the source cell changed.

That is fine when Excel, LibreOffice, or another spreadsheet application will
open the workbook later and calculate it. It is not fine when a backend route,
queue worker, or test must make a decision from the computed value in the same
Node process.

## Use a narrow recalculation bridge

If your app already has XLSX bytes from SheetJS, use the SheetJS-named
recalculation bridge at the boundary. It keeps SheetJS responsible for file I/O
and adds only the missing recalculation/readback step:

```sh
npm install @bilig/sheetjs-formula-recalc
```

One-off proof:

```sh
npx --package @bilig/sheetjs-formula-recalc sheetjs-recalc --demo --json
```

Exact reproduction for the high-view Stack Overflow question:

```sh
git clone https://github.com/proompteng/bilig.git
cd bilig
npm --prefix examples/recalc-bridge-workflows install
npm --prefix examples/recalc-bridge-workflows run so:sheetjs-63085785
```

That script mirrors the small case from
[How to recalculate all formulas in excel file through Javascript?](https://stackoverflow.com/questions/63085785/how-to-recalculate-all-formulas-in-excel-file-through-javascript):
`A1` changes from `1` to `3`, SheetJS still has the stale cached `C1 = 3`,
then `@bilig/sheetjs-formula-recalc` verifies `C1 = 5`.

For a real workbook:

```sh
npx --package @bilig/sheetjs-formula-recalc sheetjs-recalc pricing.xlsx \
  --set Inputs!B2=48 \
  --set Inputs!B3=1500 \
  --read Summary!B7 \
  --out pricing.recalculated.xlsx \
  --json
```

The command writes an updated workbook and prints the values read after
recalculation.

## Minimal API path

```ts
import { readFile, writeFile } from 'node:fs/promises'
import { recalculateSheetjsWorkbook } from '@bilig/sheetjs-formula-recalc'

const result = recalculateSheetjsWorkbook(await readFile('pricing.xlsx'), {
  fileName: 'pricing.xlsx',
  edits: [
    { target: 'Inputs!B2', value: 48 },
    { target: 'Inputs!B3', value: 1500 },
  ],
  reads: ['Summary!B7'],
})

await writeFile('pricing.recalculated.xlsx', result.xlsx)

console.log({
  value: result.reads['Summary!B7'],
  warnings: result.warnings,
  verified: result.warnings.length === 0,
})
```

The important rule is that your service owns the input cells, output cells, and
verification checks. Do not read an arbitrary formula value and assume it is
fresh just because an XLSX writer succeeded.

## Proof against the common incumbents

The repository includes a bridge proof that edits the same workbook through
SheetJS/`xlsx`, `xlsx-populate`, and ExcelJS, then verifies that Bilig refreshes
the stale `48000` result to `72000` in all three paths:

```sh
git clone https://github.com/proompteng/bilig.git
cd bilig
npm --prefix examples/recalc-bridge-workflows install
npm --prefix examples/recalc-bridge-workflows run smoke
```

Use that example when you are deciding whether to keep your current file
library and add recalculation, instead of rewriting the whole workbook pipeline.

## Decision table

| Job | Use |
| --- | --- |
| Read or write many spreadsheet formats | SheetJS / `xlsx` |
| Generate a styled XLSX report for a human to open later | SheetJS, ExcelJS, or `xlsx-populate` |
| Ask Excel to recalculate when someone opens the file | workbook calc properties or Excel itself |
| Recalculate SheetJS / `xlsx` bytes inside Node after changing inputs | `@bilig/sheetjs-formula-recalc` |
| Recalculate generic XLSX bytes from another writer | `@bilig/xlsx-formula-recalc` |
| Keep an ExcelJS workbook and add fresh formula readback | `@bilig/exceljs-formula-recalc` |
| Own formula-backed workbook state as JSON in a service | `@bilig/workpaper` |
| Need commercial SheetJS formula calculation support | evaluate SheetJS Pro |

## Production checks

Before using this on a critical path, keep fixtures for:

- the exact workbook template your service receives or emits
- every input cell your code writes
- every output cell your code reads
- unsupported formulas and import warnings
- exported workbook reimport
- an Excel or LibreOffice oracle check for representative customer files

Bilig is not full Excel. The useful promise is narrower: a Node process can
edit known input cells, recalculate supported formulas, read back known output
cells, and export a workbook with tests around that boundary.

## Related proof

- [XLSX formula recalculation in Node.js](xlsx-formula-recalculation-node.md)
- [Stale XLSX formula cache in Node.js](stale-xlsx-formula-cache-node.md)
- [SheetJS and ExcelJS boundary guide](sheetjs-exceljs-alternative-formula-workbook-api.md)
- [xlsx-populate formula results in Node.js](xlsx-populate-formula-result-node.md)
- [ExcelJS formula recalculation in Node.js](exceljs-formula-recalculation-node.md)
- [@bilig/xlsx-formula-recalc package](https://www.npmjs.com/package/@bilig/xlsx-formula-recalc)
- [@bilig/sheetjs-formula-recalc package](https://www.npmjs.com/package/@bilig/sheetjs-formula-recalc)
- [SheetJS, xlsx-populate, and ExcelJS bridge example](https://github.com/proompteng/bilig/tree/main/examples/recalc-bridge-workflows)

If this saves you from opening Excel in a backend job just to refresh formula
values, star the repository so the fix is easier for the next SheetJS user to
find: <https://github.com/proompteng/bilig/stargazers>.

## Sources

- SheetJS Cell Objects:
  <https://docs.sheetjs.com/docs/csf/cell/>
- SheetJS Formulae:
  <https://docs.sheetjs.com/docs/csf/features/formulae>
- SheetJS Parse Options:
  <https://docs.sheetjs.com/docs/api/parse-options>
- `xlsx` npm package:
  <https://www.npmjs.com/package/xlsx>
- `@bilig/sheetjs-formula-recalc` npm package:
  <https://www.npmjs.com/package/@bilig/sheetjs-formula-recalc>
- `@bilig/xlsx-formula-recalc` npm package:
  <https://www.npmjs.com/package/@bilig/xlsx-formula-recalc>
