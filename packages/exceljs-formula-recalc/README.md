# @bilig/exceljs-formula-recalc

Recalculate formulas in an ExcelJS workbook without opening Excel, LibreOffice, or a browser.

ExcelJS can read and write formula cells, but it does not run an Excel-compatible calculation engine for you after backend code edits inputs. This package bridges that gap: serialize the ExcelJS workbook, run the Bilig WorkPaper recalculation path, optionally load the recalculated XLSX back into the same ExcelJS workbook, and read proof values.

## If You Arrived From an ExcelJS Formula Issue

This package is for the recurring ExcelJS boundary behind issues and searches
like:

- `ExcelJS formula result not updating`
- `Formula based cell not updating`
- `Updating Formula Result`
- `Automatic formula calculation`
- `get computed value of Excel sheet cell in Node.js`

The important distinction is:

- `workbook.calcProperties.fullCalcOnLoad = true` asks Excel or LibreOffice to
  recalculate later, when the file is opened.
- `@bilig/exceljs-formula-recalc` recalculates before your Node process returns, then
  patches requested ExcelJS formula cells with fresh `result` values.

Use ExcelJS for workbook I/O and presentation. Use this package only for the
calculation/readback boundary that ExcelJS intentionally does not own.

For a cross-library proof, run
[`examples/recalc-bridge-workflows`](../../examples/recalc-bridge-workflows).
It edits the same workbook through SheetJS/`xlsx`, `xlsx-populate`, and
ExcelJS, then verifies that Bilig refreshes the stale formula result.

If the workflow might belong at the generic XLSX, SheetJS, template, CI,
WorkPaper, or agent boundary instead, use the
[stale formula readback chooser](https://proompteng.github.io/bilig/stale-formula-readback-chooser.html)
before adding an ExcelJS-specific bridge.

## Install

```sh
npm install exceljs @bilig/exceljs-formula-recalc
```

## CLI

If your ExcelJS workflow has already written an `.xlsx` file, the package also
ships an ExcelJS-named CLI for quick proof runs:

```sh
npx --package @bilig/exceljs-formula-recalc exceljs-recalc --demo --json
```

For a real workbook saved by ExcelJS:

```sh
npx --package @bilig/exceljs-formula-recalc exceljs-recalc quote.xlsx \
  --set Inputs!B2=48 \
  --set Inputs!B3=1500 \
  --read Summary!B7 \
  --out quote.recalculated.xlsx \
  --json
```

Use the API below when backend code needs the recalculated values patched back
onto the in-memory ExcelJS workbook object.

## Use With ExcelJS

```ts
import ExcelJS from 'exceljs'
import { recalculateExceljsWorkbook } from '@bilig/exceljs-formula-recalc'

const workbook = new ExcelJS.Workbook()
await workbook.xlsx.readFile('quote.xlsx')

const result = await recalculateExceljsWorkbook(workbook, {
  edits: [
    { target: 'Inputs!B2', value: 48 },
    { target: 'Inputs!B3', value: 1500 },
  ],
  reads: ['Summary!B7'],
})

console.log(result.reads['Summary!B7'])
await workbook.xlsx.writeFile('quote.recalculated.xlsx')
```

By default, `recalculateExceljsWorkbook` mutates the provided ExcelJS workbook by loading the recalculated XLSX bytes back into it. For targets listed in `reads`, it also patches the ExcelJS formula cell object with the recalculated `result`, so backend code can inspect proof values without reopening the file. Pass `mutateWorkbook: false` if you only need the returned `xlsx` bytes.

## Common Boundaries

| Job                                                             | Use                                     |
| --------------------------------------------------------------- | --------------------------------------- |
| Create styled XLSX reports, worksheets, tables, images, or rows | ExcelJS                                 |
| Ask Excel to recalculate after a human opens the file           | ExcelJS `fullCalcOnLoad`                |
| Read recalculated formula values before an API/job returns      | `@bilig/exceljs-formula-recalc`         |
| Recalculate raw XLSX bytes from SheetJS or xlsx-populate        | `@bilig/xlsx-formula-recalc`            |
| Keep formula-backed business state as JSON, not XLSX            | `@bilig/workpaper` or `bilig-workpaper` |

## API

```ts
import { recalculateExceljsBuffer, recalculateExceljsWorkbook } from '@bilig/exceljs-formula-recalc'
```

`recalculateExceljsWorkbook(workbook, options)` accepts any workbook-like object with `workbook.xlsx.writeBuffer()` and `workbook.xlsx.load(...)`, which matches ExcelJS workbooks.

`recalculateExceljsBuffer(input, options)` accepts XLSX bytes and returns the same result shape as `@bilig/xlsx-formula-recalc`.

Cell targets must be sheet-qualified A1 references such as `Inputs!B2` or `'Pricing Model'!F12`.

## After The Proof

If the demo or a reduced ExcelJS workbook gives the fresh formula values your
service needs, keep the repository and release feed nearby:
<https://github.com/proompteng/bilig>.

Watch releases if this package is close to a production path:
<https://github.com/proompteng/bilig/subscription>.

If a real ExcelJS workbook still blocks adoption, open the smallest fixture
with the formula, input cells, output cells, and warning output:
<https://github.com/proompteng/bilig/discussions/new?category=general>.

## Scope

Use this when a Node service already uses ExcelJS for workbook I/O but needs deterministic formula readback after changing inputs. It is not a full Excel clone: unsupported Excel functions, external workbook links, macros, and volatile functions may need review. Import warnings are returned in `result.warnings`.
