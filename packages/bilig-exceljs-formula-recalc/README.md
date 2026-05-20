# @bilig/exceljs-formula-recalc

Scoped ExcelJS formula recalculation package for Node.js without Excel, LibreOffice, or browser automation.

ExcelJS can read and write formula cells, but it does not run an
Excel-compatible calculation engine for you after backend code edits inputs.
This package bridges that gap: serialize the ExcelJS workbook, run the Bilig
WorkPaper recalculation path, optionally load the recalculated XLSX back into
the same ExcelJS workbook, and read proof values.

The unscoped `exceljs-formula-recalc` package remains published as a
compatibility and search alias.

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
- `@bilig/exceljs-formula-recalc` recalculates before your Node process returns,
  then patches requested ExcelJS formula cells with fresh `result` values.

Use ExcelJS for workbook I/O and presentation. Use this package only for the
calculation/readback boundary that ExcelJS intentionally does not own.

## Install

```sh
npm install @bilig/exceljs-formula-recalc
```

If you do not already have ExcelJS in the project:

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

By default, `recalculateExceljsWorkbook` mutates the provided ExcelJS workbook
by loading the recalculated XLSX bytes back into it. For targets listed in
`reads`, it also patches the ExcelJS formula cell object with the recalculated
`result`, so backend code can inspect proof values without reopening the file.
Pass `mutateWorkbook: false` if you only need the returned `xlsx` bytes.

## Common Boundaries

| Job                                                             | Use                             |
| --------------------------------------------------------------- | ------------------------------- |
| Create styled XLSX reports, worksheets, tables, images, or rows | ExcelJS                         |
| Ask Excel to recalculate after a human opens the file           | ExcelJS `fullCalcOnLoad`        |
| Read recalculated formula values before an API/job returns      | `@bilig/exceljs-formula-recalc` |
| Recalculate raw XLSX bytes from SheetJS or xlsx-populate        | `@bilig/xlsx-formula-recalc`    |
| Keep formula-backed business state as JSON, not XLSX            | `@bilig/workpaper`              |

## API

```ts
import { recalculateExceljsBuffer, recalculateExceljsWorkbook } from '@bilig/exceljs-formula-recalc'
```

`recalculateExceljsWorkbook(workbook, options)` accepts any workbook-like object
with `workbook.xlsx.writeBuffer()` and `workbook.xlsx.load(...)`, which matches
ExcelJS workbooks.

`recalculateExceljsBuffer(input, options)` accepts XLSX bytes and returns the
same result shape as `@bilig/xlsx-formula-recalc`.

Cell targets must be sheet-qualified A1 references such as `Inputs!B2` or
`'Pricing Model'!F12`.

## Scope

Use this when a Node service already uses ExcelJS for workbook I/O but needs
deterministic formula readback after changing inputs. It is not a full Excel
clone: unsupported Excel functions, external workbook links, macros, and
volatile functions may need review. Import warnings are returned in
`result.warnings`.

Full docs: <https://proompteng.github.io/bilig/exceljs-formula-recalculation-node.html>
