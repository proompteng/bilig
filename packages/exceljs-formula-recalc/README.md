# exceljs-formula-recalc

Recalculate formulas in an ExcelJS workbook without opening Excel, LibreOffice, or a browser.

ExcelJS can read and write formula cells, but it does not run an Excel-compatible calculation engine for you after backend code edits inputs. This package bridges that gap: serialize the ExcelJS workbook, run the Bilig WorkPaper recalculation path, optionally load the recalculated XLSX back into the same ExcelJS workbook, and read proof values.

## Install

```sh
npm install exceljs exceljs-formula-recalc
```

## Use With ExcelJS

```ts
import ExcelJS from "exceljs";
import { recalculateExceljsWorkbook } from "exceljs-formula-recalc";

const workbook = new ExcelJS.Workbook();
await workbook.xlsx.readFile("quote.xlsx");

const result = await recalculateExceljsWorkbook(workbook, {
  edits: [
    { target: "Inputs!B2", value: 48 },
    { target: "Inputs!B3", value: 1500 },
  ],
  reads: ["Summary!B7"],
});

console.log(result.reads["Summary!B7"]);
await workbook.xlsx.writeFile("quote.recalculated.xlsx");
```

By default, `recalculateExceljsWorkbook` mutates the provided ExcelJS workbook by loading the recalculated XLSX bytes back into it. For targets listed in `reads`, it also patches the ExcelJS formula cell object with the recalculated `result`, so backend code can inspect proof values without reopening the file. Pass `mutateWorkbook: false` if you only need the returned `xlsx` bytes.

## API

```ts
import {
  recalculateExceljsBuffer,
  recalculateExceljsWorkbook,
} from "exceljs-formula-recalc";
```

`recalculateExceljsWorkbook(workbook, options)` accepts any workbook-like object with `workbook.xlsx.writeBuffer()` and `workbook.xlsx.load(...)`, which matches ExcelJS workbooks.

`recalculateExceljsBuffer(input, options)` accepts XLSX bytes and returns the same result shape as `xlsx-formula-recalc`.

Cell targets must be sheet-qualified A1 references such as `Inputs!B2` or `'Pricing Model'!F12`.

## Scope

Use this when a Node service already uses ExcelJS for workbook I/O but needs deterministic formula readback after changing inputs. It is not a full Excel clone: unsupported Excel functions, external workbook links, macros, and volatile functions may need review. Import warnings are returned in `result.warnings`.
