---
name: sheetjs-formula-recalc
version: 0.1.0
description: Add fresh formula readback to SheetJS and xlsx workflows after writing XLSX inputs in Node.js.
tags:
  - sheetjs
  - xlsx
  - excel
  - formula-recalculation
  - node
  - typescript
---

# SheetJS Formula Recalculation

Use `sheetjs-formula-recalc` when an agent or Node.js service has changed XLSX
inputs through SheetJS / `xlsx` and must read recalculated formula outputs
without opening Excel, LibreOffice, or a browser.

## First Check

```sh
npx --package sheetjs-formula-recalc sheetjs-recalc --demo --json
```

The demo should print `verified: true` and a `Summary!B2` value of `72000`.

## Real Workbook

```sh
npx --package sheetjs-formula-recalc sheetjs-recalc workbook.xlsx \
  --set Inputs!B2=48 \
  --read Summary!B7 \
  --out workbook.recalculated.xlsx \
  --json
```

## TypeScript

```ts
import { recalculateSheetjsWorkbook } from 'sheetjs-formula-recalc'

const result = recalculateSheetjsWorkbook(inputXlsxBytes, {
  edits: [{ target: 'Inputs!B2', value: 48 }],
  reads: ['Summary!B7'],
})
```

Use `xlsx-formula-recalc` when the caller does not care about SheetJS naming,
and use `exceljs-formula-recalc` when the caller owns an ExcelJS `Workbook`
object and wants read results patched back into that object.
