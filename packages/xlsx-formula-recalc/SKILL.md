---
name: xlsx-formula-recalc
version: 0.1.0
description: Recalculate XLSX formula outputs in Node.js after cell edits without opening Excel, LibreOffice, or browser automation.
tags:
  - xlsx
  - excel
  - formula-recalculation
  - node
  - spreadsheet-automation
  - typescript
---

# XLSX Formula Recalculation

Use `xlsx-formula-recalc` when an agent or Node.js service has changed XLSX
inputs and must read recalculated formula outputs without opening Excel,
LibreOffice, or a browser.

## First Check

```sh
npx --package xlsx-formula-recalc xlsx-recalc --demo --json
```

The demo should print `verified: true` and a `Summary!B2` value of `72000`.
For SheetJS / `xlsx` stale-formula issues, use the SheetJS-named binary from the
same package:

```sh
npx --package xlsx-formula-recalc sheetjs-recalc --demo --json
```

## Real Workbook

```sh
npx --package xlsx-formula-recalc xlsx-recalc workbook.xlsx \
  --set Inputs!B2=48 \
  --read Summary!B7 \
  --out workbook.recalculated.xlsx \
  --json
```

## TypeScript

```ts
import { recalculateXlsx } from 'xlsx-formula-recalc'

const result = recalculateXlsx(inputXlsxBytes, {
  edits: [{ target: 'Inputs!B2', value: 48 }],
  reads: ['Summary!B7'],
})
```

Prefer `exceljs-formula-recalc` when the caller already owns an ExcelJS
`Workbook` object and wants read results patched back into that object.
