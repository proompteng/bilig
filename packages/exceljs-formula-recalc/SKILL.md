---
name: exceljs-formula-recalc
version: 0.1.0
description: Recalculate formula outputs for ExcelJS workbook flows in Node.js after agents or services edit cells.
tags:
  - exceljs
  - xlsx
  - excel
  - formula-recalculation
  - node
  - typescript
---

# ExcelJS Formula Recalculation

Use `exceljs-formula-recalc` when an agent or Node.js service already uses
ExcelJS and needs fresh formula output values after changing cells.

## First Check

```sh
npx --package exceljs-formula-recalc exceljs-recalc --demo --json
```

The demo should print `verified: true` and a `Summary!B2` value of `72000`.

## Real Workbook

```sh
npx --package exceljs-formula-recalc exceljs-recalc workbook.xlsx \
  --set Inputs!B2=48 \
  --read Summary!B7 \
  --out workbook.recalculated.xlsx \
  --json
```

## TypeScript

```ts
import { recalculateExceljsWorkbook } from 'exceljs-formula-recalc'

const result = await recalculateExceljsWorkbook(workbook, {
  edits: [{ target: 'Inputs!B2', value: 48 }],
  reads: ['Summary!B7'],
})
```

Use `xlsx-formula-recalc` instead when the caller only has raw XLSX bytes and
does not need an ExcelJS workbook object.
