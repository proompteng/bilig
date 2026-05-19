# XLSX Formula Recalculation

Use `xlsx-formula-recalc` when an agent or Node.js service has changed XLSX
inputs and must read recalculated formula outputs without opening Excel,
LibreOffice, or a browser.

## First Check

```sh
npx --package xlsx-formula-recalc xlsx-recalc --demo --json
```

The demo should print `verified: true` and a `Summary!B2` value of `72000`.

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
