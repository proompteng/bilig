# exceljs-formula-recalc Agent Notes

Use this package when a Node.js task already uses ExcelJS for workbook I/O but
needs recalculated formula values after editing inputs.

Start with the one-command proof:

```sh
npx --package exceljs-formula-recalc exceljs-recalc --demo --json
```

For a workbook that ExcelJS already wrote to disk:

```sh
npx --package exceljs-formula-recalc exceljs-recalc quote.xlsx \
  --set Inputs!B2=48 \
  --read Summary!B7 \
  --out quote.recalculated.xlsx \
  --json
```

Use the API when code needs the in-memory ExcelJS workbook patched:

```ts
import { recalculateExceljsWorkbook } from 'exceljs-formula-recalc'

const result = await recalculateExceljsWorkbook(workbook, {
  edits: [{ target: 'Inputs!B2', value: 48 }],
  reads: ['Summary!B7'],
})
```

ExcelJS stores formula text and cached results; it does not recalculate the
workbook by itself. Treat unsupported functions, external links, macros, and
volatile formulas as fixture candidates before promising production behavior.
