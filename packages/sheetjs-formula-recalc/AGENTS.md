# @bilig/sheetjs-formula-recalc Agent Notes

Use this package when a Node.js task already uses SheetJS or `xlsx` for workbook
file I/O but needs fresh formula values after editing input cells.

Start with the one-command proof:

```sh
npx --package @bilig/sheetjs-formula-recalc sheetjs-recalc --demo --json
```

For a real workbook, use sheet-qualified A1 targets:

```sh
npx --package @bilig/sheetjs-formula-recalc sheetjs-recalc quote.xlsx \
  --set Inputs!B2=48 \
  --read Summary!B7 \
  --out quote.recalculated.xlsx \
  --json
```

Use the API when code already has workbook bytes:

```ts
import { recalculateSheetjsWorkbook } from '@bilig/sheetjs-formula-recalc'

const result = recalculateSheetjsWorkbook(xlsxBytes, {
  edits: [{ target: 'Inputs!B2', value: 48 }],
  reads: ['Summary!B7'],
})
```

Do not claim this is a full Excel clone or a SheetJS replacement. Keep SheetJS
for file I/O, review `result.warnings`, and reduce unsupported functions,
external links, macros, and volatile formula cases into fixtures before
promising production behavior.
