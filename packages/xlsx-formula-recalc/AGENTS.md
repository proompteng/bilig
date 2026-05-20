# xlsx-formula-recalc Agent Notes

Use this package when a Node.js task edits an `.xlsx` workbook and needs fresh
formula results before returning the file or reading output cells.

Start with the one-command proof:

```sh
npx --package xlsx-formula-recalc xlsx-recalc --demo --json
```

If the task came from a SheetJS / `xlsx` stale-formula issue, the same package
also exposes a SheetJS-named command:

```sh
npx --package xlsx-formula-recalc sheetjs-recalc --demo --json
```

For a real workbook, use sheet-qualified A1 targets:

```sh
npx --package xlsx-formula-recalc xlsx-recalc quote.xlsx \
  --set Inputs!B2=48 \
  --read Summary!B7 \
  --out quote.recalculated.xlsx \
  --json
```

Use the API when code already has workbook bytes:

```ts
import { recalculateXlsx } from 'xlsx-formula-recalc'

const result = recalculateXlsx(xlsxBytes, {
  edits: [{ target: 'Inputs!B2', value: 48 }],
  reads: ['Summary!B7'],
})
```

Do not claim this is a full Excel clone. Review `result.warnings` and reduce
unsupported functions, external links, macros, and volatile formula cases into
fixtures before promising production behavior.
