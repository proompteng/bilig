# @bilig/excel-import

CSV/XLSX-to-`WorkbookSnapshot` import helpers and supported-subset XLSX export helpers for bilig.

## Package Status

This package is part of the `bilig` monorepo runtime package set, but the
`@bilig/excel-import` npm name is not provisioned yet. Use it from a repository
checkout for now. The external npm install path will be documented here after
the package is published.

From the repository:

```sh
pnpm install
pnpm --filter @bilig/excel-import build
pnpm exec vitest run packages/excel-import/src/__tests__/excel-import.test.ts
```

## XLSX To WorkPaper

```ts
import { readFileSync } from 'node:fs'
import { WorkPaper } from '@bilig/headless'
import { importXlsx } from '@bilig/excel-import'

const imported = importXlsx(new Uint8Array(readFileSync('model.xlsx')), 'model.xlsx')

const workbook = WorkPaper.buildFromSnapshot(imported.snapshot, {
  evaluationTimeoutMs: 30_000,
  useColumnIndex: true,
})
```

Use `WorkPaper.buildFromSnapshot()` for imported XLSX files. It preserves the
workbook metadata that Excel formulas need, including defined names, table
metadata, and structured-reference translations. `WorkPaper.buildFromSheets()`
is intentionally metadata-free.

Literal Excel error cells such as `#N/A`, `#DIV/0!`, `#REF!`, and `#VALUE!`
are imported as their display text instead of SheetJS numeric error codes.

## CSV Import

```ts
import { importCsv } from '@bilig/excel-import'

const imported = importCsv('Account;Amount\n4000;125,50', 'ledger.csv', {
  delimiter: ';',
  decimalSeparator: ',',
})
```

CSV import auto-detects comma, semicolon, and tab delimiters. Semicolon and tab
exports that contain decimal-comma values are parsed as locale accounting CSV by
default; pass explicit options when the source format is known. Integer-looking
fields with leading zeros are kept as text so account numbers, routing numbers,
invoice IDs, and similar identifiers are not silently changed.
