# @bilig/excel-import

CSV/XLSX-to-`WorkbookSnapshot` import helpers and supported-subset XLSX export helpers for bilig.

## Package Status

This package is part of the published `bilig` runtime npm package set. Install
it with `@bilig/headless` when a Node project needs XLSX import, WorkPaper
calculation, and XLSX export from the same public package path.

```sh
pnpm add @bilig/headless @bilig/excel-import
```

Repository development:

```sh
pnpm install
pnpm --filter @bilig/excel-import build
pnpm exec vitest run packages/excel-import/src/__tests__/excel-import.test.ts
```

## XLSX To WorkPaper

```ts
import { readFileSync, writeFileSync } from 'node:fs'
import { WorkPaper } from '@bilig/headless'
import { exportXlsx, importXlsx } from '@bilig/excel-import'

const imported = importXlsx(new Uint8Array(readFileSync('model.xlsx')), 'model.xlsx')
const workbook = WorkPaper.buildFromSnapshot(imported.snapshot, {
  evaluationTimeoutMs: 30_000,
  useColumnIndex: true,
})

const firstSheetName = imported.snapshot.sheets[0]?.name
const firstSheet = firstSheetName === undefined ? undefined : workbook.getSheetId(firstSheetName)
if (firstSheet === undefined) throw new Error('Workbook has no sheets')

workbook.setCellContents({ sheet: firstSheet, row: 1, col: 1 }, 150_000)
const recalculated = workbook.getCellDisplayValue({ sheet: firstSheet, row: 1, col: 1 })

writeFileSync('model-edited.xlsx', exportXlsx(workbook.exportSnapshot()))
workbook.dispose()

console.log({ recalculated })
```

Use `WorkPaper.buildFromSnapshot()` for imported XLSX files. It preserves the
workbook metadata that Excel formulas need, including defined names, table
metadata, and structured-reference translations. `WorkPaper.buildFromSheets()`
is intentionally metadata-free. Use `workbook.exportSnapshot()` with
`exportXlsx()` when exporting a WorkPaper after edits.

Literal Excel error cells such as `#N/A`, `#DIV/0!`, `#REF!`, and `#VALUE!`
are imported as their display text instead of SheetJS numeric error codes.

Hidden XLSX rows and columns are preserved in workbook metadata and exported
back into worksheet XML, including hidden column states attached to width
metadata.

Workbook calculation properties such as iterative calculation, iteration count
and delta, forced recalculation, concurrent calculation, and manual calculation
mode are preserved from XLSX `<calcPr>` metadata on roundtrip.

Worksheet protection elements preserve non-default XML attributes from source
workbooks, so protected sheets are not normalized into a different
`<sheetProtection sheet="1"/>` state during no-op roundtrips.

Worksheet printer settings preserve binary `xl/printerSettings/*.bin` parts,
worksheet relationships, and `pageSetup` relationship links during no-op XLSX
roundtrips.

Worksheet `<sheetPr>` properties preserve non-tabColor code names,
`outlinePr`, and `pageSetUpPr` metadata during no-op XLSX roundtrips.

Workbook sheet visibility preserves hidden and very hidden worksheet state
during no-op XLSX roundtrips.

Cell hyperlinks preserve external URL targets, internal workbook targets,
tooltips, and display text during no-op XLSX roundtrips.

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
