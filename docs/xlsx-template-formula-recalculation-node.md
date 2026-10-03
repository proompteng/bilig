---
title: xlsx-template formula recalculation in Node.js
published: true
description: How to handle formulas after xlsx-template substitutions when a Node.js service needs recalculated values without opening Excel.
tags: typescript, node, xlsx-template, xlsx, formulas, recalculation
canonical_url: https://proompteng.github.io/bilig/xlsx-template-formula-recalculation-node.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# xlsx-template formula recalculation in Node.js

`xlsx-template` is good at one job: take an XLSX template, substitute values,
and write the workbook back out.

It does not turn the generated workbook into a freshly calculated model. If a
template writes `Sheet1!A1` and another sheet has `=Sheet1!A1`, the generated
file can still require Excel, LibreOffice, or another formula engine before an
API can trust the formula result.

That is not a templating bug. It is a calculation-boundary problem.

## The production split

Keep the workflow explicit:

1. Use `xlsx-template` to fill the XLSX template.
2. Run a formula runtime only for the cells that the backend must read before a
   human opens the file.
3. Add a reduced fixture for the exact input cells, output cells, and formulas.
4. Export XLSX only after the backend decision has already been made from a
   verified readback path.

For tiny formula sets, `xlsx-calc` may be enough. For exact Excel behavior, use
Excel, LibreOffice, or Microsoft Graph. Use `@bilig/workpaper` when the service
can own the workbook state locally and needs write, recalculate, readback,
JSON persistence, and restore checks.

## Bilig-shaped check

After `xlsx-template` writes `result.xlsx`, import the workbook into WorkPaper
and read the output cells that matter to the API:

```ts
import { readFileSync } from 'node:fs'
import { WorkPaper } from '@bilig/workpaper'
import { importXlsx } from '@bilig/workpaper/xlsx'

const imported = importXlsx(new Uint8Array(readFileSync('result.xlsx')), 'result.xlsx')
const workbook = WorkPaper.buildFromSnapshot(imported.snapshot, {
  evaluationTimeoutMs: 30_000,
  useColumnIndex: true,
})

try {
  const sheetId = workbook.getSheetId('Sheet2')
  if (sheetId === undefined) {
    throw new Error('Missing Sheet2')
  }

  const value = workbook.getCellDisplayValue({ sheet: sheetId, row: 1, col: 1 })
  console.log({ outputCell: 'Sheet2!B2', value })
} finally {
  workbook.dispose()
}
```

The important part is the boundary: the API does not read stale XLSX cached
formula values and hope a later Excel open fixes them.

## Fixture-first debugging

Before changing a production template pipeline, reduce the workbook to:

- the template input cells;
- the formula output cells;
- the expected output from Excel, LibreOffice, or a manual check;
- any formula families that must be supported.

Then generate a local report without uploading the workbook:

```sh
curl -fsSLo formula-clinic-report.ts \
  https://proompteng.github.io/bilig/formula-clinic-report.ts
npx tsx formula-clinic-report.ts ./reduced-template-output.xlsx \
  --cells "Sheet2!B2"
```

Paste the report into the
[formula bug clinic](https://proompteng.github.io/bilig/formula-bug-clinic.html)
if the reduced case fails.

## When not to use Bilig

Do not use Bilig as a drop-in replacement for `xlsx-template`.

If your only deliverable is a styled XLSX file and a person will open it before
the value matters, keep the current template pipeline and set recalculation on
open. If exact Excel compatibility, macros, pivots, charts, or full formatting
fidelity are the product, keep Excel/LibreOffice/Microsoft Graph in the loop.

Use Bilig when the backend owns the calculation decision and stale cached XLSX
values are the thing you are trying to remove.

## Related

- [Fix stale XLSX formula values in Node.js](stale-xlsx-formula-cache-node.md)
- [XLSX formula recalculation in Node.js](xlsx-formula-recalculation-node.md)
- [Excel file as a Node calculation engine](excel-file-calculation-engine-node.md)
- [Submit a workbook fixture](submit-workbook-fixture.md)

Source issue:
<https://github.com/optilude/xlsx-template/issues/192>
