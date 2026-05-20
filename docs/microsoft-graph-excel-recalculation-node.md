---
title: Microsoft Graph Excel recalculation vs local Node WorkPaper
published: true
description: Decide when a Node service should use Microsoft Graph Excel calculation, LibreOffice automation, or @bilig/workpaper for formula-backed workbook outputs.
tags: typescript, node, excel, microsoft-graph, xlsx, formulas
canonical_url: https://proompteng.github.io/bilig/microsoft-graph-excel-recalculation-node.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Microsoft Graph Excel recalculation vs local Node WorkPaper

Microsoft Graph is a legitimate way to get Excel itself to calculate a workbook
from a Node service. It is also a heavy boundary if the service only needs to
write known inputs, recalculate known formulas, and read known outputs.

Use this page when the question is: "Should I upload the workbook to OneDrive or
SharePoint and ask Excel Online to recalculate it, or should the service own the
workbook runtime locally?"

## Short answer

Use Microsoft Graph when exact Excel Online behavior is the requirement, the
workbook already lives in OneDrive or SharePoint, and the auth/storage boundary
is acceptable.

Use LibreOffice or desktop Excel automation when the workbook depends on desktop
Excel behavior, add-ins, or manual Excel compatibility checks.

Use `@bilig/workpaper` when the workbook is service-owned state: write stable
input cells, recalculate in Node, read output cells, persist JSON, and optionally
import or export XLSX at the boundary.

## What Graph gives you

The Microsoft Graph Excel calculate endpoint recalculates currently opened
workbooks in Excel:

```http
POST /me/drive/items/{id}/workbook/application/calculate
Content-Type: application/json

{
  "calculationType": "FullRebuild"
}
```

Microsoft documents `Recalculate`, `Full`, and `FullRebuild` calculation types.
As of the current Graph v1.0 docs, the endpoint requires delegated work or
school account access with `Files.ReadWrite`; personal Microsoft accounts and
application permissions are not supported for that API.

That makes Graph a good fit when the file is already a Microsoft 365 artifact.
It is a worse fit when every request has to upload a temporary workbook, wait
for Excel Online, read values back, and clean up the file.

## The service-boundary question

Ask this before choosing the runtime:

| Requirement                                                   | Better first choice             |
| ------------------------------------------------------------- | ------------------------------- |
| Must match Excel Online calculation behavior                  | Microsoft Graph                 |
| Workbook is already in SharePoint or OneDrive                 | Microsoft Graph                 |
| Needs desktop Excel, add-ins, or macro-adjacent behavior      | Excel or LibreOffice automation |
| Needs a local deterministic Node decision path                | `@bilig/workpaper`               |
| Needs JSON persistence and restore tests                      | `@bilig/workpaper`               |
| Needs broad mature formula coverage more than WorkPaper state | HyperFormula                    |
| Needs XLSX read/write/styling, not calculation ownership      | SheetJS or ExcelJS              |

The wrong design is pretending an `.xlsx` writer is a calculator. Libraries such
as ExcelJS can preserve formula objects and cached results, but they do not
calculate new formula results for you.

## Minimal local WorkPaper shape

```ts
import { readFile, writeFile } from 'node:fs/promises'
import { WorkPaper } from '@bilig/workpaper'
import { exportXlsx, importXlsx } from '@bilig/workpaper/xlsx'

const imported = importXlsx(await readFile('quote-model.xlsx'), 'quote-model.xlsx')
const workbook = WorkPaper.buildFromSnapshot(imported.snapshot)

const inputs = workbook.getSheetId('Inputs')
const quote = workbook.getSheetId('Quote')
if (inputs === undefined || quote === undefined) {
  throw new Error('Expected Inputs and Quote sheets')
}

workbook.setCellContents({ sheet: inputs, row: 1, col: 1 }, 42_000)
workbook.setCellContents({ sheet: inputs, row: 2, col: 1 }, 0.18)

const decision = workbook.getCellValue({ sheet: quote, row: 8, col: 1 })
const auditJson = workbook.exportSnapshot()
const editedXlsx = exportXlsx(auditJson)

await writeFile('quote-model-edited.xlsx', editedXlsx)
console.log({ decision })
```

In production, keep the cell contract boring: name the sheets, input cells, and
output cells in one adapter, then test those fixtures. Do not let arbitrary
workbook layout become invisible backend logic.

## When Graph is still the right answer

Keep Graph in the loop if any of these are true:

- users already collaborate on the workbook in Microsoft 365;
- exact Excel Online formula behavior matters more than local determinism;
- the workbook uses functions Bilig does not support yet;
- your compliance model wants the workbook to stay in Microsoft storage;
- the service already has delegated Microsoft Graph auth and file lifecycle
  cleanup.

For those cases, Bilig can still be useful as a local fixture or migration path,
but it should not be presented as a drop-in Excel Online replacement.

## Related

- Microsoft Graph calculate endpoint:
  <https://learn.microsoft.com/en-us/graph/api/workbookapplication-calculate>
- ExcelJS formula note:
  <https://www.npmjs.com/package/exceljs/v/3.3.0>
- [Excel file as a Node calculation engine](excel-file-calculation-engine-node.md)
- [ExcelJS formula recalculation in Node.js](exceljs-formula-recalculation-node.md)
- [XLSX formula recalculation in Node.js](xlsx-formula-recalculation-node.md)
- [Compatibility limits](where-bilig-is-not-excel-compatible-yet.md)

If this saves you from building a temporary OneDrive calculation worker, star or
bookmark the repo:
<https://github.com/proompteng/bilig/stargazers>.
