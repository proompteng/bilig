# bilig

Formula WorkPaper runtime for Node.js services, agent tools, and server-side spreadsheet formulas.

This is the short npm name for the Bilig headless runtime. Use it when business logic is easiest to review as workbook cells and formulas, but the calculation needs to run in a backend service, queue worker, serverless route, test, or coding-agent tool.

## Install

```sh
npm install bilig
```

## Use A WorkPaper In Node

```ts
import { WorkPaper } from 'bilig'

const workbook = WorkPaper.buildFromSheets({
  Inputs: [
    ['Metric', 'Value'],
    ['Units', 40],
    ['Price', 1200],
  ],
  Summary: [
    ['Metric', 'Value'],
    ['Revenue', '=Inputs!B2*Inputs!B3'],
  ],
})

const inputs = workbook.getSheetId('Inputs')
const summary = workbook.getSheetId('Summary')

if (inputs === undefined || summary === undefined) {
  throw new Error('Expected sheets to exist')
}

workbook.setCellContents({ sheet: inputs, row: 1, col: 1 }, 48)
workbook.setCellContents({ sheet: inputs, row: 2, col: 1 }, 1500)

console.log(workbook.getCellValue({ sheet: summary, row: 1, col: 1 }))
console.log(workbook.exportSnapshot())

workbook.dispose()
```

## XLSX Import And Export

```ts
import { WorkPaper } from 'bilig'
import { exportXlsx, importXlsx } from 'bilig/xlsx'
```

Use `xlsx-formula-recalc` when you only need to edit and recalculate XLSX files. Use `exceljs-formula-recalc` when you already use ExcelJS and need recalculated formula results after changing inputs.

## Agent Tools And MCP

```ts
import { createWorkPaperMcpServer } from 'bilig/mcp'
```

For a runnable starter, use:

```sh
npm create @bilig/workpaper
```

## Scope

Bilig is not a desktop Excel clone. It is a formula workbook runtime for service-owned calculations, JSON persistence, XLSX import/export, and agent-readable readback. Unsupported Excel functions, external workbook links, macros, and volatile functions may need review.

Full docs: <https://proompteng.github.io/bilig/>
