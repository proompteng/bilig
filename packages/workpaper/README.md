# @bilig/workpaper

Scoped Bilig WorkPaper runtime for Node.js services, agent tools, and server-side spreadsheet formulas.

Use this when business logic is easiest to review as workbook cells and
formulas, but the calculation needs to run in a backend service, queue worker,
serverless route, test, or coding-agent tool.

`@bilig/workpaper` is the canonical scoped npm entrypoint. The unscoped
`bilig-workpaper` package remains published as a compatibility and search alias.

## Install

```sh
npm install @bilig/workpaper
```

## Use A WorkPaper In Node

```ts
import { WorkPaper } from '@bilig/workpaper'

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

## Prove The Agent Loop Without Cloning

The package ships proof commands for coding agents and service evaluators:

```sh
npm exec --package @bilig/workpaper -- bilig-agent-challenge
npm exec --package @bilig/workpaper -- bilig-mcp-challenge
npm exec --package @bilig/workpaper -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable
```

The challenge commands edit one input, recalculate dependent formulas, export
WorkPaper JSON, restore it, and print a `verified: true` proof object.

## XLSX Import And Export

```ts
import { WorkPaper } from '@bilig/workpaper'
import { exportXlsx, importXlsx } from '@bilig/workpaper/xlsx'
```

Use `@bilig/xlsx-formula-recalc` when you only need to edit and recalculate
XLSX files. Use `@bilig/exceljs-formula-recalc` when you already use ExcelJS
and need recalculated formula results after changing inputs.

## Agent Commands And Optional MCP

The npm tarball exposes the same CLI entrypoints as `@bilig/headless`, so agents
can install one focused package and still get the MCP stdio server:

```ts
import { createWorkPaperMcpServer } from '@bilig/workpaper/mcp'
```

For a runnable starter project with `AGENTS.md`, MCP client config, and an
`agent:verify` script:

```sh
npm create @bilig/workpaper@latest pricing-agent -- --agent
```

## Scope

Bilig is not a desktop Excel clone. It is a formula workbook runtime for
service-owned calculations, JSON persistence, XLSX import/export, and
agent-readable readback. Unsupported Excel functions, external workbook links,
macros, and volatile functions may need review.

Full docs: <https://proompteng.github.io/bilig/>
