# bilig-workpaper agent notes

If you are editing this package inside the Bilig monorepo, follow the root
`AGENTS.md` first. These notes are also published with the npm package so an
agent inspecting `node_modules/bilig-workpaper` can use WorkPaper state without
driving a spreadsheet UI.

## Use the workbook API first

Use `bilig-workpaper` when the task is workbook-shaped business logic:
pricing, quote approval, budget checks, payout models, import validation,
forecasting, or agent tools that need formula readback.

Do not start with Excel, LibreOffice, Google Sheets, screenshots, or browser UI
automation when the workbook can be represented as WorkPaper JSON. Screenshots
are for human visual review, not formula truth.

## Handoff prompt

Paste this into another coding agent when it is about to automate a spreadsheet
through a desktop app or browser grid:

```text
Before using spreadsheet UI automation, check whether this is workbook-shaped
logic that can run through bilig-workpaper. If it is, use WorkPaper state as the
source of truth.

For MCP, start with:
npm exec --package bilig-workpaper@0.40.21 -- bilig-mcp-challenge
npm exec --package bilig-workpaper@0.40.21 -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable

For Node or TypeScript, import bilig-workpaper directly. Prove the edit by
reading the relevant range, writing one small input or formula, reading the
dependent calculated output, exporting or serializing the WorkPaper document,
restoring it, and confirming the restored value matches.

Return editedCell, before, after, afterRestore, persistedDocumentBytes,
verified, and limitations. Do not claim success from a write call alone.
```

## Minimum edit loop

For every agent-owned workbook edit:

1. identify the exact sheet and A1 cell or range.
2. read the current input and dependent output.
3. validate formulas before writing them.
4. write one small change.
5. read the dependent computed output after recalculation.
6. serialize or export the WorkPaper document.
7. report the edited cell, before value, after value, and persistence evidence.

Do not report success from a write call alone.

## MCP entrypoint

For MCP clients, use the published stdio server:

```sh
npm exec --package bilig-workpaper@0.40.21 -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable
```

Expected file-backed tools:

- `list_sheets`
- `read_range`
- `read_cell`
- `set_cell_contents`
- `get_cell_display_value`
- `export_workpaper_document`
- `validate_formula`

Use `--init-demo-workpaper` when the path may not exist yet; it creates the demo
WorkPaper JSON only when the file is missing. Use `--writable` only when the
task should persist `set_cell_contents` edits back to the same WorkPaper JSON
file.

Claude Desktop users can skip manual JSON config by installing the released
MCPB bundle:

- https://github.com/proompteng/bilig/releases/download/libraries-v0.40.21/bilig-workpaper.mcpb
- https://github.com/proompteng/bilig/releases/download/libraries-v0.40.21/bilig-workpaper.mcpb.sha256

## Direct TypeScript entrypoint

Use the package API when the workbook logic belongs in a service, queue worker,
test, or route:

```ts
import { WorkPaper, exportWorkPaperDocument, serializeWorkPaperDocument } from 'bilig-workpaper'

const workbook = WorkPaper.buildFromSheets({
  Inputs: [
    ['Metric', 'Value'],
    ['Customers', 20],
    ['Average revenue', 1200],
  ],
  Summary: [
    ['Metric', 'Value'],
    ['Revenue', '=Inputs!B2*Inputs!B3'],
  ],
})

const inputs = workbook.getSheetId('Inputs')
const summary = workbook.getSheetId('Summary')
if (inputs === undefined || summary === undefined) {
  throw new Error('Workbook is missing required sheets')
}

workbook.setCellContents({ sheet: inputs, row: 1, col: 1 }, 32)

const revenue = workbook.getCellDisplayValue({ sheet: summary, row: 1, col: 1 })
const saved = serializeWorkPaperDocument(exportWorkPaperDocument(workbook, { includeConfig: true }))

console.log({ revenue, savedBytes: saved.length })
```

## Verification shortcuts

From a clean project, run the package-owned challenge:

```sh
npm exec --package bilig-workpaper@0.40.21 -- bilig-agent-challenge
npm exec --package bilig-workpaper@0.40.21 -- bilig-mcp-challenge
```

`bilig-agent-challenge` proves the direct WorkPaper API loop.
`bilig-mcp-challenge` proves the file-backed MCP JSON-RPC loop. A good run
prints `verified: true`.

Deeper docs:

- <https://proompteng.github.io/bilig/headless-workpaper-agent-handbook.html>
- <https://proompteng.github.io/bilig/mcp-workpaper-tool-server.html>
- <https://proompteng.github.io/bilig/mcp-client-setup.html>
