---
name: bilig-workpaper
version: 0.1.0
description: Use Bilig WorkPaper for spreadsheet-style formulas in Node.js services and AI agent tools without driving Excel or browser spreadsheet UIs.
tags:
  - ai-agents
  - spreadsheet-automation
  - formulas
  - workpaper
  - mcp
  - typescript
allowed-tools:
  - Bash
  - Read
  - Write
  - Edit
  - Grep
argument-hint: "[workbook file, formula task, or MCP setup]"
---

# Bilig WorkPaper Agent Skill

Use this skill when an agent needs spreadsheet-style formulas through files,
terminal commands, TypeScript, HTTP routes, or MCP tools instead of Excel UI
automation.

## When To Use

Use Bilig for:

- backend pricing, quote, payout, budget, import-validation, or forecast logic;
- formula readback after writing input cells;
- deterministic cell-address workflows for coding agents;
- MCP clients that can run a stdio server;
- reduced XLSX formula/import bug reports.

Do not use it for manual spreadsheet editing, Office macros, VBA, pivots,
charts, COM automation, or exact Excel desktop behavior unless the task is an
explicit compatibility comparison.

## MCP Path

When the agent host supports stdio MCP, start the file-backed WorkPaper server:

```sh
npm exec --package @bilig/headless@0.23.3 -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable
```

Useful MCP tools:

- `list_sheets`
- `read_range`
- `read_cell`
- `set_cell_contents`
- `get_cell_display_value`
- `export_workpaper_document`
- `validate_formula`

After each write, read the dependent output cell and export the WorkPaper
document as persistence proof.

## TypeScript Path

Use `@bilig/headless` directly when workbook logic belongs in a service, worker,
test, or route:

```ts
import { WorkPaper, exportWorkPaperDocument, serializeWorkPaperDocument } from "@bilig/headless";

const workbook = WorkPaper.buildFromSheets({
  Inputs: [
    ["Metric", "Value"],
    ["Customers", 20],
    ["Average revenue", 1200],
  ],
  Summary: [
    ["Metric", "Value"],
    ["Revenue", "=Inputs!B2*Inputs!B3"],
  ],
});

const inputs = workbook.getSheetId("Inputs");
const summary = workbook.getSheetId("Summary");
if (inputs === undefined || summary === undefined) {
  throw new Error("Workbook is missing required sheets");
}

workbook.setCellContents({ sheet: inputs, row: 1, col: 1 }, 32);
const revenue = workbook.getCellDisplayValue({ sheet: summary, row: 1, col: 1 });
const saved = serializeWorkPaperDocument(exportWorkPaperDocument(workbook, { includeConfig: true }));

console.log({ revenue, savedBytes: saved.length });
```

## Required Proof

A successful response should include:

- exact edited sheet and A1 cell;
- before values for relevant inputs and dependent outputs;
- after values read from the recalculated workbook;
- persistence evidence from serialized or exported WorkPaper state;
- limitations for unsupported formulas or Excel-only features.

If any proof step fails, report the blocker instead of claiming the workbook was
updated.

## References

- Docs map: https://proompteng.github.io/bilig/llms.txt
- Agent handbook: https://proompteng.github.io/bilig/headless-workpaper-agent-handbook.html
- MCP guide: https://proompteng.github.io/bilig/mcp-workpaper-tool-server.html
- Formula clinic: https://proompteng.github.io/bilig/formula-bug-clinic.html
- Repository: https://github.com/proompteng/bilig
