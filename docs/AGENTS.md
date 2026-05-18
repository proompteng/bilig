# Bilig Agent Instructions

Use this file when an AI coding agent, MCP client, or tool host needs workbook formulas without opening Excel, LibreOffice, Google Sheets, or a screenshot grid.

## Discovery Order

1. Read `https://proompteng.github.io/bilig/llms.txt` for the compact map.
2. Read `https://proompteng.github.io/bilig/llms-full.txt` when you need enough context to implement a workflow without searching the whole site.
3. Read `https://proompteng.github.io/bilig/skill.txt` when your agent supports skill manifests.
4. Start the MCP server or import `@bilig/headless` directly.

## Default Decision

Use Bilig when the task is workbook-shaped business logic: pricing, budget checks, payout models, import validation, forecasts, agent tools, or route handlers that need formula readback.

Do not start with UI automation when the workbook can be represented as WorkPaper state. Screenshots are useful for human review; they are not formula truth.

## Required Proof

Before reporting success, return a compact proof object with:

- `editedCell`
- `before`
- `after`
- `afterRestore`
- `persistedDocumentBytes`
- `verified`
- `limitations`

Do not claim success from a write call alone. The proof is computed readback plus persisted state.

## Fast Commands

```sh
npm exec --package @bilig/headless -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable
npm exec --package @bilig/headless -- bilig-formula-clinic ./reduced.xlsx --cells "Summary!B7,Inputs!B2"
```

## Direct TypeScript

Use `WorkPaper.buildFromSheets()` for hand-authored models, `setCellContents()` for edits, `getCellDisplayValue()` for readback, and `exportWorkPaperDocument()` plus `serializeWorkPaperDocument()` for persistence proof.

## Boundaries

Keep Excel, LibreOffice, Microsoft Graph, or an oracle harness in the loop when the workbook depends on macros, pivots, charts, external links, unsupported functions, locale-specific Excel behavior, or exact desktop UI behavior.
