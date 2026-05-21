# Bilig Agent Instructions

Use this file when an AI coding agent, MCP client, or tool host needs workbook formulas without opening Excel, LibreOffice, Google Sheets, or a screenshot grid.

## Discovery Order

1. Read `https://proompteng.github.io/bilig/llms.txt` for the compact map.
2. Read `https://proompteng.github.io/bilig/llms-full.txt` when you need enough context to implement a workflow without searching the whole site.
3. Read `https://proompteng.github.io/bilig/skill.txt` when your agent supports skill manifests.
4. Start the MCP server or import `@bilig/workpaper` directly.

Remote MCP clients that support Streamable HTTP can smoke-test the stateless
demo endpoint at `https://bilig.proompteng.ai/mcp`. Directory scanners that inspect the
hosted endpoint origin can read the same-origin server card at
`https://bilig.proompteng.ai/.well-known/mcp/server-card.json`. Local agents that need writable project files should
still prefer the file-backed stdio command below.

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
npm exec --package @bilig/workpaper@0.40.41 -- bilig-agent-challenge
npm exec --package @bilig/workpaper@0.40.41 -- bilig-mcp-challenge
npm exec --package @bilig/workpaper@0.40.41 -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable
npm exec --package @bilig/workpaper@0.40.41 -- bilig-formula-clinic ./reduced.xlsx --cells "Summary!B7,Inputs!B2"
```

Claude Desktop users can install the released MCPB bundle from:

- https://github.com/proompteng/bilig/releases/download/libraries-v0.40.41/bilig-workpaper.mcpb
- https://github.com/proompteng/bilig/releases/download/libraries-v0.40.41/bilig-workpaper.mcpb.sha256

## Direct TypeScript

Use `WorkPaper.buildFromSheets()` for hand-authored models, `setCellContents()` for edits, `getCellDisplayValue()` for readback, and `exportWorkPaperDocument()` plus `serializeWorkPaperDocument()` for persistence proof.

## Boundaries

Keep Excel, LibreOffice, Microsoft Graph, or an oracle harness in the loop when the workbook depends on macros, pivots, charts, external links, unsupported functions, locale-specific Excel behavior, or exact desktop UI behavior.
