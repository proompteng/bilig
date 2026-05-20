---
title: AI spreadsheet agent tool for Node.js
description: Build an AI agent spreadsheet tool that edits workbook inputs, recalculates formulas, verifies readback, and persists state without driving Excel or screenshots.
tags: ai-agents, spreadsheet-agent, excel-agent, nodejs, xlsx, formulas, workpaper
canonical_url: https://proompteng.github.io/bilig/ai-agent-spreadsheet-tool-node.html
image: /assets/github-social-preview.png
---

# AI spreadsheet agent tool for Node.js

If an agent needs to change workbook inputs and trust the formula output, do
not start with screenshots. Give it a small tool surface that can write cells,
recalculate, read the dependent formula values, and save a proof object.

Bilig has two entry points for that:

- `@bilig/workpaper` or `@bilig/workpaper` when the workbook can live as
  WorkPaper JSON inside the service or agent tool.
- `@bilig/xlsx-formula-recalc` or `@bilig/exceljs-formula-recalc` when the user already has
  an `.xlsx` pipeline and the immediate bug is stale formula results after
  editing inputs in Node.

## Run the agent starter first

From an empty directory:

```sh
npm create @bilig/workpaper@latest pricing-agent -- --agent
cd pricing-agent
npm install
npm run agent:verify
```

The starter builds a quote-approval workbook, writes request inputs, reads the
recalculated decision cells, persists JSON, restores the workbook, and prints a
compact `verified: true` proof. It also includes `AGENTS.md`, `CLAUDE.md`,
Cursor and VS Code MCP configs, and a generic MCP config under `mcp/`.

Use this when the agent owns the model and you want reviewable business logic,
not a hidden spreadsheet process.

## Prove the direct package path

If you do not want a generated project yet:

```sh
npm exec --package @bilig/workpaper@latest -- bilig-agent-challenge
```

That command is intentionally small. It proves the minimum loop an agent needs:

1. build or load a workbook;
2. read a formula-backed output;
3. edit an input cell;
4. read the dependent formula output again;
5. persist and restore state;
6. return `verified: true` only after readback matches.

## Tool contract

Keep the agent tool API boring. The useful surface is:

```ts
type SpreadsheetAgentTools = {
  listSheets(): Promise<string[]>
  readRange(input: { sheet: string; range: string }): Promise<unknown[][]>
  setCellContents(input: { sheet: string; cell: string; value: unknown }): Promise<{ changed: boolean }>
  getCellDisplayValue(input: { sheet: string; cell: string }): Promise<string>
  exportWorkpaperDocument(): Promise<{ json: string; bytes: number }>
}
```

The agent should not report success from `setCellContents` alone. The return
path should include the edited cell, formula readback before and after the
edit, persisted document size, and known limitations.

## Existing Excel or XLSX files

When the product already uses ExcelJS, SheetJS, `xlsx-populate`, or a template
library, keep that file-writing layer. Add a recalculation step before reading
formula outputs or sending the workbook.

For raw XLSX bytes:

```sh
npm install @bilig/xlsx-formula-recalc
npx --package @bilig/xlsx-formula-recalc xlsx-recalc quote.xlsx \
  --set Inputs!B2=48 \
  --read Summary!B7 \
  --out quote.recalculated.xlsx \
  --json
```

For ExcelJS:

```sh
npm install exceljs @bilig/exceljs-formula-recalc
npx --package @bilig/exceljs-formula-recalc exceljs-recalc --demo --json
```

Use this path for the common support-ticket shape: "my Node service changed
inputs in an XLSX file, but the formula value I read is still the old cached
value."

## Framework adapters

The same tool contract works in the usual agent stacks:

- OpenAI Agents SDK function tools;
- OpenAI Responses API function calling;
- Vercel AI SDK tools;
- LangChain.js tools and LangGraph.js `ToolNode`;
- LlamaIndex.TS tools;
- CrewAI or other Python agents through a small Node worker or MCP bridge.

The important part is not the framework. The important part is making the
spreadsheet state explicit: write input cells, recalculate, read formula
outputs, and persist the model.

## When not to use this

Keep Excel, LibreOffice, Microsoft Graph, or a human review step in the loop
when the workbook depends on macros, pivots, charts, external links, desktop
Excel add-ins, unsupported functions, or exact visual layout behavior.

Bilig is for workbook-shaped logic that can be represented as cells and
formulas. It is not a replacement for every Excel feature.

## Links

- [Agent WorkPaper tool-calling recipe](agent-workpaper-tool-calling-recipe.md)
- [Headless WorkPaper agent handbook](headless-workpaper-agent-handbook.md)
- [OpenAI Agents SDK WorkPaper tool](openai-agents-sdk-workpaper-tool.md)
- [OpenAI Responses WorkPaper tool call](openai-responses-workpaper-tool-call.md)
- [Vercel AI SDK and LangChain spreadsheet tools](vercel-ai-sdk-langchain-spreadsheet-tool.md)
- [XLSX formula recalculation in Node.js](xlsx-formula-recalculation-node.md)
- [ExcelJS formula recalculation in Node.js](exceljs-formula-recalculation-node.md)
- [GitHub repo](https://github.com/proompteng/bilig)
- [Star or bookmark Bilig](https://github.com/proompteng/bilig/stargazers)
