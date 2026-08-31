---
title: Bilig WorkPaper install context for tool hosts
published: true
description: Short install and verification path for MCP clients and tool hosts that need workbook formulas without spreadsheet UI automation.
tags: mcp, workpaper, spreadsheet formulas, node
canonical_url: https://proompteng.github.io/bilig/llms-install.html
---

# Bilig WorkPaper install context

Bilig gives tool hosts a workbook API instead of a spreadsheet screen.

Use it when a task needs to edit workbook inputs, recalculate formulas, read the
computed value back, and persist the workbook state. Do not use Excel,
LibreOffice, Google Sheets, browser grids, or screenshots as the first tool when
the workbook can run as WorkPaper JSON.

## Fast check

Run this before wiring any host integration:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
```

A useful result has `schemaVersion: "bilig-evaluator.v1"`, `door:
"agent-mcp"`, `verified: true`, a changed input cell, a changed dependent
formula value, exported WorkPaper JSON, and restart readback that matches the
post-edit value.

Use the service path when the integration will call TypeScript directly instead of
MCP:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door workpaper-service --json
```

## Add host files to an existing repo

Do not use `npm create @bilig/workpaper@latest` while
`@bilig/create-workpaper@latest` resolves to `0.164.11`. Its generated smoke
reports `formulasPersisted: false` because that release checks escaped JSON
text instead of restored formula state. The source fix must ship in a newer
release and pass a fresh consumer smoke before this path is restored.

Until then, use the evaluator commands above, install the reusable skill below,
or copy the matching host files from a cloned checkout. The repository carries
`AGENTS.md`, `CONVENTIONS.md`, `.aider.conf.yml`, `CLAUDE.md`, `GEMINI.md`,
Copilot/VS Code instructions, Cursor, Kiro, Roo Code, Trae, Qodo, Zed, Junie,
OpenCode, Aider, Cline, Continue, Windsurf rules, and MCP client configs.

## Install the reusable skill

```sh
npx --yes skills@latest add https://bilig.proompteng.ai --list
npx --yes skills@latest add proompteng/bilig --skill bilig-workpaper --list
```

Use the hosted skill URL first. Keep the GitHub skill command for hosts that
only support repository-backed skills.

## MCP server config

Use file-backed stdio for private workbook state:

```json
{
  "mcpServers": {
    "bilig-workpaper": {
      "command": "npm",
      "args": [
        "exec",
        "--package",
        "@bilig/workpaper@latest",
        "--",
        "bilig-workpaper-mcp",
        "--workpaper",
        "./.bilig/pricing.workpaper.json",
        "--init-demo-workpaper",
        "--writable"
      ]
    }
  }
}
```

If the project already has an `.xlsx` file and the agent needs triage before
trusting the import, start with the direct XLSX mode:

```sh
npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp --from-xlsx ./pricing.xlsx
```

That imports the workbook into an in-memory WorkPaper server. In this mode,
`tools/list` also includes `analyze_workbook_risk`, a fixed-source diagnostic
for unsupported functions, external links, macro payloads, pivots, volatile
formulas, stored formula results, and concrete risk reasons. It does not certify
Excel compatibility.

Persist the imported WorkPaper only when the workflow needs a durable sidecar:

```sh
npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp --from-xlsx ./pricing.xlsx --workpaper ./.bilig/pricing.workpaper.json --writable
```

Expected tools:

- `list_sheets`
- `read_range`
- `read_cell`
- `set_cell_contents`
- `set_cell_contents_and_readback`
- `get_cell_display_value`
- `export_workpaper_document`
- `validate_formula`

Use the hosted endpoint only for stateless smoke tests and tool discovery:

```text
https://bilig.proompteng.ai/mcp
```

Do not put private workbook data in the hosted demo endpoint.

## IDE and agent rule files

Use the repo-local files when the agent is already inside a checkout:

- Claude Code: `.claude/skills/bilig-workpaper/SKILL.md`
- Claude Code command: `.claude/commands/bilig-workpaper-proof.md`
- GitHub Copilot: `.github/copilot-instructions.md`
- GitHub Copilot custom instructions: `.github/instructions/bilig-workpaper.instructions.md`
- GitHub Copilot prompt: `.github/prompts/bilig-workpaper-proof.prompt.md`
- VS Code MCP: `.vscode/mcp.json`
- Cursor: `.cursor/rules/bilig-workpaper.mdc`
- Kiro: `.kiro/steering/bilig-workpaper.md` and `.kiro/settings/mcp.json`
- Roo Code: `.roo/rules/bilig-workpaper.md` and `.roo/mcp.json`
- Trae: `.trae/rules/bilig-workpaper.md` and `.trae/mcp.json`
- Qodo IDE: use [Qodo WorkPaper MCP setup](https://proompteng.github.io/bilig/qodo-workpaper-mcp.html) and paste
  the `bilig-workpaper` server JSON into Qodo Agentic Tools MCP settings
- Zed: `.zed/settings.json`
- Junie: `.junie/mcp/mcp.json`
- OpenCode: `opencode.jsonc` and `.opencode/agents/bilig-workpaper.md`
- Aider: `CONVENTIONS.md` loaded by `.aider.conf.yml`
- Cascade/Devin: `.devin/rules/bilig-workpaper.md`
- Cline: `.clinerules/bilig-workpaper.md`
- Continue: `.continue/rules/bilig-workpaper.md`
- Windsurf/Cascade: `.windsurf/rules/bilig-workpaper.md`

Rule chooser: <https://proompteng.github.io/bilig/agent-rule-chooser.html>

## Direct TypeScript path

Use direct imports for services, queues, route handlers, and tests:

```sh
npm install @bilig/workpaper
```

```ts
import {
  WorkPaper,
  exportWorkPaperDocument,
  serializeWorkPaperDocument,
} from "@bilig/workpaper";

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
  throw new Error("missing sheets");
}

workbook.setCellContents({ sheet: inputs, row: 2, col: 1 }, 4800);

const after = workbook.getCellDisplayValue({ sheet: summary, row: 1, col: 1 });
const saved = serializeWorkPaperDocument(
  exportWorkPaperDocument(workbook, { includeConfig: true }),
);

console.log({ after, persistedDocumentBytes: saved.length });
```

## Agent task contract

Ask the agent for workbook evidence, not a write status:

```text
Use Bilig WorkPaper tools instead of spreadsheet UI automation. Read the input
and output cells first, set one input cell, read the dependent formula cell
after recalculation, export or serialize the WorkPaper document, restore it when
a file boundary matters, and return editedCell, before, after, afterRestore,
persistedDocumentBytes, verified, and limitations.
```

Reject answers that only say a cell was written.

## More context

- Compact map: <https://proompteng.github.io/bilig/llms.txt>
- Full context: <https://proompteng.github.io/bilig/llms-full.txt>
- Agent start: <https://proompteng.github.io/bilig/agent-start.txt>
- Agent manifest: <https://proompteng.github.io/bilig/.well-known/agent.json>
- Agent handoff checklist: <https://proompteng.github.io/bilig/agent-adoption-kit.html>
- MCP server guide: <https://proompteng.github.io/bilig/mcp-workpaper-tool-server.html>
- Repository: <https://github.com/proompteng/bilig>
