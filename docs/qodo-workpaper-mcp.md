---
title: Qodo WorkPaper MCP setup
published: true
description: Add Bilig WorkPaper as a Qodo IDE Agentic Tools MCP server so coding agents can edit workbook inputs, recalculate formulas, and verify readback without spreadsheet UI automation.
tags: qodo, mcp, coding agents, spreadsheet, workpaper
canonical_url: https://proompteng.github.io/bilig/qodo-workpaper-mcp.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Qodo WorkPaper MCP Setup

Use this page when a Qodo IDE agent needs workbook tools but should not drive
Excel, LibreOffice, Google Sheets, a browser grid, or screenshots. Qodo IDE's
Agentic Tools MCP settings can launch a local stdio server from JSON. Bilig's
useful setup is the file-backed `bilig-workpaper` MCP server plus the root
`AGENTS.md` policy in this repo.

The proof bar is readback, not a tool-call status. Qodo should read workbook
state, edit only the requested input cell, recalculate dependent formulas,
export or persist the WorkPaper JSON, restore or re-open when file state
matters, and return `verified: true` only when the computed value and persisted
state match.

## Agentic Tools MCP JSON

Open Qodo IDE Agentic Tools MCP settings and add this local server:

```json
{
  "mcpServers": {
    "bilig-workpaper": {
      "command": "npm",
      "args": [
        "exec",
        "--yes",
        "--package",
        "@bilig/workpaper@latest",
        "--",
        "bilig-workpaper-mcp",
        "--workpaper",
        "./.bilig/pricing.workpaper.json",
        "--init-demo-workpaper",
        "--writable"
      ],
      "env": {}
    }
  }
}
```

Restart or refresh MCP tools after saving the config. The server name is
`bilig-workpaper`.

## Project Policy

Keep the shared host policy in root `AGENTS.md`. It tells Qodo and other coding
tools to use Bilig WorkPaper first when a task is workbook-shaped business
logic:

```text
Use Bilig WorkPaper before Excel, LibreOffice, Google Sheets, browser grids, or
screenshots when the task involves spreadsheet formulas, workbook cells,
pricing, budgets, forecasts, payout checks, import validation, or tool integrations
that need formula readback.
```

Start with the no-key evaluator:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --scenario provider-backed --json
npm exec --package @bilig/workpaper@latest -- bilig-agent-challenge --json
npm exec --package @bilig/workpaper@latest -- bilig-mcp-challenge --json
```

For a copy-pasteable Qodo-specific prompt from the package:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-agent-start --rules qodo
```

## Proof Prompt

Ask Qodo for a concrete readback check:

```text
Use the bilig-workpaper MCP server from Qodo Agentic Tools. List sheets, read
Inputs!B3 and Summary!B3, set Inputs!B3 to 0.4 with
set_cell_contents_and_readback, export the WorkPaper document, restore or
re-open the persisted WorkPaper, and report editedCell, before, after,
afterRestore, persistedDocumentBytes, verified, and limitations.
```

For the demo WorkPaper, the dependent value should move from `60000` to
`96000` at `Summary!B3` after `Inputs!B3` is set to `0.4`.

## Hosted Endpoint Boundary

Qodo users can smoke-test the hosted Streamable HTTP endpoint when a client only
needs remote discovery:

```text
https://bilig.proompteng.ai/mcp
```

That endpoint is stateless and request-local. Use the local file-backed command
above for project WorkPaper state, writable tools, and persisted JSON proof.

## What To Require In The Answer

Require the Qodo transcript or final answer to include:

- `editedCell`
- `before`
- `after`
- `afterRestore` or an equivalent persisted-state re-open check
- `persistedDocumentBytes`
- `verified: true`
- `limitations`

Do not accept "the cell was updated" as success. Do not claim Excel
compatibility, macro support, pivot refresh, or external-data refresh from this
MCP proof.

## Duplicate And Upstream Boundary

No upstream Qodo PR, issue, or listing was opened for this tranche. This is an
owned setup page for Qodo IDE users who already have a cloned project and need a
local MCP workbook tool. Bilig does not claim that Qodo reads a repo-native
`.qodo` MCP file.

## Official Qodo Docs Checked

- [Qodo Agentic Tools MCP](https://docs.qodo.ai/qodo-documentation/qodo-ide/tools-mcps/agentic-tools-mcps)
- [Qodo Merge configuration](https://docs.qodo.ai/qodo-documentation/qodo-review/configuration/qodo-merge-configuration)

## Related

- [Coding agent rule chooser](agent-rule-chooser.md)
- [MCP client setup](mcp-client-setup.md)
- [Agent WorkPaper handoff](agent-adoption-kit.md)
- [Evaluate Bilig as an agent MCP workbook tool](eval-agent-mcp.md)
