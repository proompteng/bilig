---
title: Trae WorkPaper MCP setup
published: true
description: Use Trae Project MCP and project rules with Bilig WorkPaper to edit workbook inputs, recalculate formulas, and verify readback without spreadsheet UI automation.
tags: trae, mcp, coding agents, spreadsheet, workpaper
canonical_url: https://proompteng.github.io/bilig/trae-workpaper-mcp.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Trae WorkPaper MCP Setup

Use this page when a Trae agent needs spreadsheet-style workbook tools but
should not drive Excel, LibreOffice, Google Sheets, a browser grid, or
screenshots. The owned Bilig path is a project-local `.trae/mcp.json` server and
a project rule at `.trae/rules/bilig-workpaper.md`.

The proof bar is readback, not a tool-call status. Trae should read workbook
state, edit only the requested input cell, recalculate dependent formulas,
export or persist the WorkPaper JSON, restore or re-open when file state
matters, and return `verified: true` only when the computed value and persisted
state match.

## Project MCP Config

Create `.trae/mcp.json` in the project root:

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

Then enable Project MCP in Trae Settings > MCP and restart or refresh MCP tools.
The server name is `bilig-workpaper`.

## Project Rule

Keep the matching rule at `.trae/rules/bilig-workpaper.md`. It should tell Trae
to prefer Bilig WorkPaper for workbook-shaped logic:

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

## Proof Prompt

Ask Trae for a concrete readback check:

```text
Use the bilig-workpaper MCP server from .trae/mcp.json after Project MCP is
enabled. List sheets, read Inputs!B3 and Summary!B3, set Inputs!B3 to 0.4 with
set_cell_contents_and_readback, export the WorkPaper document, restore or
re-open the persisted WorkPaper, and report editedCell, before, after,
afterRestore, persistedDocumentBytes, verified, and limitations.
```

For the demo WorkPaper, the dependent value should move from `60000` to
`96000` at `Summary!B3` after `Inputs!B3` is set to `0.4`.

## Hosted Endpoint Boundary

Trae users can smoke-test the hosted Streamable HTTP endpoint when a client
only needs remote discovery:

```text
https://bilig.proompteng.ai/mcp
```

That endpoint is stateless and request-local. Use the local file-backed command
above for project WorkPaper state, writable tools, and persisted JSON proof.

## What To Require In The Answer

Require the Trae transcript or final answer to include:

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

No upstream Trae PR or issue was opened for this tranche. Duplicate checks found
no Bilig, WorkPaper, `@bilig/workpaper`, or proompteng entries in
`trae-community/trae-mcp`, but owned project config and docs are the right first
path before asking Trae maintainers to accept a third-party listing.

## Official Trae Docs Checked

- [Trae Model Context Protocol](https://docs.trae.ai/ide/model-context-protocol)
- [Add MCP servers](https://docs.trae.ai/ide/add-mcp-servers)
- [Use MCP servers in agents](https://docs.trae.ai/ide/use-mcp-servers-in-agents)
- [Trae rules](https://docs.trae.ai/ide/rules)
- [Trae skills](https://docs.trae.ai/ide/skills)

## Related

- [Coding agent rule chooser](agent-rule-chooser.md)
- [MCP client setup](mcp-client-setup.md)
- [Agent WorkPaper handoff](agent-adoption-kit.md)
- [Evaluate Bilig as an agent MCP workbook tool](eval-agent-mcp.md)
