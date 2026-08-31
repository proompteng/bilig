---
title: OpenCode WorkPaper MCP setup
published: true
description: Configure OpenCode with Bilig's WorkPaper MCP server and project subagent so workbook edits return formula readback instead of spreadsheet UI status.
tags: opencode, agents, mcp, workbook-api, spreadsheet-automation
canonical_url: https://proompteng.github.io/bilig/opencode-workpaper-mcp.html
image: /assets/github-social-preview.png
---

# OpenCode WorkPaper MCP setup

Use this when an OpenCode agent needs spreadsheet formulas while coding in a
repo. OpenCode should own the code task; Bilig should own workbook truth:
read a range, write one cell, read the dependent formula output, persist
WorkPaper JSON, and return proof.

Official OpenCode references:

- <https://opencode.ai/docs/config/>
- <https://opencode.ai/docs/mcp-servers/>
- <https://opencode.ai/docs/agents/>

## First Proof Command

Before changing OpenCode config, prove the published WorkPaper MCP door:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
```

Trust the path only when the result includes `verified: true`, edited cell
evidence, dependent formula readback, exported WorkPaper JSON, and restart
readback.

## Add The MCP Server

OpenCode reads project config from `opencode.json` or `opencode.jsonc` and
supports local MCP servers under the `mcp` option. Use this project-local
`opencode.jsonc`:

```jsonc
{
  "$schema": "https://opencode.ai/config.json",
  "instructions": ["AGENTS.md"],
  "mcp": {
    "bilig-workpaper": {
      "type": "local",
      "command": [
        "npm",
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
      "enabled": true
    },
    "bilig-workpaper-demo": {
      "type": "remote",
      "url": "https://bilig.proompteng.ai/mcp",
      "enabled": false
    }
  }
}
```

Use the hidden `.bilig/` path for existing repos so the overlay does not add a
noisy root workbook file. For a blank generated project, the starter uses
`./pricing.workpaper.json`.

Enable `bilig-workpaper-demo` only for remote, stateless tool-discovery smoke
tests. Use `bilig-workpaper` for private project workbook state.

## Add The Project Agent

OpenCode supports markdown agents under `.opencode/agents/`. Add:

```text
.opencode/agents/bilig-workpaper.md
```

The generated file tells OpenCode to invoke the WorkPaper proof path when a
task mentions pricing, quotes, payouts, budgets, imports, forecasts, workbook
validation, or formula readback after cell edits. It also allows
`bilig-workpaper_*` MCP tools while keeping shell commands reviewable.

The published Bilig starter is release-pending while
`@bilig/create-workpaper@latest` resolves to `0.164.11`; that release's smoke
reports `formulasPersisted: false`. Until a newer release passes a fresh
consumer smoke, copy `opencode.jsonc` and
`.opencode/agents/bilig-workpaper.md` from a cloned checkout and verify the
public MCP path directly:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
```

## OpenCode Task Prompt

```text
Use the Bilig WorkPaper OpenCode agent and MCP server instead of spreadsheet UI
automation. Read the relevant range first, write one precise input or formula
change, read the dependent calculated output after recalculation, export or
serialize the WorkPaper document, restart or restore when the file boundary
matters, and return editedCell, before, after, afterRestore,
persistedDocumentBytes, verified, and limitations. Do not claim success from a
write call alone.
```

Expected MCP tools:

- `list_sheets`
- `read_range`
- `read_cell`
- `set_cell_contents`
- `set_cell_contents_and_readback`
- `get_cell_display_value`
- `export_workpaper_document`
- `validate_formula`

## Boundary

This is for code-owned workbook logic where JSON state can represent the
workbook. It is a fit for quote approval, payout checks, pricing rules, import
validation, and forecast gates.

It is not a claim that Bilig replaces desktop Excel for macros, add-ins, pivot
tables, or visual workbook review. For raw `.xlsx` files, start with:

```sh
npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- \
  bilig-evaluate --door xlsx-cache --json
```

No upstream OpenCode PR or issue was opened for this guide. It is an owned
Bilig integration surface backed by public OpenCode config, MCP, and agent docs
plus a no-key WorkPaper readback proof.
