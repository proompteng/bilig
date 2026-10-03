---
title: OpenHands WorkPaper MCP setup
published: true
description: Configure OpenHands with Bilig's WorkPaper MCP server and repository skill so workbook edits return formula readback instead of spreadsheet UI status.
tags: openhands, agents, mcp, workbook-api, spreadsheet-automation
canonical_url: https://proompteng.github.io/bilig/openhands-workpaper-mcp.html
image: /assets/github-social-preview.png
---

# OpenHands WorkPaper MCP setup

Use this when an OpenHands agent needs spreadsheet formulas while coding in a
repo. OpenHands should own the code task; Bilig should own workbook truth:
read a range, write one cell, read the dependent formula output, persist
WorkPaper JSON, and return proof.

Official OpenHands references:

- <https://docs.openhands.dev/openhands/usage/cli/mcp-servers>
- <https://docs.openhands.dev/overview/skills>
- <https://docs.openhands.dev/sdk/arch/skill>

## First Proof Command

Before changing an OpenHands config, prove the published WorkPaper MCP door:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
```

Trust the path only when the result includes `verified: true`, edited cell
evidence, dependent formula readback, exported WorkPaper JSON, and restart
readback.

## Add The MCP Server

OpenHands CLI MCP setup uses `openhands mcp add <name> --transport stdio
<command> -- [args...]`. Add Bilig's file-backed WorkPaper server like this:

```sh
openhands mcp add bilig-workpaper --transport stdio npm -- \
  exec --yes --package @bilig/workpaper@latest -- \
  bilig-workpaper-mcp \
  --workpaper ./.bilig/pricing.workpaper.json \
  --init-demo-workpaper \
  --writable
```

Check it before starting or restarting the conversation:

```sh
openhands mcp list
openhands mcp get bilig-workpaper
```

Inside an OpenHands conversation, use `/mcp` to confirm the active server. New
or edited MCP config is loaded on conversation restart.

## Manual MCP Config

OpenHands also reads `~/.openhands/mcp.json`. The equivalent config is:

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
      ]
    }
  }
}
```

Use the hidden `.bilig/` path for existing repos so the overlay does not add a
noisy root workbook file. For a blank generated project, the starter uses
`./pricing.workpaper.json`.

## Repository Skill

OpenHands prefers always-on repository instructions in `AGENTS.md` and supports
project skills under `.agents/skills/`. A Bilig-aware repo should include:

```text
AGENTS.md
.agents/skills/bilig-workpaper/SKILL.md
.mcp.json
mcp/bilig-workpaper.mcp.json
```

The published Bilig starter is release-pending while
`@bilig/create-workpaper@latest` resolves to `0.164.11`; that release's smoke
reports `formulasPersisted: false`. Until a newer release passes a fresh
consumer smoke, copy `AGENTS.md` and
`.agents/skills/bilig-workpaper/SKILL.md` from a cloned checkout and verify the
public MCP path directly:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
```

## OpenHands Task Prompt

```text
Use Bilig WorkPaper MCP tools instead of spreadsheet UI automation. Read the
relevant range first, write one precise input or formula change, read the
dependent calculated output after recalculation, export or serialize the
WorkPaper document, restart or restore when the file boundary matters, and
return editedCell, before, after, afterRestore, persistedDocumentBytes,
verified, and limitations. Do not claim success from a write call alone.
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

No upstream OpenHands PR or issue was opened for this guide. It is an owned
Bilig integration surface backed by public OpenHands MCP and skill docs plus a
no-key WorkPaper readback proof.
