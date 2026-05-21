---
title: Agent workbook challenge
published: true
description: A copy-paste challenge for coding agents: solve a workbook task without spreadsheet UI automation, then prove the edit with Bilig WorkPaper readback and persistence.
tags: ai agents, mcp, spreadsheet automation, workbook formulas, node
canonical_url: https://proompteng.github.io/bilig/agent-workbook-challenge.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Agent Workbook Challenge

Use this page when an agent, MCP client, or coding assistant is about to solve a
spreadsheet task by clicking through Excel, LibreOffice, Google Sheets, or a
screenshot grid.

The challenge is simple: do the workbook edit through an API, not through a UI,
and return proof that the calculated value changed and survived persistence.

## The Task

Paste this into an agent:

```text
You need to solve a workbook-shaped task without opening Excel, LibreOffice,
Google Sheets, or a screenshot UI unless the user explicitly asks for a visual
review.

Use @bilig/headless or the Bilig WorkPaper MCP server. Build or load a workbook
with these sheets:

Inputs
- A1: Metric
- B1: Value
- A2: Customers
- B2: 20
- A3: Average revenue
- B3: 1200

Summary
- A1: Metric
- B1: Value
- A2: Revenue
- B2: =Inputs!B2*Inputs!B3

Then change Inputs!B2 from 20 to 32. Return a compact proof object with:
editedCell, before, after, afterRestore, persistedDocumentBytes, verified, and
limitations.

Do not claim success from the write call alone. Success requires computed
readback after the edit and restore proof from serialized WorkPaper JSON.
```

Expected outcome:

```json
{
  "editedCell": "Inputs!B2",
  "before": 24000,
  "after": 38400,
  "afterRestore": 38400,
  "verified": true
}
```

The exact byte count can change between package versions. The invariant is that
the edited input changes the dependent formula result, and the restored document
keeps the same result.

## Fastest Path: Published Package

This uses the package-owned challenge command. It does not clone the repo, curl
a TypeScript file, or require a spreadsheet UI:

```sh
npm exec --package @bilig/headless@0.40.35 -- bilig-agent-challenge
npm exec --package @bilig/headless@0.40.35 -- bilig-mcp-challenge
```

A passing run prints `verified: true`.
Use `--markdown` when you want a paste-ready report for an issue, PR, or agent
eval transcript.

Use `bilig-agent-challenge` for the direct WorkPaper API loop. Use
`bilig-mcp-challenge` when the evaluator cares about the actual MCP path:
JSON-RPC initialize, tool/resource/prompt discovery, `set_cell_contents`,
dependent formula readback, WorkPaper JSON export, and restart readback from the
same persisted file.

## MCP Path

Use this when the host supports MCP servers:

```sh
npm exec --package @bilig/headless@0.40.35 -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable
```

Required tool sequence:

1. `list_sheets`
2. `read_range` for the input and summary ranges
3. `set_cell_contents` for `Inputs!B2`
4. `get_cell_display_value` for the dependent summary cell
5. `export_workpaper_document`

That sequence is the point of the challenge. It keeps the agent honest about
what changed, what recalculated, and what can be saved.

## Why This Beats Screenshot Automation

Screenshot automation can be useful for final human review, but it is a weak
primary interface for agents:

- screenshots hide formula text and typed cell addresses;
- clicks can land on the wrong sheet, row, or browser state;
- cached XLSX formula values can look valid while being stale;
- a visual grid does not prove the workbook can be persisted and restored.

WorkPaper state gives the agent a smaller contract: read cells, write cells,
recalculate formulas, export JSON, and report the proof object.

## Pass/Fail Rubric

Pass:

- the answer names the exact edited cell;
- the answer includes the before and after calculated values;
- the after value is read from the dependent formula cell;
- the workbook document is serialized or exported;
- restore or reimport gives the same calculated value;
- limitations are named instead of hidden.

Fail:

- the answer only says that a cell was written;
- the agent relies on a screenshot as formula truth;
- the agent reports cached XLSX values as recalculated values;
- the answer omits persistence proof;
- unsupported formulas are silently skipped.

## Shareable Prompt

Use this shorter version in an issue, discussion, or agent-tool eval:

```text
Try the Bilig agent workbook challenge: update one input cell, read the
dependent formula result, serialize the WorkPaper JSON, restore it, and return
verified: true. Do it without spreadsheet UI automation unless visual review is
explicitly required.

Start here:
https://proompteng.github.io/bilig/agent-workbook-challenge.html
```

## Where To Go Next

- For a broader agent playbook, use the
  [Headless WorkPaper agent handbook](headless-workpaper-agent-handbook.md).
- For MCP client setup, use the
  [MCP client setup guide](mcp-client-setup.md).
- For direct tool wrappers, use the
  [WorkPaper tool-calling recipe](agent-workpaper-tool-calling-recipe.md).
- If the challenge almost works but a real workbook blocks adoption, use the
  [formula bug clinic](formula-bug-clinic.md) or
  [submit a workbook fixture](submit-workbook-fixture.md).
