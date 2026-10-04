---
title: Evaluate Bilig as an agent MCP workbook tool
published: true
description: Evaluate MCP workbook tools for agents that need workbook writes, formula readback, JSON persistence, and restart proof.
tags: agents, mcp, workpaper, spreadsheet, evaluator
canonical_url: https://proompteng.github.io/bilig/eval-agent-mcp.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Evaluate Bilig as an agent MCP workbook tool

Use this when an agent is about to drive a spreadsheet UI by screenshots or
clicks. The narrower contract is better: list workbook tools, write one input
cell, read the dependent formula output, export WorkPaper JSON, restart from the
persisted file, and return proof.

## One command

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
```

For a less toy-like workbook, run the revenue-plan scenario:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --scenario revenue-plan --json
```

That path edits `Deals!C2` and verifies `SUM`, `SUMIF`, `XLOOKUP`, a `FILTER`
spill, a named expression, JSON persistence, and restart readback.

If the workbook uses a provider-backed formula such as `IMPORTRANGE`, run the
adapter-boundary scenario:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --scenario provider-backed --json
```

That scenario reads an `IMPORTRANGE` cell with no adapter and expects
`#BLOCKED!` plus a `provider-backed-adapter-missing` diagnostic. It then
installs a local synthetic adapter, recalculates a fresh `96000` readback,
exports WorkPaper JSON, restarts from disk, and verifies the diagnostic is
gone. It does not call Google Sheets or prove remote authorization.

If you are handing this to another coding agent, start from the
[Agent WorkPaper handoff](agent-adoption-kit.md). It includes the installable skill,
one MCP config, a workbook task, and the pass/fail proof object.

If you need the raw JSON-RPC challenge output, run:

```sh
npm exec --package @bilig/workpaper@latest -- bilig-mcp-challenge --json
```

## Expected proof

The evaluator prints this shape:

```json
{
  "schemaVersion": "bilig-evaluator.v1",
  "door": "agent-mcp",
  "doorName": "Agent MCP proof",
  "verified": true,
  "packageVersions": {
    "@bilig/workpaper": "1.1.1",
    "@bilig/xlsx-formula-recalc": "1.1.1"
  },
  "evidence": {
    "editedCell": "Inputs!B3",
    "dependentCell": "Summary!B3",
    "before": 60000,
    "after": 96000,
    "afterRestore": 96000,
    "afterRestart": 96000,
    "persistedDocumentBytes": 1162,
    "toolCount": 8,
    "tools": [
      "list_sheets",
      "read_range",
      "read_cell",
      "set_cell_contents",
      "set_cell_contents_and_readback",
      "get_cell_display_value",
      "export_workpaper_document",
      "validate_formula"
    ],
    "checks": {
      "listedFileBackedTools": true,
      "listedResourcesAndPrompts": true,
      "formulaValidationPassed": true,
      "dependentCellChanged": true,
      "persistedToDisk": true,
      "exportContainsWorkPaperDocument": true,
      "restartReadbackMatchesAfter": true,
      "displayValueRead": true
    }
  }
}
```

The exact package versions, byte count, and duration can change. The invariants
are `door: "agent-mcp"`, `dependentCellChanged`, `persistedToDisk`,
`restartReadbackMatchesAfter`, `displayValueRead`, and `verified: true`.

For `--scenario revenue-plan`, the invariants are `scenario: "revenue-plan"`,
`editedCell: "Deals!C2"`, `readbackRange: "Summary!B2:B8"`,
`totalRevenueRecalculated`, `sumifReadbackChanged`, `xlookupReadbackStable`,
`filterSpillUpdated`, `namedExpressionApplied`, `persistedToDisk`,
`restartReadbackMatchesAfter`, and `verified: true`.

For `--scenario provider-backed`, the invariants are
`scenario: "provider-backed"`, `providerFunction: "IMPORTRANGE"`,
`adapterSurface: "web"`, `before.displayValue: "#BLOCKED!"`,
`provider-backed-adapter-missing`, `after.displayValue: "96000"`,
`adapterBackedDiagnosticsCleared`, `restartReadbackMatchesAfter`, and
`verified: true`.

## What this proves

- the published package exposes a file-backed MCP stdio server
- an agent can discover spreadsheet tools and prompts
- an input edit changes a dependent formula result
- the updated WorkPaper document can be exported and persisted
- restart readback matches the calculated value after the edit
- provider-backed formulas fail closed with actionable diagnostics until the
  host supplies an adapter

## What this does not prove

This does not prove arbitrary workbook compatibility, macros, pivots, charts,
live Google Sheets authorization, external links, unsupported formulas, or
desktop Excel parity. It proves the agent tool contract: no screenshot truth,
no blind write-only success, and no missing persistence proof.

## After the proof

- Repository:
  <https://github.com/proompteng/bilig>
- Watch releases for MCP and agent-tool updates:
  <https://github.com/proompteng/bilig/subscription>
- Report the exact implementation gap:
  <https://github.com/proompteng/bilig/discussions/new?category=general>

## Related

- [Agent workbook challenge](agent-workbook-challenge.md)
- [MCP WorkPaper tool server](mcp-workpaper-tool-server.md)
- [WorkPaper agent handbook](headless-workpaper-agent-handbook.md)
- [MCP client setup](mcp-client-setup.md)
