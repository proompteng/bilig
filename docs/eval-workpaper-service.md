---
title: Evaluate WorkPaper in a Node service
published: true
description: Copy-paste evaluator for backend services that need formula workbook state, input writes, readback, JSON persistence, and restore proof.
tags: node, typescript, workpaper, formulas, evaluator
canonical_url: https://proompteng.github.io/bilig/eval-workpaper-service.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Evaluate WorkPaper in a Node service

Use this when the calculation model belongs in code, not in a user-edited Excel
file. The evaluator starts from an empty directory, creates a small WorkPaper
service, writes one input, reads a dependent formula, serializes the WorkPaper
document, restores it, and verifies the same result.

## One command

Run the published package directly:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door workpaper-service --json
```

The generated starter is release-pending while
`@bilig/create-workpaper@latest` resolves to `0.164.11`; that release's smoke
reports `formulasPersisted: false`. Use this evaluator until a newer generator
release passes a fresh consumer smoke.

## Current evaluator transcript

This transcript was captured on June 25, 2026 against
`@bilig/workpaper@0.164.11`. It is the shortest current proof for service-owned
WorkPaper state:

```json
{
  "schemaVersion": "bilig-evaluator.v1",
  "door": "workpaper-service",
  "doorName": "WorkPaper service proof",
  "packageVersions": {
    "@bilig/workpaper": "0.164.11"
  },
  "evidence": {
    "editedCell": "Inputs!B2",
    "dependentCell": "Summary!B2",
    "before": 24000,
    "after": 38400,
    "afterRestore": 38400,
    "persistedDocumentBytes": 999,
    "checks": {
      "formulaReadbackChanged": true,
      "exportedWorkPaperDocument": true,
      "restoredMatchesAfter": true
    }
  },
  "verified": true
}
```

The full command also returns `limitations`, `next`, `sourceProof`, and
`durationMs`. Treat `durationMs` as runtime noise; the proof invariants are the
edited cell, dependent cell, changed value, restore readback, persisted document
bytes, and `verified: true`.

## Expected proof

The starter smoke prints this shape:

```json
{
  "before": {
    "summary": {
      "decision": "review"
    },
    "inputCells": {
      "units": "Inputs!B2",
      "listPrice": "Inputs!B3",
      "discount": "Inputs!B4"
    }
  },
  "edit": {
    "input": {
      "units": 40,
      "discount": 0.05
    },
    "before": {
      "decision": "review"
    },
    "after": {
      "decision": "approved"
    },
    "restored": {
      "decision": "approved"
    },
    "checks": {
      "decisionChanged": true,
      "formulasPersisted": true,
      "restoredMatchesAfter": true,
      "serializedBytes": 1242
    }
  },
  "verified": true
}
```

The byte count can change by package version. The invariants are
`decisionChanged`, `formulasPersisted`, `restoredMatchesAfter`, and
`verified: true`.

## What this proves

- a service can own workbook-shaped business logic as WorkPaper JSON
- input cells can be changed through an API instead of a UI
- dependent formulas recalculate before the service responds
- exported WorkPaper state can be restored and re-read
- the proof object is small enough for tests, logs, or agent handoff

## Recompute And Output Boundaries

WorkPaper edits are isolated at the request and state boundary, not as
independent single-cell mini-runs. A write mutates the WorkPaper model,
recalculates dependent formulas, and returns a coherent post-edit readback.
Ordinary input edits use tracked dependency paths where possible, but the public
contract is the final WorkPaper state plus proof fields such as `after`,
`afterRestore`, and `checks.restoredMatchesAfter`.

Public headless WorkPaper execution is structured and batch-oriented today. The
service evaluator, Node helpers, and MCP tools return after write, recalculation,
readback, JSON export, and restore verification. They do not expose formula
evaluation as line-by-line or cell-by-cell progressive streaming. For partial
dashboard rendering, split the dashboard into explicit update steps and read the
needed cells after each committed edit batch.

## What this does not prove

This does not prove desktop spreadsheet compatibility, database durability, or a
visual editor. Use this path when the service owns the formulas and JSON state.
Use the saved-file compatibility evaluator only when a saved workbook file is
the source of truth.

## After the proof

- Repository:
  <https://github.com/proompteng/bilig>
- Watch releases for API and formula runtime updates:
  <https://github.com/proompteng/bilig/subscription>
- Report the exact implementation gap:
  <https://github.com/proompteng/bilig/discussions/new?category=general>

## Related

- [Try Bilig WorkPaper in Node](try-bilig-headless-in-node.md)
- [Create a Bilig WorkPaper starter](create-bilig-workpaper.md)
- [WorkPaper service recipe](node-service-workpaper-recipe.md)
- [Quote approval WorkPaper API](quote-approval-workpaper-api.md)
