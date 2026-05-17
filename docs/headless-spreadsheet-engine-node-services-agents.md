---
title: Headless spreadsheet engine for Node services and agents
published: true
description: Use @bilig/headless as a TypeScript WorkPaper runtime for Node services and agent tools that need cell edits, formula recalculation, JSON persistence, and verified readback.
tags: typescript, node, spreadsheet, agents, mcp
canonical_url: https://proompteng.github.io/bilig/headless-spreadsheet-engine-node-services-agents.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Headless spreadsheet engine for Node services and agents

`@bilig/headless` is a TypeScript WorkPaper runtime for backend workflows that
still make the most sense as cells and formulas. It gives a Node service or
agent tool an API for writing cells, recalculating formulas, reading computed
values, and saving the workbook as JSON.

Use it when a workbook is the business-logic shape, but a browser spreadsheet
UI is the wrong runtime boundary.

## The job it fits

Good fits:

- pricing, quote approval, payout, budget, or revenue-model rules that are
  easier to review as workbook formulas
- import validation where inputs arrive as JSON or CSV-shaped records and the
  service needs formula readback
- agent tools that must prove what changed after writing a cell
- serverless or queue-worker workflows that need local formula execution
  without OAuth, screenshots, or a hosted spreadsheet document
- persisted WorkPaper JSON that can be saved, restored, tested, and diffed

Poor fits:

- a full Excel-compatible desktop spreadsheet
- XLSX formatting, images, charts, and workbook-file generation as the primary
  product
- a collaborative browser grid where users directly edit the UI
- unbounded Excel formula parity without checking the documented compatibility
  gaps

## Try the package first

Start from an empty directory and run the maintained quickstart when you want
the shortest package sanity check:

```sh
mkdir bilig-headless-eval
cd bilig-headless-eval
npm init -y
npm pkg set type=module
npm install @bilig/headless
npm install -D tsx typescript @types/node
curl -fsSLo quickstart.ts https://proompteng.github.io/bilig/npm-eval.ts
npx tsx quickstart.ts
```

Expected shape:

```json
{
  "before": 24000,
  "after": 38400,
  "afterRestore": 38400,
  "sheets": ["Inputs", "Summary"],
  "bytes": 1000,
  "verified": true,
  "nextStep": "If this proof matches your service or agent workflow, star or bookmark Bilig: https://github.com/proompteng/bilig/stargazers"
}
```

The byte count can move between releases. The important signal is
`verified: true`: an input changed, a dependent formula recalculated, the
document serialized to JSON, and the restored workbook produced the same value.

For a more production-shaped evaluator path, run the
[quote approval WorkPaper API proof](quote-approval-workpaper-api.md). It still
starts from an empty Node directory, but downloads one maintained TypeScript
route smoke, writes quote inputs into known cells, recalculates net revenue,
gross margin, and approval decision, serializes the WorkPaper JSON, restores
it, and verifies `restoredMatchesAfter: true`.

## If you arrived from HN or LibHunt

The useful question is not whether Bilig is a smaller Excel clone. It is whether
your service needs a workbook-shaped calculation boundary that can be tested
without opening a browser spreadsheet.

Use the smallest proof that matches your blocker:

- If you need formulas inside a backend service or agent tool, run the
  90-second npm quickstart above and check for `verified: true`.
- If you are blocked by stale XLSX cached formula results, run the
  [XLSX recalculation proof](xlsx-recalculation-proof.md) and compare the
  before/after/exported readback.
- If you are comparing engines, start with the
  [headless spreadsheet engine comparison](headless-spreadsheet-engine-comparison.md)
  before looking at benchmark cards.

Bilig is indexed on the
[LibHunt headless-spreadsheet topic](https://www.libhunt.com/topic/headless-spreadsheet),
but the repo is the current source of truth for stars, releases, limitations,
and API examples. If one proof maps to your use case, star the repo as a public
bookmark. If it almost maps but fails on a real formula, import, or workflow
boundary, open an adoption blocker with the smallest reproducer you can share.

## Backend service shape

A service should keep the WorkPaper behind a narrow business boundary:

1. Load or build the workbook.
2. Accept typed business inputs.
3. Write those inputs into known cells.
4. Read computed values from known output cells.
5. Persist the exported WorkPaper document.

The canonical service example is the
[serverless WorkPaper API](../examples/serverless-workpaper-api). It includes a
quote approval API, framework adapters, and persistence adapters for common
Node deployment shapes.

Use the [Node service WorkPaper recipe](node-service-workpaper-recipe.md) when
you want the smallest local service boundary before adopting a framework.

## Agent tool shape

An agent should not infer workbook state from screenshots when it needs a
durable result. Expose explicit tools instead:

- `list_sheets`
- `read_cell`
- `read_range`
- `set_cell_contents`
- `get_cell_display_value`
- `export_workpaper_document`

That makes writeback testable: the tool call writes a cell, the runtime
recalculates, and the agent reads the exact output cell before claiming success.

Start with the [agent tool-calling recipe](agent-workpaper-tool-calling-recipe.md)
or the [MCP spreadsheet tool server](mcp-workpaper-tool-server.md).

## Decision boundary

Use `@bilig/headless` when the runtime requirement is formula-backed workbook
state inside TypeScript. Use other tools when the primary requirement is
different:

| Requirement                                                    | Start with                                              |
| -------------------------------------------------------------- | ------------------------------------------------------- |
| Formula-backed business logic inside a Node service            | `@bilig/headless` WorkPaper                             |
| Agent writeback with verified readback                         | WorkPaper tools or the MCP server                       |
| XLSX file import, export, styling, and reports                 | SheetJS, ExcelJS, or the `@bilig/headless/xlsx` subpath |
| Broad formula coverage in an established headless engine       | HyperFormula comparison first                           |
| Shared Google Workspace document with permissions and comments | Google Sheets API                                       |

Read the [headless spreadsheet engine comparison](headless-spreadsheet-engine-comparison.md)
when the boundary is not obvious.

## Proof links

- [90-second Node quickstart](try-bilig-headless-in-node.md)
- [Quote approval WorkPaper API proof](quote-approval-workpaper-api.md)
- [Node spreadsheet formula engine guide](node-spreadsheet-formula-engine.md)
- [Server-side spreadsheet automation in Node](server-side-spreadsheet-automation-node.md)
- [Google Sheets API boundary](google-sheets-api-alternative-node-workpaper.md)
- [SheetJS and ExcelJS boundary](sheetjs-exceljs-alternative-formula-workbook-api.md)
- [HyperFormula alternative notes](hyperformula-alternative-headless-workpaper.md)
- [What the benchmark proves](what-workpaper-benchmark-proves.md)
- [Compatibility limits](where-bilig-is-not-excel-compatible-yet.md)

If this matches a backend or agent workflow you are evaluating, star the repo as
a bookmark: <https://github.com/proompteng/bilig/stargazers>.

If it almost matches but a gap blocks adoption, use the adoption blocker form:
<https://github.com/proompteng/bilig/discussions/new?category=general>.
