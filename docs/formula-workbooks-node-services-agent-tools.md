---
title: Formula workbooks for Node services and agent tools
published: true
description: A runnable @bilig/headless proof for backend services and agent tools that need formula cell edits, recalculation, JSON persistence, and restored readback.
tags: typescript, node, spreadsheet, agents, formulas
canonical_url: https://proompteng.github.io/bilig/formula-workbooks-node-services-agent-tools.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Formula workbooks for Node services and agent tools

`@bilig/headless` is for one specific job: keep workbook-shaped business logic
as cells and formulas, but run it from a Node service, queue worker, serverless
route, test, or agent tool.

The runtime boundary is a `WorkPaper`: write known input cells, recalculate,
read known output cells, and persist the workbook as JSON. No browser grid,
spreadsheet account, screenshot automation, or hosted sheet is required.

## The evaluator path

Run the quote approval proof from an empty directory:

```sh
mkdir bilig-quote-approval
cd bilig-quote-approval
npm init -y
npm pkg set type=module
npm install @bilig/headless
npm install -D tsx typescript @types/node
curl -fsSLo quote-approval-api.ts \
  https://raw.githubusercontent.com/proompteng/bilig/main/examples/serverless-workpaper-api/quote-approval-api.ts
npx tsx quote-approval-api.ts
```

Expected shape:

```json
{
  "route": "Quote approval WorkPaper API",
  "before": {
    "netRevenue": 43200,
    "grossMargin": 0.2963,
    "decision": "review"
  },
  "edit": {
    "after": {
      "netRevenue": 45600,
      "grossMargin": 0.3333,
      "decision": "approved"
    },
    "checks": {
      "decisionChanged": true,
      "formulasPersisted": true,
      "inputPersisted": true,
      "restoredMatchesAfter": true
    }
  },
  "verified": true
}
```

The important check is `restoredMatchesAfter: true`: the service wrote inputs
into cells, formulas recalculated, the WorkPaper serialized to JSON, and a
restored WorkPaper produced the same decision.

## Why this is not just a formula parser

A backend evaluator usually needs more than `=A1+B1`:

- typed business input maps to stable cell addresses
- dependent formulas update after each write
- output values are read back from the workbook runtime
- formulas survive persistence
- restored state can be tested without Excel, Google Sheets, or a browser

That shape fits pricing rules, quote approval, payout checks, budget guardrails,
import validation, and agent tools where the user wants a durable result rather
than a plausible screenshot.

## Agent tool boundary

For agents, expose small tools:

- `list_sheets`
- `read_cell`
- `read_range`
- `set_cell_contents`
- `get_cell_display_value`
- `export_workpaper_document`

The published MCP server can run against a local WorkPaper JSON file:

```sh
npm exec --package @bilig/headless@0.40.21 -- \
  bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable
```

With `--writable`, `set_cell_contents` persists the updated WorkPaper JSON back
to the file. The point is not to let an agent "look" at a spreadsheet. The point
is to let the agent make a workbook edit and prove the recalculated value it
used.

## Where other tools may be better

Use HyperFormula first when you need a mature, broad formula engine with a large
function surface and the WorkPaper persistence shape is not the differentiator.

Use SheetJS or ExcelJS first when the primary job is reading, writing, styling,
or generating spreadsheet files. Bilig's XLSX paths are secondary to the
WorkPaper runtime boundary.

Use Google Sheets API first when a shared hosted spreadsheet, permissions,
comments, and human collaboration are the actual product requirement.

Use `@bilig/headless` when a Node service or agent tool owns the workbook state
and needs formula readback, persistence, and restore proof.

## Evidence and limits

The current checked benchmark artifact records `100/100` comparable
mean-latency wins against HyperFormula-style workloads. The narrowest p95 row is
public: `aggregate-overlapping-sliding-window-wide` is a `0.946x`
WorkPaper-to-HyperFormula p95 ratio, and browser UI rendering is outside the
benchmark.

This page does not claim full Excel compatibility. It does not claim chart,
macro, formatting, collaborative editing, or complete XLSX fidelity. Start with
the compatibility page before importing real workbooks.

## Start here

- [Quote approval WorkPaper API proof](quote-approval-workpaper-api.md)
- [Try `@bilig/headless` in Node](try-bilig-headless-in-node.md)
- [MCP spreadsheet tool server](mcp-workpaper-tool-server.md)
- [What the WorkPaper benchmark proves](what-workpaper-benchmark-proves.md)
- [Where bilig is not Excel-compatible yet](where-bilig-is-not-excel-compatible-yet.md)

If this matches a backend or agent workflow you are evaluating, star the repo as
a bookmark: <https://github.com/proompteng/bilig/stargazers>.

If it does not match yet, use the adoption-blocker form instead of a vague
comment:
<https://github.com/proompteng/bilig/discussions/new?category=general>. The
most useful feedback names the workflow, blocker type, proof already tried, and
the smallest example, benchmark, or compatibility note that would unblock a
real trial.
