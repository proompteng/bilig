---
title: Spreadsheet MCP server comparison
published: true
description: 'Compare spreadsheet MCP server choices for agents: Excel file tools, Google Sheets tools, read-only workbook inspection, and Bilig WorkPaper formula readback.'
tags: mcp, model context protocol, spreadsheet, excel, agents
canonical_url: https://proompteng.github.io/bilig/spreadsheet-mcp-server-comparison.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Spreadsheet MCP Server Comparison

Spreadsheet MCP servers are not one category. Some are file editors. Some are
Google Sheets API wrappers. Some inspect workbooks for an agent without writing
anything. Bilig WorkPaper is narrower: a local formula-backed workbook runtime
that lets an agent write known input cells, recalculate, and return structured
readback.

Use this page when you are choosing an MCP tool surface for agent workflows that
touch spreadsheet-shaped business logic.

## Quick Decision Table

| Need                                                                                                     | Better starting point                                       |
| -------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------- |
| Read and write arbitrary `.xlsx` files with formatting, charts, and workbook layout                      | Excel-focused MCP server or an Office automation workflow   |
| Read and update Google Sheets through a live cloud spreadsheet                                           | Google Sheets MCP server                                    |
| Let an agent inspect workbook structure, formulas, and cached values without mutating files              | Read-only spreadsheet inspection MCP server                 |
| Mutate service-owned workbook inputs, recalculate formulas, verify before/after values, and persist JSON | Bilig WorkPaper MCP                                         |
| Exact Excel compatibility across macros, pivots, charts, external links, and every function              | Excel, LibreOffice, Graph API, or a dedicated Excel runtime |

## Named Public Alternatives

Use the existing spreadsheet MCP ecosystem when the source of truth is already
somewhere else:

| Server or path                                                          | Best fit                                                                                 | Boundary to check before adopting                                                               |
| ----------------------------------------------------------------------- | ---------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------- |
| [Google Sheets MCP](https://github.com/henilcalagiya/google-sheets-mcp) | Agents that need CRUD operations against live Google Sheets through a service account    | Requires Google Cloud, Sheets API, Drive API, and service-account setup                         |
| [Univer MCP](https://github.com/dream-num/univer-mcp)                   | Agents that operate a Univer spreadsheet runtime through an MCP session                  | Requires an API key and a running Univer instance; the repo labels plain-text mode experimental |
| [GRID MCP](https://github.com/GRID-is/claude-mcp)                       | Claude Desktop workflows against spreadsheets uploaded to GRID                           | Requires a GRID account, uploaded workbook, and API key                                         |
| Excel file or SheetJS-style tooling                                     | Creating, reading, or preserving `.xlsx` files                                           | A file library can preserve formulas without recalculating fresh results in Node                |
| Bilig WorkPaper MCP                                                     | Local agent tools that own WorkPaper JSON and need write, recalculate, readback, restore | Not a full Excel editor; use it when formula readback is the product                            |

That split is useful for outreach too. Do not pitch Bilig as "another Google
Sheets MCP server" or "another Excel file editor." Pitch it where the agent
needs a local formula runtime and a machine-checkable proof object after an
edit.

## Where Bilig Fits

The Bilig MCP server is for workflows where the workbook is the service model,
not merely a file attachment. The useful loop is:

1. load a WorkPaper JSON document or the built-in demo workbook;
2. list sheets or read a range;
3. write one input cell;
4. read the recalculated display value;
5. export or persist the updated WorkPaper document.

That makes it a fit for quote approvals, payout checks, budget alerts,
import-validation workbooks, and agent tools that need proof of what changed.

It is not a replacement for a full Excel file editor. It should not be sold as
one.

## Formula Recalculation Is The Split

The important question is not "does this MCP server work with spreadsheets?"
It is "can the agent trust a formula result immediately after it writes an
input?"

Many spreadsheet MCP servers are intentionally file-oriented. That is useful
when the job is report generation, workbook inspection, or careful `.xlsx`
mutation. It is not the same as a formula-runtime loop. For example, SheetForge
MCP documents that its read tools do not recalculate Excel formulas and instead
surface formula cells as formula text. That is the right behavior for a file
editor that should not invent fresh values.

The same user pain shows up outside MCP. A long-running SheetJS issue asks
whether a formula value can be refreshed after changing an input cell, and an
ExcelJS discussion describes JSON-driven workbook edits where shared formulas
and calculated results only become trustworthy after opening and saving in a
spreadsheet application. Those threads are not Bilig marketing claims; they are
evidence that "write XLSX" and "trust a recalculated value in Node" are separate
requirements.

Bilig takes the opposite boundary for service-owned workbooks:

- the persisted artifact is WorkPaper JSON, not an opaque Excel cache;
- the agent writes a known input cell;
- formulas recalculate inside the runtime;
- the agent reads a display value or raw value after the edit;
- the updated WorkPaper document can be exported and restored for audit.

That makes the comparison less about "best spreadsheet MCP server" and more
about the source of truth. Use file-first MCP tools when Excel fidelity is the
product. Use Bilig WorkPaper MCP when recalculated readback is the product.

## Verify The Bilig MCP Path

Install and list the packaged server:

```sh
npm exec --package @bilig/headless@0.40.21 -- bilig-workpaper-mcp
```

Run the maintained JSON-RPC transcript from a clone:

```sh
git clone --depth 1 https://github.com/proompteng/bilig.git
cd bilig
pnpm --dir examples/headless-workpaper install --ignore-workspace
pnpm --dir examples/headless-workpaper run agent:mcp-transcript
```

The transcript edits `Inputs!B3`, recalculates dependent formulas, serializes
the WorkPaper document, restores it, and verifies that the restored values match
the post-edit values.

For a persisted workbook file:

```sh
npm exec --package @bilig/headless@0.40.21 -- \
  bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable
```

File-backed mode exposes tools such as `list_sheets`, `read_range`,
`set_cell_contents`, `get_cell_display_value`, `export_workpaper_document`, and
`validate_formula`.

## What To Ask Before Choosing A Spreadsheet MCP Server

- Is the source of truth an Excel file, a Google Sheet, or service-owned
  workbook state?
- Does the agent need to write cells, or only inspect them?
- Is stale cached formula data acceptable, or must the tool recalculate before
  responding?
- Does the workflow need exact file fidelity, or only auditable formula
  readback?
- What artifact proves the agent's edit: a screenshot, a saved file, or
  machine-checkable before/after values?

If the answer is "the backend must trust a recalculated value before it returns
or persists anything," choose a formula runtime path and keep the MCP layer thin.

## Related Bilig Pages

- [MCP spreadsheet tool server for WorkPaper agents](mcp-workpaper-tool-server.md)
- [MCP client setup](mcp-client-setup.md)
- [MCP spreadsheet server directory status](mcp-spreadsheet-server-directory.md)
- [Agent spreadsheet tool call loop](agent-spreadsheet-tool-call-loop.md)
- [Why agents need workbook APIs](why-agents-need-workbook-apis.md)
- [Stop driving spreadsheets with screenshots](stop-driving-spreadsheets-with-screenshots.md)

## Public Directory References

- [SheetForge MCP](https://mcpservers.org/servers/iheldan/sheetforge-mcp)
- [Excel file manipulation MCP](https://mcp.directory/servers/excel-file-manipulation)
- [Bilig WorkPaper MCP registry search](https://registry.modelcontextprotocol.io/v0.1/servers?search=io.github.proompteng%2Fbilig-workpaper)
- [Bilig WorkPaper on Glama](https://glama.ai/mcp/servers/proompteng/bilig)

If this is the MCP boundary you were looking for, star or bookmark Bilig so the
server is easier for other agent builders to find:
<https://github.com/proompteng/bilig/stargazers>.
