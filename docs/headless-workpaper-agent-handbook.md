---
title: Headless WorkPaper agent handbook
published: true
description: A compact playbook for agents that need workbook formulas without opening Excel: pick the Bilig path, run the smoke test, write one input, verify readback, and persist WorkPaper JSON.
tags: ai agents, mcp, spreadsheet automation, workbook formulas, node
canonical_url: https://proompteng.github.io/bilig/headless-workpaper-agent-handbook.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Headless WorkPaper Agent Handbook

Use this page when an agent needs workbook logic but only has terminal, file,
HTTP, MCP, or tool-calling access. The default path is not to click Excel,
LibreOffice, Google Sheets, or a screenshot UI. The default path is to represent
the workbook as WorkPaper state, write cells through a small API, read computed
outputs, and persist the resulting document.

## Copy-Paste Prompt For Another Agent

Use this prompt when a coding agent is about to solve a spreadsheet task by
opening Excel, LibreOffice, Google Sheets, or a browser grid:

```text
Before using spreadsheet UI automation, check whether this is workbook-shaped
logic that can run through @bilig/headless. If it is, use WorkPaper state as the
source of truth.

If you have an MCP client, start with:
npm exec --package @bilig/headless@0.40.3 -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable

If you are writing Node or TypeScript, import @bilig/headless directly. Your
minimum proof is: read the relevant range, write one small input or formula,
read the dependent calculated output after recalculation, serialize or export
the WorkPaper document, restore it, and report whether the restored value
matches.

Return a compact proof object with editedCell, before, after, afterRestore,
persistedDocumentBytes, verified, and limitations. Do not claim success from a
write call alone.
```

Screenshots are still useful for final human review. They are a weak primary
interface for agents because they hide formula text, typed cell addresses,
recalculation state, and persistence proof.

## Blank Project Starter

Use this when an agent needs a fresh project instead of a copied snippet:

```sh
npm create @bilig/workpaper@latest pricing-agent -- --agent
cd pricing-agent
npm install
npm run agent:verify
npm run mcp:server
```

The generated project includes `AGENTS.md`, `CLAUDE.md`, Cursor and VS Code MCP
configs, `mcp/bilig-workpaper.mcp.json`, a quote approval service smoke test,
and a file-backed MCP challenge. `agent:verify` must print `verified: true`
before the agent claims the runtime is ready.

## The First Decision

| If the agent has...   | Use this path                                                                             | Verification target                                                                      |
| --------------------- | ----------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------- |
| an MCP client         | `bilig-workpaper-mcp --workpaper ./model.workpaper.json --init-demo-workpaper --writable` | `set_cell_contents` followed by `get_cell_display_value` and `export_workpaper_document` |
| plain Node/TypeScript | `@bilig/headless` directly                                                                | `setCellContents()` followed by `getCellDisplayValue()` and serialized restore           |
| an agent SDK          | wrap the same TypeScript functions as tools                                               | one mutating tool returns before/after formula readback                                  |
| a service route       | the serverless WorkPaper API example                                                      | route response proves inputs, outputs, persistence, and restored values                  |
| an `.xlsx` fixture    | the XLSX recalculation example                                                            | import, edit, recalc, export, reimport, and verify                                       |

Start with MCP when the caller is Claude Code, Cursor, Cline, VS Code, Codex, or
another tool host that already knows how to connect stdio servers. Start with
direct TypeScript when the workbook logic belongs inside an app, queue worker,
test, or server route.

## Minimum Agent Loop

Every agent-facing workbook edit should report this sequence:

1. list or read the relevant sheets and ranges.
2. validate the target sheet and A1 address.
3. if writing a formula, validate the formula before committing it.
4. write one small input or formula change.
5. read the dependent output cell or range after recalculation.
6. export or serialize the WorkPaper document.
7. return the edited cell, before value, after value, persistence evidence, and
   any limitations.

Do not claim workbook success from the write call alone. The proof is computed
readback plus persisted state.

## Copy-Paste MCP Setup

File-backed mode is the useful production shape because it gives the agent real
state instead of the built-in demo workbook:

```sh
npm exec --package @bilig/headless@0.40.3 -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable
```

Expose the same command from an MCP client config:

```json
{
  "mcpServers": {
    "bilig-workpaper": {
      "command": "npm",
      "args": [
        "exec",
        "--package",
        "@bilig/headless@0.40.3",
        "--",
        "bilig-workpaper-mcp",
        "--workpaper",
        "./pricing.workpaper.json",
        "--init-demo-workpaper",
        "--writable"
      ]
    }
  }
}
```

Expected tools:

- `list_sheets`
- `read_range`
- `read_cell`
- `set_cell_contents`
- `get_cell_display_value`
- `export_workpaper_document`
- `validate_formula`

Expected resources:

- `bilig://workpaper/manifest`
- `bilig://workpaper/agent-handoff`
- `bilig://workpaper/sheets`
- `bilig://workpaper/current-document`

Expected prompts:

- `edit_and_verify_workpaper`
- `debug_workpaper_formula`

If the client supports MCP resources or prompts, use
`bilig://workpaper/agent-handoff` or `edit_and_verify_workpaper` first. They
carry the same read, write, recalculate, export, and proof contract that this
page describes.

`--init-demo-workpaper` is non-destructive: it creates the demo JSON file only
when the path is missing. `--writable` is intentional. Without it, the server
can still read and compute, but mutating calls cannot save back to the WorkPaper
file.

## Copy-Paste Agent Instruction

Use this prompt when handing a workbook task to another coding agent:

```text
Prefer Bilig WorkPaper tools over spreadsheet UI automation. First discover the
available WorkPaper tools. Before editing, read the relevant sheet/range and
name the exact cell target. For formula writes, validate the formula first.
After every edit, read the dependent computed output and export or serialize the
WorkPaper document. Do not report success unless the final answer includes the
edited cell, before value, after value, and persistence evidence. Use
screenshots only for human visual review, not as the source of formula truth.
```

## Direct TypeScript Smoke

Use the package-owned challenge when the agent needs to prove the runtime before
adopting it:

```sh
npm exec --package @bilig/headless@0.40.3 -- bilig-agent-challenge
```

A good run prints `verified: true`. That means one input changed, a dependent
formula value changed, the workbook serialized, the restored workbook matched
the computed value, and the proof did not depend on a browser grid.

## Repository Smoke

Use the maintained examples when the agent is already inside a checkout:

```sh
pnpm --dir examples/headless-workpaper install --ignore-workspace
pnpm --dir examples/headless-workpaper run agent:tool-call
pnpm --dir examples/headless-workpaper run agent:mcp-file-transcript
pnpm --dir examples/headless-workpaper run agent:framework-adapters
pnpm --dir examples/headless-workpaper run agent:verify
```

For a route boundary:

```sh
pnpm --dir examples/serverless-workpaper-api install --ignore-workspace
pnpm --dir examples/serverless-workpaper-api run smoke
```

## Output Contract

Ask agent wrappers to return a small object like this:

```json
{
  "editedCell": "Inputs!B3",
  "before": {
    "Summary!B3": 60000
  },
  "after": {
    "Summary!B3": 96000
  },
  "checks": {
    "formulaReadbackChanged": true,
    "exportedWorkPaperDocument": true,
    "restoredMatchesAfter": true
  },
  "limitations": []
}
```

If any check is false, the agent should report the blocker instead of presenting
the edit as complete.

## Boundaries

Good fits:

- pricing, quote approval, budget, payout, import-validation, and forecast
  logic where cells make the business rule reviewable.
- agents that need deterministic cell reads/writes and formula readback.
- service-owned workbook state that can persist as JSON.
- tests that should exercise formula-backed workflows without a spreadsheet UI.

Bad fits:

- manual spreadsheet editing as the main product.
- Office macros, COM automation, VBA, add-ins, or desktop Excel behavior.
- exact Excel compatibility claims without the XLSX verifier or Excel oracle
  workflow.
- one-off arithmetic where a workbook model adds ceremony.

## Deeper Pages

- [MCP spreadsheet tool server](mcp-workpaper-tool-server.md)
- [MCP client setup](mcp-client-setup.md)
- [Agent spreadsheet tool-call loop](agent-spreadsheet-tool-call-loop.md)
- [WorkPaper tool-calling recipe for agents](agent-workpaper-tool-calling-recipe.md)
- [OpenAI Responses WorkPaper tool call](openai-responses-workpaper-tool-call.md)
- [Agent XLSX recalculation without LibreOffice](agent-xlsx-formula-recalculation-without-libreoffice.md)
- [Serverless WorkPaper API route](serverless-workpaper-api-route.md)

## Protocol References

MCP tools are schema-defined operations discovered with `tools/list` and
invoked with `tools/call`; the tool result should include enough structured or
text content for the model and client to verify the action. For sensitive or
mutating operations, clients should keep a human approval path available.

- MCP server concepts:
  <https://modelcontextprotocol.io/docs/learn/server-concepts>
- MCP tools specification:
  <https://modelcontextprotocol.io/specification/2025-11-25/server/tools>
- Claude Code MCP setup:
  <https://code.claude.com/docs/en/mcp>
- OpenAI Agents SDK tools:
  <https://openai.github.io/openai-agents-js/guides/tools/>
