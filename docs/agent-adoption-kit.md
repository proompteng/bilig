---
title: Agent WorkPaper handoff for formulas
published: true
description: One no-key handoff path for coding agents: install the Bilig WorkPaper skill, run the agent MCP evaluator, paste the workbook task, and return verified formula readback.
tags: agents, mcp, workbook formulas, spreadsheet automation, node
canonical_url: https://proompteng.github.io/bilig/agent-adoption-kit.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Tool Host WorkPaper Handoff

Use this page when you are deciding whether Bilig should be the workbook tool
inside a coding agent, MCP client, or agent framework. The path is deliberately
short: install the instructions, run one no-key evaluator, paste one workbook
task, and require formula readback before calling the job done.

## Choose Bilig When Readback Matters

Choose Bilig instead of Excel or Sheets UI automation when the agent needs a
machine-checkable workbook proof, not a visual session transcript.

| Agent need                 | Use Bilig WorkPaper                                                                                                               | Keep Excel, Sheets, or browser UI automation                                                                     |
| -------------------------- | --------------------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------- |
| Cell-addressed automation  | The task names exact sheets, ranges, or cells such as `Inputs!B3` and `Summary!B3`.                                               | The task depends on a person inspecting or editing the live grid.                                                |
| Formula proof after writes | Success requires formula readback after writes through `set_cell_contents_and_readback`, not a write-call status.                 | The spreadsheet product owns recalculation and the agent only needs to click through the product UI.             |
| Durable agent state        | The agent must export, persist, restore, or restart from persisted WorkPaper JSON and prove the same calculated value comes back. | The durable artifact is an existing Excel or Sheets file with macros, charts, pivots, add-ins, or visual layout. |
| CI and tool-call evidence  | The result must fit logs, MCP transcripts, CI checks, or a compact proof object with `verified: true`.                            | The output is a screenshot, screen recording, manual review note, or collaboration comment in a spreadsheet app. |

Run the no-key proof before wiring the path:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
```

For the browser/agent boundary, pair this page with the
[Browser Use WorkPaper formula tool](browser-use-workpaper-formula-tool.md), the
[headless WorkPaper agent handbook](headless-workpaper-agent-handbook.md), and
the [agent framework map](agent-framework-workbook-tools.md).

## Avoid Stored-Result False Negatives

Agents often test spreadsheet code by writing an `.xlsx` file with ExcelJS,
SheetJS, openpyxl, pandas, or another file library, then reading the formula
cell back in the same process. That is a false-negative trap: many file
libraries can preserve or write formula records without recalculating the stored
result. The agent may blame its own code, or claim a workbook failed, when it
only read unchecked file metadata.

Use Bilig at the decision boundary when the agent, CI job, route handler, or
queue worker must trust the value before Excel or Sheets opens the file:

- If the workbook can stay headless, use `@bilig/workpaper` and require
  `set_cell_contents_and_readback`, export or restore evidence, and
  `verified: true`.
- If the workflow must keep an `.xlsx` file, run
  `@bilig/xlsx-formula-recalc` or `@bilig/exceljs-formula-recalc` at the
  recalculation boundary, then read the calculated cells.
- If the task is a visual layout, chart, pivot, macro, add-in, or human review
  problem, keep Excel, Sheets, or browser UI automation in the loop and do not
  overclaim Bilig compatibility.

This is the shortest rule to paste into another agent:

```text
Do not grade workbook automation from unchecked stored formula results. After a formula or
input edit, prove the dependent calculated cell changed through WorkPaper
readback or an XLSX recalculation bridge before reporting success or failure.
```

## Install The Agent Instructions

If your agent supports installable skills, start here:

```sh
npx --yes skills@latest add https://bilig.proompteng.ai --list
npx --yes skills@latest add proompteng/bilig --skill bilig-workpaper --list
```

Use the app-host discovery URL first. Keep the GitHub repo skill command as a
fallback for hosts that only support GitHub skill sources.

If the agent is already inside a cloned Bilig checkout, do not copy a host list
from this page. Use the [coding agent rule chooser](agent-rule-chooser.md); it
maps each host to the current repo-local rule and MCP config. Common anchors are
`CLAUDE.md`, `.claude/skills/bilig-workpaper/SKILL.md`,
`.claude/commands/bilig-workpaper-proof.md`,
`.github/instructions/bilig-workpaper.instructions.md`,
`.github/prompts/bilig-workpaper-proof.prompt.md`, `.vscode/mcp.json`,
`.cursor/rules/bilig-workpaper.mdc`, `.kiro/steering/bilig-workpaper.md`,
`.trae/mcp.json`, `.trae/rules/bilig-workpaper.md`, `.zed/settings.json`,
`opencode.jsonc`, and `.continue/mcpServers/bilig-workpaper.yaml`.

The published starter overlay is release-pending. Do not use
`npm create @bilig/workpaper@latest` while `@bilig/create-workpaper@latest`
resolves to `0.164.11`; that release's generated smoke reports
`formulasPersisted: false`. Until a newer release passes a fresh consumer
smoke, use the rule chooser to copy only the host files you need and run the
direct `npm exec` MCP commands above.

For web fetch, give the agent the compact map first:

```text
https://proompteng.github.io/bilig/llms.txt
```

When a reviewer wants to see a successful run before adopting the path, require
fresh evaluator JSON from `bilig-evaluate --door agent-mcp --json`.

## Agent Manifest Gate

When an agent host, directory scanner, or internal platform wants machine-readable
entrypoints, start with:

```text
https://proompteng.github.io/bilig/.well-known/agent.json
```

Accept the integration only if the manifest exposes `public_entrypoints`,
`evaluator_doors`, `proof_contract`, and the `mcp` server block. Those fields
let the host find the compact start file, choose the right evaluator door, and
verify that success means computed readback plus persisted state.

For a human-readable decision object, run:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-agent-start --json
```

Then validate the same boundary with the agent MCP evaluator:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
```

Do not treat a manifest link, installed rule file, or MCP server listing as
success by itself. The gate passes only when the evaluator returns the proof
fields in `proof_contract`, including `editedCell`, `before`, `after`,
`afterRestore`, `persistedDocumentBytes`, and `verified`.

## Run The No-Key Check

This checks the published package and the file-backed MCP tool path without
cloning the repo or using an API key:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
```

For a richer workbook check, use the revenue-plan scenario:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --scenario revenue-plan --json
```

That scenario proves `SUM`, `SUMIF`, `XLOOKUP`, `FILTER`, a named expression,
JSON persistence, and restart readback through the same MCP door.

If the workbook includes provider-backed formulas such as `IMPORTRANGE`, run the
adapter-boundary check:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --scenario provider-backed --json
```

That check should show `#BLOCKED!` and `provider-backed-adapter-missing` before
a local synthetic adapter is installed, then a fresh `96000` readback with
diagnostics cleared after the adapter path runs. It does not call Google Sheets.

A passing run must return `schemaVersion: "bilig-evaluator.v1"`,
`door: "agent-mcp"`, `verified: true`, and these checks:

- tools, resources, and prompts were discovered;
- one input cell changed;
- a dependent formula cell changed after recalculation;
- WorkPaper JSON was exported and persisted;
- restart readback matched the post-edit value.

Use the raw MCP challenge only when you need the lower-level JSON-RPC proof:

```sh
npm exec --package @bilig/workpaper@latest -- bilig-mcp-challenge --json
```

Use the service evaluator when the agent will import `@bilig/workpaper` instead
of using MCP:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door workpaper-service --json
```

## Wire The Local MCP Server

Use file-backed stdio for private project state:

```json
{
  "mcpServers": {
    "bilig-workpaper": {
      "command": "npm",
      "args": [
        "exec",
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

If the repository already has a workbook, create the file-backed WorkPaper from
that XLSX first:

```sh
npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp --from-xlsx ./pricing.xlsx --workpaper ./.bilig/pricing.workpaper.json --writable
```

That command imports the XLSX once, refuses to replace an existing WorkPaper
JSON unless `--overwrite-workpaper` is present, and then exposes the same
`read_cell`, `set_cell_contents_and_readback`, `export_workpaper_document`, and
read-only `analyze_workbook_risk` tools.

Use the hosted endpoint only for smoke tests and tool discovery:

```text
https://bilig.proompteng.ai/mcp
```

The hosted endpoint is stateless. It is not where private workbook files live.

## Paste This Task Into An Agent

```text
Use Bilig WorkPaper tools instead of spreadsheet UI automation. Build or load a
small workbook with Inputs!B2 as customers, Inputs!B3 as average revenue, and
Summary!B3 as the revenue formula. First read the relevant input and summary
range. Then set Inputs!B3 to 4800, read Summary!B3 after recalculation, export
or serialize the WorkPaper document, restore it, and return editedCell, before,
after, afterRestore, persistedDocumentBytes, verified, and limitations.

Do not claim success from a write call alone. Success requires computed
readback plus persisted or restored state.
```

## Expected Result

The exact values can change, but the evaluator result should look like this:

```json
{
  "schemaVersion": "bilig-evaluator.v1",
  "door": "agent-mcp",
  "verified": true,
  "packageVersions": {
    "@bilig/workpaper": "0.164.11",
    "xlsx-formula-recalc": "0.164.11"
  },
  "evidence": {
    "editedCell": "Inputs!B3",
    "dependentCell": "Summary!B3",
    "before": 60000,
    "after": 96000,
    "afterRestore": 96000,
    "afterRestart": 96000,
    "checks": {
      "listedFileBackedTools": true,
      "listedResourcesAndPrompts": true,
      "formulaValidationPassed": true,
      "dependentCellChanged": true,
      "persistedToDisk": true,
      "exportContainsWorkPaperDocument": true,
      "restartReadbackMatchesAfter": true
    }
  }
}
```

Reject answers that only say a cell was written. The point of Bilig is that the
agent returns the calculated result and saved state from the workbook itself.

## Upstream Maintainer Notes

When you want a third-party agent host, MCP client, framework, or docs site to
add Bilig, use the same proof bar and avoid duplicate threads.

Before opening anything upstream:

- search that project for existing Bilig issues, PRs, examples, and docs links;
- run the no-key `agent-mcp` evaluator against the currently published package;
- decide whether the host needs a local file-backed MCP config, a hosted
  stateless smoke endpoint, an installable rule file, or only a short docs note;
- keep one thread per project and update it in place when proof changes.

After publishing new `bilig-agent-start --rules` targets, run the public-latest
smoke before pointing maintainers at the page:

```sh
pnpm agent:public-rules:check
```

The first upstream message should be a maintainer question, not a drive-by
listing:

```text
Would you accept a small docs example for deterministic spreadsheet formula
readback in this agent host? The no-key proof is:

npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json

The useful evidence is `verified: true`, `editedCell`, `before`, `after`,
`afterRestore` or `afterRestart`, and persisted WorkPaper JSON bytes. I can keep
the PR limited to this host's documented MCP/rules surface and close it if it is
out of scope.
```

Do not open duplicate issues, duplicate PRs, or broad directory submissions
when a project already has an active Bilig thread. A merged integration, an
accepted issue, or a maintainer-requested PR is useful evidence; a submitted
form by itself is not.

## After The Check

If the check matches your workflow, keep the repository nearby:
<https://github.com/proompteng/bilig>.

If you need release notifications for agent or MCP changes, watch releases:
<https://github.com/proompteng/bilig/subscription>.

If it almost works but the workflow is blocked, open the concrete blocker:
<https://github.com/proompteng/bilig/discussions/new?category=general>.

## Next Pages

- [Evaluate Bilig as an agent MCP workbook tool](eval-agent-mcp.md)
- [WorkPaper agent handbook](headless-workpaper-agent-handbook.md)
- [MCP client setup](mcp-client-setup.md)
- [OpenHands WorkPaper MCP setup](openhands-workpaper-mcp.md)
- [Trae WorkPaper MCP setup](trae-workpaper-mcp.md)
- [Qodo WorkPaper MCP setup](qodo-workpaper-mcp.md)
- [OpenCode WorkPaper MCP setup](opencode-workpaper-mcp.md)
- [Agent workbook challenge](agent-workbook-challenge.md)
- [Workbook tools for agent frameworks](agent-framework-workbook-tools.md)
