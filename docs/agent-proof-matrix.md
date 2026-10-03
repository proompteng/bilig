---
title: Agent WorkPaper evaluator matrix
published: true
description: Pick the smallest Bilig WorkPaper proof for coding agents, MCP clients, AI SDK tools, OpenAI tool calls, LangGraph, Microsoft Agent Framework, Semantic Kernel, Mastra, and XLSX recalculation.
tags: agents, mcp, workpaper, spreadsheet automation, proof
canonical_url: https://proompteng.github.io/bilig/agent-proof-matrix.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Agent WorkPaper Evaluator Matrix

Use this page before an agent drives Excel, LibreOffice, Google Sheets, or a
browser grid. Pick the smallest proof that writes an input, recalculates a
dependent formula, reads the value back, and preserves enough state for another
process to check the result.

If you only run one command, run the agent MCP evaluator:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
```

Expected invariants:

```json
{
  "schemaVersion": "bilig-evaluator.v1",
  "door": "agent-mcp",
  "verified": true,
  "evidence": {
    "editedCell": "Inputs!B3",
    "dependentCell": "Summary!B3",
    "before": 60000,
    "after": 96000,
    "afterRestore": 96000,
    "afterRestart": 96000,
    "persistedDocumentBytes": 1162,
    "checks": {
      "listedFileBackedTools": true,
      "listedResourcesAndPrompts": true,
      "dependentCellChanged": true,
      "persistedToDisk": true,
      "restartReadbackMatchesAfter": true
    }
  }
}
```

## Evaluator Matrix

| Evaluator or contract            | Command or asset                                                                                                                                  | Expected JSON field                                                                                                    | What it verifies                                                                                                                                                                    | Explicit limits                                                                                                                                            |
| -------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------- |
| WorkPaper service                | `npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door workpaper-service --json`                                              | `door: "workpaper-service"`, `verified: true`                                                                          | Node can edit a WorkPaper input, recalculate a formula, export JSON, restore it, and verify readback.                                                                               | MCP discovery, private workbook compatibility, macros, pivots, charts, or Excel UI behavior.                                                               |
| Agent MCP evaluator              | `npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json`                                                      | `door: "agent-mcp"`, `listedResourcesAndPrompts`, `restartReadbackMatchesAfter`                                        | A coding agent or MCP client can discover workbook tools, write a cell, read a formula value, persist state, and restart from disk.                                                 | Hosted auth, arbitrary client UX, or full workbook compatibility.                                                                                          |
| Provider-backed formula boundary | `npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --scenario provider-backed --json`                           | `scenario: "provider-backed"`, `provider-backed-adapter-missing`, `adapterBackedDiagnosticsCleared`                    | Provider formulas such as `IMPORTRANGE` fail closed until the host supplies an adapter, then verify readback.                                                                       | Live Google Sheets authorization or remote provider availability.                                                                                          |
| Workbook Compatibility Report    | `npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- bilig-evaluate --door workbook-compatibility --json`                               | `door: "workbook-compatibility"`, `riskLevel`, `unsupportedFunctions`, `noCompatibilityScore`                          | A saved `.xlsx` can be inspected for unsupported functions, external links, macros, pivots, volatile functions, stored formula results, and risk reasons before an agent trusts it. | Excel compatibility certification, macro execution, pivot refresh, or a defensible compatibility percentage.                                               |
| Agent XLSX risk preflight        | `pnpm --dir examples/headless-workpaper run agent:mcp-xlsx-risk-preflight`                                                                        | `schemaVersion: "bilig-agent-xlsx-risk-preflight.v1"`, `analyze_workbook_risk`, `afterExpectedArr: 96000`              | A local MCP client can inspect real XLSX risk, then edit an imported WorkPaper, read a dependent formula back, persist state, and export the WorkPaper JSON.                        | Excel compatibility certification, desktop Excel UI behavior, or safe continuation when risk findings require Excel, Graph, LibreOffice, or oracle review. |
| XLSX recalculation               | `npm exec --package @bilig/xlsx-formula-recalc@latest -- xlsx-recalc --demo --json`                                                               | `recalculationCompleted: true`                                                                                         | An XLSX file boundary can be edited, recalculated, exported, and reimported for readback.                                                                                           | A full Excel clone, macro execution, charts, pivots, or desktop layout fidelity.                                                                           |
| ExcelJS recalculation            | `npx --package @bilig/exceljs-formula-recalc exceljs-recalc --demo --json`                                                                        | `commandSucceeded: true`, `recalculationCompleted: true`, `expectedValueMatched: true`                                 | An existing ExcelJS workbook can get fresh formula readback after Node edits.                                                                                                       | ExcelJS styling/export behavior, desktop Excel parity, or every Excel formula.                                                                             |
| MCP Inspector                    | `npx -y @modelcontextprotocol/inspector@latest --cli npm exec --yes --package @bilig/workpaper@latest -- bilig-workpaper-mcp --method tools/list` | tool names such as `read_workpaper_summary`, `set_workpaper_input_cell`                                                | A neutral MCP client can inspect the packaged stdio server before a user adds it to an agent host.                                                                                  | Private workbook persistence unless the file-backed config is used.                                                                                        |
| File-backed MCP server           | `npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable`         | `set_cell_contents_and_readback`, `export_workpaper_document`; `analyze_workbook_risk` when started with `--from-xlsx` | A local agent can use a persistent WorkPaper JSON file and inspect imported XLSX risk indicators before trusting the WorkPaper.                                                     | Hosted multi-user storage, secret management, or Excel compatibility certification.                                                                        |
| Vercel AI SDK `generateText()`   | `pnpm --dir examples/headless-workpaper run agent:ai-sdk-generate-text`                                                                           | `apiShape: "AI SDK generateText -> tool -> execute"`                                                                   | AI SDK tools can return before/after/restore WorkPaper proof from a `generateText()` loop.                                                                                          | Provider model quality or production prompt behavior.                                                                                                      |
| Vercel AI SDK `streamText()`     | `pnpm --dir examples/headless-workpaper run agent:ai-sdk-stream-text`                                                                             | `apiShape: "AI SDK streamText -> tool -> execute"`, `streamChunkTypes`                                                 | Streaming tool calls can carry the same WorkPaper proof while the model streams final text.                                                                                         | Browser UI streaming, telemetry retention, or non-deterministic provider output.                                                                           |
| OpenAI Responses function call   | `pnpm --dir examples/headless-workpaper run agent:openai-responses`                                                                               | `function_call_output`, `verified: true`                                                                               | OpenAI tool calling can wrap WorkPaper readback as a structured function result.                                                                                                    | Hosted remote MCP app review or ChatGPT UI behavior.                                                                                                       |
| OpenAI Agents SDK hosted MCP     | `pnpm --dir examples/headless-workpaper run agent:openai-agents-sdk-hosted-mcp`                                                                   | `MCPServerStreamableHttp`, `set_cell_contents_and_readback`                                                            | An OpenAI Agents SDK agent can call the hosted Streamable HTTP MCP endpoint.                                                                                                        | Private writable workbook state. Use local stdio for that.                                                                                                 |

## Selection Rules

Use `agent-mcp` first when the caller is an agent, MCP client, tool host, or
integration reviewer. It proves discovery, write/readback, resources, prompts,
and restart state in one command.

Use the Vercel AI SDK and OpenAI examples only after the generic evaluator
passes. Those examples verify host fit, not a stronger workbook runtime.

Use the XLSX and ExcelJS paths when a saved file or ExcelJS object is already
the contract. Do not force a WorkPaper model when the job is mostly workbook
formatting, image embedding, or file metadata.

Use the Agent XLSX risk preflight when an agent has a real `.xlsx` file and the
next action would otherwise be UI automation. It keeps the workbook local,
calls `analyze_workbook_risk` first, then requires `set_cell_contents_and_readback`
and `export_workpaper_document` before the agent reports success.

If the reviewer asks what a successful agent session looks like, send them to
the evaluator commands above and require fresh JSON output with `verified: true`.

## Limits

Bilig is not a desktop Excel replacement. Keep Excel, LibreOffice, Microsoft
Graph, or a spreadsheet-specific oracle in the loop for macros, pivots, charts,
external links, unsupported formulas, locale-specific Excel behavior, or exact
manual UI workflows.

## Related

- [Evaluate Bilig as an agent MCP workbook tool](eval-agent-mcp.md)
- [Agent XLSX risk preflight](agent-xlsx-risk-preflight.md)
- [MCP spreadsheet formula server for coding agents](mcp-spreadsheet-formula-server-for-coding-agents.md)
- [Vercel AI SDK spreadsheet tool: generateText and streamText with formula readback](vercel-ai-sdk-spreadsheet-tool-formula-readback.md)
- [ExcelJS formula result not updating after Node edits](exceljs-formula-result-not-updating-after-node-edits.md)
- [Workbook tools for agent frameworks](agent-framework-workbook-tools.md)
- [Compatibility limits](where-bilig-is-not-excel-compatible-yet.md)
