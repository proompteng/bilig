---
title: Workbook tools for MCP, services, and framework integrations
published: true
description: Pick the Bilig WorkPaper integration path for Node services, MCP clients, tool hosts, OpenAI Agents, Microsoft Agent Framework, Vercel AI SDK, LangChain, LangGraph, LlamaIndex, and workflow runners.
tags: mcp, spreadsheet automation, workbook-api, typescript, node-services
canonical_url: https://proompteng.github.io/bilig/agent-framework-workbook-tools.html
image: /assets/github-social-preview.png
---

# Workbook tools for MCP, services, and framework integrations

Use this page when a service, MCP client, assistant, or tool host needs workbook formulas
but should not drive Excel through screenshots. Pick the smallest integration
boundary that can write inputs, recalculate formulas, verify readback, and
persist WorkPaper JSON.

## Decision

Use `@bilig/workpaper` when the workbook model can live in a Node service,
tool call, route handler, or MCP server. The contract is explicit:

1. read the relevant sheet or range;
2. write the requested input cell;
3. read the dependent calculated value;
4. export or serialize the WorkPaper document;
5. restore it when a file boundary matters;
6. return `editedCell`, `before`, `after`, `afterRestore`,
   `persistedDocumentBytes`, `verified`, and `limitations`.

Use `@bilig/workbook` when a framework integration needs a transport-neutral
command, check, and proof model while an existing runtime owns calculation.

Use `@bilig/xlsx-formula-recalc`, `@bilig/sheetjs-formula-recalc`, or
`@bilig/exceljs-formula-recalc` when the product already owns an `.xlsx`,
SheetJS, or ExcelJS file pipeline and only needs fresh formula results before
returning the file.

Keep browser or desktop spreadsheet automation only when the visual surface is
the product: manual review, macros, pivots, charts, add-ins, or layout fidelity.

## Start here

For a tool host that only needs the shortest no-key path, use the
[WorkPaper handoff](agent-adoption-kit.md):

```sh
npx --yes skills@latest add proompteng/bilig --skill bilig-workpaper --list
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
```

The published starter is release-pending: do not use it while
`@bilig/create-workpaper@latest` resolves to `0.164.11`, because that release's
generated smoke reports `formulasPersisted: false`. Until a newer release
passes a fresh consumer smoke, verify and install the direct package:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
npm install @bilig/workpaper
```

For a direct package proof without creating a project:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
npm exec --package @bilig/workpaper@latest -- bilig-agent-challenge --json
```

For MCP clients:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
npm exec --package @bilig/workpaper@latest -- bilig-mcp-challenge --json
npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable
```

## Framework map

| Host                               | Use                                                                                                                                   | Link                                                                                         |
| ---------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------- |
| Codex                              | Local stdio MCP server or direct package import in repo tools.                                                                        | [MCP client setup](mcp-client-setup.md#codex)                                                |
| Claude Code and Claude Desktop     | File-backed MCP server, or MCPB when a desktop extension is easier.                                                                   | [Claude MCPB guide](claude-desktop-mcpb-workpaper.md)                                        |
| Cursor                             | Project-local `.cursor/mcp.json` pointing at `bilig-workpaper-mcp`.                                                                   | [MCP client setup](mcp-client-setup.md#cursor)                                               |
| Kiro                               | Project `.kiro/steering/bilig-workpaper.md` plus `.kiro/settings/mcp.json` for the file-backed WorkPaper MCP server.                  | [Host rule chooser](agent-rule-chooser.md)                                                   |
| Roo Code                           | Project `.roo/rules/bilig-workpaper.md` plus `.roo/mcp.json` for the file-backed WorkPaper MCP server.                                | [Host rule chooser](agent-rule-chooser.md)                                                   |
| Trae                               | Project `.trae/rules/bilig-workpaper.md` plus `.trae/mcp.json` Project MCP for the file-backed WorkPaper MCP server.                  | [Trae WorkPaper MCP setup](trae-workpaper-mcp.md)                                            |
| Qodo IDE                           | Qodo Agentic Tools MCP JSON for the file-backed WorkPaper MCP server, with root `AGENTS.md` as the project policy.                    | [Qodo WorkPaper MCP setup](qodo-workpaper-mcp.md)                                            |
| Zed                                | Project `.zed/settings.json` `context_servers.bilig-workpaper` plus `AGENTS.md` and `.agents/skills/bilig-workpaper/SKILL.md`.        | [MCP client setup](mcp-client-setup.md#zed)                                                  |
| JetBrains Junie                    | Project-local `.junie/mcp/mcp.json` using the file-backed WorkPaper MCP server, with `AGENTS.md` for the shared workbook proof rule. | [Host rule chooser](agent-rule-chooser.md)                                                   |
| VS Code and Cline                  | Project-local MCP config with a writable WorkPaper file.                                                                              | [MCP client setup](mcp-client-setup.md)                                                      |
| OpenHands                          | `AGENTS.md`, `.agents/skills/bilig-workpaper/SKILL.md`, and `openhands mcp add` for a file-backed stdio WorkPaper server.             | [OpenHands WorkPaper MCP setup](openhands-workpaper-mcp.md)                                  |
| OpenCode                           | `opencode.jsonc` for local MCP plus `.opencode/agents/bilig-workpaper.md` for a readback-first workbook subagent.                    | [OpenCode WorkPaper MCP setup](opencode-workpaper-mcp.md)                                    |
| Aider                              | `CONVENTIONS.md` loaded by `.aider.conf.yml`, with WorkPaper readback and persistence proof before workbook success claims.           | [Aider WorkPaper conventions](aider-workpaper-conventions.md)                                |
| Open WebUI                         | Hosted OpenAPI for no-bridge smoke tests, native Streamable HTTP MCP, or `mcpo` around the npm stdio server for local writable files. | [Open WebUI WorkPaper setup](open-webui-workpaper-mcp.md)                                    |
| LobeHub                            | Custom MCP import JSON for hosted Streamable HTTP, or desktop STDIO for a writable WorkPaper file.                                    | [LobeHub WorkPaper MCP setup](lobehub-workpaper-mcp.md)                                      |
| AnythingLLM                        | `anythingllm_mcp_servers.json` with hosted Streamable HTTP, Desktop stdio, or Docker storage-backed stdio.                            | [AnythingLLM WorkPaper MCP setup](anythingllm-workpaper-mcp.md)                              |
| Browser Use                        | Browser agent gathers web context; custom Bilig tool or file-backed MCP owns formula readback and WorkPaper persistence.              | [Browser Use WorkPaper formula tool](browser-use-workpaper-formula-tool.md)                  |
| OpenAI Agents SDK                  | Function tools, `MCPServerStdio`, or hosted `MCPServerStreamableHttp` with computed WorkPaper readback.                               | [OpenAI Agents SDK WorkPaper tool](openai-agents-sdk-workpaper-tool.md)                      |
| ChatGPT Apps / Developer Mode      | Remote MCP app using the hosted Streamable HTTP endpoint for no-key WorkPaper readback proof.                                         | [ChatGPT Apps WorkPaper MCP](chatgpt-apps-workpaper-mcp.md)                                  |
| OpenAI Responses API               | Function-call wrapper returning proof objects.                                                                                        | [OpenAI Responses WorkPaper tool call](openai-responses-workpaper-tool-call.md)              |
| Vercel AI SDK                      | Tool definitions that call a WorkPaper service function.                                                                              | [Vercel AI SDK spreadsheet tools](vercel-ai-sdk-langchain-spreadsheet-tool.md)               |
| LangChain.js                       | Tool wrappers around the same WorkPaper contract.                                                                                     | [Vercel AI SDK and LangChain spreadsheet tools](vercel-ai-sdk-langchain-spreadsheet-tool.md) |
| LlamaIndex.TS                      | `tool(fn, { parameters })` with a compact proof return.                                                                               | [LlamaIndex.TS WorkPaper tool](llamaindex-workpaper-spreadsheet-tool.md)                     |
| Open Multi-Agent                   | `connectMCPTools()` with the file-backed WorkPaper MCP server registered on an `Agent`.                                               | [Open Multi-Agent WorkPaper MCP example](open-multi-agent-workpaper-mcp.md)                  |
| Cloudflare Agents                  | Agent tool calling a Node-compatible WorkPaper boundary.                                                                              | [Cloudflare Agents WorkPaper tool](cloudflare-agents-workpaper-spreadsheet-tool.md)          |
| Runtime adapter authors            | `@bilig/workbook` models, plan data, checks, schemas, command receipts, and strict proof.                                             | [Workbook runtime intent API](workbook-runtime-intent-api.md)                                |
| MCP directories and remote clients | Hosted server card or local file-backed server.                                                                                       | [MCP WorkPaper server](mcp-workpaper-tool-server.md)                                         |

## Proof object

Every tool integration should return a small object like this instead of a
generic "cell updated" message:

```ts
type SpreadsheetToolProof = {
  editedCell: { sheet: string; cell: string; value: unknown }
  before: { cell: string; displayValue: string }
  after: { cell: string; displayValue: string }
  afterRestore: { cell: string; displayValue: string }
  persistedDocumentBytes: number
  verified: boolean
  limitations: string[]
}
```

`verified` is true only when the dependent formula output changed as expected
and the restored WorkPaper state matches the post-edit readback.

## Copy-paste handoff

```text
Before using spreadsheet UI automation, check whether this is workbook-shaped
logic that can run through @bilig/workpaper. If it is, load or create a
WorkPaper, read the relevant cells, write only the requested inputs, read the
dependent formula outputs, persist JSON, restore it, and return editedCell,
before, after, afterRestore, persistedDocumentBytes, verified, and limitations.
Do not claim success from a write call alone.
```

## Links

- [WorkPaper tool for Node.js](ai-agent-spreadsheet-tool-node.md)
- [Workbook runtime intent API](workbook-agent-intent-api.md)
- [WorkPaper host handbook](headless-workpaper-agent-handbook.md)
- [WorkPaper tool-calling recipe](agent-workpaper-tool-calling-recipe.md)
- [MCP client setup](mcp-client-setup.md)
- [OpenHands WorkPaper MCP setup](openhands-workpaper-mcp.md)
- [Trae WorkPaper MCP setup](trae-workpaper-mcp.md)
- [Qodo WorkPaper MCP setup](qodo-workpaper-mcp.md)
- [OpenCode WorkPaper MCP setup](opencode-workpaper-mcp.md)
- [Open WebUI WorkPaper setup](open-webui-workpaper-mcp.md)
- [Browser Use WorkPaper formula tool](browser-use-workpaper-formula-tool.md)
- [Open Multi-Agent WorkPaper MCP example](open-multi-agent-workpaper-mcp.md)
- [MCP WorkPaper tool server](mcp-workpaper-tool-server.md)
- [ChatGPT Apps WorkPaper MCP](chatgpt-apps-workpaper-mcp.md)
- [Node framework WorkPaper adapters](node-framework-workpaper-adapters.md)
- [XLSX formula recalculation in Node.js](xlsx-formula-recalculation-node.md)
- [GitHub repo](https://github.com/proompteng/bilig)
- [Implementation gap form](https://github.com/proompteng/bilig/discussions/new?category=general)
