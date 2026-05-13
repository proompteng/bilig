---
title: Agent framework spreadsheet tools
published: true
description: Wrap @bilig/headless WorkPaper reads, verified edits, formula contracts, and persistence checks as AI SDK, LangChain, Mastra, LlamaIndex.TS, LangGraph.js, CopilotKit, and Cloudflare Agents tools.
tags: vercel ai sdk, langchain, mastra, llamaindex, langgraph, copilotkit, cloudflare agents, spreadsheet, node
canonical_url: https://proompteng.github.io/bilig/vercel-ai-sdk-langchain-spreadsheet-tool.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Agent Framework Spreadsheet Tools

This page is for agent builders who already have an AI SDK, LangChain,
Mastra, LlamaIndex.TS, LangGraph.js, CopilotKit, or Cloudflare Agents loop and
need a spreadsheet tool that can do more than return a screenshot.

`@bilig/headless` gives the agent a WorkPaper object: sheets, addresses,
formulas, computed readback, and JSON persistence. The framework wrapper should
stay thin. Keep the workbook behavior in ordinary Node functions, then expose
those functions through the tool shape your agent framework expects.

## Runnable Adapter Example

Run the dependency-free adapter example from a clean checkout:

```sh
git clone https://github.com/proompteng/bilig.git
cd bilig/examples/headless-workpaper
npm install
npm run agent:framework-adapters
```

The script builds the same workbook once per adapter family and exposes the
same operations in the shapes those frameworks expect:

- `readWorkPaperSummary` and `setWorkPaperInputCell` for an AI SDK-style tool
  map with `inputSchema` and `execute`
- `read_workpaper_summary` and `set_workpaper_input_cell` for a
  LangChain-style tool list with `schema` and `invoke`
- Mastra-style `createTool({ id, inputSchema, outputSchema, execute })`
- LlamaIndex.TS `tool(fn, { parameters })` / `FunctionTool` style functions
- LangGraph.js `ToolNode`-style dispatch over LangChain tools
- CopilotKit `useCopilotAction({ parameters, handler })` action objects
- Cloudflare Agents `AIChatAgent` / `streamText({ tools })` style tools

The example installs `zod` for real schemas, but it does not install the agent
frameworks. That is deliberate. It keeps the WorkPaper contract visible and
avoids hiding workbook logic behind framework setup.

## What A Passing Run Proves

The mutating tool edits `Inputs!B3` and then verifies the dependent summary
formulas:

```json
{
  "editedCell": "Inputs!B3",
  "before": {
    "expectedCustomers": 5,
    "expectedArr": 60000,
    "expansionArr": 66000,
    "targetGap": -34000
  },
  "after": {
    "expectedCustomers": 8,
    "expectedArr": 96000,
    "expansionArr": 105600,
    "targetGap": 5600
  },
  "checks": {
    "previousValue": 0.25,
    "newValue": 0.4,
    "formulasPersisted": true,
    "restoredMatchesAfter": true,
    "expectedArrChanged": true
  }
}
```

That is the useful part for agents. The tool result names the exact edited
cell, returns before and after computed values, preserves formula contracts,
serializes the workbook, restores it, and proves the restored output still
matches the post-write state.

## Adapter Boundary

Keep the adapter boring:

1. Build small SDK-neutral WorkPaper functions first.
2. Validate the sheet name and A1 address before writing.
3. Read dependent formulas before and after the edit.
4. Serialize and restore the WorkPaper document.
5. Return formula contracts and restored readback in the tool result.

The framework wrapper can then expose those functions with the local tool shape:
`inputSchema` and `execute`, `schema` and `invoke`, `parameters` and `handler`,
or a ToolNode-style dispatch wrapper. The workbook behavior should not care
which framework called it.

Official docs for the framework shapes:

- Vercel AI SDK tool calling:
  <https://ai-sdk.dev/docs/ai-sdk-core/tools-and-tool-calling>
- AI SDK `tool` reference:
  <https://ai-sdk.dev/docs/reference/ai-sdk-core/tool>
- LangChain JavaScript tools:
  <https://docs.langchain.com/oss/javascript/langchain/tools>
- Mastra `createTool()`:
  <https://mastra.ai/reference/tools/create-tool>
- LlamaIndex.TS tools:
  <https://developers.llamaindex.ai/typescript/framework/modules/agents/tool/>
- LangGraph.js `ToolNode`:
  <https://langchain-ai.github.io/langgraphjs/reference/classes/langgraph.prebuilt.ToolNode.html>
- CopilotKit `useCopilotAction`:
  <https://docs.copilotkit.ai/reference/hooks/useCopilotAction>
- Cloudflare Agents API and agent tools:
  <https://developers.cloudflare.com/agents/api-reference/agents-api/>
  <https://developers.cloudflare.com/agents/api-reference/agent-tools/>

## Files To Inspect

- adapter script:
  [`examples/headless-workpaper/agent-framework-adapters.ts`](../examples/headless-workpaper/agent-framework-adapters.ts)
- example README:
  [`examples/headless-workpaper/README.md#agent-framework-adapters`](../examples/headless-workpaper/README.md#agent-framework-adapters)
- longer tool-calling recipe:
  [`docs/agent-workpaper-tool-calling-recipe.md`](agent-workpaper-tool-calling-recipe.md)
- agent writeback verification:
  [`examples/headless-workpaper/agent-writeback-verification.ts`](../examples/headless-workpaper/agent-writeback-verification.ts)

## When This Is A Good Fit

Use this pattern when the agent needs to change a forecast, pricing model,
pipeline summary, budget check, or workbook-backed business rule and then prove
the formulas reacted. If the tool only says "I updated the spreadsheet" without
computed readback, it is not enough for production workflows.

Start with the adapter command above. If it saves you an agent-tooling spike,
star the repository so the next person searching for spreadsheet tools can
find it:
<https://github.com/proompteng/bilig/stargazers>.
