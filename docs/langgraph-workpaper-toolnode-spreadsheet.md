---
title: LangGraph.js WorkPaper ToolNode spreadsheet tool
published: true
description: Route @bilig/headless read and write tools through a LangGraph.js ToolNode-style workflow with formula readback.
tags: langgraph, toolnode, langchain, spreadsheet, workpaper
canonical_url: https://proompteng.github.io/bilig/langgraph-workpaper-toolnode-spreadsheet.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# LangGraph.js WorkPaper ToolNode Spreadsheet Tool

LangGraph.js workflows often route model tool calls through a `ToolNode`. That
is a good place for WorkPaper tools when the graph needs a number it can trust:
read a summary range, write one input, then read the dependent formulas again.

The checked example keeps the WorkPaper functions framework-neutral. It exposes
them as LangChain-style tools, then dispatches those calls through a small
ToolNode-style wrapper.

## Run the checked adapter

```sh
git clone https://github.com/proompteng/bilig.git
cd bilig/examples/headless-workpaper
npm install
npm run agent:framework-adapters
```

The LangGraph lane returns the node name, tool names, and verified write:

```json
{
  "nodeName": "tools",
  "toolNames": ["read_workpaper_summary", "set_workpaper_input_cell"],
  "writeResult": {
    "editedCell": "Inputs!B3",
    "checks": {
      "formulasPersisted": true,
      "restoredMatchesAfter": true,
      "expectedArrChanged": true
    }
  }
}
```

## ToolNode shape

```ts
const tools = createLangChainTools(workPaperTools)
const toolNode = new ToolNode(tools)
```

The repository example does not import LangGraph. It shows the same dispatch
shape without adding the framework dependency to the standalone smoke test.
In an app that already uses LangGraph.js, keep the WorkPaper read/write
functions the same and pass the actual tool objects into `ToolNode`.

## What to copy

- Use separate read and write tools so graph state stays easy to inspect.
- Return exact tool messages with the edited cell and formula readback.
- Keep persistence verification in the tool result when the graph will resume
  later.
- Keep the compatibility caveat visible: this is a WorkPaper API, not full Excel
  UI automation.

Official LangGraph.js `ToolNode` reference:
<https://langchain-ai.github.io/langgraphjs/reference/classes/langgraph.prebuilt.ToolNode.html>.

Runnable source:
[`examples/headless-workpaper/agent-framework-adapters.ts`](../examples/headless-workpaper/agent-framework-adapters.ts).
