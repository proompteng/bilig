---
title: MCP spreadsheet tool server for WorkPaper agents
published: true
description: Expose @bilig/headless workbook reads, verified edits, formula contracts, and persistence checks through MCP-style tools/list and tools/call handlers.
tags: mcp, model context protocol, spreadsheet, tool calling, node
canonical_url: https://proompteng.github.io/bilig/mcp-workpaper-tool-server.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# MCP Spreadsheet Tool Server For WorkPaper Agents

This page is for agent builders who want workbook formulas behind a Model
Context Protocol tool surface. The useful boundary is small: list the tools,
call one tool, return exact cell readback, and include enough structured output
for the agent to verify the edit.

`@bilig/headless` owns the workbook behavior. MCP should stay as the transport
and discovery layer around ordinary Node functions.

## Runnable MCP-Style Example

Run the dependency-free example from a clean checkout:

```sh
git clone https://github.com/proompteng/bilig.git
cd bilig/examples/headless-workpaper
npm install
npm run agent:mcp-tools
```

For a local stdio transport, pipe newline-delimited JSON-RPC requests into the
stdio entrypoint:

```sh
printf '%s\n' \
  '{"jsonrpc":"2.0","id":1,"method":"initialize"}' \
  '{"jsonrpc":"2.0","method":"notifications/initialized"}' \
  '{"jsonrpc":"2.0","id":2,"method":"tools/list"}' |
  npm run --silent agent:mcp-stdio
```

The script implements two JSON-RPC methods shaped around the MCP tool model:

- `tools/list` returns `read_workpaper_summary` and
  `set_workpaper_input_cell` with JSON Schema inputs.
- `tools/call` invokes the requested WorkPaper tool and returns text content
  plus structured formula readback.

The example deliberately avoids an MCP SDK dependency so the workbook contract
is visible. Put the same handlers behind stdio, HTTP, or your MCP SDK adapter
when you wire it into a production agent host.

## What A Passing Run Proves

The write tool edits `Inputs!B3`, recalculates dependent formulas, serializes
the WorkPaper document, restores it, and checks that formulas and computed
values survived the round trip:

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

That is the part spreadsheet agents need. A tool that only says "updated" is
not enough. Return the edited address, previous value, new value, before/after
computed values, formula contracts, and persistence proof.

## Tool Boundary

Expose only the minimum useful surface first:

1. `read_workpaper_summary` reads a bounded range and returns computed values
   plus serialized cell contents.
2. `set_workpaper_input_cell` validates the sheet and A1 address before a
   write, then returns formula readback and persistence checks.
3. Everything outside that boundary stays in your MCP host: auth, transport,
   rate limits, logging, and user approval policy.

The official MCP specification describes tool discovery through `tools/list`
and tool invocation through `tools/call`, with input schemas on each tool:
<https://modelcontextprotocol.io/specification/2025-06-18/server/tools>.

## Files To Inspect

- MCP-style adapter script:
  [`examples/headless-workpaper/mcp-tool-server.mjs`](https://github.com/proompteng/bilig/blob/main/examples/headless-workpaper/mcp-tool-server.mjs)
- stdio adapter script:
  [`examples/headless-workpaper/mcp-stdio-server.mjs`](https://github.com/proompteng/bilig/blob/main/examples/headless-workpaper/mcp-stdio-server.mjs)
- example README:
  [`examples/headless-workpaper/README.md#mcp-tool-server-shape`](https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#mcp-tool-server-shape)
- SDK-neutral tool-calling recipe:
  [`docs/agent-workpaper-tool-calling-recipe.md`](agent-workpaper-tool-calling-recipe.md)
- Vercel AI SDK and LangChain wrappers:
  [`docs/vercel-ai-sdk-langchain-spreadsheet-tool.md`](vercel-ai-sdk-langchain-spreadsheet-tool.md)

## Feedback Thread

Use the
[MCP spreadsheet tool server discussion](https://github.com/proompteng/bilig/discussions/230)
for adapter feedback. The open questions are deliberately concrete: stdio,
HTTP/SSE, or SDK adapter next; which spreadsheet workflow should be proven
next; and which structured fields every write tool should return.

## When This Is A Good Fit

Use this pattern when an agent needs to edit a forecast, pricing workbook,
quote approval rule, budget check, or service-side spreadsheet model and prove
the formulas reacted. Keep the MCP layer thin, keep the workbook logic
testable, and make every write return structured verification.

Start with the adapter command above. If it saves you a spreadsheet-tooling
spike, star the repository so the next person searching for MCP spreadsheet
tools can find it:
<https://github.com/proompteng/bilig/stargazers>.
