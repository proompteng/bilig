---
title: Agent spreadsheet tool-call loop
published: true
description: A runnable @bilig/headless loop where an agent writes one workbook input, checks formula outputs, and persists the verified result.
tags: ai agents, tool calling, node, spreadsheet, workpaper
canonical_url: https://proompteng.github.io/bilig/agent-spreadsheet-tool-call-loop.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Agent Spreadsheet Tool Call Loop

This page is for people building coding agents or backend workers that need a
spreadsheet engine but do not want to drive a spreadsheet UI.

`@bilig/headless` lets the agent expose workbook operations as ordinary Node
tools: read a range, set a cell, recalculate formulas, verify the result, and
serialize the workbook for persistence. The important part is the loop, not the
agent framework around it.

## Runnable Example

Run the maintained example from a clean checkout:

```sh
cd examples/headless-workpaper
npm install
npm run agent:tool-call
```

The example builds a workbook with `Inputs` and `Summary` sheets, then executes
one agent-style tool call:

```json
{
  "toolName": "setInputCell",
  "arguments": {
    "sheetName": "Inputs",
    "address": "B3",
    "value": 0.4,
    "reason": "Use the latest qualified pipeline conversion estimate."
  }
}
```

The tool result reports the edited cell, before and after formula outputs,
restored workbook outputs, and persistence checks. A passing run proves that:

- the write targeted `Inputs!B3`
- `Expected ARR` moved from `60000` to `96000`
- `Target gap` moved from `-34000` to `5600`
- formulas persisted after serialize and restore
- restored values matched the post-write values

That is the contract an agent needs. It should not have to infer formula
success from a screenshot.

## Tool Shape

Keep the first tool surface small:

- `readRange(range)` returns computed values and serialized cell contents.
- `setInputCell({ sheetName, address, value, reason })` validates the sheet and
  A1 address before writing.
- every mutating tool returns before and after computed values.
- persistence happens only after the verification readback succeeds.

This shape works with OpenAI tool calls, local coding-agent tools, queue
workers, and normal service endpoints because the WorkPaper API is the boundary.
The agent SDK can change without rewriting the workbook logic.

## Files To Inspect

- runnable script:
  [`examples/headless-workpaper/agent-tool-call-loop.mjs`](../examples/headless-workpaper/agent-tool-call-loop.mjs)
- external smoke check:
  [`scripts/workpaper-external-smoke.ts`](../scripts/workpaper-external-smoke.ts)
- package contract:
  [`packages/headless/README.md`](../packages/headless/README.md)
- longer recipe:
  [`docs/agent-workpaper-tool-calling-recipe.md`](agent-workpaper-tool-calling-recipe.md)

## When This Is A Good Fit

Use this pattern when a product asks an agent to edit a forecast, refresh a
scenario, check a formula-backed model, or generate a workbook artifact from
service data. Keep the model responsible for choosing the next action, and keep
the WorkPaper tool responsible for addresses, formulas, recalculation, and
persistence proof.

If the workflow needs HTTP boundaries instead of direct function calls, start
with the [Node service recipe](node-service-workpaper-recipe.md). If the input
is JSON records rather than arrays, use the
[`json-records-input.mjs`](../examples/headless-workpaper/json-records-input.mjs)
example.
