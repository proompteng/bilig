---
title: OpenAI Responses WorkPaper tool calls
published: true
description: Run @bilig/workpaper behind OpenAI Responses function calls, return function_call_output items, and verify formula readback after a workbook edit.
tags: openai responses, function calling, tool calling, spreadsheet, node, typescript
canonical_url: https://proompteng.github.io/bilig/openai-responses-workpaper-tool-call.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# OpenAI Responses WorkPaper Tool Calls

Use this when an OpenAI Responses agent needs to change workbook inputs and
then explain the number the workbook calculated.

The model should not write workbook JSON. Give it two small function tools:
read a bounded range, and set one validated input cell. Your Node process runs
those tools against `@bilig/workpaper`, returns `function_call_output` items, and
asks the model to answer from the computed readback.

## Runnable TypeScript Example

Run the dependency-light example from a checkout:

```sh
git clone https://github.com/proompteng/bilig.git
cd bilig
pnpm --dir examples/headless-workpaper install --ignore-workspace
pnpm --dir examples/headless-workpaper run agent:openai-responses
```

The script does not call the OpenAI API. It gives you the application-side code
you run between Responses API turns:

1. define `read_workpaper_summary` and `set_workpaper_input_cell` function
   tools.
2. receive `function_call` items from the model.
3. parse and validate arguments with `zod`.
4. edit the WorkPaper.
5. return matching `function_call_output` items.

Expected proof:

```json
{
  "apiShape": "OpenAI Responses function_call -> function_call_output",
  "toolNames": ["read_workpaper_summary", "set_workpaper_input_cell"],
  "followupInputTypes": ["user", "function_call", "function_call", "function_call_output", "function_call_output"],
  "writeResult": {
    "editedCell": "Inputs!B3",
    "before": {
      "expectedArr": 60000,
      "targetGap": -34000
    },
    "after": {
      "expectedArr": 96000,
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
}
```

The full output also includes the exact model-style function calls, the
serialized `function_call_output` strings, formula contracts, restored summary,
and a deterministic final message built from tool output.

## The OpenAI Boundary

The official Responses function-calling flow is a loop: send tools, receive
`function_call` items, run your code, append `function_call_output` items, then
send the updated input back to the model. The WorkPaper part is the dispatcher:

```ts
function dispatchOpenAiResponsesCall(call: OpenAiResponsesFunctionCall) {
  if (call.name === 'read_workpaper_summary') {
    const args = readSummaryInputSchema.parse(JSON.parse(call.arguments))
    return tools.readWorkPaperSummary(args.range)
  }

  if (call.name === 'set_workpaper_input_cell') {
    const args = setInputCellInputSchema.parse(JSON.parse(call.arguments))
    return tools.setWorkPaperInputCell(args)
  }

  throw new Error(`unknown WorkPaper tool: ${call.name}`)
}
```

Return JSON from the tool, not prose. The next model turn can then say:
`Edited Inputs!B3. Expected ARR moved from 60000 to 96000.` That sentence is
grounded in formula readback, not in a guess.

OpenAI's current function-calling guide covers the Responses API item types and
the `function_call_output` handoff:
<https://platform.openai.com/docs/guides/function-calling?api-mode=responses>.

## Why This Shape Works

- the model chooses tools, but WorkPaper owns cells, formulas, and persistence.
- the write tool validates the sheet and A1 address before mutation.
- the result includes before and after computed values.
- formulas are serialized, restored, and compared before the tool result is
  accepted.
- the final model answer can cite the edited cell and computed values directly.

That is the useful contract for workbook automation. A response that says
"updated" is not enough unless the tool result proves what changed.

## Files To Inspect

- runnable OpenAI Responses dispatcher:
  [`examples/headless-workpaper/openai-responses-tool-wrapper.ts`](../examples/headless-workpaper/openai-responses-tool-wrapper.ts)
- example README section:
  [`examples/headless-workpaper/README.md#openai-responses-tool-wrapper`](../examples/headless-workpaper/README.md#openai-responses-tool-wrapper)
- broader tool-calling recipe:
  [`docs/agent-workpaper-tool-calling-recipe.md`](agent-workpaper-tool-calling-recipe.md)
- framework adapters:
  [`examples/headless-workpaper/agent-framework-adapters.ts`](../examples/headless-workpaper/agent-framework-adapters.ts)

Repository and release notes:
<https://github.com/proompteng/bilig>.
