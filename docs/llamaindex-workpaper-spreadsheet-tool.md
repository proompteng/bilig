---
title: LlamaIndex.TS WorkPaper spreadsheet tool
published: true
description: Add @bilig/headless to a LlamaIndex.TS agent as narrow workbook tools with formula readback and persistence checks.
tags: llamaindex, llamaindex ts, spreadsheet, workpaper, typescript
canonical_url: https://proompteng.github.io/bilig/llamaindex-workpaper-spreadsheet-tool.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# LlamaIndex.TS WorkPaper Spreadsheet Tool

Use a LlamaIndex.TS tool when an agent should change workbook assumptions but
not freehand-edit a file. The useful shape is small: read a summary range,
write one allowed input, and return the cells and formula values that changed.

The LlamaIndex.TS `tool(fn, { parameters })` shape takes a function plus a
configuration object with `name`, `description`, and `parameters`. The
WorkPaper adapter keeps the same pattern: Zod validates the arguments, and
`@bilig/headless` does the spreadsheet work.

## Run the checked adapter

```sh
git clone https://github.com/proompteng/bilig.git
cd bilig/examples/headless-workpaper
npm install
npm run agent:framework-adapters
```

The LlamaIndex.TS lane proves the same WorkPaper functions are exposed as
tool-style calls:

```json
{
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

## LlamaIndex.TS shape

```ts
const setInputTool = tool(setWorkPaperInputCell, {
  name: 'set_workpaper_input_cell',
  description: 'Set one validated WorkPaper input and return formula readback.',
  parameters: setInputCellInputSchema,
})
```

The important boundary is the function behind the tool. It should validate the
sheet and A1 address, apply one write, read dependent formulas before and after
the write, serialize the WorkPaper document, restore it, and return the
verification result.

## What to copy

- Use Zod schemas for agent-generated arguments.
- Keep workbook state in your app or workflow context.
- Return exact cells and computed values so the agent can decide the next step.
- Prefer a small `set_workpaper_input_cell` tool over a broad "edit workbook"
  tool.

Official LlamaIndex.TS tools docs:
<https://developers.llamaindex.ai/typescript/framework/modules/agents/tool/>.

Runnable source:
[`examples/headless-workpaper/agent-framework-adapters.ts`](../examples/headless-workpaper/agent-framework-adapters.ts).
