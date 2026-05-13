---
title: Mastra WorkPaper spreadsheet tool
published: true
description: Use @bilig/headless as the workbook logic behind a Mastra createTool: read a range, write one input, and return formula readback.
tags: mastra, createTool, spreadsheet, workpaper, typescript
canonical_url: https://proompteng.github.io/bilig/mastra-workpaper-spreadsheet-tool.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Mastra WorkPaper Spreadsheet Tool

If a Mastra agent needs spreadsheet math, keep the workbook code in ordinary
TypeScript. Mastra should get small tool wrappers: one tool reads a summary
range, and one tool writes a validated input cell and returns the formula
readback.

That keeps the agent boundary boring. `@bilig/headless` owns formulas,
serialization, and restore checks; `createTool` owns the schema and the tool
name the model sees.

## Run the checked adapter

```sh
git clone https://github.com/proompteng/bilig.git
cd bilig/examples/headless-workpaper
npm install
npm run agent:framework-adapters
```

The Mastra lane returns tool IDs and a verified write result:

```json
{
  "toolIds": ["read-workpaper-summary", "set-workpaper-input-cell"],
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

## Mastra shape

The example mirrors the `createTool({ id, description, inputSchema,
outputSchema, execute })` shape from the Mastra docs:

```ts
export const setWorkPaperInputCell = createTool({
  id: 'set-workpaper-input-cell',
  description: 'Set one WorkPaper input cell and return formula readback.',
  inputSchema: setInputCellInputSchema,
  outputSchema: workPaperWriteOutputSchema,
  execute: async ({ context }) => setWorkPaperInputCellInWorkbook(context),
})
```

Use a narrow input schema. For the demo, the write tool accepts only the
`Inputs` sheet and an A1-style address. That keeps an agent from treating the
workbook like an arbitrary mutation surface.

## What to copy

- Keep `@bilig/headless` WorkPaper construction in your application code.
- Validate tool arguments before writing.
- Return before/after formula readback, not just an "updated" message.
- Serialize and restore the WorkPaper document inside the tool result when the
  workflow depends on persistence.

Official Mastra reference: <https://mastra.ai/reference/tools/create-tool>.

Runnable source:
[`examples/headless-workpaper/agent-framework-adapters.ts`](../examples/headless-workpaper/agent-framework-adapters.ts).
