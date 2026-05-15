---
title: CrewAI WorkPaper spreadsheet tool
published: true
description: Expose @bilig/headless WorkPaper calculations to CrewAI as a small JSON tool contract with formula readback.
tags: crewai, spreadsheet, workpaper, agent tools, typescript
canonical_url: https://proompteng.github.io/bilig/crewai-workpaper-spreadsheet-tool.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# CrewAI WorkPaper Spreadsheet Tool

CrewAI workflows can call a WorkPaper-backed TypeScript service when an agent
needs spreadsheet math, formula readback, or workbook persistence. Keep the
CrewAI side as the orchestration layer; keep workbook construction, validation,
formula calculation, and serialization in `@bilig/headless`.

This is an interop recipe, not an official CrewAI adapter. The useful boundary
is a small JSON contract:

- input payload: `sheetName`, `address`, and `value`
- formula readback: before/after computed `Summary` values
- error shape: `{ ok: false, error: string }`

## Run the checked adapter

```sh
git clone https://github.com/proompteng/bilig.git
cd bilig/examples/headless-workpaper
npm install
npm run agent:framework-adapters
```

The CrewAI lane returns plain JSON tool metadata plus a verified WorkPaper write
result:

```json
{
  "toolNames": ["read_workpaper_summary", "set_workpaper_input_cell"],
  "contract": {
    "inputPayload": "validated JSON args",
    "formulaReadback": "before/after computed Summary values",
    "errorShape": "{ ok: false, error: string }"
  },
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

## TypeScript service shape

Expose narrow WorkPaper functions from a Node service and let CrewAI call them
over HTTP, a queue, or any other app-owned transport:

```ts
import { z } from 'zod'

const setInputCellInputSchema = z.object({
  sheetName: z.literal('Inputs'),
  address: z.string().regex(/^[A-Z]+[1-9][0-9]*$/),
  value: z.union([z.string(), z.number(), z.boolean(), z.null()]),
})

export function runCrewAiWorkPaperTool(payload: unknown) {
  const args = setInputCellInputSchema.safeParse(payload)
  if (!args.success) {
    return {
      ok: false,
      error: args.error.issues.map((issue) => issue.message).join('; '),
    }
  }

  const result = setWorkPaperInputCell(args.data)
  return {
    ok: true,
    result,
  }
}
```

The WorkPaper function behind `setWorkPaperInputCell` should build or load the
workbook, write one validated input, read dependent formulas before and after
the edit, and return a plain JSON result. The agent should receive evidence,
not just an "updated" string.

## What to copy

- Validate agent-generated JSON before writing to the workbook.
- Return the edited cell, before/after formula values, and persistence checks.
- Keep the tool contract small enough for a CrewAI task to reason about.
- Do not require CrewAI for normal `bilig` usage; the same WorkPaper functions
  also work from Node services, queues, tests, and other agent frameworks.

Runnable source:
[`examples/headless-workpaper/agent-framework-adapters.ts`](../examples/headless-workpaper/agent-framework-adapters.ts).
