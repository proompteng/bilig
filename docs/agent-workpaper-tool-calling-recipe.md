---
title: WorkPaper tool-calling recipe for AI agents
published: true
description: Wrap @bilig/headless workbook reads, writes, formula readback, and persistence as deterministic Node.js tools for coding agents.
tags: ai agents, tool calling, node, spreadsheet, typescript
canonical_url: https://proompteng.github.io/bilig/agent-workpaper-tool-calling-recipe.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# WorkPaper Tool-Calling Recipe For Agents

This recipe shows how to wrap `@bilig/headless` WorkPaper operations as
agent-callable functions without binding the workflow to one agent SDK.

Use this pattern when an agent needs to inspect, edit, verify, and persist a
formula-backed workbook from Node. Do not screen scrape a spreadsheet UI when
the WorkPaper API is available. Screenshots are useful for final human review,
but they hide formulas, typed addresses, recalculation state, and persistence
contracts.

Start with the package README for the public API contract:
[`packages/headless/README.md`](../packages/headless/README.md).

For a runnable external example, use
[`examples/headless-workpaper`](../examples/headless-workpaper) and run
`npm run agent:tool-call`. If your app calls OpenAI Responses directly, run
`npm run agent:openai-responses` and read the
[OpenAI Responses WorkPaper tool-call guide](openai-responses-workpaper-tool-call.md).
For a smaller writeback-only proof, run
`npm run agent:verify`. For framework-shaped wrappers that do not pull Vercel
AI SDK or LangChain into this repository, run
`npm run agent:framework-adapters`. For a CrewAI interop shape, use the
[CrewAI WorkPaper spreadsheet tool](crewai-workpaper-spreadsheet-tool.md)
recipe; it keeps the WorkPaper code in TypeScript and exposes a small JSON
contract to the agent workflow.
If you want the real AI SDK loop, run `npm run agent:ai-sdk-generate-text`.
That script calls `generateText()` and `tool()` from `ai`, using `ai/test` as a
deterministic provider so no API key is needed.
For the streaming path, run `npm run agent:ai-sdk-stream-text`. That script
calls `streamText()` from `ai`, streams tool-call chunks and final text, and
keeps the WorkPaper read/write verification in ordinary TypeScript.

If your app calls OpenAI directly, start with the
[Responses API function-calling guide](https://developers.openai.com/api/docs/guides/function-calling)
and keep the WorkPaper functions below as your application-side tool handlers.
If this is the path you are trying, use the
[OpenAI Responses tool-call discussion](https://github.com/proompteng/bilig/discussions/335)
to say what readback or streaming transcript shape would make the example more
useful.

## Tool Contract

Expose a small, boring tool surface first:

- `readSummary(range)` returns computed values and serialized inputs for a
  summary range.
- `setInputCell(sheetName, address, value)` validates the target sheet and A1
  address, writes one value, and returns before/after computed verification.
- `serializeWorkbook()` exports a persisted WorkPaper document only after the
  edit succeeds.

Keep each tool deterministic. Let the agent choose the next action, but make the
tool result carry enough evidence for verification.

## Complete Node Example

```ts
import { WorkPaper, exportWorkPaperDocument, serializeWorkPaperDocument, type WorkPaperCellAddress } from '@bilig/headless'

type CellInputValue = string | number | boolean | null

type SummaryReadback = {
  currentMrr: number
  nextMonthMrr: number
}

type SetInputCellArgs = {
  sheetName: string
  address: string
  value: CellInputValue
}

const workbook = WorkPaper.buildFromSheets({
  Assumptions: [
    ['Metric', 'Value'],
    ['Growth rate', 0.1],
  ],
  Revenue: [
    ['Segment', 'Customers', 'ARPA', 'MRR'],
    ['Self serve', 200, 30, '=B2*C2'],
    ['Sales', 15, 300, '=B3*C3'],
  ],
  Summary: [
    ['Metric', 'Value'],
    ['Current MRR', '=SUM(Revenue!D2:D3)'],
    ['Next month MRR', '=B2*(1+Assumptions!B2)'],
  ],
})

const summarySheet = requireSheet('Summary')
const currentMrrAddress = requireCellAddress('Summary', 'B2')
const nextMonthMrrAddress = requireCellAddress('Summary', 'B3')

const tools = {
  readSummary(range: string = 'Summary!A1:B3') {
    const parsedRange = workbook.simpleCellRangeFromString(range, summarySheet)
    if (parsedRange === undefined) {
      throw new Error(`invalid summary range: ${range}`)
    }

    return {
      range,
      values: workbook.getRangeValues(parsedRange),
      serialized: workbook.getRangeSerialized(parsedRange),
    }
  },

  setInputCell({ sheetName, address, value }: SetInputCellArgs) {
    const target = requireCellAddress(sheetName, address)
    const before = readComputedSummary()

    workbook.setCellContents(target, value)

    const after = readComputedSummary()
    const serializedWorkbook = serializeWorkbook()

    return {
      editedCell: workbook.simpleCellAddressToString(target, {
        includeSheetName: true,
      }),
      before,
      after,
      checks: {
        currentMrrChanged: before.currentMrr !== after.currentMrr,
        nextMonthMrrChanged: before.nextMonthMrr !== after.nextMonthMrr,
        serializedBytes: Buffer.byteLength(serializedWorkbook, 'utf8'),
      },
    }
  },

  serializeWorkbook,
}

console.log(tools.readSummary())
console.log(
  tools.setInputCell({
    sheetName: 'Revenue',
    address: 'B3',
    value: 25,
  }),
)

function requireSheet(sheetName: string): number {
  const sheetId = workbook.getSheetId(sheetName)
  if (sheetId === undefined) {
    throw new Error(`unknown sheet: ${sheetName}`)
  }
  return sheetId
}

function requireCellAddress(sheetName: string, a1Address: string): WorkPaperCellAddress {
  const sheetId = requireSheet(sheetName)
  const parsed = workbook.simpleCellAddressFromString(a1Address, sheetId)

  if (parsed === undefined) {
    throw new Error(`invalid cell address: ${sheetName}!${a1Address}`)
  }

  if (parsed.sheet !== sheetId) {
    throw new Error(`address ${a1Address} does not belong to ${sheetName}`)
  }

  return parsed
}

function readComputedSummary(): SummaryReadback {
  return {
    currentMrr: readNumber(currentMrrAddress, 'Current MRR'),
    nextMonthMrr: readNumber(nextMonthMrrAddress, 'Next month MRR'),
  }
}

function readNumber(address: WorkPaperCellAddress, label: string): number {
  const value = workbook.getCellValue(address) as unknown
  if (typeof value !== 'object' || value === null || !('value' in value) || typeof value.value !== 'number') {
    throw new Error(`expected ${label} to be numeric, received ${JSON.stringify(value)}`)
  }
  return Math.round(value.value * 100) / 100
}

function serializeWorkbook(): string {
  return serializeWorkPaperDocument(
    exportWorkPaperDocument(workbook, {
      includeConfig: true,
    }),
  )
}
```

The important check is not that the write call returned. It is that the computed
summary changed as expected:

```json
{
  "editedCell": "Revenue!B3",
  "before": {
    "currentMrr": 10500,
    "nextMonthMrr": 11550
  },
  "after": {
    "currentMrr": 13500,
    "nextMonthMrr": 14850
  },
  "checks": {
    "currentMrrChanged": true,
    "nextMonthMrrChanged": true,
    "serializedBytes": 1155
  }
}
```

`serializedBytes` will vary as the document schema evolves. Treat it as a
positive persistence check, not a stable snapshot value.

## OpenAI Responses API Tool Wrapper

OpenAI function tools should stay thin. The model chooses a tool call; your
Node process parses the arguments, runs the WorkPaper function, and sends the
structured result back as a `function_call_output`. Do not ask the model to
modify workbook JSON by hand.

The maintained repository script for this section is
[`examples/headless-workpaper/openai-responses-tool-wrapper.ts`](../examples/headless-workpaper/openai-responses-tool-wrapper.ts):

```sh
cd examples/headless-workpaper
npm run agent:openai-responses
```

The official Responses API function-calling flow preserves the model output,
executes every `function_call`, appends `function_call_output` items, and sends
that input back to the model. The WorkPaper-specific part is the dispatcher:

```ts
import OpenAI from 'openai'

type OpenAiToolResult = ReturnType<typeof tools.readSummary> | ReturnType<typeof tools.setInputCell>

type OpenAiWorkPaperCall = {
  name: string
  arguments: string
}

const openai = new OpenAI()

const openAiWorkPaperTools = [
  {
    type: 'function',
    name: 'read_workpaper_summary',
    description: 'Read computed WorkPaper summary values and serialized inputs for a small A1 range.',
    parameters: {
      type: 'object',
      properties: {
        range: {
          type: 'string',
          description: 'A small A1 range including the sheet name.',
          default: 'Summary!A1:B3',
        },
      },
      required: ['range'],
      additionalProperties: false,
    },
    strict: true,
  },
  {
    type: 'function',
    name: 'set_workpaper_input_cell',
    description: 'Set one validated WorkPaper input cell and return before/after formula readback.',
    parameters: {
      type: 'object',
      properties: {
        sheetName: {
          type: 'string',
          description: 'Target sheet name, for example Revenue.',
        },
        address: {
          type: 'string',
          description: 'A1 address inside the target sheet, for example B3.',
        },
        value: {
          type: ['string', 'number', 'boolean', 'null'],
          description: 'Literal input value. Use a separate tool for formulas.',
        },
      },
      required: ['sheetName', 'address', 'value'],
      additionalProperties: false,
    },
    strict: true,
  },
] as const

const input: Array<Record<string, unknown>> = [
  {
    role: 'user',
    content: 'Set Sales customers to 25, then tell me the current MRR and next month MRR.',
  },
]

let response = await openai.responses.create({
  model: process.env.OPENAI_MODEL ?? 'gpt-5',
  tools: openAiWorkPaperTools,
  input,
})

input.push(...response.output)

for (const item of response.output) {
  if (item.type !== 'function_call') {
    continue
  }

  const result = dispatchOpenAiWorkPaperCall({
    name: item.name,
    arguments: item.arguments,
  })

  input.push({
    type: 'function_call_output',
    call_id: item.call_id,
    output: JSON.stringify(result),
  })
}

response = await openai.responses.create({
  model: process.env.OPENAI_MODEL ?? 'gpt-5',
  instructions: 'Answer from WorkPaper tool output only. Mention the edited cell and computed readback.',
  tools: openAiWorkPaperTools,
  input,
})

console.log(response.output_text)

function dispatchOpenAiWorkPaperCall(call: OpenAiWorkPaperCall): OpenAiToolResult {
  if (call.name === 'read_workpaper_summary') {
    const args = JSON.parse(call.arguments) as { range?: string }
    return tools.readSummary(args.range ?? 'Summary!A1:B3')
  }

  if (call.name === 'set_workpaper_input_cell') {
    const args = JSON.parse(call.arguments) as SetInputCellArgs
    const result = tools.setInputCell(args)

    if (!result.checks.currentMrrChanged || !result.checks.nextMonthMrrChanged) {
      throw new Error(`WorkPaper edit did not change the dependent summary: ${JSON.stringify(result.checks)}`)
    }

    return result
  }

  throw new Error(`unknown WorkPaper tool: ${call.name}`)
}
```

The object returned to OpenAI should be the same object you would log in a local
smoke test: `editedCell`, `before`, `after`, and `checks`. That makes the final
assistant message explain the workbook change from computed readback instead of
from a guess.

## Vercel AI SDK Tool Wrapper

Vercel AI SDK users can expose the same WorkPaper operations through an
AI-SDK-shaped `tools` object. This repository does not need the AI SDK as a
dependency; the snippet is for applications that already use `ai` and want a
familiar `tool()` wrapper:

```ts
import { tool } from 'ai'
import { z } from 'zod'

type WorkPaperToolValue = string | number | boolean | null

export const workPaperTools = {
  readWorkPaperSummary: tool({
    description: 'Read computed WorkPaper summary values and serialized inputs for a small range.',
    inputSchema: z.object({
      range: z.string().default('Summary!A1:B3').describe('A small A1 range, including the sheet name.'),
    }),
    execute: async ({ range = 'Summary!A1:B3' }: { range?: string }) => tools.readSummary(range),
  }),

  setWorkPaperInputCell: tool({
    description: 'Set one validated WorkPaper input cell and return before/after formula readback.',
    inputSchema: z.object({
      sheetName: z.string().describe('Target sheet name, for example Revenue.'),
      address: z.string().describe('A1 cell address inside the target sheet.'),
      value: z
        .union([z.string(), z.number(), z.boolean(), z.null()])
        .describe('Literal cell value. Use a separate formula tool for formulas.'),
    }),
    execute: async ({ sheetName, address, value }: { sheetName: string; address: string; value: WorkPaperToolValue }) => {
      const result = tools.setInputCell({ sheetName, address, value })

      if (!result.checks.currentMrrChanged || !result.checks.nextMonthMrrChanged) {
        throw new Error(`WorkPaper edit did not change the dependent summary: ${JSON.stringify(result.checks)}`)
      }

      return result
    },
  }),
}
```

Pass `workPaperTools` to `generateText()` or `streamText()` from your AI SDK
application. Keep the model-facing result structured: the mutating tool should
return `editedCell`, `before`, `after`, and `checks` so the next model step can
explain exactly what changed. Persist the serialized workbook only after these
computed readback checks pass.

For a dependency-free runnable version of this shape, use
[`examples/headless-workpaper/agent-framework-adapters.ts`](../examples/headless-workpaper/agent-framework-adapters.ts):

```sh
cd examples/headless-workpaper
npm run agent:framework-adapters
```

For the actual AI SDK `generateText()` loop, use
[`examples/headless-workpaper/ai-sdk-generate-text-tool-smoke.ts`](../examples/headless-workpaper/ai-sdk-generate-text-tool-smoke.ts):

```sh
cd examples/headless-workpaper
npm run agent:ai-sdk-generate-text
```

For the actual AI SDK `streamText()` loop, use
[`examples/headless-workpaper/ai-sdk-stream-text-tool-smoke.ts`](../examples/headless-workpaper/ai-sdk-stream-text-tool-smoke.ts):

```sh
cd examples/headless-workpaper
npm run agent:ai-sdk-stream-text
```

## LangChain Tool Wrapper

LangChain users can wrap the same SDK-neutral WorkPaper functions without adding
a LangChain dependency to this repository. In an app that already uses
LangChain, define thin tools around the `tools` object from the example above:

```ts
import { tool } from 'langchain'
import * as z from 'zod'

type WorkPaperToolValue = string | number | boolean | null

const readWorkPaperSummary = tool(({ range = 'Summary!A1:B3' }: { range?: string }) => tools.readSummary(range), {
  name: 'read_workpaper_summary',
  description: 'Read computed WorkPaper summary values and serialized inputs for a small range.',
  schema: z.object({
    range: z.string().default('Summary!A1:B3').describe('A small A1 range, including the sheet name.'),
  }),
})

const setWorkPaperInputCell = tool(
  async ({ sheetName, address, value }: { sheetName: string; address: string; value: WorkPaperToolValue }) => {
    const result = tools.setInputCell({ sheetName, address, value })

    if (!result.checks.currentMrrChanged || !result.checks.nextMonthMrrChanged) {
      throw new Error(`WorkPaper edit did not change the dependent summary: ${JSON.stringify(result.checks)}`)
    }

    return result
  },
  {
    name: 'set_workpaper_input_cell',
    description: 'Set one validated WorkPaper input cell and return before/after formula readback.',
    schema: z.object({
      sheetName: z.string().describe('Target sheet name, for example Revenue.'),
      address: z.string().describe('A1 cell address inside the target sheet.'),
      value: z
        .union([z.string(), z.number(), z.boolean(), z.null()])
        .describe('Literal cell value. Use a separate formula tool for formulas.'),
    }),
  },
)

export const workPaperTools = [readWorkPaperSummary, setWorkPaperInputCell]
```

Return structured objects, not prose. LangChain will pass the returned object
back to the model as tool output, so keep the WorkPaper result explicit:
`editedCell`, `before`, `after`, and `checks`. In a durable app, write the
serialized workbook to external storage only after these computed readback
checks pass.

## Agent Guardrails

- Validate sheet names with `getSheetId()` before parsing a target address.
- Parse user-facing addresses through `simpleCellAddressFromString()` or
  `simpleCellRangeFromString()` instead of building `{ row, col }` objects from
  ad hoc string splits.
- Return computed values after every write; do not ask the agent to infer
  success from a rendered grid.
- Serialize only after a successful write and verification readback.
- Keep tool results small. Return the range, changed cell, before/after values,
  and persistence check; do not dump the whole workbook unless the agent asks
  for it.
- Use public `@bilig/headless` exports and WorkPaper methods only. Do not import
  from internal `src/`, `dist/`, or monorepo package internals in an external
  agent workflow.

## When To Add More Tools

Add tools only after the agent has a repeated need for them:

- `readRange(range)` for broader model inspection
- `setFormula(sheetName, address, formula)` when formulas are first-class agent
  outputs
- `validateFormula(address)` when the workflow needs structured diagnostics
- `persistAndRestore()` when the workflow must prove round-trip safety before
  committing output

The same rule holds: every mutating tool should return computed verification
and enough context for the caller to explain what changed.
