# Headless WorkPaper Example

This example shows `@bilig/headless` running as a Node library with no browser
UI. It builds a small revenue workbook, evaluates formulas, uses a named
expression in the revenue plan, applies an agent-style edit, persists the
workbook, restores it, and prints the verified summary.

Run it outside the monorepo with the published package:

```sh
npm install
npm start
```

Expected output:

```json
{
  "initial": {
    "totalRevenue": 27300,
    "westCustomers": 30,
    "targetRevenue": 30576
  },
  "afterAgentEdit": {
    "totalRevenue": 36900,
    "westCustomers": 38,
    "enterpriseArpa": 1200,
    "targetRevenue": 41328,
    "qualifiedCustomerCounts": [20, 30, 18]
  },
  "persistedSheets": ["Deals", "Summary"],
  "persistedNamedExpressions": ["GrowthRatePercent"],
  "restoredGrowthRatePercent": 12
}
```

The repository smoke test runs this same example against packed local runtime
packages through `pnpm workpaper:smoke:external`.

## Command Index

| Use case                 | Command                            | What it proves                                                                                                    |
| ------------------------ | ---------------------------------- | ----------------------------------------------------------------------------------------------------------------- |
| Quick revenue workbook   | `npm start`                        | formulas, named expressions, persistence                                                                          |
| Agent tool call loop     | `npm run agent:tool-call`          | read, edit, verify, serialize, restore                                                                            |
| OpenAI Responses wrapper | `npm run agent:openai-responses`   | `function_call` dispatch, `function_call_output`, verified WorkPaper readback                                     |
| AI SDK generateText      | `npm run agent:ai-sdk-generate-text` | real `generateText()` and `tool()` calls with verified WorkPaper readback                                         |
| AI SDK streamText        | `npm run agent:ai-sdk-stream-text` | real `streamText()` and streamed tool calls with verified WorkPaper readback                                      |
| Agent framework adapters | `npm run agent:framework-adapters` | TypeScript wrappers for AI SDK, LangChain, Mastra, LlamaIndex.TS, LangGraph.js, CopilotKit, and Cloudflare Agents |
| MCP tool server shape    | `npm run agent:mcp-tools`          | `tools/list`, `tools/call`, verified edits                                                                        |
| MCP stdio server         | `npm run agent:mcp-stdio`          | newline-delimited JSON-RPC over stdin/stdout                                                                      |
| npm package eval         | `npm run npm-eval`                 | the same `.ts` file used by the npm-only smoke test                                                               |
| Agent writeback check    | `npm run agent:verify`             | exact input edits and formula preservation                                                                        |
| Budget variance alerts   | `npm run budget-variance`          | budget, actuals, variance, alert formulas                                                                         |
| Fulfillment capacity     | `npm run fulfillment-capacity`     | orders, labor hours, capacity gap                                                                                 |
| Quote approval           | `npm run quote-approval`           | quote total, discount, approval threshold                                                                         |
| Subscription MRR         | `npm run subscription-mrr`         | churn, expansion, ending MRR forecast                                                                             |
| Revenue scenarios        | `npm run scenarios`                | multi-sheet formulas and planning edits                                                                           |
| Persistence round trip   | `npm run persistence`              | save, restore, edit, and export                                                                                   |
| Named expression update  | `npm run named-expression`         | workbook-scoped names and dependent formulas                                                                      |
| CSV-shaped input         | `npm run csv-shaped`               | CSV-shaped data plus formula summary                                                                              |
| Invoice totals           | `npm run invoice-totals`           | line items, subtotal, tax, total                                                                                  |
| JSON records input       | `npm run json-records`             | API records to formula-backed workbook                                                                            |
| JSON file input          | `npm run json-file`                | disk JSON records to verified summary                                                                             |
| Formula diagnostics      | `npm run formula-diagnostics`      | display errors and structured diagnostics                                                                         |
| Markdown report output   | `npm run markdown-report`          | calculated plain-text report generation                                                                           |
| Snapshot diff            | `npm run snapshot-diff`            | persisted before/after input and outputs                                                                          |
| Range readback           | `npm run range-readback`           | computed values and serialized formulas                                                                           |
| Sheet inspection         | `npm run sheet-inspection`         | restored sheet names, IDs, and dimensions                                                                         |
| HTTP JSON summary        | `npm run http-json-summary`        | no-framework Node HTTP service boundary                                                                           |

For durable service storage, see the docs recipe for
[plain node-postgres (`pg`) WorkPaper JSON persistence](../../docs/node-service-workpaper-recipe.md#plain-node-postgres-pg-json-persistence).
It is the low-level Postgres path for teams not using Prisma, Drizzle, or
Kysely, and includes save/load SQL plus restored WorkPaper readback
verification.

## npm Package Eval

Run this when you want the smallest maintained TypeScript file for checking the
published package. It creates two sheets, edits one input cell, serializes the
document, restores it, and verifies the recalculated value:

```sh
npm run npm-eval
```

Expected output:

```json
{
  "before": 24000,
  "after": 38400,
  "afterRestore": 38400,
  "sheets": ["Inputs", "Summary"],
  "bytes": 1000,
  "verified": true
}
```

The exact byte count can move between package versions. The important part is
that `verified` is `true` and `afterRestore` matches `after`.

## Agent Tool Call Loop

Run the tool-call loop example when you want a small SDK-neutral artifact for
wrapping WorkPaper operations as agent tools. It reads a summary range, applies
a planned input edit through a `setInputCell` tool, verifies formula readback,
persists the workbook, restores it, and checks that computed outputs survive
the round trip:

```sh
npm run agent:tool-call
```

Expected output:

```json
{
  "toolCall": {
    "toolName": "setInputCell",
    "arguments": {
      "sheetName": "Inputs",
      "address": "B3",
      "value": 0.4,
      "reason": "Use the latest qualified pipeline conversion estimate."
    }
  },
  "toolResult": {
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
    "verified": {
      "previousValue": 0.25,
      "newValue": 0.4,
      "formulasPersisted": true,
      "restoredMatchesAfter": true,
      "expectedArrImproved": true,
      "targetGapClosed": true
    }
  }
}
```

The actual output also includes the initial range read, formula contracts, the
restored summary, and serialized byte count.

For agent frameworks, the
[`WorkPaper tool-calling recipe`](../../docs/agent-workpaper-tool-calling-recipe.md)
also links to wrappers that keep the same validation and computed readback
contract across the OpenAI Responses API, AI SDK, LangChain, Mastra,
LlamaIndex.TS, LangGraph.js, CopilotKit, and Cloudflare Agents.

## OpenAI Responses Tool Wrapper

Run this when your app calls OpenAI Responses directly and you want the
application-side WorkPaper dispatcher without an API key:

```sh
npm run agent:openai-responses
```

The example mirrors the Responses tool loop: model output contains
`function_call` items, the Node process runs the WorkPaper tools, and the next
input includes matching `function_call_output` items.

Expected proof:

```json
{
  "apiShape": "OpenAI Responses function_call -> function_call_output",
  "toolNames": ["read_workpaper_summary", "set_workpaper_input_cell"],
  "followupInputTypes": ["user", "function_call", "function_call", "function_call_output", "function_call_output"],
  "writeResult": {
    "editedCell": "Inputs!B3",
    "before": { "expectedArr": 60000, "targetGap": -34000 },
    "after": { "expectedArr": 96000, "targetGap": 5600 },
    "checks": {
      "formulasPersisted": true,
      "restoredMatchesAfter": true,
      "expectedArrChanged": true
    }
  }
}
```

Use this file as the local dispatcher around the official OpenAI Responses API
call. The workbook logic stays in TypeScript functions; the model only sees the
tool schema and structured tool output.

## AI SDK GenerateText Tool Smoke

Run this when your app uses the Vercel AI SDK and you want the actual
`generateText()` loop, not just a framework-shaped object:

```sh
npm run agent:ai-sdk-generate-text
```

The script imports `generateText`, `stepCountIs`, and `tool` from `ai`. It uses
`MockLanguageModelV3` from `ai/test` so the smoke test is deterministic and does
not need a provider key. The mocked model asks for two tools:

- `readWorkPaperSummary` reads `Summary!A1:B5`.
- `setWorkPaperInputCell` writes `Inputs!B3 = 0.4` and returns computed
  readback.

Expected proof:

```json
{
  "apiShape": "AI SDK generateText -> tool -> execute",
  "modelCallCount": 2,
  "toolNames": ["readWorkPaperSummary", "setWorkPaperInputCell"],
  "writeResult": {
    "editedCell": "Inputs!B3",
    "before": { "expectedArr": 60000, "targetGap": -34000 },
    "after": { "expectedArr": 96000, "targetGap": 5600 },
    "checks": {
      "formulasPersisted": true,
      "restoredMatchesAfter": true,
      "expectedArrChanged": true
    }
  }
}
```

Use [`ai-sdk-generate-text-tool-smoke.ts`](ai-sdk-generate-text-tool-smoke.ts)
when you want a copyable TypeScript file that proves the AI SDK can call the
WorkPaper tools and receive structured results.

## AI SDK StreamText Tool Smoke

Run this when your app uses the Vercel AI SDK streaming path and you want the
same WorkPaper read/write proof:

```sh
npm run agent:ai-sdk-stream-text
```

The script imports `streamText`, `stepCountIs`, and `simulateReadableStream`
from `ai`. It uses `MockLanguageModelV3` from `ai/test`, so the example stays
provider-free. The model stream emits tool calls, the AI SDK executes the
WorkPaper tools, and the final answer is streamed as text deltas.

Expected proof:

```json
{
  "apiShape": "AI SDK streamText -> tool -> execute",
  "modelStreamCallCount": 2,
  "streamChunkTypes": ["tool-call", "tool-result", "tool-call", "tool-result", "text-delta", "text-delta"],
  "toolNames": ["readWorkPaperSummary", "setWorkPaperInputCell"],
  "writeResult": {
    "editedCell": "Inputs!B3",
    "before": { "expectedArr": 60000, "targetGap": -34000 },
    "after": { "expectedArr": 96000, "targetGap": 5600 },
    "checks": {
      "formulasPersisted": true,
      "restoredMatchesAfter": true,
      "expectedArrChanged": true
    }
  }
}
```

Use [`ai-sdk-stream-text-tool-smoke.ts`](ai-sdk-stream-text-tool-smoke.ts)
when you want a copyable TypeScript file for the AI SDK streaming loop.

## Agent Framework Adapters

Run this when you want copyable TypeScript wrapper shapes for common agent
frameworks without adding those frameworks to the standalone example:

```sh
npm run agent:framework-adapters
```

Expected output:

```json
{
  "aiSdk": {
    "toolNames": ["readWorkPaperSummary", "setWorkPaperInputCell"],
    "writeResult": {
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
  },
  "openAiResponses": {
    "toolNames": ["read_workpaper_summary", "set_workpaper_input_cell"],
    "toolOutputTypes": ["function_call_output", "function_call_output"],
    "writeResult": {
      "editedCell": "Inputs!B3",
      "checks": {
        "formulasPersisted": true,
        "restoredMatchesAfter": true,
        "expectedArrChanged": true
      }
    }
  },
  "langChain": {
    "toolNames": ["read_workpaper_summary", "set_workpaper_input_cell"],
    "writeResult": {
      "editedCell": "Inputs!B3",
      "checks": {
        "formulasPersisted": true,
        "restoredMatchesAfter": true,
        "expectedArrChanged": true
      }
    }
  },
  "mastra": {
    "toolIds": ["read-workpaper-summary", "set-workpaper-input-cell"],
    "writeResult": {
      "editedCell": "Inputs!B3",
      "checks": {
        "formulasPersisted": true,
        "restoredMatchesAfter": true,
        "expectedArrChanged": true
      }
    }
  },
  "llamaIndex": {
    "toolNames": ["read_workpaper_summary", "set_workpaper_input_cell"]
  },
  "langGraph": {
    "nodeName": "tools",
    "toolNames": ["read_workpaper_summary", "set_workpaper_input_cell"]
  },
  "copilotKit": {
    "actionNames": ["readWorkPaperSummary", "setWorkPaperInputCell"]
  },
  "cloudflareAgents": {
    "toolNames": ["readWorkPaperSummary", "setWorkPaperInputCell"]
  }
}
```

The script uses real `zod` schemas and one WorkPaper tool implementation, then
adapts it to:

- AI SDK-style `execute({ ... })` tools
- OpenAI Responses API `function_call` and `function_call_output` messages
- LangChain `tool(..., { schema })` / LangGraph `ToolNode` shapes
- Mastra `createTool({ id, inputSchema, outputSchema, execute })`
- LlamaIndex.TS `tool(fn, { parameters })` / `FunctionTool` shapes
- CopilotKit `useCopilotAction({ parameters, handler })` actions
- Cloudflare Agents `AIChatAgent` / `streamText({ tools })` style tools

The actual output also includes read results, formula contracts, restored
summary, and serialized byte counts for every write path.

## MCP Tool Server Shape

Run this when you want an MCP-style tool surface without pulling in an MCP SDK
or transport dependency:

```sh
npm run agent:mcp-tools
```

The script exposes the same WorkPaper functions through two JSON-RPC methods:

- `tools/list` returns `read_workpaper_summary` and
  `set_workpaper_input_cell` with JSON Schema input definitions and MCP tool
  annotations.
- `tools/call` runs the selected tool and returns both text content and
  structured output for computed readback.

The read tool is annotated as read-only, idempotent, and closed-world. The
write tool is annotated as mutating local WorkPaper state, idempotent for the
same cell/value arguments, and closed-world.

Expected write output:

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

The actual output also includes the `tools/list` response, read response,
formula contracts, restored summary, and serialized byte count.

## MCP Stdio Server

Run this when you want the same tools behind a local stdio transport:

```sh
printf '%s\n' \
  '{"jsonrpc":"2.0","id":1,"method":"initialize"}' \
  '{"jsonrpc":"2.0","method":"notifications/initialized"}' \
  '{"jsonrpc":"2.0","id":2,"method":"tools/list"}' \
  '{"jsonrpc":"2.0","id":3,"method":"tools/call","params":{"name":"set_workpaper_input_cell","arguments":{"sheetName":"Inputs","address":"B3","value":0.4}}}' |
  npm run --silent agent:mcp-stdio
```

The server reads newline-delimited JSON-RPC requests from stdin and writes one
JSON-RPC response per line to stdout. It supports `initialize`,
`notifications/initialized`, `tools/list`, and `tools/call` without adding a
transport package or MCP SDK dependency.

### Vercel AI SDK MCP Client Recipe

Use the published stdio command when you want an AI SDK agent to call the same
MCP tools from a TypeScript workflow:

```ts
import { createMCPClient } from '@ai-sdk/mcp'
import { Experimental_StdioMCPTransport } from '@ai-sdk/mcp/mcp-stdio'
import { generateText } from 'ai'

const client = await createMCPClient({
  transport: new Experimental_StdioMCPTransport({
    command: 'npm',
    args: ['exec', '--package', '@bilig/headless', '--', 'bilig-workpaper-mcp'],
  }),
})

try {
  const tools = await client.tools()
  const { text } = await generateText({
    model: 'your-model',
    tools,
    prompt: [
      'Read the WorkPaper summary with read_workpaper_summary for Summary!A1:B5.',
      'Then set Inputs!B3 to 0.4 with set_workpaper_input_cell.',
      'Return editedCell plus the before and after expectedArr values.',
    ].join('\n'),
  })

  console.log(text)
} finally {
  await client.close()
}
```

The important calls are still the MCP `read_workpaper_summary` read and the
`set_workpaper_input_cell` write. The server command is `bilig-workpaper-mcp`;
`npm exec --package @bilig/headless -- bilig-workpaper-mcp` just resolves the
published package for a clean local recipe. The stdio transport receives `npm`
as the command and the rest as `args`, so the SDK launches the process directly.

Verify this docs recipe from the repo root with:

```sh
pnpm docs:discovery:check
```

### Published Package MCP Client Config

Use this when you want an MCP client to start the published WorkPaper server,
not the copy from your local checkout.

Claude Desktop:

```json
{
  "mcpServers": {
    "bilig-workpaper": {
      "type": "stdio",
      "command": "npm",
      "args": ["exec", "--package", "@bilig/headless", "--", "bilig-workpaper-mcp"],
      "env": {}
    }
  }
}
```

Cline:

```json
{
  "mcpServers": {
    "bilig-workpaper": {
      "command": "npm",
      "args": ["exec", "--package", "@bilig/headless", "--", "bilig-workpaper-mcp"],
      "env": {},
      "disabled": false,
      "autoApprove": []
    }
  }
}
```

After the client shows the tools, use the same small writeback check:

```text
Call read_workpaper_summary for Summary!A1:B5.
Then call set_workpaper_input_cell on Inputs!B3 with value 0.4.
Return editedCell, before.expectedArr, after.expectedArr, and checks.
```

### MCP Stdio Troubleshooting

| Symptom                        | What to check                                                                                   |
| ------------------------------ | ----------------------------------------------------------------------------------------------- |
| `Parse error` response         | Make sure each stdin line is valid JSON before it reaches the server.                           |
| No response appears            | End each JSON-RPC message with a newline; the server waits for newline-delimited input.         |
| Notification has no output     | `notifications/initialized` is intentionally one-way and does not produce a JSON-RPC response.  |
| `Invalid params` or tool error | Check that `tools/call` includes a supported `name` and the required `arguments` for that tool. |

### Local MCP Client Config

From a clean checkout, install the example dependencies first:

```sh
cd examples/headless-workpaper
npm install
```

Then point your local MCP client at the stdio entrypoint. Replace the path with
the absolute path to your checkout:

```json
{
  "mcpServers": {
    "bilig-workpaper": {
      "command": "npm",
      "args": ["--prefix", "/absolute/path/to/bilig/examples/headless-workpaper", "run", "--silent", "agent:mcp-stdio"]
    }
  }
}
```

## Agent Writeback Verification

Run the agent verification demo when you want a small artifact for the claim
that spreadsheet agents need workbook APIs, not screenshots. It applies an
agent-style assumption edit, records the exact input cells changed, verifies the
dependent formulas and readback values, persists the workbook, restores it, and
checks that formulas and outputs survived the round trip:

```sh
npm run agent:verify
```

Expected output:

```json
{
  "edits": [
    { "cell": "Assumptions!B2", "before": 500, "after": 650 },
    { "cell": "Assumptions!B3", "before": 0.08, "after": 0.1 },
    { "cell": "Assumptions!B5", "before": 1.1, "after": 1.2 }
  ],
  "before": {
    "customers": 40,
    "grossMrr": 9600,
    "expansionMrr": 10560,
    "annualizedArr": 126720,
    "arrTargetDelta": -23280
  },
  "after": {
    "customers": 65,
    "grossMrr": 15600,
    "expansionMrr": 18720,
    "annualizedArr": 224640,
    "arrTargetDelta": 74640
  },
  "restored": {
    "customers": 65,
    "grossMrr": 15600,
    "expansionMrr": 18720,
    "annualizedArr": 224640,
    "arrTargetDelta": 74640
  },
  "formulaContracts": {
    "customers": "=Assumptions!B2*Assumptions!B3",
    "grossMrr": "=B2*Assumptions!B4",
    "expansionMrr": "=B3*Assumptions!B5",
    "annualizedArr": "=B4*12",
    "arrTargetDelta": "=Plan!B5-150000"
  },
  "verified": {
    "formulasUnchanged": true,
    "formulasPersisted": true,
    "restoredMatchesAfter": true,
    "serializedBytes": 1237
  }
}
```

## Revenue Scenarios

Run the scenario model when you want to see a multi-sheet revenue workbook,
formula-backed projections, an agent-style planning edit, and persistence
readback:

```sh
npm run scenarios
```

Expected output:

```json
{
  "beforeEdit": {
    "totalNetMrr": 119267.2,
    "annualRunRate": 1431206.4,
    "enterpriseNetMrr": 57456,
    "expansionTarget": 1688823.55,
    "scenarios": {
      "conservativeNetMrr": 107340.48,
      "expansionNetMrr": 137157.28,
      "stretchNetMrr": 161010.72
    }
  },
  "afterEdit": {
    "totalNetMrr": 136791.2,
    "annualRunRate": 1641494.4,
    "enterpriseNetMrr": 66074.4,
    "expansionTarget": 1936963.39,
    "scenarios": {
      "conservativeNetMrr": 123112.08,
      "expansionNetMrr": 157309.88,
      "stretchNetMrr": 184668.12
    }
  },
  "persistedSheets": ["Pipeline", "Summary", "Scenarios"],
  "serializedBytes": 1594
}
```

## Subscription MRR Forecast

Run the subscription MRR example when you want a compact service-side forecast
for plan price, churn, expansion, new customers, and ending recurring revenue:

```sh
npm run subscription-mrr
```

Expected output:

```json
{
  "months": 4,
  "startingMrr": 5880,
  "endingMrr": 9604.03,
  "endingCustomers": 181.48,
  "netExpansionMrr": 711.41,
  "fourMonthNetMrr": 33044.9,
  "mrrDelta": 3724.03,
  "firstForecastRow": [
    "January",
    18,
    "=Assumptions!B2",
    "=C2*Assumptions!B4",
    "=C2-D2+B2",
    "=E2*Assumptions!B3",
    "=F2*Assumptions!B5",
    "=F2+G2"
  ],
  "verified": true
}
```

## Quote Approval Threshold

Run the quote approval example when you want a compact sales-ops workflow that
calculates line totals, discount amount, quote total, and an approval flag for
discounts above the threshold:

```sh
npm run quote-approval
```

Expected output:

```json
{
  "quoteId": "Q-2026-041",
  "lineItems": 4,
  "listTotal": 6980,
  "discountAmount": 993,
  "quoteTotal": 5987,
  "discountPercent": 0.1423,
  "maxLineDiscount": 0.25,
  "approvalRequired": "Review",
  "reviewedSku": "SETUP",
  "firstQuoteRow": ["PRO-ANNUAL", 12, 240, 0.1, "=B2*C2", "=E2*D2", "=E2-F2", "=IF(D2>0.2,\"Review\",\"OK\")"],
  "verified": true
}
```

## Fulfillment Capacity Plan

Run the fulfillment capacity example when you want a compact operations
workflow that compares forecast order volume with available labor hours and
reports the capacity gap:

```sh
npm run fulfillment-capacity
```

Expected output:

```json
{
  "days": 4,
  "forecastOrders": 2020,
  "requiredHours": 61.0318,
  "availableHours": 60,
  "capacityGap": -1.0318,
  "status": "Short",
  "shortDays": 2,
  "largestDailyShortfall": -1.5667,
  "bottleneckDay": "Thursday",
  "firstCapacityRow": ["Monday", 420, 1.8, 55, 14, "=B2*C2/D2", "=E2-F2", "=IF(G2<0,\"Short\",\"Ready\")"],
  "verified": true
}
```

## Budget Variance Alerts

Run the budget variance example when you want a compact service-side reporting
workflow. It compares budget and actual rows, calculates dollar variance,
variance percent, and an alert formula for rows that are more than 10 percent
over budget:

```sh
npm run budget-variance
```

Expected output:

```json
{
  "rows": 4,
  "flaggedDepartment": "Marketing",
  "varianceAmount": 7500,
  "variancePercent": 0.15,
  "summary": {
    "totalBudget": 185000,
    "totalActual": 196600,
    "totalVariance": 11600,
    "largestOverage": 7500,
    "largestVariancePercent": 0.15,
    "reviewCount": 1
  },
  "firstVarianceRow": ["Marketing", 50000, 57500, "=C2-B2", "=D2/B2", "=IF(E2>0.1,\"Review\",\"OK\")"],
  "verified": true
}
```

## Invoice Totals

Run the invoice totals example when you want a compact service-side billing
workflow. It builds invoice line items, calculates line totals, subtotal, tax,
and grand total formulas, then verifies both computed values and serialized
formula readback:

```sh
npm run invoice-totals
```

Expected output:

```json
{
  "invoiceNumber": "INV-2026-001",
  "lineItems": 4,
  "subtotal": 1890,
  "taxRate": 0.08,
  "tax": 151.2,
  "total": 2041.2,
  "formulas": [["=SUM(Invoice!D2:D5)"], [0.08], ["=B2*B3"], ["=B2+B4"]],
  "firstLineItem": ["Implementation workshop", 5, 120, "=B2*C2"],
  "verified": true
}
```

## Persistence Round Trip

Run the focused persistence example when you want to see a WorkPaper document
written to disk, restored, edited, and exported again:

```sh
npm run persistence
```

Expected output:

```json
{
  "beforeSave": {
    "quarterNetMrr": 42100,
    "annualizedRunRate": 505200,
    "expansionAdjustedArr": 545616
  },
  "afterRestoreAndEdit": {
    "quarterNetMrr": 45100,
    "annualizedRunRate": 541200,
    "expansionAdjustedArr": 584496
  },
  "persistedSheets": ["Plan", "Summary"],
  "persistedNamedExpressions": ["ExpansionRatePercent"],
  "saveFileBytes": 1209
}
```

## Named Expression Update

Run the named expression example when you want to see a service or agent change
a workbook-scoped named expression, recalculate dependent formulas, persist the
workbook, restore it, and verify the restored value still matches the edited
state:

```sh
npm run named-expression
```

Expected output:

```json
{
  "verified": true,
  "namedExpression": "GrowthRatePercent",
  "before": {
    "baseRevenue": 36000,
    "growthAdjustedRevenue": 39600
  },
  "after": {
    "baseRevenue": 36000,
    "growthAdjustedRevenue": 45000
  },
  "restored": {
    "baseRevenue": 36000,
    "growthAdjustedRevenue": 45000
  },
  "namedExpressionValues": {
    "before": 10,
    "after": 25,
    "restored": 25
  },
  "persistedNamedExpressions": ["GrowthRatePercent"],
  "restoredMatchesAfter": true
}
```

## CSV Shaped Input

Run the CSV shaped input example when a service receives a small tabular
payload, normalizes it into a WorkPaper, and needs formula-backed totals rather
than hand-coded arithmetic:

```sh
npm run csv-shaped
```

Expected output:

```json
{
  "sourceRows": 3,
  "computed": {
    "totalRevenue": 36900,
    "westCustomers": 20,
    "largestDeal": 24000
  },
  "serializedFirstDataRow": ["West", 20, 1200, "=B2*C2"],
  "verified": true
}
```

## JSON Records Input

Run the JSON records input example when a Node service or agent already has an
array of API records and needs to turn it into a formula-backed WorkPaper
without writing an import subsystem:

```sh
npm run json-records
```

Expected output:

```json
{
  "sourceRecords": 3,
  "computed": {
    "committedMrr": 39600,
    "weightedPipelineMrr": 43400,
    "westSeats": 27,
    "largestOpportunityMrr": 21600
  },
  "serializedFirstDataRow": ["Acme Manufacturing", "West", "Committed", 12, 1800, 1, "=D2*E2", "=G2*F2"],
  "verified": true
}
```

## JSON File Input

Run the JSON file input example when a script or service already has exported
records on disk. It reads `fixtures/opportunities.json`, builds the same
formula-backed WorkPaper summary, verifies the expected output, and prints a
compact JSON result:

```sh
npm run json-file
```

Expected output:

```json
{
  "verified": true,
  "source": "fixtures/opportunities.json",
  "sourceRecords": 3,
  "computed": {
    "committedMrr": 39600,
    "weightedPipelineMrr": 43400,
    "westSeats": 27,
    "largestOpportunityMrr": 21600
  }
}
```

## Formula Diagnostics

Run the formula diagnostics example when a Node service or agent needs to turn a
visible workbook error into a structured response. It builds a WorkPaper with
one invalid `XIRR()` formula and one valid `XIRR()` formula, reads the display
value with `getCellDisplayValue()`, reads structured diagnostics with
`getCellFormulaDiagnostics()`, verifies the diagnostic code and references, and
prints a compact JSON result:

```sh
npm run formula-diagnostics
```

Expected output:

```json
{
  "verified": true,
  "invalidDisplay": "#VALUE!",
  "invalidDiagnostics": [
    {
      "code": "financial-unsupported-date-coercion",
      "functionName": "XIRR",
      "errorText": "#VALUE!",
      "references": ["Tax!D2:D5", "Tax!D2"]
    }
  ],
  "validDisplay": "0.02256857579464",
  "validValue": 0.02256857579463996
}
```

## Markdown Report Output

Run the Markdown report example when a service or agent needs a plain-text
artifact for a pull request, job summary, Slack draft, or email body. It builds
a WorkPaper workbook, reads calculated summary cells, formats the result as a
Markdown table, verifies the exact text, and prints the table inside JSON:

```sh
npm run markdown-report
```

Expected output:

```json
{
  "verified": true,
  "report": "| Metric | Value |\n| --- | ---: |\n| Committed MRR | $39,600 |\n| Weighted pipeline MRR | $43,400 |\n| Target gap | $10,400 |"
}
```

Generated report:

```md
| Metric                |   Value |
| --------------------- | ------: |
| Committed MRR         | $39,600 |
| Weighted pipeline MRR | $43,400 |
| Target gap            | $10,400 |
```

## Snapshot Diff

Run the snapshot diff example when a service or agent needs to show how a
programmatic edit changed both persisted workbook input and dependent summary
values. It exports the WorkPaper document before and after a cell edit, compares
the edited input cell, reads formula-backed summary values, verifies the exact
diff, and prints a compact JSON result:

```sh
npm run snapshot-diff
```

Expected output:

```json
{
  "verified": true,
  "changedCell": "Revenue!B2",
  "beforeSerializedInput": 12000,
  "afterSerializedInput": 15000,
  "changedSummaryValues": {
    "before": {
      "netMrr": 14200,
      "annualizedArr": 170400
    },
    "after": {
      "netMrr": 17200,
      "annualizedArr": 206400
    }
  },
  "documentBytes": {
    "before": 1058,
    "after": 1058
  }
}
```

## Range Readback

Run the range readback example when a service, agent, or test needs both
calculated cell values and the source formulas for the same WorkPaper range. It
builds a tiny revenue workbook, reads `Summary!A1:B3` with `getRangeValues()`,
reads the same range with `getRangeSerialized()`, verifies both views, and
prints a compact JSON result:

```sh
npm run range-readback
```

Expected output:

```json
{
  "verified": true,
  "range": "Summary!A1:B3",
  "valueReadback": [
    ["Metric", "Value"],
    ["Total MRR", 31500],
    ["West Customers", 20]
  ],
  "serializedReadback": [
    ["Metric", "Value"],
    ["Total MRR", "=SUM(Revenue!D2:D3)"],
    ["West Customers", "=Revenue!B2"]
  ]
}
```

## Sheet Inspection

Run the sheet inspection example when a service or agent should check workbook
shape before writing cells. It builds a two-sheet WorkPaper, persists and
restores it, reads the restored sheet names with `getSheetNames()`, verifies a
`Summary` sheet lookup with `getSheetId()`, and prints a compact JSON result:

```sh
npm run sheet-inspection
```

Expected output:

```json
{
  "verified": true,
  "restoredSheets": ["Inputs", "Summary"],
  "lookup": {
    "query": "Summary",
    "sheetId": 2,
    "sheetName": "Summary",
    "dimensions": {
      "width": 2,
      "height": 3
    }
  }
}
```

## HTTP JSON Summary

Run the HTTP JSON summary example when you want the same record-to-WorkPaper
pattern behind a tiny Node service boundary. The script starts a local
`node:http` server on an ephemeral port, posts opportunity records with
`fetch`, builds a WorkPaper from the posted JSON, reads formula-backed summary
cells, verifies the exact response, prints the response, and closes the server:

```sh
npm run http-json-summary
```

Expected output:

```json
{
  "verified": true,
  "sourceRecords": 3,
  "computed": {
    "committedMrr": 39600,
    "weightedPipelineMrr": 43400,
    "westSeats": 27,
    "largestOpportunityMrr": 21600
  }
}
```
