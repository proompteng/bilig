---
title: Stop driving spreadsheets with screenshots. Run formula workbooks in Node.
description: Replace spreadsheet screen automation with @bilig/workpaper tools that edit inputs, recalculate formulas, read proof cells, and persist workbook state.
canonical_url: https://proompteng.github.io/bilig/stop-driving-spreadsheets-with-screenshots.html
image: /assets/github-social-preview.png
---

# Stop driving spreadsheets with screenshots. Run formula workbooks in Node.

Research refreshed: 2026-06-03.

Spreadsheets are still where teams keep a lot of operational logic: pricing
rules, revenue models, quote approvals, capacity plans, billing checks, and
import validation. The problem is not that spreadsheets exist. The problem is
that automation often treats the grid as pixels instead of as state.

If a backend job or coding agent has to click cells, infer formulas from a
rendered view, and trust a screenshot after the edit, the verification boundary
is weak. The automation can look right while still failing to prove which input
changed, whether a dependent formula recalculated, whether the workbook state
persisted, or whether a later restore returns the same computed value.

`@bilig/workpaper` is built around the smaller claim: keep spreadsheet-shaped
business logic in a workbook model, but run it through a TypeScript API, MCP
tool server, or agent tool in Node services and coding-agent workflows.

## The Failure Mode

Screenshot-driven spreadsheet automation is brittle because the visible grid is
not the whole workbook contract.

A serious workflow usually needs to answer these questions:

- Which cells were changed?
- Which cells are formulas rather than literals?
- Did dependent formulas recalculate after the edit?
- Did the value survive export and restore?
- Can the same operation run in CI without a browser session?
- Can an agent return exact readback instead of a plausible visual answer?
- Did the tool call expose the formula result, or only the fact that a write
  call returned?

Those are state questions. A screenshot can help a human inspect the final
shape, but it should not be the only proof that a calculation is correct.

## The API Boundary

A WorkPaper gives code explicit operations:

- build sheets from arrays or records
- write cells and formulas
- read calculated values and display values
- export a JSON workbook document
- restore that document and re-read the result
- expose the same operations as agent or MCP tools

That makes the workbook a reviewable calculation artifact instead of a browser
grid that an automation script has to push around.

## Run The No-Key Proof First

For an agent or MCP client, start with the maintained evaluator:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
```

A useful result includes:

```json
{
  "schemaVersion": "bilig-evaluator.v1",
  "door": "agent-mcp",
  "editedCell": "Inputs!B3",
  "before": 60000,
  "after": 96000,
  "restoredMatchesAfter": true,
  "verified": true
}
```

The exact cells and serialized byte count can change between releases. The
important part is the shape: tool discovery, input edit, dependent formula
readback, exported or persisted WorkPaper state, restore or restart readback,
and `verified: true`.

For a direct Node service instead of MCP, use:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door workpaper-service --json
```

## Minimal TypeScript

```ts
import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from '@bilig/workpaper'

const workbook = WorkPaper.buildFromSheets({
  Inputs: [
    ['Metric', 'Value'],
    ['Seats', 25],
    ['Price', 147],
  ],
  Summary: [
    ['Metric', 'Value'],
    ['Total', '=Inputs!B2*Inputs!B3'],
  ],
})

const inputs = workbook.getSheetId('Inputs')
const summary = workbook.getSheetId('Summary')
if (inputs === undefined || summary === undefined) {
  throw new Error('Workbook did not create the expected sheets')
}

const before = readNumber(workbook.getCellValue({ sheet: summary, row: 1, col: 1 }))
workbook.setCellContents({ sheet: inputs, row: 1, col: 1 }, 40)
const after = readNumber(workbook.getCellValue({ sheet: summary, row: 1, col: 1 }))

const saved = serializeWorkPaperDocument(exportWorkPaperDocument(workbook, { includeConfig: true }))
const restored = createWorkPaperFromDocument(parseWorkPaperDocument(saved))
const restoredSummary = restored.getSheetId('Summary')
if (restoredSummary === undefined) {
  throw new Error('Restored workbook did not create the Summary sheet')
}

const afterRestore = readNumber(
  restored.getCellValue({
    sheet: restoredSummary,
    row: 1,
    col: 1,
  }),
)

console.log({
  before,
  after,
  afterRestore,
  verified: after === afterRestore,
})

function readNumber(cell: unknown): number {
  if (typeof cell === 'object' && cell !== null && typeof (cell as { value: unknown }).value === 'number') {
    return (cell as { value: number }).value
  }
  if (typeof cell === 'number') {
    return cell
  }
  throw new Error(`Expected numeric cell value, got ${JSON.stringify(cell)}`)
}
```

Expected output:

```json
{
  "before": 3675,
  "after": 5880,
  "afterRestore": 5880,
  "verified": true
}
```

## Agent Host Boundary

MCP is a tool and context boundary, not a spreadsheet screen. The agent host can
make WorkPaper tools available, but the proof still has to come from the
returned workbook state.

Use the boundary that matches the host:

- OpenAI Agents SDK and Responses API: expose WorkPaper as local stdio,
  Streamable HTTP, or hosted MCP when that trust model fits. Keep `allowed_tools`
  narrow for workbook tasks, and treat computer-use UI control as a fallback for
  visual-only workflows.
- GitHub Copilot in VS Code: use `.vscode/mcp.json`, trust the server, and
  enable only the WorkPaper tools needed for the current task. Use
  `biligWorkpaperFile` when project-local persistence matters; use
  `biligWorkpaperDemo` only for stateless hosted smoke tests.
- GitHub Copilot cloud agent: configure repository MCP settings on GitHub, not
  only `.vscode/mcp.json`. Cloud-agent MCP config needs explicit tool
  allowlists and cannot rely on MCP resources or prompts, so keep the WorkPaper
  proof as tool calls and terminal evaluator output.
- Browser Use: let browser context inform the task, then return WorkPaper
  readback as the proof. The browser DOM, click success, and screenshots are context, not formula truth.

The useful agent result is not "clicked", "typed", or "cell updated". It is a
small object with `editedCell`, `before`, `after`, `afterRestore`,
`persistedDocumentBytes`, `verified`, and limitations.

## Where This Fits

Use a WorkPaper when the code owns the workflow:

- quote approval endpoints
- pricing and discount rules
- finance checks and payout validation
- import validation that needs formula readback
- agent tools that must prove the value after editing inputs
- MCP servers that need workbook state without opening Excel or Sheets
- Browser Use, computer-use, or RPA flows where the browser gathers context but
  WorkPaper owns the calculation proof

Keep Google Sheets or Excel when the primary job is human collaboration,
desktop workbook authoring, or full XLSX compatibility. `@bilig/workpaper` is a
formula-backed runtime boundary, not a finished Excel clone.

## Source Checks

- [OpenAI Agents SDK MCP](https://openai.github.io/openai-agents-js/guides/mcp/)
- [OpenAI computer use guide](https://developers.openai.com/api/docs/guides/tools-computer-use)
- [GitHub Copilot MCP and coding agent](https://docs.github.com/en/copilot/concepts/agents/cloud-agent/mcp-and-cloud-agent)
- [VS Code MCP servers](https://code.visualstudio.com/docs/copilot/chat/mcp-servers)
- [Browser Use custom tools](https://docs.browser-use.com/open-source/customize/tools/add)
- [Microsoft Excel calculation mode](https://learn.microsoft.com/en-us/troubleshoot/microsoft-365-apps/excel/current-mode-of-calculation)

## Useful Next Pages

- [Try `@bilig/workpaper` in Node](try-bilig-headless-in-node.md)
- [Node service WorkPaper recipe](node-service-workpaper-recipe.md)
- [ExcelJS formula recalculation in Node.js](exceljs-formula-recalculation-node.md)
- [Agent tool-calling recipe](agent-workpaper-tool-calling-recipe.md)
- [MCP spreadsheet tool server](mcp-workpaper-tool-server.md)
- [Headless spreadsheet engine comparison](headless-spreadsheet-engine-comparison.md)
- [Agent WorkPaper evaluator matrix](agent-proof-matrix.md)
- [Spreadsheet MCP server comparison](spreadsheet-mcp-server-comparison.md)

If this solves a workflow you have, the most useful signal is a star on the
repository or a concrete issue with the workbook shape you need:

<https://github.com/proompteng/bilig>
