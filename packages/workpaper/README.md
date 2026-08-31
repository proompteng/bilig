# @bilig/workpaper

Bilig WorkPaper is an API, CLI evaluator, and optional MCP server for
workbook-shaped business logic in Node.js.

Use this when business logic is easiest to review as workbook cells and
formulas, but the calculation needs to run in a backend service, queue worker,
serverless route, test, or tool.

`@bilig/workpaper` is the canonical scoped npm entrypoint. The unscoped
`bilig-workpaper` package remains published as a compatibility and search alias.

## Install

```sh
npm install @bilig/workpaper
```

## Start Here

Pick the door that matches the state you own:

| Door                    | Run first                                                                                            | What it proves                                                                                     |
| ----------------------- | ---------------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------- |
| Node service or test    | `npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door workpaper-service --json` | edit input, recalculate output, persist JSON, restore, and return `verified: true`.                |
| Tool host or MCP client | `npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json`         | tool discovery, cell mutation, formula readback, JSON export, restart proof, and `verified: true`. |
| Unsure which proof fits | `npm exec --yes --package @bilig/workpaper@latest -- bilig-agent-start --json`                       | compact routing card with proof commands, evidence fields, and public links.                       |

`bilig-agent-start --json` is intentionally small. It prints first proof
commands, required evidence fields, expected MCP tools, and public discovery
links without asking a tool host to read the whole site.

## What Success Looks Like

Run the service proof without cloning the repo:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door workpaper-service --json
```

The useful output is not a write-call status. It is readback proof:

```json
{
  "schemaVersion": "bilig-evaluator.v1",
  "door": "workpaper-service",
  "verified": true,
  "packageVersions": {
    "@bilig/workpaper": "0.164.11"
  },
  "evidence": {
    "editedCell": "Inputs!B2",
    "dependentCell": "Summary!B2",
    "before": 24000,
    "after": 38400,
    "afterRestore": 38400,
    "persistedDocumentBytes": 999
  }
}
```

For recompute and output boundaries, see
<https://proompteng.github.io/bilig/eval-workpaper-service.html#recompute-and-output-boundaries>.

For a richer tool check, add `--scenario revenue-plan` to the `agent-mcp`
evaluator. It proves `SUM`, `SUMIF`, `XLOOKUP`, `FILTER`, a named expression,
JSON persistence, and restart readback.

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --scenario revenue-plan --json
```

If the workbook has provider-backed formulas such as `IMPORTRANGE`, run
`npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --scenario provider-backed --json`.
That proves the formula fails closed with an adapter diagnostic, then verifies a
local synthetic adapter readback. It does not call Google Sheets.

Framework examples live in the repo instead of this first screen. Use the owned
examples after one evaluator passes:

- Tool runtimes: Vercel AI SDK, OpenAI Agents SDK, OpenAI Responses, Open WebUI,
  and the retained adapter smoke paths under `examples/headless-workpaper`.
- Workflow engines: the retained serverless and n8n package examples.
- Saved workbook files: use the saved-file boundary section only when a file is
  the contract.

## Searchable Example Guides

These are retained guide names that users search for on npm. They are links,
not the first-run path:

| Guide need                                     | Start here                                                                |
| ---------------------------------------------- | ------------------------------------------------------------------------- |
| n8n formula readback for self-hosted workflows | <https://proompteng.github.io/bilig/n8n-workpaper-formula-readback.html>  |
| Serverless API route shape                     | <https://proompteng.github.io/bilig/serverless-workpaper-api-route.html>  |
| Saved XLSX formula recalculation               | <https://proompteng.github.io/bilig/xlsx-formula-recalculation-node.html> |

## Use A WorkPaper In Node

```ts
import { buildA1WorkPaper } from '@bilig/workpaper'

const book = buildA1WorkPaper({
  Inputs: [
    ['Metric', 'Value'],
    ['Units', 40],
    ['Price', 1200],
  ],
  Summary: [
    ['Metric', 'Value'],
    ['Revenue', '=Inputs!B2*Inputs!B3'],
  ],
})

const proof = book.editAndReadback('Inputs!B2', 48, {
  readbackRange: 'Summary!B2',
})

console.log({
  editedCell: proof.editedCell,
  before: proof.beforeReadback.displayValues,
  after: proof.afterReadback.displayValues,
  afterRestore: proof.restoredReadback.displayValues,
  persistedDocumentBytes: proof.persistedDocumentBytes,
  verified: proof.verified,
})

book.dispose()
```

Use `book.set('Inputs!B2', 48)`, `book.setMany({ 'Inputs!B3': 1500 })`,
`book.readMany(['Inputs!B2', 'Summary!B2'])`, `book.display('Summary!B2')`,
and `book.saveJson()` when you do not need the full proof object. Use
`book.editManyAndReadback()` when several inputs should commit as one atomic
proof with typed readback comparison, formula diagnostics, persistence, and
restore checks.

## Use WorkPaper Tools With The Vercel AI SDK

Install the AI SDK and Zod in the application that owns the agent loop:

```sh
npm install @bilig/workpaper ai zod
```

Then expose a WorkPaper as normal AI SDK tools:

```ts
import { generateText, stepCountIs } from 'ai'
import { WorkPaper } from '@bilig/workpaper'
import { createAiSdkWorkPaperTools } from '@bilig/workpaper/ai-sdk'

const workpaper = WorkPaper.buildFromSheets({
  Inputs: [
    ['Metric', 'Value'],
    ['Qualified opportunities', 20],
    ['Win rate', 0.25],
    ['Average ARR', 12000],
  ],
  Summary: [
    ['Metric', 'Value'],
    ['Expected customers', '=Inputs!B2*Inputs!B3'],
    ['Expected ARR', '=B2*Inputs!B4'],
  ],
})

const tools = createAiSdkWorkPaperTools({
  workpaper,
  defaultReadRange: 'Summary!A1:B3',
  proofRange: 'Summary!A1:B3',
  writableSheets: ['Inputs'],
})

const result = await generateText({
  model,
  tools,
  stopWhen: stepCountIs(2),
  prompt: 'Read the summary, set Inputs!B3 to 0.4, then report the computed ARR change.',
})

console.log(result.text)
```

The mutating tool returns `editedCell`, `before`, `after`, `restored`, and
`checks`. Keep `writableSheets` narrow so the model can edit inputs without
rewriting formula sheets.

## Verify Without Cloning

The public package ships three no-clone checks. Start with the smallest one that
matches the state owner:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door workpaper-service --json
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
npm exec --yes --package @bilig/workpaper@latest -- bilig-agent-start --json
```

`bilig-evaluate` prints a `bilig-evaluator.v1` object with `door`, `evidence`,
`verified`, `limitations`, and the source command output.

Use the raw challenge commands only when you need a lower-level transcript for
debugging:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-agent-challenge --json
npm exec --package @bilig/workpaper@latest -- bilig-mcp-challenge --json
```

Those commands edit one input, recalculate dependent formulas, export WorkPaper
JSON, restore it, and print a `verified: true` proof object.

## Tool Host WorkPaper Handoff

When a tool host is about to solve a spreadsheet task by opening Excel,
LibreOffice, Google Sheets, or a screenshot grid, hand it the WorkPaper checklist
instead:

```sh
npx --yes skills@latest add https://bilig.proompteng.ai --list
npx --yes skills@latest add proompteng/bilig --skill bilig-workpaper --list
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
```

The kit gives the host one path: install the instructions, run the no-key MCP
evaluator, paste a workbook edit task, and require computed readback plus
persisted state before reporting success. Use `bilig-mcp-challenge --json` only
when debugging the lower-level MCP transcript.

Docs: <https://proompteng.github.io/bilig/agent-adoption-kit.html>

## Workflow Builders

Use the local formula-readback server when a workflow platform should
orchestrate the task but Bilig should own workbook state:

```sh
npm exec --package @bilig/workpaper@latest -- bilig-n8n-formula-server --port 4321
```

The retained owned example is the n8n community-node package:
`integrations/n8n-nodes-workpaper`.

Docs: <https://proompteng.github.io/bilig/n8n-workpaper-formula-readback.html>

## Saved File Boundaries

```ts
import { WorkPaper } from '@bilig/workpaper'
import { exportXlsx, importXlsx } from '@bilig/workpaper/xlsx'
```

Use saved-file commands only when a workbook file is the integration contract:

```sh
npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp --from-xlsx ./pricing.xlsx
npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp --from-xlsx ./pricing.xlsx --workpaper ./.bilig/pricing.workpaper.json --writable
```

The `--from-xlsx` path imports the file once into an in-memory MCP server by
default, or into persisted WorkPaper JSON when `--workpaper --writable` is also
supplied. It also exposes `analyze_workbook_risk`, a read-only tool fixed to the
source workbook passed at startup. That report surfaces workbook risk indicators
before a workflow trusts the imported WorkPaper; it does not certify Excel compatibility.

Use `@bilig/xlsx-formula-recalc` when the job is only to edit and recalculate
XLSX files. Use `@bilig/exceljs-formula-recalc` when an existing ExcelJS
workflow needs recalculated formula results after changing inputs.

## Tool Commands And Optional MCP

The npm tarball exposes the same CLI entrypoints through the canonical scoped
package, so tool hosts can install one focused package and still get the MCP
stdio server:

```ts
import { createWorkPaperMcpServer } from '@bilig/workpaper/mcp'
```

The source tree also maintains a starter project with `AGENTS.md`, MCP client
config, and an `agent:verify` script. Do not use
`npm create @bilig/workpaper@latest` while `@bilig/create-workpaper@latest`
resolves to `0.164.11`: that release's generated smoke reports
`formulasPersisted: false`. Use the evaluators above until a newer generator
release passes a fresh consumer smoke.

## Scope

Bilig is not a desktop Excel clone. It is a formula workbook runtime for
service-owned calculations, JSON persistence, XLSX import/export, and verified
readback. Unsupported Excel functions, external workbook links,
macros, and volatile functions may need review.

## After The Proof

If the starter or challenge output matches your service or tool workflow,
keep the repository nearby for release notes and public limits:
<https://github.com/proompteng/bilig>.

Watch releases if this is close to a production path:
<https://github.com/proompteng/bilig/subscription>.

If the model is close but blocked by a formula, import/export, persistence,
framework, MCP, or package-boundary gap, open the smallest implementation gap:
<https://github.com/proompteng/bilig/discussions/new?category=general>.

Full docs: <https://proompteng.github.io/bilig/>
