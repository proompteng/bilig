# bilig

[![MCP Toplist](https://mcptoplist.com/badge/io.github.proompteng%2Fbilig-workpaper.svg)](https://mcptoplist.com/server/io.github.proompteng%2Fbilig-workpaper)

[![CI](https://github.com/proompteng/bilig/actions/workflows/ci.yml/badge.svg)](https://github.com/proompteng/bilig/actions/workflows/ci.yml)
[![npm: @bilig/workpaper](https://img.shields.io/npm/v/@bilig/workpaper?label=%40bilig%2Fworkpaper)](https://www.npmjs.com/package/@bilig/workpaper)
[![CodeQL](https://github.com/proompteng/bilig/actions/workflows/codeql.yml/badge.svg)](https://github.com/proompteng/bilig/actions/workflows/codeql.yml)
[![OpenSSF Scorecard](https://api.scorecard.dev/projects/github.com/proompteng/bilig/badge)](https://scorecard.dev/viewer/?uri=github.com/proompteng/bilig)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

**Run workbook-shaped business rules inside Node.**

Bilig gives services, queue workers, tests, MCP servers, and tool integrations a
typed WorkPaper object: write inputs, recalculate formulas, read outputs,
persist JSON, restore, and verify. It fits pricing models, quote approval,
payout checks, import validation, forecasts, and formula-backed workflow steps.

<p align="center">
  <img src="docs/assets/github-social-preview.png" alt="bilig WorkPaper runtime preview" />
</p>

Run the no-project service check from any Node project:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door workpaper-service --json
```

Expected WorkPaper service result:

```json
{
  "schemaVersion": "bilig-evaluator.v1",
  "door": "workpaper-service",
  "verified": true,
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

For TypeScript services that should own the workbook model:

```sh
npm create @bilig/workpaper@latest pricing-workpaper
cd pricing-workpaper
npm install
npm run smoke
```

For MCP clients or other tool integrations, run the same proof loop through the
MCP evaluator before adding host-specific config:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
```

Evaluator examples live in
[`examples/bilig-evaluator-proof`](examples/bilig-evaluator-proof). Integration
matrix docs and host-specific config files are available when a team needs
them, but the public proof starts with WorkPaper service readback. File
compatibility diagnostics are a separate path for import/export boundaries, not
the default runtime story.

Project site: <https://proompteng.github.io/bilig/>

## Start Here

Pick the path that matches the job:

| You have...                                                      | Start with                                                             | You should see                                                                                         |
| ---------------------------------------------------------------- | ---------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------ |
| A Node service, route, queue, test, or tool needs workbook logic | [Node service WorkPaper evaluator](docs/eval-workpaper-service.md)     | input edit, recalculated output, serialized JSON, restore check, and `verified: true`.                 |
| An MCP client or tool integration needs workbook commands        | [MCP tool evaluator](docs/eval-agent-mcp.md)                           | tool discovery, cell edit, formula readback, export, restart check, and `verified: true`.              |
| You want a starter project with the runtime installed            | [90-second Node quickstart](docs/try-bilig-headless-in-node.md)        | a local package smoke test that edits one input, recalculates, saves JSON, and restores the WorkPaper. |
| An imported file is the integration boundary                     | [Workbook Compatibility Report](docs/workbook-compatibility-report.md) | unsupported functions, external links, macros, pivots, volatile formulas, and import/export risks.     |

If you are not sure which one fits, start with the thing that owns state. Use
WorkPaper when your service or tool should own the workbook model. Use file
diagnostics only when import/export compatibility is the actual contract.

Good fits: pricing, quote approval, payout checks, import validation, forecasts,
CI fixtures, formula-backed workflow steps, and tool integrations that need exact
cell addresses plus readback. Bad fits: manual spreadsheet editing, Office
macros, desktop Excel automation, or one-off arithmetic where a workbook would
be ceremony.

## If You Only Try One Thing

Run the WorkPaper service proof at the top of this README first. It is the
shortest proof that Bilig gives backend code a workbook object it can change,
recalculate, read back, save, and restore without driving Excel, LibreOffice,
Google Sheets, or a browser grid.

If an MCP client or tool integration owns the workflow, run the MCP door:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
```

Trust boundaries:

- Runs locally in Node or in your GitHub Actions runner; no hosted workbook
  upload is required.
- Does not claim Excel parity. Start with
  [where Bilig is not Excel-compatible yet](docs/where-bilig-is-not-excel-compatible-yet.md)
  before using it for irreversible workflows.
- File import/export diagnostics are available when an imported file is the
  contract, but they are a separate path from service-owned WorkPaper state.

## Which Path Should I Install?

| Problem you have right now                                                        | Install or use                 | First proof                                                            |
| --------------------------------------------------------------------------------- | ------------------------------ | ---------------------------------------------------------------------- |
| Formula workbook state belongs inside a Node service, route, queue, test, or tool | `npm install @bilig/workpaper` | [Node service WorkPaper evaluator](docs/eval-workpaper-service.md)     |
| An MCP client or tool integration needs workbook tools with computed readback     | `npm install @bilig/workpaper` | [MCP tool evaluator](docs/eval-agent-mcp.md)                           |
| Import/export compatibility is the integration boundary                           | compatibility diagnostics      | [Workbook Compatibility Report](docs/workbook-compatibility-report.md) |

Advanced adapters are still available when the boundary is already specific:
[SheetJS](docs/sheetjs-formula-result-not-updating-node.md),
[ExcelJS](docs/exceljs-formula-recalculation-node.md),
[external workbooks](docs/external-workbook-recalc-proof.md),
[MCP/tool integrations](docs/ai-agent-spreadsheet-tool-node.md),
[`@bilig/workbook`](docs/workbook-runtime-intent-api.md) when a runtime needs
transport-neutral plan data and command receipts, and
[runtime provenance](docs/npm-provenance-package-trust.md).

## MCP And Tool Integrations

Use the [WorkPaper host handoff](docs/agent-adoption-kit.md) when a tool host
needs workbook reads, writes, recalculation, JSON export, and restore proof.
The first check is always the no-key MCP evaluator:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
```

That evaluator starts the published WorkPaper tool server, discovers tools,
edits an input, reads the dependent formula, exports JSON, restarts, restores,
and returns `verified: true`. Use
[`llms-install.md`](llms-install.md) when a host wants one install file, and use
the [tool-host evaluator matrix](docs/agent-proof-matrix.md) when the host matters
more than the package boundary.

The published package also carries `AGENTS.md` and `SKILL.md` so hosts
inspecting `node_modules/@bilig/workpaper` can find the same proof locally.
Cloned checkouts keep host-specific config indexed in
[agent rule chooser](docs/agent-rule-chooser.md): `CLAUDE.md`,
`.claude/skills/bilig-workpaper/SKILL.md`,
`.claude/commands/bilig-workpaper-proof.md`,
`.cursor/rules/bilig-workpaper.mdc`, `.devin/rules/bilig-workpaper.md`,
`.windsurf/rules/bilig-workpaper.md`, `.clinerules/bilig-workpaper.md`,
`.continue/rules/bilig-workpaper.md`, `.zed/settings.json`, `opencode.jsonc`,
and `.opencode/agents/bilig-workpaper.md`. The public manifest is
`docs/.well-known/agent.json`.

```sh
npx --yes skills@latest add https://bilig.proompteng.ai --list
npx --yes skills@latest add proompteng/bilig --skill bilig-workpaper --list
```

## Integration Recipes After The Proof

Run one evaluator first. Then use the recipe that matches the platform boundary:

- Open WebUI WorkPaper MCP:
  <https://proompteng.github.io/bilig/open-webui-workpaper-mcp.html>.
- OpenAI Agents SDK with direct tools, stdio MCP, or
  `MCPServerStreamableHttp`:
  <https://proompteng.github.io/bilig/openai-agents-sdk-workpaper-tool.html>.
- ChatGPT Apps WorkPaper MCP:
  <https://proompteng.github.io/bilig/chatgpt-apps-workpaper-mcp.html>.
- n8n self-hosted workflows can use the owned community node package:
  `@bilig/n8n-nodes-workpaper`.

## Choose An Evaluation Path

| If you are evaluating...          | Start here                                                                                                                                                                                                                                                                                                   | What should be true before you adopt                                                                      |
| --------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ | --------------------------------------------------------------------------------------------------------- |
| Node service formulas             | [Node service WorkPaper evaluator](docs/eval-workpaper-service.md)                                                                                                                                                                                                                                           | A starter writes one input, recalculates, persists JSON, restores, and prints `verified: true`.           |
| MCP tool contract                 | [MCP workbook evaluator](docs/eval-agent-mcp.md)                                                                                                                                                                                                                                                             | MCP tool discovery, input edit, formula readback, persistence, and restart proof all pass.                |
| Integration proof chooser         | [Tool-host evaluator matrix](docs/agent-proof-matrix.md), [MCP spreadsheet tool server](docs/mcp-spreadsheet-formula-server-for-coding-agents.md), and [Vercel AI SDK formula readback](docs/vercel-ai-sdk-spreadsheet-tool-formula-readback.md) | The integration path starts with the smallest verified proof and avoids write-only or UI-only claims.     |
| Runtime intent adapters           | [Workbook runtime intent API](docs/workbook-runtime-intent-api.md)                                                                                                                                                                                                                                           | A model prepares transport-neutral plan data, strict runtime proof, command receipts, and check evidence. |
| Basic fit                         | [Why use Bilig?](docs/why-use-bilig.md)                                                                                                                                                                                                                                                                      | The problem is workbook-shaped business logic that needs API readback and persistence.                    |
| Published npm package             | [90-second Node quickstart](docs/try-bilig-headless-in-node.md)                                                                                                                                                                                                                                              | `@bilig/workpaper` edits one input, recalculates, persists JSON, restores, and prints `verified: true`.   |
| Backend service shape             | [Quote approval WorkPaper API](docs/quote-approval-workpaper-api.md)                                                                                                                                                                                                                                         | A realistic route-style workflow returns formula readback and `restoredMatchesAfter: true`.               |
| MCP clients and host integrations | [WorkPaper host handbook](docs/headless-workpaper-agent-handbook.md), [MCP spreadsheet tool server](docs/mcp-workpaper-tool-server.md), [Gemini CLI extension](docs/gemini-cli-workpaper-extension.md), and [Claude Desktop MCPB bundle](docs/claude-desktop-mcpb-workpaper.md)                              | The host installs a tool path, follows the handoff guide, then proves write/readback/persist.             |
| Technical WorkPaper review        | [WorkPaper maintainer proof note](docs/show-hn-formula-workbooks-node-services.md)                                                                                                                                                                                                                           | One compact page has the npm check, benchmark caveat, known limits, and open questions.                   |
| Trust and compatibility           | [npm provenance](docs/npm-provenance-package-trust.md), [compatibility limits](docs/where-bilig-is-not-excel-compatible-yet.md), and [Workbook Compatibility Report](docs/workbook-compatibility-report.md)                                                                                                  | npm shows SLSA provenance, and workbook behavior is checked through deterministic evaluator and XLSX gates. |
| Imported files                    | [Workbook Compatibility Report](docs/workbook-compatibility-report.md), [file formula recalculation](docs/xlsx-formula-recalculation-node.md), and [ExcelJS formula recalculation](docs/exceljs-formula-recalculation-node.md)                                                                               | The file boundary is inspected before a service, CI job, or workflow trusts imported formulas.            |
| Almost a fit                      | [implementation gap discussion](https://github.com/proompteng/bilig/discussions/new?category=general)                                                                                                                                                                                                        | Name the formula, import/export, persistence, framework, MCP, package, or benchmark gap.                  |
| Formula or import bug             | [formula bug clinic](docs/formula-bug-clinic.md) and [submit a workbook fixture](docs/submit-workbook-fixture.md)                                                                                                                                                                                            | Share one reduced public case that can become a fixture.                                                  |

Reduced workbook already in hand? If the blocker is an import, formula, or
persistence gap, generate the fixture report:

```sh
npm exec --package @bilig/workpaper@latest -- bilig-formula-clinic ./reduced.xlsx --cells "Summary!B7,Inputs!B2"
```

Handing a workbook task to an MCP client or host integration? Start with the
[host handoff guide](docs/headless-workpaper-agent-handbook.md#copy-paste-prompt-for-another-agent)
before opening Excel, LibreOffice, Google Sheets, or a screenshot UI. That
section keeps the host handoff prompt for clients that require copy-paste
instructions.
To prove the package-owned MCP loop without cloning the repo:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door workpaper-service --json
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
npm exec --package @bilig/workpaper@latest -- bilig-mcp-challenge --json
```

Tool hosts that support skill manifests can start from
[`skill.md`](docs/skill.md) or the well-known index at
[`docs/.well-known/agent-skills/index.json`](docs/.well-known/agent-skills/index.json).
Claude Desktop users can install the released MCPB bundle directly:
<https://github.com/proompteng/bilig/releases/latest/download/bilig-workpaper.mcpb>.
For host-specific project files, use the
[agent rule chooser](docs/agent-rule-chooser.md).

## Try It In 90 Seconds

This uses the published npm package. It builds a workbook, changes one input,
reads the calculated value, saves JSON, restores the workbook, and prints the
same value again.

```sh
npm create @bilig/workpaper@latest pricing-workpaper
cd pricing-workpaper
npm install
npm run smoke
```

Expected output includes these fields:

```json
{
  "before": {
    "summary": {
      "decision": "review"
    },
    "inputCells": {
      "units": "Inputs!B2",
      "listPrice": "Inputs!B3"
    }
  },
  "edit": {
    "before": {
      "decision": "review"
    },
    "after": {
      "decision": "approved"
    },
    "restored": {
      "decision": "approved"
    },
    "checks": {
      "decisionChanged": true,
      "formulasPersisted": true,
      "restoredMatchesAfter": true,
      "serializedBytes": 1242
    }
  },
  "verified": true
}
```

The generated starter uses the same WorkPaper fields as the
public mirror at <https://proompteng.github.io/bilig/npm-eval.ts> and
[`examples/headless-workpaper/npm-eval.ts`](examples/headless-workpaper/npm-eval.ts).
The exact byte count can change between package versions; `verified: true`,
`decisionChanged`, `formulasPersisted`, and `restoredMatchesAfter` are the
checks.

For a route-shaped quote approval API today, run the maintained example:

```sh
git clone --depth 1 https://github.com/proompteng/bilig.git
cd bilig
pnpm --dir examples/serverless-workpaper-api install --ignore-workspace
pnpm --dir examples/serverless-workpaper-api run smoke
```

For a generated project from a blank directory, run
`npm create @bilig/workpaper@latest pricing-workpaper` through the
`@bilig/create-workpaper` package. The package source lives in
[`packages/create-workpaper`](packages/create-workpaper), and the publish gate
is documented in [create a Bilig WorkPaper starter](docs/create-bilig-workpaper.md).
For an MCP-enabled project with host integration files, MCP client configs, and
an `agent:verify` script, run
`npm create @bilig/workpaper@latest pricing-agent -- --agent`.
Representative host files include `AGENTS.md`, `CLAUDE.md`, `GEMINI.md`,
`.claude/skills/bilig-workpaper/SKILL.md`, `.cursor/rules/bilig-workpaper.mdc`,
`.trae/mcp.json`, and `.zed/settings.json`.
For an existing repo, run
`npm create @bilig/workpaper@latest . -- --add-agent`; it adds Bilig MCP and
host instructions without replacing your app template or editing
`package.json`. If a host policy already exists, it writes
`BILIG_WORKPAPER_INSTALL.md` with the skipped paths and a short handoff block.

If that proof almost matches a service or integration workflow you maintain, the
useful next step is a concrete gap report in
[Discussions](https://github.com/proompteng/bilig/discussions/new?category=general):
formula coverage, service persistence, MCP setup, agent writeback, import/export
boundary, or benchmark coverage.

## TypeScript API Shape

Most integrations are just this: build a workbook, write an input, read the
calculated value, and save the workbook state. When a workflow writes more than
one input, use `editManyAndReadback()` so the edits are applied atomically and
the proof compares typed readback values, persisted restore output, and formula
diagnostics.

```ts
import { buildA1WorkPaper } from '@bilig/workpaper'

const book = buildA1WorkPaper({
  Inputs: [
    ['Metric', 'Value'],
    ['Customers', 20],
    ['Average revenue', 1200],
  ],
  Summary: [
    ['Metric', 'Value'],
    ['Revenue', '=Inputs!B2*Inputs!B3'],
  ],
})

const proof = book.editAndReadback('Inputs!B2', 32, {
  readbackRange: 'Summary!B2',
})

console.log({
  editedCell: proof.editedCell,
  revenue: proof.afterReadback.displayValues[0]?.[0],
  persistedDocumentBytes: proof.persistedDocumentBytes,
  verified: proof.verified,
})

book.dispose()
```

The lower-level `WorkPaper` runtime is still exported for engine integrations,
but the A1 facade is the default service and agent path. Use
`book.set()`, `book.setMany()`, `book.readMany()`, `book.display()`, and
`book.saveJson()` when a full readback proof is not needed. Use
`book.editManyAndReadback()` when several inputs should be committed and proven
as one atomic workbook edit.

## When To Reach For It

Use `@bilig/workpaper` when:

- a Node service owns a workbook-shaped calculation;
- an agent needs tools such as `readRange` and `setInputCell`, with computed
  before/after values instead of screenshots;
- tests need deterministic spreadsheet state and formula readback;
- a workflow needs to save the edited workbook as JSON and restore it later.

Use something else when you need a visual spreadsheet grid, Office macros,
desktop Excel automation, or a one-off arithmetic helper. Do not treat embedded
XLSX stored formula results as truth; use the Excel oracle workflow when
accuracy matters.

## Package Boundary

Current checked npm metadata for `@bilig/workpaper@latest`:

- Published package: `57.7 kB` unpacked, `49` package entries.
- Boundary: the public package owns WorkPaper starters, evaluators, MCP command
  wrappers, formula clinic reports, JSON persistence, and restored readback.
- Runtime: Node `>=22.0.0`; Node 22 compatibility is covered by the runtime
  package workflow.

## Published Package Trust

`@bilig/workpaper` is published with npm registry signatures and SLSA provenance
attestations. Verify the package version you are about to adopt:

```sh
npm view @bilig/workpaper version dist.attestations dist.signatures --json
```

After installing, npm can verify the current dependency tree:

```sh
npm audit signatures
```

The current package trust path is documented in
[npm provenance and package trust](docs/npm-provenance-package-trust.md).
Repository security posture is tracked by
[OpenSSF Scorecard](https://scorecard.dev/viewer/?uri=github.com/proompteng/bilig)
and uploaded to GitHub code scanning on every `main` update.

## Deeper Evaluation Paths

After the first proof in [Start Here](#start-here), use the deeper guide that
matches the next job.

1. Run the [90-second npm eval](#try-it-in-90-seconds) in a blank project.
2. Run the flagship
   [serverless WorkPaper API](examples/serverless-workpaper-api) example:
   `npm run quote-approval-api`.
3. If the workflow starts with a saved workbook file, run the
   [XLSX formula recalculation in Node](examples/xlsx-recalculation-node):
   `npm start`.
4. If a tool host needs workbook tools, start with the
   [headless WorkPaper host handbook](docs/headless-workpaper-agent-handbook.md),
   then use the [MCP server guide](docs/mcp-workpaper-tool-server.md) when the
   caller is an MCP client.
5. If a real workbook almost works, start with the
   [formula bug clinic](docs/formula-bug-clinic.md). Include the exact cell,
   expected value, actual value, and command output.
   If the fixture is already reduced, submit the structured
   [fixture form](docs/submit-workbook-fixture.md) so the blocker can become a
   test, example, or corpus case instead of private feedback.
   <https://github.com/proompteng/bilig/issues/new?template=workbook_fixture.yml>.
   If you are still reducing the case, discuss the shape first:
   <https://github.com/proompteng/bilig/discussions/414>.

The rest of the docs are an index, not a prerequisite.

For comparison and integration details, use the
[plain-language fit guide](docs/why-use-bilig.md),
[screenshot automation boundary](docs/stop-driving-spreadsheets-with-screenshots.md),
[Google Sheets API boundary](docs/google-sheets-api-alternative-node-workpaper.md),
[Google Sheets QUERY/SORTN in Node](docs/google-sheets-query-sortn-node-workpaper.md),
[workbook automation examples](docs/workbook-automation-examples-node.md),
the [formula workbooks proof page](docs/formula-workbooks-node-services-agent-tools.md),
the [Node spreadsheet formula engine guide](docs/node-spreadsheet-formula-engine.md),
[server-side spreadsheet automation](docs/server-side-spreadsheet-automation-node.md),
[framework adapters](docs/node-framework-workpaper-adapters.md),
[formula bug clinic](docs/formula-bug-clinic.md),
[workbook fixture submissions](docs/submit-workbook-fixture.md),
[OpenAI Agents SDK tools](docs/openai-agents-sdk-workpaper-tool.md),
[Browser Use formula tool](docs/browser-use-workpaper-formula-tool.md),
[OpenHands MCP setup](docs/openhands-workpaper-mcp.md),
[OpenCode MCP setup](docs/opencode-workpaper-mcp.md),
[tool-host evaluator matrix](docs/agent-proof-matrix.md),
[MCP spreadsheet formula server for tool hosts](docs/mcp-spreadsheet-formula-server-for-coding-agents.md),
[Vercel AI SDK formula readback](docs/vercel-ai-sdk-spreadsheet-tool-formula-readback.md),
[CrewAI adapter](docs/crewai-workpaper-spreadsheet-tool.md),
the [WorkPaper host handbook](docs/headless-workpaper-agent-handbook.md),
the [MCP server guide](docs/mcp-workpaper-tool-server.md),
[spreadsheet MCP server comparison](docs/spreadsheet-mcp-server-comparison.md),
[MCP directory status](docs/mcp-spreadsheet-server-directory.md),
[MCP client setup](docs/mcp-client-setup.md),
[Gemini CLI extension](docs/gemini-cli-workpaper-extension.md),
[Claude Desktop MCPB bundle](docs/claude-desktop-mcpb-workpaper.md),
[npm provenance and package trust](docs/npm-provenance-package-trust.md),
[JavaScript library comparison](docs/javascript-spreadsheet-library-headless-node.md),
[Node spreadsheet formula engine guide](docs/node-spreadsheet-formula-engine.md),
[server-side spreadsheet automation](docs/server-side-spreadsheet-automation-node.md),
[saved-workbook formula recalculation](docs/xlsx-formula-recalculation-node.md),
[XLSX formula support answers](docs/xlsx-formula-support-answers.md),
[SheetJS/ExcelJS boundary](docs/sheetjs-exceljs-alternative-formula-workbook-api.md),
[ExcelJS formula result boundary](docs/exceljs-formula-result-not-updating-after-node-edits.md),
[Microsoft Graph Excel boundary](docs/microsoft-graph-excel-recalculation-node.md),
and [engine comparison](docs/headless-spreadsheet-engine-comparison.md).

Useful deeper examples: [invoice totals](examples/headless-workpaper#invoice-totals),
[budget variance alerts](examples/headless-workpaper#budget-variance-alerts),
[fulfillment capacity plan](examples/headless-workpaper#fulfillment-capacity-plan),
[quote approval threshold](examples/headless-workpaper#quote-approval-threshold),
[subscription MRR forecast](examples/headless-workpaper#subscription-mrr-forecast),
[agent framework adapters](examples/headless-workpaper#agent-framework-adapters),
[MCP tool server shape](examples/headless-workpaper#mcp-tool-server-shape),
[XLSX formula recalculation in Node](examples/xlsx-recalculation-node),
and [serverless quote approval](examples/serverless-workpaper-api). Run
`npm run quote-approval-api`, `npm run agent:openai-agents-sdk`,
`npm run agent:framework-adapters`,
`npm run agent:mcp-tools`, `npm run agent:mcp-transcript`,
`npm run agent:mcp-file-transcript`, `npm run agent:mcp-stdio`, or
`npm exec --package @bilig/workpaper -- bilig-workpaper-mcp` when that is the
path you are evaluating.

Saved workbook diagnostics stay available when a file is the integration
boundary:

```sh
npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- bilig-evaluate --door workbook-compatibility --json
npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- workbook-compatibility-report workbook.xlsx --json
npm exec --package @bilig/xlsx-formula-recalc@latest -- xlsx-recalc --demo --json
npm exec --package @bilig/sheetjs-formula-recalc@latest -- sheetjs-recalc --demo --json
```

The serverless example also includes `npm run next-route-handler`,
`npm run next-server-action`, `npm run next-server-action-formdata`,
`npm run hono-route`, `npm run framework-adapters`, and
`npm run persistence-adapters` for
framework-specific boundary checks.

The MCP server is also listed in the official registry:
<https://registry.modelcontextprotocol.io/v0.1/servers?search=io.github.proompteng%2Fbilig-workpaper>.
Clients that support Streamable HTTP MCP can also smoke-test the stateless
hosted endpoint at `https://bilig.proompteng.ai/mcp`; use the local stdio server
when the agent needs to persist a project WorkPaper JSON file.

## Examples You Can Run

The runnable examples are TypeScript files. Some source imports end in `.js`
because Node ESM resolves compiled package output that way; the files you edit
and run are still `.ts`.

From a cloned checkout:

```sh
pnpm --dir examples/headless-workpaper install --ignore-workspace
pnpm --dir examples/headless-workpaper run start
pnpm --dir examples/headless-workpaper run json-records
pnpm --dir examples/headless-workpaper run csv-shaped
pnpm --dir examples/headless-workpaper run invoice-totals
pnpm --dir examples/headless-workpaper run budget-variance
pnpm --dir examples/headless-workpaper run fulfillment-capacity
pnpm --dir examples/headless-workpaper run quote-approval
pnpm --dir examples/headless-workpaper run subscription-mrr
pnpm --dir examples/headless-workpaper run persistence
```

The most useful entry points:

- [JSON records input](examples/headless-workpaper#json-records-input)
- [CSV shaped input](examples/headless-workpaper#csv-shaped-input)
- [invoice totals](examples/headless-workpaper#invoice-totals)
- [budget variance alerts](examples/headless-workpaper#budget-variance-alerts)
- [fulfillment capacity plan](examples/headless-workpaper#fulfillment-capacity-plan)
- [quote approval threshold](examples/headless-workpaper#quote-approval-threshold)
- [subscription MRR forecast](examples/headless-workpaper#subscription-mrr-forecast)
- [SheetJS, xlsx-populate, and ExcelJS recalculation bridge](examples/recalc-bridge-workflows)

For tool integrations:

```sh
pnpm --dir examples/headless-workpaper run agent:verify
pnpm --dir examples/headless-workpaper run agent:tool-call
pnpm --dir examples/headless-workpaper run agent:openai-agents-sdk
pnpm --dir examples/headless-workpaper run agent:openai-agents-sdk-mcp
pnpm --dir examples/headless-workpaper run agent:openai-agents-sdk-hosted-mcp
pnpm --dir examples/headless-workpaper run agent:openai-responses
pnpm --dir examples/headless-workpaper run agent:ai-sdk-generate-text
pnpm --dir examples/headless-workpaper run agent:ai-sdk-stream-text
pnpm --dir examples/headless-workpaper run agent:framework-adapters
pnpm --dir examples/serverless-workpaper-api run hono-route
pnpm --dir examples/headless-workpaper run agent:mcp-tools
pnpm --dir examples/headless-workpaper run agent:mcp-file-transcript
pnpm --dir examples/headless-workpaper run agent:mcp-xlsx-risk-preflight
pnpm --dir examples/headless-workpaper run agent:mcp-stdio
```

The AI SDK example uses
[`ai-sdk-generate-text-tool-smoke.ts`](examples/headless-workpaper/ai-sdk-generate-text-tool-smoke.ts).
The OpenAI Agents SDK guide is
[`docs/openai-agents-sdk-workpaper-tool.md`](docs/openai-agents-sdk-workpaper-tool.md).
It includes direct `tool()` wrapping, private `MCPServerStdio` discovery, and
hosted stateless `MCPServerStreamableHttp` discovery through the WorkPaper MCP
tool loop.
The ChatGPT Apps Developer Mode setup is
[`docs/chatgpt-apps-workpaper-mcp.md`](docs/chatgpt-apps-workpaper-mcp.md).
It shows the public `/mcp` endpoint as a data/tool-only remote MCP app and keeps
custom Apps SDK component UI as future scope.
The OpenAI Responses guide is
[`docs/openai-responses-workpaper-tool-call.md`](docs/openai-responses-workpaper-tool-call.md).
The retained agent framework examples live in `examples/headless-workpaper` and
are covered by its local smoke scripts.

The package also ships the MCP stdio binary:

```sh
npm exec --package @bilig/workpaper@latest -- bilig-agent-challenge --json
npm exec --package @bilig/workpaper@latest -- bilig-formula-clinic ./reduced.xlsx --cells "Summary!B7,Inputs!B2"
npm exec --package @bilig/workpaper@latest -- bilig-mcp-challenge --json
npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp
npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable
npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp --from-xlsx ./pricing.xlsx
npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp --from-xlsx ./pricing.xlsx --workpaper ./.bilig/pricing.workpaper.json --writable
pnpm --dir examples/headless-workpaper run agent:mcp-xlsx-risk-preflight
docker build --target bilig-workpaper-mcp -t bilig-workpaper-mcp:local .
```

`bilig-agent-challenge` prints the same edit, formula readback, WorkPaper JSON
export, restore, and `verified: true` proof object used by the agent workbook
challenge page.

`bilig-mcp-challenge` proves the file-backed MCP path end to end: initialize
JSON-RPC, list tools/resources/prompts, edit `Inputs!B3`, read recalculated
`Summary!B3`, export the WorkPaper JSON, restart from disk, and return
`verified: true`.

`bilig-formula-clinic` imports a reduced XLSX locally, samples formulas, reads
requested cells through WorkPaper, and prints a Markdown issue body. It does not
upload workbook contents.

Without `--workpaper`, the binary starts the built-in demo workbook. With
`--workpaper`, it loads your persisted WorkPaper JSON and exposes
`list_sheets`, `read_range`, `read_cell`, `set_cell_contents`,
`set_cell_contents_and_readback`, `get_cell_display_value`,
`export_workpaper_document`, and `validate_formula`; `--writable` persists
`set_cell_contents` or `set_cell_contents_and_readback` edits back to the same
file. If you already have an XLSX, `--from-xlsx ./pricing.xlsx` imports it into
an in-memory WorkPaper server for readback, throwaway edits, and
`analyze_workbook_risk` without writing a sidecar. Add `--workpaper ... --writable`
only when the agent needs persisted file state. It also
exposes MCP resources and prompts for `bilig://workpaper/agent-handoff`,
`bilig://workpaper/current-document`, `edit_and_verify_workpaper`, and
`debug_workpaper_formula`, so capable clients can discover the workflow before
calling tools.
For a maintained real-XLSX transcript, run
`pnpm --dir examples/headless-workpaper run agent:mcp-xlsx-risk-preflight`.
It calls `analyze_workbook_risk`, edits `Inputs!B3`, verifies `Summary!B3`
changes from `60000` to `96000`, exports WorkPaper JSON, and keeps
`excelParity: "not_proven"`.
The Docker target is for MCP directory scanners: it seeds a demo WorkPaper JSON
inside the image and starts the file-backed `--writable` tool surface so
`tools/list`, `resources/list`, and `prompts/list` return the general WorkPaper
agent surface without cloning this monorepo. For remote MCP clients, the app
runtime exposes `https://bilig.proompteng.ai/mcp` as a stateless JSON-only
Streamable HTTP endpoint for tool discovery and write/readback smoke tests.

It is published in the official MCP Registry as
`io.github.proompteng/bilig-workpaper`:
<https://registry.modelcontextprotocol.io/v0.1/servers?search=io.github.proompteng%2Fbilig-workpaper>.
It is also live on Glama with `Try in Browser`, A-grade tool pages, and the
file-backed WorkPaper tools:
<https://glama.ai/mcp/servers/proompteng/bilig>.

## Proof You Can Reproduce

- The 90-second TypeScript check above edits one input, restores the saved JSON
  document, and verifies the dependent formula result.
- For a service evaluator path, run the
  [quote approval WorkPaper API proof](docs/quote-approval-workpaper-api.md).
  It starts from an empty Node directory, downloads one maintained TypeScript
  route smoke, writes quote inputs, recalculates an approval decision, persists
  JSON, and verifies restored readback.
- For a shorter evaluation page, read
  [formula workbooks for Node services and tool integrations](docs/formula-workbooks-node-services-agent-tools.md).
  It compresses the WorkPaper boundary, MCP file-backed mode, benchmark caveat,
  and alternative-tool guidance into one evaluation path.
- For a compact review note, use the
  [WorkPaper maintainer proof note](docs/show-hn-formula-workbooks-node-services.md).
  It keeps the empty npm-project command, `verified: true` output, benchmark
  caveat, known limits, and open questions together.
- For saved-file integration, run the XLSX formula recalculation example:
  [`examples/xlsx-recalculation-node`](examples/xlsx-recalculation-node). It
  imports a generated XLSX pricing workbook, edits input cells, reads the
  recalculated approval decision, exports XLSX, reimports it, and verifies the
  formulas survived the round trip. The public decision page is
  [XLSX formula recalculation in Node.js](docs/xlsx-formula-recalculation-node.md).
- Read the [compatibility limits](docs/where-bilig-is-not-excel-compatible-yet.md)
  before importing real Excel workbooks.
- Use the
  [production adoption checklist](docs/production-adoption-checklist-headless-workpaper.md)
  before promoting a WorkPaper-backed workflow beyond evaluation.
- For XLSX accuracy audits, use the
  [Excel oracle harness](docs/xlsx-corpus-verifier-walkthrough.md#run-the-excel-oracle-harness).
  It separates import success, timeouts, cached workbook values, and fresh
  Microsoft Excel recalculation results.
- The WorkPaper MCP server is listed in the
  [official MCP Registry](https://registry.modelcontextprotocol.io/v0.1/servers?search=io.github.proompteng%2Fbilig-workpaper)
  and on [Glama](https://glama.ai/mcp/servers/proompteng/bilig). The
  [directory status page](docs/mcp-spreadsheet-server-directory.md) keeps the
  npm command, remote endpoint, static MCP server card, and directory evidence
  in one place.
- Public feedback threads:
  [workflow questions](https://github.com/proompteng/bilig/discussions/157),
  [service examples](https://github.com/proompteng/bilig/discussions/213),
  [persistence adapters](https://github.com/proompteng/bilig/discussions/307),
  [JavaScript spreadsheet library guide](https://github.com/proompteng/bilig/discussions/308),
  [OpenAI Responses tool calls](https://github.com/proompteng/bilig/discussions/335),
  and [benchmark critique](https://github.com/proompteng/bilig/discussions/340).

If you are evaluating Bilig runtime packages for production and want release
notifications, watch releases:
<https://github.com/proompteng/bilig/subscription>.

## XLSX Accuracy Policy

Cached formula values embedded in `.xlsx` files are cache diagnostics, not an
accuracy verdict. A Bilig correctness bug should only be claimed when the
expected value came from a fresh Excel recalculation oracle.

```sh
OUT=.cache/excel-oracle-evaluation
pnpm workpaper:xlsx-oracle -- prepare-oracle /path/to/xlsx-corpus "$OUT"
pnpm workpaper:xlsx-oracle -- evaluate-cache /path/to/xlsx-corpus "$OUT"
pnpm workpaper:xlsx-oracle -- evaluate-oracle /path/to/xlsx-corpus "$OUT/recalculated" "$OUT"
pnpm workpaper:xlsx-oracle -- summarize "$OUT"
```

`evaluate-cache` writes `cache-diagnostic.json` and stays non-authoritative.
`evaluate-oracle` writes `excel-oracle-report.json`, and `summarize` writes
`summary.md`. If Excel automation is unavailable, cells are classified as
`missing_excel_oracle` instead of being promoted to bugs.

## What Is In This Repo

- `packages/workpaper`: public WorkPaper package, starters, evaluator binaries,
  and MCP wrappers.
- `packages/headless`: lower-level WorkPaper runtime that backs the public
  package.
- `packages/excel-import`: saved-workbook import/export boundary.
- `packages/formula`: formula parser, binder, compiler, and evaluator.
- `packages/core`: workbook engine, snapshots, mutation flow, and scheduler.
- `packages/grid` and `apps/web`: browser spreadsheet shell.
- `apps/bilig`: fullstack monolith runtime, API surface, and static asset
  server.
- `packages/renderer`: React workbook renderer.
- `packages/protocol`, `packages/binary-protocol`, `packages/agent-api`, and
  `packages/worker-transport`: protocol and integration boundaries.
- `packages/wasm-kernel`: AssemblyScript/WASM numeric fast path.
- `packages/benchmarks`: benchmark harness and performance contracts.

For XLSX import/export from TypeScript:

```ts
import { WorkPaper } from '@bilig/workpaper'
import { exportXlsx, importXlsx } from '@bilig/workpaper/xlsx'
```

Use `WorkPaper.buildFromSnapshot(imported.snapshot)` after import and
`workbook.exportSnapshot()` before `exportXlsx()`.

## Local Development

Use Node `24+`, Bun, and `pnpm@10.32.1`.

```sh
pnpm install
pnpm dev:web
pnpm dev:web-local
pnpm dev:sync
```

For a full local preflight:

```sh
pnpm lint
pnpm typecheck
pnpm test
pnpm test:browser
pnpm run ci
```

Generated sources and deterministic public docs are checked:

```sh
pnpm protocol:check
pnpm formula-inventory:check
pnpm workspace-resolution:check
pnpm docs:discovery:check
```

## For Coding Agents

Start with the public package boundary unless the task is explicitly engine
work.

1. Read `packages/workpaper/README.md` before touching public WorkPaper behavior.
2. Read `docs/AGENTS.md`, `docs/skill.md`, or `docs/llms-full.txt` when
   building an agent-facing integration from outside the repo.
3. Use public exports from `@bilig/workpaper`; do not reach into `src/` or
   `dist/` when writing consumer examples.
4. Keep examples TypeScript-first.
5. Do not call embedded XLSX stored formula results an accuracy oracle.
6. Add focused tests before changing formulas, persistence, range bounds,
   config rebuilds, events, row/column moves, or sheet lifecycle.
7. Run the focused package tests first, then broaden to `pnpm run ci`.

## Contributing

Read [CONTRIBUTING.md](CONTRIBUTING.md) before opening a PR. If this is your
first patch, start with the
[new contributor guide](docs/new-contributor-guide.md) and then claim a scoped
starter issue.

Good first patches usually fit one of these shapes:

- formula fixtures with clear expected behavior;
- small WorkPaper examples that prove a real service or agent workflow;
- focused correctness fixes with regression tests;
- grid accessibility and keyboard-behavior improvements;
- docs that turn an existing architecture note into a runnable command.

The shortest public on-ramp is the
[`starter issues`](docs/starter-issues.md) queue. It keeps code/test picks,
example tasks, adapters, and focused docs work in one current list, with small
acceptance commands for first patches.

If this is your first contribution to `bilig`, use the
[`first-timers-only`](https://github.com/proompteng/bilig/issues?q=is%3Aissue%20state%3Aopen%20label%3Afirst-timers-only)
filter.

## Security And Support

Read [SECURITY.md](SECURITY.md) before sharing vulnerability details, private
workbook data, tokens, credentials, or exploit reproductions. Security reports
should use GitHub private vulnerability reporting when available, or
<security@proompteng.ai> when the private flow is not visible.

Use [SUPPORT.md](SUPPORT.md) for the fastest public support path. Good reports
include the package version, Node version, OS, exact formula or workbook input,
expected value, actual value, and the smallest command or script that reproduces
the issue.

## CI

Forgejo Actions is the primary CI surface via
`.forgejo/workflows/forgejo-ci.yml`. GitHub Actions mirrors the verification
contract in `.github/workflows/ci.yml`.

The strict gate includes frozen lockfile install, full `pnpm run ci`, artifact
budget checks, browser smoke, and tracked-file cleanliness checks.

## License

MIT.
