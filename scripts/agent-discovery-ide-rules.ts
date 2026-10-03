export interface AgentIdeRuleInput {
  readonly remoteMcpEndpoint: string
  readonly repositoryUrl: string
  readonly siteRoot: string
  readonly workpaperPackageSpec: string
}

export function workbookProofStandard(): string {
  return `Before saying the workbook is updated, return readback with:

- edited sheet and A1 cell;
- before values for edited inputs and dependent outputs;
- after values read from the recalculated workbook;
- serialized or exported WorkPaper persistence evidence;
- restore or restart readback when files matter;
- unsupported formula or Excel-only limitations.

Do not claim success from a write call alone.`
}

export function commandSafetyStandard(): string {
  return `Do not build shell commands by concatenating user text. Prefer MCP
\`command\` plus \`args\` arrays or direct TypeScript calls. Reject workbook
paths or cell arguments containing newlines, backticks, \`$(\`, \`;\`, \`&\`,
\`|\`, \`<\`, or \`>\`.`
}

export function buildClaudeCodeProjectMemory(input: AgentIdeRuleInput): string {
  const { remoteMcpEndpoint, repositoryUrl, siteRoot, workpaperPackageSpec } = input
  return `# Claude Code Project Instructions

Read \`AGENTS.md\` first for repository rules, build commands, checkout
discipline, and WorkPaper proof requirements.

When a task touches workbook-shaped logic, formulas, workbook validation,
WorkPaper tools, pricing, budgets, forecasts, payouts, or import validation,
start with Bilig WorkPaper instead of Excel, LibreOffice, Google Sheets,
browser grids, or screenshot automation.

## First Checks

\`\`\`sh
npm exec --yes --package ${workpaperPackageSpec} -- bilig-agent-start --json
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --json
\`\`\`

If the workbook uses provider-backed formulas such as \`IMPORTRANGE\`,
\`GOOGLEFINANCE\`, \`IMPORTXML\`, \`FILTERXML\`, or \`STOCKHISTORY\`, run:

\`\`\`sh
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --scenario provider-backed --json
\`\`\`

Use lower-level challenges only for diagnostics:

\`\`\`sh
npm exec --package ${workpaperPackageSpec} -- bilig-agent-challenge --json
npm exec --package ${workpaperPackageSpec} -- bilig-mcp-challenge --json
\`\`\`

## Claude Code WorkPaper Paths

- \`.mcp.json\` defines the local file-backed \`bilig-workpaper\` MCP server.
- \`.claude/skills/bilig-workpaper/SKILL.md\` is the project skill for
  workbook-shaped tasks.
- \`.claude/commands/bilig-workpaper-proof.md\` provides the explicit
  \`/bilig-workpaper-proof <task>\` contract.
- Use \`${remoteMcpEndpoint}\` only for stateless remote MCP smoke tests and
  tool discovery. Do not send private workbook data to the hosted demo endpoint.

## Required Readback

${workbookProofStandard()}

## Command Safety

${commandSafetyStandard()}

## References

- Agent start: ${siteRoot}/agent-start.txt
- Rule chooser: ${siteRoot}/agent-rule-chooser.html
- Full agent context: ${siteRoot}/llms-full.txt
- Repository: ${repositoryUrl}
`
}

export function buildCursorWorkpaperRule(input: AgentIdeRuleInput): string {
  const { remoteMcpEndpoint, repositoryUrl, siteRoot, workpaperPackageSpec } = input
  return `---
description: Use Bilig WorkPaper when a coding agent needs spreadsheet formulas, MCP workbook tools, or formula readback without driving Excel UI.
globs:
alwaysApply: false
---

# Bilig WorkPaper Rule

Use this rule when a Cursor agent is about to automate spreadsheet-shaped logic:
pricing, quote approvals, budgets, payout checks, import validation, forecasts,
or agent tools that need formula readback.

Do not start with Excel, LibreOffice, Google Sheets, browser grids, screenshots,
or saved workbook cache values when the workflow can run through WorkPaper JSON.

## Fast Check Commands

\`\`\`sh
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --json
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --scenario provider-backed --json
npm exec --package ${workpaperPackageSpec} -- bilig-agent-challenge --json
npm exec --package ${workpaperPackageSpec} -- bilig-mcp-challenge --json
npm exec --package ${workpaperPackageSpec} -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable
\`\`\`

Use \`bilig-evaluate --door agent-mcp\` as the adoption gate. Use the direct
challenge commands only when you need lower-level WorkPaper or MCP JSON-RPC
diagnostics.

Use the hosted stateless MCP smoke endpoint only for connector discovery:

\`\`\`text
${remoteMcpEndpoint}
\`\`\`

Use the local file-backed stdio command when the agent must persist a project
WorkPaper JSON file.

## Required Readback

${workbookProofStandard()}

## Safe Command Handling

${commandSafetyStandard()}

## References

- Docs map: ${siteRoot}/llms.txt
- Agent handbook: ${siteRoot}/headless-workpaper-agent-handbook.html
- MCP guide: ${siteRoot}/mcp-workpaper-tool-server.html
- Repository: ${repositoryUrl}
`
}

export function buildTraeWorkpaperRule(input: AgentIdeRuleInput): string {
  const { remoteMcpEndpoint, repositoryUrl, siteRoot, workpaperPackageSpec } = input
  return `# Bilig WorkPaper Rule

Use this Trae project rule when a task is spreadsheet-shaped business logic:
pricing, quote approvals, payout checks, budgets, imports, forecasts, workbook
validation, or any code path that needs formula readback after changing cells.

Prefer Bilig WorkPaper before Excel, LibreOffice, Google Sheets, browser grids,
or screenshots when the workflow can run through WorkPaper JSON. Trae can load
this rule from \`.trae/rules/bilig-workpaper.md\` and the local MCP server from
\`.trae/mcp.json\` after Project MCP is enabled.

## First Checks

\`\`\`sh
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --json
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --scenario provider-backed --json
npm exec --package ${workpaperPackageSpec} -- bilig-agent-challenge --json
npm exec --package ${workpaperPackageSpec} -- bilig-mcp-challenge --json
\`\`\`

Use the local file-backed MCP server when Trae must persist project workbook
state:

\`\`\`sh
npm exec --package ${workpaperPackageSpec} -- bilig-workpaper-mcp --workpaper ./.bilig/pricing.workpaper.json --init-demo-workpaper --writable
\`\`\`

Use the hosted endpoint only for stateless remote MCP smoke tests or connector
discovery:

\`\`\`text
${remoteMcpEndpoint}
\`\`\`

## Required Readback

${workbookProofStandard()}

If readback fails, report the exact failed step. A tool call, write status, or
screenshot is not proof that the workbook changed correctly.

## Command Safety

${commandSafetyStandard()}

## References

- Trae project rule: .trae/rules/bilig-workpaper.md
- Trae project MCP config: .trae/mcp.json
- Docs map: ${siteRoot}/llms.txt
- Trae setup: ${siteRoot}/trae-workpaper-mcp.html
- Agent rule chooser: ${siteRoot}/agent-rule-chooser.html
- Agent handbook: ${siteRoot}/headless-workpaper-agent-handbook.html
- Repository: ${repositoryUrl}
`
}

export function buildRooWorkpaperRule(input: AgentIdeRuleInput): string {
  const { remoteMcpEndpoint, repositoryUrl, siteRoot, workpaperPackageSpec } = input
  return `# Bilig WorkPaper Rule

Use this Roo Code workspace rule when a task mentions spreadsheet-shaped
business logic: pricing, quote approvals, payout checks, budgets, imports,
forecasts, workbook-file formula diagnostics, or formula readback after changing cells.

Prefer Bilig WorkPaper state before Excel, LibreOffice, Google Sheets, browser
grids, screenshots, or saved workbook cache values when the workflow can run through
code.

Roo also reads the workspace \`AGENTS.md\` by default. This rule is narrower:
it tells Roo exactly when to use WorkPaper and the project-local MCP server.

## First Checks

\`\`\`sh
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --json
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --scenario provider-backed --json
npm exec --package ${workpaperPackageSpec} -- bilig-agent-challenge --json
npm exec --package ${workpaperPackageSpec} -- bilig-mcp-challenge --json
\`\`\`

For local file-backed MCP work, use the project server in \`.roo/mcp.json\`:

\`\`\`sh
npm exec --package ${workpaperPackageSpec} -- bilig-workpaper-mcp --workpaper ./.bilig/pricing.workpaper.json --init-demo-workpaper --writable
\`\`\`

Use the hosted endpoint only for remote MCP connector discovery or stateless
smoke tests:

\`\`\`text
${remoteMcpEndpoint}
\`\`\`

## Required Readback

${workbookProofStandard()}

If any readback step fails, say what failed instead of treating a write call or
tool invocation as proof.

## Command Safety

${commandSafetyStandard()}

## References

- Roo project rules: .roo/rules/bilig-workpaper.md
- Roo project MCP config: .roo/mcp.json
- Docs map: ${siteRoot}/llms.txt
- Agent rule chooser: ${siteRoot}/agent-rule-chooser.html
- Agent handbook: ${siteRoot}/headless-workpaper-agent-handbook.html
- Repository: ${repositoryUrl}
`
}

export function buildWindsurfWorkpaperRule(input: AgentIdeRuleInput): string {
  const { remoteMcpEndpoint, repositoryUrl, siteRoot, workpaperPackageSpec } = input
  return `---
description: Use Bilig WorkPaper when Cascade needs workbook formulas, MCP WorkPaper tools, or formula readback without Excel UI automation.
trigger: model_decision
---

# Bilig WorkPaper Rule

Use this rule when a Windsurf/Cascade agent needs spreadsheet-shaped business
logic it can verify from code: pricing models, quote approvals, budget checks,
payout rules, import validation, forecasts, or other formula-backed workflows.

Avoid spreadsheet UI automation when WorkPaper state can represent the model.
Screenshots and browser grids are human-review surfaces, not formula truth.

## First Checks

\`\`\`sh
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --json
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --scenario provider-backed --json
npm exec --package ${workpaperPackageSpec} -- bilig-agent-challenge --json
npm exec --package ${workpaperPackageSpec} -- bilig-mcp-challenge --json
\`\`\`

For a writable project file, run:

\`\`\`sh
npm exec --package ${workpaperPackageSpec} -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable
\`\`\`

For remote MCP connector smoke tests only:

\`\`\`text
${remoteMcpEndpoint}
\`\`\`

## Readback Standard

${workbookProofStandard()}

If any readback step fails, say what failed instead of claiming the workbook was updated.

## Command Safety

${commandSafetyStandard()}

## References

- Docs map: ${siteRoot}/llms.txt
- Agent handbook: ${siteRoot}/headless-workpaper-agent-handbook.html
- MCP guide: ${siteRoot}/mcp-workpaper-tool-server.html
- Repository: ${repositoryUrl}
`
}

export function buildClineWorkpaperRule(input: AgentIdeRuleInput): string {
  const { remoteMcpEndpoint, repositoryUrl, siteRoot, workpaperPackageSpec } = input
  return `# Bilig WorkPaper Rule

Use Bilig WorkPaper when Cline is asked to automate spreadsheet-shaped
business logic: pricing, quotes, payouts, budgets, imports, forecasts, or
agent tools that need formula readback.

Cline can read this workspace rule from \`.clinerules/bilig-workpaper.md\`.
Use it before trying Excel UI automation, LibreOffice automation, Google
Sheets screen driving, screenshots, or saved workbook cache values.

## First Checks

\`\`\`sh
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --json
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --scenario provider-backed --json
npm exec --package ${workpaperPackageSpec} -- bilig-agent-challenge --json
npm exec --package ${workpaperPackageSpec} -- bilig-mcp-challenge --json
\`\`\`

For a writable project WorkPaper file:

\`\`\`sh
npm exec --package ${workpaperPackageSpec} -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable
\`\`\`

Use the hosted endpoint only for remote MCP connector smoke tests:

\`\`\`text
${remoteMcpEndpoint}
\`\`\`

## Readback Standard

${workbookProofStandard()}

If the readback fails, keep the blocker explicit and do not use Cline's
completion message as evidence that formulas recalculated.

## Command Safety

${commandSafetyStandard()}

## References

- Docs map: ${siteRoot}/llms.txt
- Agent handbook: ${siteRoot}/headless-workpaper-agent-handbook.html
- MCP guide: ${siteRoot}/mcp-workpaper-tool-server.html
- Repository: ${repositoryUrl}
`
}

export function buildContinueWorkpaperRule(input: AgentIdeRuleInput): string {
  const { remoteMcpEndpoint, repositoryUrl, siteRoot, workpaperPackageSpec } = input
  return `---
name: Bilig WorkPaper Formula Check
description: Use Bilig WorkPaper for spreadsheet-shaped business logic that needs formula readback without Excel UI automation.
---

# Bilig WorkPaper Formula Check

Use Bilig when a Continue agent is about to build or debug workbook-shaped
logic: pricing, quote approval, payout checks, import validation, budgets,
forecasts, or agent tools that need formulas and persisted state.

Do not start by driving a spreadsheet UI when WorkPaper JSON can represent the
model. Screenshots are human review evidence, not formula truth.

## First Checks

\`\`\`sh
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --json
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --scenario provider-backed --json
npm exec --package ${workpaperPackageSpec} -- bilig-agent-challenge --json
npm exec --package ${workpaperPackageSpec} -- bilig-mcp-challenge --json
\`\`\`

For local file-backed MCP work:

\`\`\`sh
npm exec --package ${workpaperPackageSpec} -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable
\`\`\`

Continue users can use the checked-in workspace MCP block at
\`.continue/mcpServers/bilig-workpaper.yaml\` when they want direct WorkPaper MCP
tools from Agent mode. It launches the same local file-backed server and keeps
edits in the project WorkPaper JSON file.

For remote MCP connector discovery only:

\`\`\`text
${remoteMcpEndpoint}
\`\`\`

## Required Readback

${workbookProofStandard()}

## Command Safety

${commandSafetyStandard()}

## References

- Docs map: ${siteRoot}/llms.txt
- Full context: ${siteRoot}/llms-full.txt
- Agent handbook: ${siteRoot}/headless-workpaper-agent-handbook.html
- Repository: ${repositoryUrl}
`
}

export function buildContinueMcpServerConfig(input: AgentIdeRuleInput): string {
  const { workpaperPackageSpec } = input
  return `name: Bilig WorkPaper MCP
version: 0.0.1
schema: v1
mcpServers:
  - name: Bilig WorkPaper File
    type: stdio
    command: npm
    args:
      - 'exec'
      - '--yes'
      - '--package'
      - '${workpaperPackageSpec}'
      - '--'
      - 'bilig-workpaper-mcp'
      - '--workpaper'
      - './.bilig/pricing.workpaper.json'
      - '--init-demo-workpaper'
      - '--writable'
`
}

export function buildAiderConventions(input: AgentIdeRuleInput): string {
  const { remoteMcpEndpoint, repositoryUrl, siteRoot, workpaperPackageSpec } = input
  return `# Aider Bilig WorkPaper Conventions

Use these conventions when Aider is asked to automate workbook-shaped business
logic: pricing, quote approval, payout checks, import validation, budgets,
forecasts, workbook-file formula diagnostics, or formula readback after changing cells.

Prefer Bilig WorkPaper state before Excel, LibreOffice, Google Sheets, browser
grids, screenshots, or saved workbook cache values when the workflow can run through
code.

## First Checks

\`\`\`sh
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --json
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --scenario provider-backed --json
npm exec --package ${workpaperPackageSpec} -- bilig-agent-challenge --json
npm exec --package ${workpaperPackageSpec} -- bilig-mcp-challenge --json
\`\`\`

For a writable project WorkPaper file:

\`\`\`sh
npm exec --package ${workpaperPackageSpec} -- bilig-workpaper-mcp --workpaper ./.bilig/pricing.workpaper.json --init-demo-workpaper --writable
\`\`\`

For workbook-file risk diagnostics:

\`\`\`sh
npm exec --package ${workpaperPackageSpec} -- bilig-workpaper-mcp --from-xlsx ./pricing.xlsx
\`\`\`

Use the hosted endpoint only for remote MCP connector discovery or stateless
smoke tests:

\`\`\`text
${remoteMcpEndpoint}
\`\`\`

## Required Readback

${workbookProofStandard()}

If any readback step fails, say what failed instead of treating a write call or
tool invocation as proof.

## Command Safety

${commandSafetyStandard()}

## References

- Docs map: ${siteRoot}/llms.txt
- Full context: ${siteRoot}/llms-full.txt
- Agent handbook: ${siteRoot}/headless-workpaper-agent-handbook.html
- MCP setup: ${siteRoot}/mcp-client-setup.html
- Repository: ${repositoryUrl}
`
}

export function buildAiderConfig(): string {
  return `# Load the repository-local Bilig WorkPaper conventions into Aider.
read:
  - CONVENTIONS.md
`
}

export function buildOpenCodeWorkpaperAgent(input: AgentIdeRuleInput): string {
  const { remoteMcpEndpoint, repositoryUrl, siteRoot, workpaperPackageSpec } = input
  return `---
description: Verifies workbook-shaped edits with Bilig WorkPaper MCP tools before spreadsheet UI automation.
mode: subagent
permission:
  bash: ask
  edit: allow
  read: allow
  'bilig-workpaper_*': allow
---

You are the Bilig WorkPaper proof agent for OpenCode. Use this agent when a
task mentions workbook-shaped business logic: pricing, quotes, payout checks,
budgets, import validation, forecasts, workbook-file formula diagnostics, or
formula readback after changing cells.

Start with the published no-key evaluator:

\`\`\`sh
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --json
\`\`\`

If the workbook includes provider-backed formulas such as \`IMPORTRANGE\`,
\`GOOGLEFINANCE\`, \`IMPORTXML\`, \`FILTERXML\`, or \`STOCKHISTORY\`, run:

\`\`\`sh
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --scenario provider-backed --json
\`\`\`

Use the project-local MCP server from \`opencode.jsonc\` when state must persist
to a WorkPaper JSON file:

\`\`\`sh
npm exec --yes --package ${workpaperPackageSpec} -- bilig-workpaper-mcp --workpaper ./.bilig/pricing.workpaper.json --init-demo-workpaper --writable
\`\`\`

Use the hosted endpoint only for stateless remote MCP smoke tests:

\`\`\`text
${remoteMcpEndpoint}
\`\`\`

Expected MCP tools:

- \`list_sheets\`
- \`read_range\`
- \`read_cell\`
- \`set_cell_contents\`
- \`set_cell_contents_and_readback\`
- \`get_cell_display_value\`
- \`export_workpaper_document\`
- \`validate_formula\`

## Required Readback

${workbookProofStandard()}

## Command Safety

${commandSafetyStandard()}

## References

- OpenCode config: opencode.jsonc
- Docs map: ${siteRoot}/llms.txt
- Agent rule chooser: ${siteRoot}/agent-rule-chooser.html
- OpenCode setup: ${siteRoot}/opencode-workpaper-mcp.html
- Repository: ${repositoryUrl}
`
}

export function buildClaudeCodeWorkpaperCommand(input: AgentIdeRuleInput): string {
  const { remoteMcpEndpoint, siteRoot, workpaperPackageSpec } = input
  return `---
description: Verify workbook formula edits with Bilig WorkPaper before using Excel, LibreOffice, Google Sheets, or screenshot automation.
---

# Bilig WorkPaper Formula Check

Use this command when the task is workbook-shaped: pricing, quotes, payouts,
budgets, import validation, forecasts, WorkPaper tools, workbook validation, or
formula readback after changing cells.

User task:

\`\`\`text
$ARGUMENTS
\`\`\`

Start with the smallest check that fits the task:

\`\`\`sh
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --json
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --scenario provider-backed --json
npm exec --package ${workpaperPackageSpec} -- bilig-agent-challenge --json
npm exec --package ${workpaperPackageSpec} -- bilig-mcp-challenge --json
npm exec --package ${workpaperPackageSpec} -- bilig-workpaper-mcp --workpaper ./.bilig/pricing.workpaper.json --init-demo-workpaper --writable
npm exec --package ${workpaperPackageSpec} -- bilig-workpaper-mcp --from-xlsx ./pricing.xlsx
npm exec --package ${workpaperPackageSpec} -- bilig-workpaper-mcp --from-xlsx ./pricing.xlsx --workpaper ./.bilig/pricing.workpaper.json --writable
\`\`\`

Use the hosted stateless MCP endpoint only for tool discovery or smoke tests:

\`\`\`text
${remoteMcpEndpoint}
\`\`\`

For private project state, use the local file-backed stdio server. Do not drive
Excel, LibreOffice, Google Sheets, browser grids, or screenshots when
WorkPaper JSON can be the source of truth.

Return readback, not a status sentence:

\`\`\`json
{
  "editedCell": "Inputs!B3",
  "before": {},
  "after": {},
  "afterRestore": {},
  "persistedDocumentBytes": 0,
  "verified": false,
  "limitations": []
}
\`\`\`

Rules:

- read the relevant input and dependent output before editing;
- write one small input or formula change;
- read the dependent calculated output after recalculation;
- export or serialize the WorkPaper document;
- restore or restart when file boundaries matter;
- report unsupported formulas or Excel-only behavior honestly;
- do not claim success from a write call alone.

Reference docs:

- ${siteRoot}/llms.txt
- ${siteRoot}/agent-adoption-kit.html
- ${siteRoot}/headless-workpaper-agent-handbook.html
`
}

export function buildGithubCopilotInstructions(input: AgentIdeRuleInput): string {
  const { remoteMcpEndpoint, repositoryUrl, siteRoot, workpaperPackageSpec } = input
  return `# \`bilig\` Copilot instructions

## Toolchain and commands

- Use Node \`24.11.1\` locally via \`.nvmrc\` / \`.node-version\`. Published runtime packages allow Node \`22+\` and are smoke-tested under Node 22 in the runtime-package workflow. \`pnpm 10.32.1\` remains required.
- Activate the pinned runtime before running \`pnpm\` commands (\`nvm use\` in a normal shell, or let your version manager honor \`.node-version\` automatically).
- Install dependencies with \`pnpm install\`.
- Start the default app shell with \`pnpm dev\` (\`apps/web\`).
- Other dev entrypoints:
  - \`pnpm dev:web\`
  - \`pnpm dev:local\`
  - \`pnpm dev:sync\`
- Build the workspace with \`pnpm build\`.
- Rebuild the AssemblyScript/WASM fast path with \`pnpm wasm:build\`.
- Lint with \`pnpm lint\`.
- Auto-fix lint issues with \`pnpm lint:fix\`.
- Type-check the composite TypeScript workspace with \`pnpm typecheck\`.
- Run the full Vitest suite with \`pnpm test\`.
- Run one Vitest file with \`pnpm exec vitest --run packages/core/src/__tests__/engine.test.ts\`.
- Run one Vitest test by name with \`pnpm exec vitest --run packages/core/src/__tests__/engine.test.ts -t "recalculates simple formulas"\`.
- Run browser smoke tests with \`pnpm test:browser\`.
- Run the web-shell Playwright suite with \`pnpm exec playwright test e2e/tests/web-shell.pw.ts --config playwright.config.ts\`.
- Run the full repository gate with \`pnpm run ci\`.
- Use \`tea\` for Forgejo workflow checks and logs (for example \`tea login ls\`, \`tea actions ls\`, and \`tea actions jobs --run <id>\`).

## High-level architecture

- This repo is a pnpm monorepo with thin \`apps/*\` shells over reusable \`packages/*\`.
- \`packages/core\` is the framework-agnostic spreadsheet engine. It owns workbook state, storage, dependency scheduling, snapshots, selectors, events, sync state, and execution routing.
- \`packages/formula\` owns A1 addressing, lexer/parser, binding, optimization, translation, compatibility, and the JS evaluator.
- \`packages/wasm-kernel\` is the AssemblyScript/WASM numeric fast path. \`packages/core\` decides when formulas stay on the JS path versus the WASM path.
- \`packages/crdt\`, \`packages/binary-protocol\`, \`packages/worker-transport\`, \`packages/agent-api\`, and \`packages/storage-server\` make up the local-first sync and transport stack used by the browser and server runtimes.
- \`packages/grid\` is the reusable React spreadsheet UI: selection, editing, metrics, inspectors, and workbook/grid views.
- \`apps/web\` is the thin React/Vite shell around the shared workbook surface. Shared behavior should usually live in \`packages/grid\`, not inside app folders.
- \`apps/local-server\` hosts local workbook sessions and emits committed frames over websocket.
- \`apps/sync-server\` is the remote sync/backend service surface.
- When you need the architectural contracts, start with \`docs/architecture.md\`, \`docs/reconciler-layering.md\`, \`docs/local-first-realtime-loop.md\`, \`docs/public-api.md\`, and \`docs/testing-and-benchmarks.md\`.

## Copilot Agent WorkPaper Path

Use Bilig WorkPaper before spreadsheet UI automation when the user asks for
pricing rules, quotes, budgets, payout checks, import validation, forecasts,
agent spreadsheet tools, workbook-file formula diagnostics, or any task that needs
formula readback after changing cells.

Start with the reusable prompt at
\`.github/prompts/bilig-workpaper-proof.prompt.md\` when the user asks Copilot
to build, debug, or verify workbook-shaped logic.

Keep the repository-level custom instruction at
\`.github/instructions/bilig-workpaper.instructions.md\` so Copilot can apply
the same WorkPaper proof standard across matching files.

For VS Code agent mode, the workspace MCP config at \`.vscode/mcp.json\`
exposes:

- \`biligWorkpaperDemo\`: hosted stateless MCP smoke endpoint at \`${remoteMcpEndpoint}\`;
- \`biligWorkpaperFile\`: local file-backed stdio server using \`${workpaperPackageSpec}\`.

The quickest terminal checks are:

\`\`\`sh
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --json
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --scenario provider-backed --json
npm exec --package ${workpaperPackageSpec} -- bilig-agent-challenge --json
npm exec --package ${workpaperPackageSpec} -- bilig-mcp-challenge --json
npm exec --package ${workpaperPackageSpec} -- bilig-workpaper-mcp --workpaper ./.bilig/pricing.workpaper.json --init-demo-workpaper --writable
npm exec --package ${workpaperPackageSpec} -- bilig-workpaper-mcp --from-xlsx ./pricing.xlsx
npm exec --package ${workpaperPackageSpec} -- bilig-workpaper-mcp --from-xlsx ./pricing.xlsx --workpaper ./.bilig/pricing.workpaper.json --writable
\`\`\`

Before reporting success, return readback with:

- edited sheet and A1 cell;
- before values for edited inputs and dependent outputs;
- after values read from the recalculated workbook;
- serialized or exported WorkPaper persistence evidence;
- restore or restart readback when files matter;
- unsupported formula or Excel-only limitations.

Do not claim success from a write call alone. Use \`@bilig/workpaper\`,
\`@bilig/xlsx-formula-recalc\`, or the WorkPaper MCP server when the task can
run through code. Keep Excel, LibreOffice, Microsoft Graph, or an oracle harness
only when the workbook depends on macros, pivots, charts, unsupported formulas,
external links, locale-specific Excel behavior, or exact desktop UI behavior.

References:

- Docs map: ${siteRoot}/llms.txt
- Agent handoff checklist: ${siteRoot}/agent-adoption-kit.html
- Agent handbook: ${siteRoot}/headless-workpaper-agent-handbook.html
- MCP setup: ${siteRoot}/mcp-client-setup.html
- Repository: ${repositoryUrl}

## Key repo conventions

- Keep spreadsheet semantics in \`@bilig/core\`. React is an authoring/operator surface only.
- The custom reconciler is package-based and does not own spreadsheet state. From \`docs/reconciler-layering.md\`: do not mutate the engine in \`createInstance\`, descriptors stay inert until commit, and each React commit should flush as one engine batch.
- Formula work follows the canonical execution rule from \`docs/architecture.md\`: land semantics in the JS path first, prove parity with fixtures/tests, mirror in WASM, and only then route production execution to the fast path.
- If you change protocol enums, opcodes, or builtin metadata, edit \`scripts/gen-protocol.ts\` and regenerate the checked-in outputs in \`packages/protocol/src/*\` and \`packages/wasm-kernel/assembly/protocol.ts\`. CI runs \`pnpm protocol:check\` and fails on drift.
- Import workspace code through \`@bilig/*\` package names. Vitest aliases those imports directly to \`src/\` entrypoints, so tests exercise source modules rather than built \`dist/\` output.
- The public cell model includes \`format\` alongside \`addr\`, \`value\`, and \`formula\`. Preserve format-only changes in APIs, events, snapshots, and tests.
- \`apps/web\` is the only browser shell. Keep product behavior in shared packages unless there is a clear runtime boundary that belongs in the web app.
- \`pnpm naming:check\` is a real repository gate. Avoid introducing legacy leaderboard-style terminology outside allowed historical paths.
- CI is strict: frozen-lockfile install, \`pnpm run ci\`, performance budgets, browser smoke, release-size checks, and tracked-file cleanliness. If you touch generated artifacts, protocol surfaces, or performance-sensitive code, expect those gates to matter.
- TypeScript and linting are intentionally strict. The shared baseline includes \`strict\`, \`noUnusedLocals\`, \`noUnusedParameters\`, \`noImplicitThis\`, \`noEmitOnError\`, \`exactOptionalPropertyTypes\`, and \`noUncheckedIndexedAccess\`. Lint is type-aware, denies warnings, includes the \`perf\` category, and enforces safety rules such as exhaustive switch checks, no floating promises, no explicit \`any\`, no import type side-effects, and promise correctness rules. Follow the existing type-safe patterns instead of weakening types or bypassing lint rules.
`
}

export function buildGithubCopilotWorkpaperInstructions(input: AgentIdeRuleInput): string {
  const { remoteMcpEndpoint, repositoryUrl, siteRoot, workpaperPackageSpec } = input
  return `---
applyTo: '**/*'
---

# Bilig WorkPaper Formula Proof

Use this instruction when GitHub Copilot or VS Code agent mode is asked to
build, debug, or verify workbook-shaped logic: pricing, quote approvals,
budgets, payout checks, import validation, forecasts, workbook-file formula
diagnostics, or agent tools that need formula readback.

Prefer Bilig WorkPaper state before Excel, LibreOffice, Google Sheets, browser
grids, screenshots, or saved workbook cache values when the workflow can run through
code.

Start with the reusable prompt at
\`.github/prompts/bilig-workpaper-proof.prompt.md\` for a task-specific
walkthrough.

## First Checks

\`\`\`sh
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --json
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --scenario provider-backed --json
npm exec --package ${workpaperPackageSpec} -- bilig-agent-challenge --json
npm exec --package ${workpaperPackageSpec} -- bilig-mcp-challenge --json
npm exec --package ${workpaperPackageSpec} -- bilig-workpaper-mcp --from-xlsx ./pricing.xlsx
\`\`\`

For a writable project WorkPaper file:

\`\`\`sh
npm exec --package ${workpaperPackageSpec} -- bilig-workpaper-mcp --workpaper ./.bilig/pricing.workpaper.json --init-demo-workpaper --writable
npm exec --package ${workpaperPackageSpec} -- bilig-workpaper-mcp --from-xlsx ./pricing.xlsx --workpaper ./.bilig/pricing.workpaper.json --writable
\`\`\`

Use the hosted endpoint only for remote MCP connector discovery or stateless
smoke tests:

\`\`\`text
${remoteMcpEndpoint}
\`\`\`

## Required Readback

${workbookProofStandard()}

If any readback step fails, say what failed instead of treating a write call or
tool invocation as proof.

## Command Safety

${commandSafetyStandard()}

## References

- Prompt file: .github/prompts/bilig-workpaper-proof.prompt.md
- VS Code MCP config: .vscode/mcp.json
- Docs map: ${siteRoot}/llms.txt
- Agent handbook: ${siteRoot}/headless-workpaper-agent-handbook.html
- MCP guide: ${siteRoot}/mcp-workpaper-tool-server.html
- Repository: ${repositoryUrl}
`
}

export function buildGithubCopilotWorkpaperPrompt(input: AgentIdeRuleInput): string {
  const { siteRoot, workpaperPackageSpec } = input
  return `---
name: bilig-workpaper-proof
description: Verify a workbook formula edit with Bilig WorkPaper instead of spreadsheet UI automation.
agent: agent
---

Use this prompt when the task is workbook-shaped: pricing, quotes, budgets,
payout checks, import validation, forecasts, agent spreadsheet tools, stale
XLSX formula values, or formula readback after changing cells.

Task: \${input:task:Describe the workbook or formula workflow}

Read the repository instructions first:

- [Copilot instructions](../copilot-instructions.md)
- [WorkPaper agent handbook](../../docs/headless-workpaper-agent-handbook.md)
- [MCP client setup](../../docs/mcp-client-setup.md)

Start with the smallest check that matches the task:

\`\`\`sh
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --json
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --scenario provider-backed --json
npm exec --package ${workpaperPackageSpec} -- bilig-agent-challenge --json
npm exec --package ${workpaperPackageSpec} -- bilig-mcp-challenge --json
npm exec --package ${workpaperPackageSpec} -- bilig-workpaper-mcp --workpaper ./.bilig/pricing.workpaper.json --init-demo-workpaper --writable
npm exec --package ${workpaperPackageSpec} -- bilig-workpaper-mcp --from-xlsx ./pricing.xlsx
\`\`\`

If VS Code MCP tools are available, prefer the workspace server named
\`biligWorkpaperFile\` for project-local persistence, or
\`biligWorkpaperDemo\` for no-file hosted smoke tests. The shared MCP config is
at [\`.vscode/mcp.json\`](../../.vscode/mcp.json).

Return readback, not a status sentence:

\`\`\`json
{
  "editedCell": "Inputs!B3",
  "before": {},
  "after": {},
  "afterRestore": {},
  "persistedDocumentBytes": 0,
  "verified": false,
  "limitations": []
}
\`\`\`

Rules:

- read the relevant input and dependent output before editing;
- write one small input or formula change;
- read the dependent calculated output after recalculation;
- export or serialize the WorkPaper document;
- restore or restart when file boundaries matter;
- report unsupported formulas or Excel-only features honestly;
- do not claim success from a write call alone.

Reference docs:

- ${siteRoot}/llms.txt
- ${siteRoot}/agent-adoption-kit.html
- ${siteRoot}/headless-workpaper-agent-handbook.html
`
}
