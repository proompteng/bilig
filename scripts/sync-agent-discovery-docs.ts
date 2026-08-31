import { createHash } from 'node:crypto'
import { mkdir, readFile, writeFile } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { buildAgentJsonManifest } from './agent-discovery-agent-json.ts'
import { buildDocsAgentInstructions, buildDocsAgentStart } from './agent-discovery-agent-instructions.ts'
import { buildLlmsFullSources } from './agent-discovery-llms-full-sources.ts'
import { buildKiroMcpConfig, buildKiroWorkpaperSteering } from './agent-discovery-kiro-rules.ts'
import {
  buildClaudeCodeMcpConfig,
  buildCursorMcpConfig,
  buildJunieMcpConfig,
  buildOpenCodeMcpConfig,
  buildReusableMcpConfig,
  buildRooMcpConfig,
  buildTraeMcpConfig,
  buildVscodeMcpConfig,
  buildZedSettingsConfig,
} from './agent-discovery-mcp-configs.ts'
import { mcpServerCardManifest } from './agent-discovery-mcp-card.ts'
import { buildWorkpaperPackageAgentInstructions, buildWorkpaperPackageSkillDocument } from './agent-discovery-package-docs.ts'
import {
  buildStarterAgentOverlayInstructions,
  buildStarterClaudeInstructions,
  buildStarterGeminiInstructions,
  buildStarterOverlayPackageJson,
  buildStarterOverlayReadme,
  withStarterWorkpaperPath,
} from './agent-discovery-starter-overlay.ts'
import { readTextFileIfExists } from './read-if-exists.ts'
import { syncVersionedStaticReferences } from './sync-agent-static-references.ts'
import {
  buildAiderConfig,
  buildAiderConventions,
  buildClineWorkpaperRule,
  buildClaudeCodeProjectMemory,
  buildClaudeCodeWorkpaperCommand,
  buildContinueMcpServerConfig,
  buildContinueWorkpaperRule,
  buildCursorWorkpaperRule,
  buildGithubCopilotInstructions,
  buildGithubCopilotWorkpaperInstructions,
  buildGithubCopilotWorkpaperPrompt,
  buildOpenCodeWorkpaperAgent,
  buildRooWorkpaperRule,
  buildTraeWorkpaperRule,
  buildWindsurfWorkpaperRule,
} from './agent-discovery-ide-rules.ts'

const repoRoot = join(dirname(fileURLToPath(import.meta.url)), '..')
const siteRoot = 'https://proompteng.github.io/bilig'
const skillDiscoveryRoot = 'https://bilig.proompteng.ai'
const remoteMcpEndpoint = 'https://bilig.proompteng.ai/mcp'
const remoteMcpAliasEndpoint = 'https://bilig.proompteng.ai/mcp/workpaper'
const remoteMcpServerCard = 'https://bilig.proompteng.ai/.well-known/mcp/server-card.json'
const repositoryUrl = 'https://github.com/proompteng/bilig'
const skillName = 'bilig-workpaper'
const skillManifestUrl = `${skillDiscoveryRoot}/.well-known/agent-skills/${skillName}/SKILL.txt`
const skillDiscoverySchemaUrl = 'https://schemas.agentskills.io/discovery/0.2.0/schema.json'
const headlessPackageVersion = parsePackageVersion(await readFile(join(repoRoot, 'packages', 'headless', 'package.json'), 'utf8'))
const headlessPackageSpec = `@bilig/headless@${headlessPackageVersion}`
const workpaperPackageSpec = '@bilig/workpaper@latest'
const workbookPackageSpec = '@bilig/workbook@latest'
const unscopedWorkpaperPackageSpec = 'bilig-workpaper@latest'
const mcpbReleaseAssetUrl = `${repositoryUrl}/releases/latest/download/bilig-workpaper.mcpb`
const mcpbReleaseChecksumUrl = `${mcpbReleaseAssetUrl}.sha256`
const checkOnly = process.argv.includes('--check')

function parsePackageVersion(packageJson: string): string {
  const parsed: unknown = JSON.parse(packageJson)
  if (typeof parsed !== 'object' || parsed === null || Array.isArray(parsed)) {
    throw new Error('packages/headless/package.json must be an object')
  }
  const version = Reflect.get(parsed, 'version')
  if (typeof version !== 'string') {
    throw new Error('packages/headless/package.json must define a string version')
  }
  return version
}

const docsAgentInstructions = buildDocsAgentInstructions({
  mcpbReleaseAssetUrl,
  mcpbReleaseChecksumUrl,
  remoteMcpEndpoint,
  remoteMcpServerCard,
  siteRoot,
  skillDiscoveryRoot,
  skillManifestUrl,
  workpaperPackageSpec,
})

const docsAgentStart = buildDocsAgentStart({
  mcpbReleaseAssetUrl,
  mcpbReleaseChecksumUrl,
  remoteMcpEndpoint,
  remoteMcpServerCard,
  siteRoot,
  skillDiscoveryRoot,
  skillManifestUrl,
  workpaperPackageSpec,
})

const headlessPackageAgentInstructions = `# @bilig/headless agent notes

If you are editing this package inside the Bilig monorepo, follow the root
\`AGENTS.md\` first. These notes are also published with the npm package so an
agent inspecting \`node_modules/@bilig/headless\` can use WorkPaper state without
driving a spreadsheet UI.

## Use WorkPaper first

Use \`@bilig/workpaper\` first when the task is workbook-shaped business logic:
pricing, quote approval, budget checks, payout models, import validation,
forecasting, or tool integrations that need formula readback.

Use \`@bilig/headless\` only when the task explicitly targets lower-level
runtime subpaths, package-boundary audits, provenance checks, or internals of
this package.

Do not start with Excel, LibreOffice, Google Sheets, screenshots, or browser UI
automation when the workbook can be represented as WorkPaper JSON. Screenshots
are for human visual review, not formula truth.

## Handoff prompt

Paste this into a tool host when it is about to automate a spreadsheet
through a desktop app or browser grid:

\`\`\`text
Before using spreadsheet UI automation, check whether this is workbook-shaped
logic that can run through @bilig/workpaper. If it is, use WorkPaper state as the
source of truth.

For triage, start with:
npm exec --yes --package ${workpaperPackageSpec} -- bilig-agent-start --json

For MCP proof, run:
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --json
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --scenario provider-backed --json
npm exec --package ${workpaperPackageSpec} -- bilig-mcp-challenge --json
npm exec --package ${workpaperPackageSpec} -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable
npm exec --package ${workpaperPackageSpec} -- bilig-workpaper-mcp --from-xlsx ./pricing.xlsx
pnpm --dir examples/headless-workpaper run agent:mcp-xlsx-risk-preflight

For Node or TypeScript, import @bilig/workpaper directly. Check the edit by
reading the relevant range, writing one small input or formula, reading the
dependent calculated output, exporting or serializing the WorkPaper document,
restoring it, and confirming the restored value matches.

Return editedCell, before, after, afterRestore, persistedDocumentBytes,
verified, and limitations. Do not claim success from a write call alone.
\`\`\`

## Minimum edit loop

For every tool-owned workbook edit:

1. identify the exact sheet and A1 cell or range.
2. read the current input and dependent output.
3. validate formulas before writing them.
4. write one small change.
5. read the dependent computed output after recalculation.
6. serialize or export the WorkPaper document.
7. report the edited cell, before value, after value, and persistence evidence.

Do not report success from a write call alone.

## MCP entrypoint

For MCP clients, use the published stdio server:

\`\`\`sh
npm exec --package ${workpaperPackageSpec} -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable
\`\`\`

Expected file-backed tools:

- \`list_sheets\`
- \`read_range\`
- \`read_cell\`
- \`set_cell_contents\`
- \`set_cell_contents_and_readback\`
- \`get_cell_display_value\`
- \`export_workpaper_document\`
- \`validate_formula\`

Use \`--init-demo-workpaper\` when the path may not exist yet; it creates the demo
WorkPaper JSON only when the file is missing. Use \`--writable\` only when the
task should persist \`set_cell_contents\` edits back to the same WorkPaper JSON
file.

When the server is started through \`${workpaperPackageSpec}\` with
\`--from-xlsx ./pricing.xlsx\`, \`tools/list\` also includes
\`analyze_workbook_risk\`. That tool is fixed to the source XLSX passed at
startup and reports workbook risk indicators before a workflow trusts the imported
WorkPaper. Without \`--workpaper --writable\`, edits stay in memory; add a
WorkPaper JSON path only when the task needs persisted file state. It does not
certify Excel compatibility.

For a maintained transcript that starts from a real XLSX, call
\`analyze_workbook_risk\`, then prove \`Inputs!B3\` -> \`Summary!B3\` readback
and export with:

\`\`\`sh
pnpm --dir examples/headless-workpaper run agent:mcp-xlsx-risk-preflight
\`\`\`

Claude Desktop users can skip manual JSON config by installing the released
MCPB bundle:

- ${mcpbReleaseAssetUrl}
- ${mcpbReleaseChecksumUrl}

## Direct TypeScript entrypoint

Use \`@bilig/workpaper\` when the workbook logic belongs in a service, queue
worker, test, or route:

\`\`\`ts
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
  after: proof.afterReadback.displayValues,
  afterRestore: proof.restoredReadback.displayValues,
  persistedDocumentBytes: proof.persistedDocumentBytes,
  verified: proof.verified,
})

book.dispose()
\`\`\`

## Verification shortcuts

From a clean project, run the package-owned check:

\`\`\`sh
npm exec --yes --package ${workpaperPackageSpec} -- bilig-agent-start --json
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door workpaper-service --json
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --json
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --scenario provider-backed --json
npm exec --package ${workpaperPackageSpec} -- bilig-agent-challenge --json
npm exec --package ${workpaperPackageSpec} -- bilig-mcp-challenge --json
pnpm --dir examples/headless-workpaper run agent:mcp-xlsx-risk-preflight
\`\`\`

\`bilig-agent-challenge\` checks the direct WorkPaper API loop.
\`bilig-mcp-challenge\` checks the file-backed MCP JSON-RPC loop. A good run
prints \`verified: true\`.

When the task explicitly targets this lower-level \`@bilig/headless\` package,
run the same checks against this package boundary:

\`\`\`sh
npm exec --package ${headlessPackageSpec} -- bilig-agent-challenge --json
npm exec --package ${headlessPackageSpec} -- bilig-mcp-challenge --json
npm exec --package ${headlessPackageSpec} -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable
\`\`\`

Deeper docs:

- <https://proompteng.github.io/bilig/headless-workpaper-agent-handbook.html>
- <https://proompteng.github.io/bilig/mcp-workpaper-tool-server.html>
- <https://proompteng.github.io/bilig/mcp-client-setup.html>
`

const skillDocument = `---
name: bilig-workpaper
version: 0.1.0
description: Use @bilig/workpaper WorkPaper state for workbook formulas, MCP editing, and tool integrations without driving spreadsheet UI.
tags:
  - agents
  - workbook-runtime
  - formulas
  - workpaper
  - mcp
  - typescript
---

# Bilig WorkPaper Agent Skill

Use this skill when an agent needs spreadsheet-style formulas but the work should run through files, terminal commands, TypeScript, HTTP routes, or MCP tools instead of Excel UI automation.

## When To Trigger

Trigger this skill for tasks involving:

- workbook-shaped business logic in Node.js services;
- formula readback after writing cells;
- quote, budget, payout, pricing, import-validation, or forecast models;
- agent spreadsheet tools that need deterministic cell addresses;
- MCP clients that can run a stdio server or call a Streamable HTTP endpoint;
- reduced formula/import bugs that need a local report.

Do not trigger it for manual spreadsheet editing, Office macros, VBA, pivots, charts, COM automation, or exact Excel desktop behavior unless the user explicitly asks to compare Bilig against an Excel oracle.

## Command Safety

Do not build shell commands by concatenating user text. Treat the commands below as literal templates, validate workbook paths before use, and reject values containing newlines, backticks, \`$(\`, \`;\`, \`&\`, \`|\`, \`<\`, or \`>\`. Prefer MCP client \`command\` plus \`args\` arrays or direct TypeScript calls when inserting user-provided paths or cell references.

## First Check: Agent Triage

Before wiring a client or opening a spreadsheet UI, print the compact decision
card:

\`\`\`json
{
  "command": "npm",
  "args": ["exec", "--yes", "--package", "${workpaperPackageSpec}", "--", "bilig-agent-start", "--json"]
}
\`\`\`

## First Check: Agent Evaluator

Before wiring a client, prove the published agent door with the package-owned evaluator.
It exercises MCP discovery, cell mutation, formula readback, JSON export, restart restore, and returns \`verified: true\`:

\`\`\`json
{
  "command": "npm",
  "args": ["exec", "--yes", "--package", "${workpaperPackageSpec}", "--", "bilig-evaluate", "--door", "agent-mcp", "--json"]
}
\`\`\`

For service-owned WorkPaper logic without MCP, run \`bilig-evaluate --door workpaper-service --json\`.
Use the lower-level challenge commands only when debugging the direct API loop or file-backed MCP JSON-RPC transcript:

\`\`\`json
[
  { "command": "npm", "args": ["exec", "--package", "${workpaperPackageSpec}", "--", "bilig-agent-challenge", "--json"] },
  { "command": "npm", "args": ["exec", "--package", "${workpaperPackageSpec}", "--", "bilig-mcp-challenge", "--json"] }
]
\`\`\`

## First Choice: MCP

Use MCP when the host can run a stdio server or call a Streamable HTTP server.
Configure stdio as an argument array, not a shell-concatenated string:

If the host supports installable skills, first check that the public skill
package is discoverable:

\`\`\`sh
npx --yes skills@latest add ${skillDiscoveryRoot} --list
npx --yes skills@latest add proompteng/bilig --skill bilig-workpaper --list
\`\`\`

\`\`\`json
{
  "command": "npm",
  "args": [
    "exec",
    "--package",
    "${workpaperPackageSpec}",
    "--",
    "bilig-workpaper-mcp",
    "--workpaper",
    "./pricing.workpaper.json",
    "--init-demo-workpaper",
    "--writable"
  ]
}
\`\`\`

Run \`bilig-evaluate --door agent-mcp --json\` first. If the workbook contains
provider-backed formulas such as \`IMPORTRANGE\`, run
\`bilig-evaluate --door agent-mcp --scenario provider-backed --json\` to confirm
the adapter boundary. If the evaluator fails, run \`bilig-mcp-challenge\` and
treat its returned \`tools\` array as the source of truth for the currently published package. The core file-backed tools are:

- \`list_sheets\`
- \`read_range\`
- \`read_cell\`
- \`set_cell_contents\`
- \`set_cell_contents_and_readback\`
- \`get_cell_display_value\`
- \`export_workpaper_document\`
- \`validate_formula\`

When the server is started through \`${workpaperPackageSpec}\` with
\`--from-xlsx ./pricing.xlsx\`, \`tools/list\` also includes
\`analyze_workbook_risk\`. That tool is fixed to the source XLSX passed at
startup and reports workbook risk indicators before a workflow trusts the imported
WorkPaper. Without \`--workpaper --writable\`, edits stay in memory; add a
WorkPaper JSON path only when the task needs persisted file state. It does not
certify Excel compatibility.

For a maintained XLSX preflight transcript, run
\`pnpm --dir examples/headless-workpaper run agent:mcp-xlsx-risk-preflight\`.
It requires \`analyze_workbook_risk\`, \`set_cell_contents_and_readback\`,
\`export_workpaper_document\`, \`Inputs!B3\`, \`Summary!B3\`, \`60000 -> 96000\`,
and \`verified: true\`.

After a write, always read the dependent output cell and export the WorkPaper
document. If the listed tool set includes \`set_cell_contents_and_readback\`,
prefer it for stateless clients because the edit and dependent readback happen
in one tool call. If it is absent, call \`set_cell_contents\`, then \`read_cell\`
or \`read_range\`, then \`export_workpaper_document\`.

For remote MCP clients, use the stateless demo endpoint when the client supports
Streamable HTTP:

\`\`\`text
${remoteMcpEndpoint}
${remoteMcpAliasEndpoint}
\`\`\`

The remote endpoint is request-local and does not write user files. Use it for
connector smoke tests, tool discovery, and agent onboarding; use the file-backed
stdio command when the workflow must persist a project WorkPaper JSON file.

## Second Choice: Direct TypeScript

Use \`@bilig/workpaper\` directly when workbook logic belongs in a service, queue worker, test, or route:

\`\`\`ts
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
  after: proof.afterReadback.displayValues,
  afterRestore: proof.restoredReadback.displayValues,
  persistedDocumentBytes: proof.persistedDocumentBytes,
  verified: proof.verified,
})

book.dispose()
\`\`\`

## Formula Clinic

When the user has a reduced workbook formula/import bug, generate a local report through an argument array:

\`\`\`json
{
  "command": "npm",
  "args": [
    "exec",
    "--package",
    "${workpaperPackageSpec}",
    "--",
    "bilig-formula-clinic",
    "./reduced.xlsx",
    "--cells",
    "Summary!B7,Inputs!B2"
  ]
}
\`\`\`

The report is local. It does not upload workbook contents. Ask for a reduced public fixture rather than private customer spreadsheets.

## Required Verification

Return readback, not a write-only claim. A successful agent response should include:

- the exact edited sheet and A1 cell;
- before values for relevant inputs and dependent outputs;
- after values read from the recalculated workbook;
- persistence evidence from serialized or exported WorkPaper state;
- restore or reimport checks when file boundaries matter;
- limitations for unsupported formulas or Excel-only features.

If any readback step fails, report the blocker instead of claiming the workbook was updated.

## Reference URLs

- Compact docs map: ${siteRoot}/llms.txt
- Full host context: ${siteRoot}/llms-full.txt
- Host handbook: ${siteRoot}/headless-workpaper-agent-handbook.html
- Agent workbook challenge: ${siteRoot}/agent-workbook-challenge.html
- MCP server guide: ${siteRoot}/mcp-workpaper-tool-server.html
- OpenHands MCP setup: ${siteRoot}/openhands-workpaper-mcp.html
- OpenCode MCP setup: ${siteRoot}/opencode-workpaper-mcp.html
- Open WebUI tool setup: ${siteRoot}/open-webui-workpaper-mcp.html
- LobeHub MCP setup: ${siteRoot}/lobehub-workpaper-mcp.html
- AnythingLLM MCP setup: ${siteRoot}/anythingllm-workpaper-mcp.html
- Sim MCP setup: ${siteRoot}/sim-workpaper-mcp.html
- Formula clinic: ${siteRoot}/formula-bug-clinic.html
- Compatibility limits: ${siteRoot}/where-bilig-is-not-excel-compatible-yet.html
- Repository: ${repositoryUrl}
`

const workpaperPackageAgentInstructions = buildWorkpaperPackageAgentInstructions({
  headlessPackageAgentInstructions,
  headlessPackageSpec,
  unscopedWorkpaperPackageSpec,
  workpaperPackageSpec,
})

const workpaperPackageSkillDocument = buildWorkpaperPackageSkillDocument({
  skillDocument,
  workpaperPackageSpec,
  unscopedWorkpaperPackageSpec,
})

const llmsFullSources = buildLlmsFullSources(repositoryUrl)

function skillIndexJson(): string {
  const skillDigest = createHash('sha256').update(skillDocument).digest('hex')
  return `${JSON.stringify(
    {
      $schema: skillDiscoverySchemaUrl,
      skills: [
        {
          name: skillName,
          type: 'skill-md',
          description:
            'Use @bilig/workpaper WorkPaper state, MCP tools, and formula-clinic reports instead of spreadsheet UI automation when an agent needs formula readback.',
          url: skillManifestUrl,
          digest: `sha256:${skillDigest}`,
        },
      ],
    },
    null,
    2,
  )}\n`
}

function stripFrontmatter(content: string): string {
  if (!content.startsWith('---\n')) {
    return content.trim()
  }
  return content.replace(/^---\n[\s\S]*?\n---\n+/, '').trim()
}

async function buildLlmsFull(): Promise<string> {
  const sections: string[] = [
    '# Bilig llms-full',
    '',
    '> Full host context for Bilig, a WorkPaper formula readback runtime for Node services, MCP clients, and tool integrations.',
    '',
    `Repository: ${repositoryUrl}`,
    `Site: ${siteRoot}/`,
    `npm: https://www.npmjs.com/package/@bilig/workpaper`,
    `npm workbook: https://www.npmjs.com/package/@bilig/workbook`,
    `Agent start: ${siteRoot}/agent-start.txt`,
    `Agent instructions: ${siteRoot}/AGENTS.md`,
    `Agent install context: ${siteRoot}/llms-install.html`,
    `Skill manifest: ${skillManifestUrl}`,
    `Compact index: ${siteRoot}/llms.txt`,
    '',
    '## Published Starter Status',
    '',
    'Do not use `npm create @bilig/workpaper@latest` while `@bilig/create-workpaper@latest` resolves to `0.164.11`. Its generated smoke reports `formulasPersisted: false` because that release checks escaped JSON text instead of restored formula state. The source fix must ship in a newer release and pass a fresh consumer smoke before this path is restored.',
    '',
    '## Generated Agent Instructions',
    docsAgentInstructions.trim(),
    '',
    '## Generated Skill Manifest',
    skillDocument.trim(),
  ]

  const sourceSections = await Promise.all(
    llmsFullSources.map(async (source): Promise<string[]> => {
      const content =
        source.relativePath === 'packages/headless/AGENTS.md'
          ? headlessPackageAgentInstructions
          : await readFile(join(repoRoot, source.relativePath), 'utf8')
      return ['', '---', '', `## ${source.title}`, '', `Source: ${source.url}`, '', stripFrontmatter(content)]
    }),
  )

  sourceSections.forEach((section) => sections.push(...section))

  return `${sections.join('\n')}\n`
}

async function generatedTargets(): Promise<ReadonlyArray<readonly [string, string]>> {
  const llmsFull = await buildLlmsFull()
  const llms = await readFile(join(repoRoot, 'docs', 'llms.txt'), 'utf8')
  const llmsInstall = await readFile(join(repoRoot, 'llms-install.md'), 'utf8')
  const agentJson = buildAgentJsonManifest({
    mcpbReleaseAssetUrl,
    mcpbReleaseChecksumUrl,
    remoteMcpAliasEndpoint,
    remoteMcpEndpoint,
    remoteMcpServerCard,
    repositoryUrl,
    siteRoot,
    skillDiscoveryRoot,
    skillManifestUrl,
    skillName,
    workpaperPackageSpec,
  })
  const ideRuleInput = { remoteMcpEndpoint, repositoryUrl, siteRoot, workpaperPackageSpec }
  const mcpServerCard = mcpServerCardManifest({
    headlessPackageSpec: workpaperPackageSpec,
    headlessPackageVersion,
    remoteMcpEndpoint,
    repositoryUrl,
    siteRoot,
  })
  return [
    ['docs/AGENTS.md', docsAgentInstructions],
    ['docs/agent-start.txt', docsAgentStart],
    ['docs/agent.json', agentJson],
    ['docs/skill.md', skillDocument],
    ['docs/skill.txt', skillDocument],
    ['docs/llms-install.md', llmsInstall],
    ['docs/llms-full.txt', llmsFull],
    ['docs/.well-known/agent.json', agentJson],
    ['docs/.well-known/agent-start.txt', docsAgentStart],
    ['docs/.well-known/llms.txt', llms],
    ['docs/.well-known/llms-full.txt', llmsFull],
    ['docs/.well-known/agent-skills/index.json', skillIndexJson()],
    ['docs/.well-known/agent-skills/bilig-workpaper/SKILL.md', skillDocument],
    ['docs/.well-known/agent-skills/bilig-workpaper/SKILL.txt', skillDocument],
    ['docs/.well-known/skills/index.json', skillIndexJson()],
    ['docs/.well-known/skills/bilig-workpaper/SKILL.md', skillDocument],
    ['docs/.well-known/skills/bilig-workpaper/SKILL.txt', skillDocument],
    ['docs/.well-known/mcp/server-card.json', mcpServerCard],
    ['docs/.well-known/mcp.json', mcpServerCard],
    ['docs/.well-known/mcp-server-card.json', mcpServerCard],
    ['CONVENTIONS.md', buildAiderConventions(ideRuleInput)],
    ['.aider.conf.yml', buildAiderConfig()],
    ['.cursor/rules/bilig-workpaper.mdc', buildCursorWorkpaperRule(ideRuleInput)],
    ['.kiro/steering/bilig-workpaper.md', buildKiroWorkpaperSteering(ideRuleInput)],
    ['.roo/rules/bilig-workpaper.md', buildRooWorkpaperRule(ideRuleInput)],
    ['.trae/rules/bilig-workpaper.md', buildTraeWorkpaperRule(ideRuleInput)],
    ['.devin/rules/bilig-workpaper.md', buildWindsurfWorkpaperRule(ideRuleInput)],
    ['.windsurf/rules/bilig-workpaper.md', buildWindsurfWorkpaperRule(ideRuleInput)],
    ['.clinerules/bilig-workpaper.md', buildClineWorkpaperRule(ideRuleInput)],
    ['.continue/rules/bilig-workpaper.md', buildContinueWorkpaperRule(ideRuleInput)],
    ['.continue/mcpServers/bilig-workpaper.yaml', buildContinueMcpServerConfig(ideRuleInput)],
    ['.github/copilot-instructions.md', buildGithubCopilotInstructions(ideRuleInput)],
    ['.github/instructions/bilig-workpaper.instructions.md', buildGithubCopilotWorkpaperInstructions(ideRuleInput)],
    ['.github/prompts/bilig-workpaper-proof.prompt.md', buildGithubCopilotWorkpaperPrompt(ideRuleInput)],
    ['.opencode/agents/bilig-workpaper.md', buildOpenCodeWorkpaperAgent(ideRuleInput)],
    ['CLAUDE.md', buildClaudeCodeProjectMemory(ideRuleInput)],
    ['.mcp.json', buildClaudeCodeMcpConfig(ideRuleInput)],
    ['.cursor/mcp.json', buildCursorMcpConfig(ideRuleInput)],
    ['.kiro/settings/mcp.json', buildKiroMcpConfig(ideRuleInput)],
    ['.junie/mcp/mcp.json', buildJunieMcpConfig(ideRuleInput)],
    ['.roo/mcp.json', buildRooMcpConfig(ideRuleInput)],
    ['.trae/mcp.json', buildTraeMcpConfig(ideRuleInput)],
    ['.zed/settings.json', buildZedSettingsConfig(ideRuleInput)],
    ['.vscode/mcp.json', buildVscodeMcpConfig(ideRuleInput)],
    ['opencode.jsonc', buildOpenCodeMcpConfig(ideRuleInput)],
    ['mcp/bilig-workpaper.mcp.json', buildReusableMcpConfig(ideRuleInput)],
    ['.claude/commands/bilig-workpaper-proof.md', buildClaudeCodeWorkpaperCommand(ideRuleInput)],
    ['.claude/skills/bilig-workpaper/SKILL.md', skillDocument],
    ['.agents/skills/bilig-workpaper/SKILL.md', skillDocument],
    ['skills/bilig-workpaper/SKILL.md', skillDocument],
    ['packages/workpaper/SKILL.md', skillDocument],
    ['packages/workpaper/AGENTS.md', docsAgentInstructions],
    ['packages/create-workpaper/agent-overlay/AGENTS.md', buildStarterAgentOverlayInstructions()],
    ['packages/create-workpaper/agent-overlay/CLAUDE.md', buildStarterClaudeInstructions()],
    ['packages/create-workpaper/agent-overlay/GEMINI.md', buildStarterGeminiInstructions()],
    ['packages/create-workpaper/agent-overlay/README.md', buildStarterOverlayReadme()],
    ['packages/create-workpaper/agent-overlay/package.json', buildStarterOverlayPackageJson()],
    ['packages/create-workpaper/agent-overlay/CONVENTIONS.md', withStarterWorkpaperPath(buildAiderConventions(ideRuleInput))],
    ['packages/create-workpaper/agent-overlay/.aider.conf.yml', buildAiderConfig()],
    ['packages/create-workpaper/agent-overlay/.agents/skills/bilig-workpaper/SKILL.md', withStarterWorkpaperPath(skillDocument)],
    ['packages/create-workpaper/agent-overlay/.claude/skills/bilig-workpaper/SKILL.md', withStarterWorkpaperPath(skillDocument)],
    [
      'packages/create-workpaper/agent-overlay/.claude/commands/bilig-workpaper-proof.md',
      withStarterWorkpaperPath(buildClaudeCodeWorkpaperCommand(ideRuleInput)),
    ],
    [
      'packages/create-workpaper/agent-overlay/.clinerules/bilig-workpaper.md',
      withStarterWorkpaperPath(buildClineWorkpaperRule(ideRuleInput)),
    ],
    [
      'packages/create-workpaper/agent-overlay/.continue/rules/bilig-workpaper.md',
      withStarterWorkpaperPath(buildContinueWorkpaperRule(ideRuleInput)),
    ],
    [
      'packages/create-workpaper/agent-overlay/.cursor/rules/bilig-workpaper.mdc',
      withStarterWorkpaperPath(buildCursorWorkpaperRule(ideRuleInput)),
    ],
    [
      'packages/create-workpaper/agent-overlay/.devin/rules/bilig-workpaper.md',
      withStarterWorkpaperPath(buildWindsurfWorkpaperRule(ideRuleInput)),
    ],
    [
      'packages/create-workpaper/agent-overlay/.github/copilot-instructions.md',
      withStarterWorkpaperPath(buildGithubCopilotInstructions(ideRuleInput)),
    ],
    [
      'packages/create-workpaper/agent-overlay/.github/instructions/bilig-workpaper.instructions.md',
      withStarterWorkpaperPath(buildGithubCopilotWorkpaperInstructions(ideRuleInput)),
    ],
    [
      'packages/create-workpaper/agent-overlay/.github/prompts/bilig-workpaper-proof.prompt.md',
      withStarterWorkpaperPath(buildGithubCopilotWorkpaperPrompt(ideRuleInput)),
    ],
    [
      'packages/create-workpaper/agent-overlay/.opencode/agents/bilig-workpaper.md',
      withStarterWorkpaperPath(buildOpenCodeWorkpaperAgent(ideRuleInput)),
    ],
    [
      'packages/create-workpaper/agent-overlay/.roo/rules/bilig-workpaper.md',
      withStarterWorkpaperPath(buildRooWorkpaperRule(ideRuleInput)),
    ],
    [
      'packages/create-workpaper/agent-overlay/.kiro/steering/bilig-workpaper.md',
      withStarterWorkpaperPath(buildKiroWorkpaperSteering(ideRuleInput)),
    ],
    [
      'packages/create-workpaper/agent-overlay/.windsurf/rules/bilig-workpaper.md',
      withStarterWorkpaperPath(buildWindsurfWorkpaperRule(ideRuleInput)),
    ],
    ['packages/create-workpaper/agent-overlay/.kiro/settings/mcp.json', withStarterWorkpaperPath(buildKiroMcpConfig(ideRuleInput))],
    ['packages/create-workpaper/agent-overlay/.mcp.json', withStarterWorkpaperPath(buildClaudeCodeMcpConfig(ideRuleInput))],
    ['packages/create-workpaper/agent-overlay/.cursor/mcp.json', withStarterWorkpaperPath(buildCursorMcpConfig(ideRuleInput))],
    ['packages/create-workpaper/agent-overlay/.junie/mcp/mcp.json', withStarterWorkpaperPath(buildJunieMcpConfig(ideRuleInput))],
    ['packages/create-workpaper/agent-overlay/.roo/mcp.json', withStarterWorkpaperPath(buildRooMcpConfig(ideRuleInput))],
    ['packages/create-workpaper/agent-overlay/.trae/mcp.json', withStarterWorkpaperPath(buildTraeMcpConfig(ideRuleInput))],
    [
      'packages/create-workpaper/agent-overlay/.trae/rules/bilig-workpaper.md',
      withStarterWorkpaperPath(buildTraeWorkpaperRule(ideRuleInput)),
    ],
    ['packages/create-workpaper/agent-overlay/.zed/settings.json', withStarterWorkpaperPath(buildZedSettingsConfig(ideRuleInput))],
    ['packages/create-workpaper/agent-overlay/.vscode/mcp.json', withStarterWorkpaperPath(buildVscodeMcpConfig(ideRuleInput))],
    ['packages/create-workpaper/agent-overlay/opencode.jsonc', withStarterWorkpaperPath(buildOpenCodeMcpConfig(ideRuleInput))],
    [
      'packages/create-workpaper/agent-overlay/mcp/bilig-workpaper.mcp.json',
      withStarterWorkpaperPath(buildReusableMcpConfig(ideRuleInput)),
    ],
    [
      'packages/create-workpaper/agent-overlay/.continue/mcpServers/bilig-workpaper.yaml',
      withStarterWorkpaperPath(buildContinueMcpServerConfig(ideRuleInput)),
    ],
    ['packages/headless/SKILL.md', skillDocument],
    ['packages/headless/AGENTS.md', headlessPackageAgentInstructions],
    ['packages/bilig/SKILL.md', workpaperPackageSkillDocument],
    ['packages/bilig/AGENTS.md', workpaperPackageAgentInstructions],
  ] as const
}

const staticReferenceMismatches = await syncVersionedStaticReferences({
  checkOnly,
  headlessPackageSpec,
  headlessPackageVersion,
  mcpbReleaseAssetUrl,
  mcpbReleaseChecksumUrl,
  repoRoot,
  workbookPackageSpec,
  workpaperPackageSpec,
})
const targetResults = await Promise.all(
  (await generatedTargets()).map(async ([relativePath, content]): Promise<string | undefined> => {
    const absolutePath = join(repoRoot, relativePath)
    const existing = await readTextFileIfExists(absolutePath)
    if (existing === content) {
      return undefined
    }

    if (checkOnly) {
      return relativePath
    }

    await mkdir(dirname(absolutePath), { recursive: true })
    await writeFile(absolutePath, content)
    return undefined
  }),
)

const mismatchedTargets = [...staticReferenceMismatches, ...targetResults.filter((target): target is string => target !== undefined)]

if (mismatchedTargets.length > 0) {
  console.error(`Agent discovery docs are stale:\n${mismatchedTargets.map((target) => `- ${target}`).join('\n')}`)
  console.error('Run `pnpm agent:discovery:generate`.')
  process.exitCode = 1
}
