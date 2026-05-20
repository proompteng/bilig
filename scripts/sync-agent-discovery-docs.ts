import { mkdir, readdir, readFile, writeFile } from 'node:fs/promises'
import { dirname, extname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { mcpServerCardManifest } from './agent-discovery-mcp-card.ts'

const repoRoot = join(dirname(fileURLToPath(import.meta.url)), '..')
const siteRoot = 'https://proompteng.github.io/bilig'
const remoteMcpEndpoint = 'https://bilig.proompteng.ai/mcp'
const remoteMcpAliasEndpoint = 'https://bilig.proompteng.ai/mcp/workpaper'
const remoteMcpServerCard = 'https://bilig.proompteng.ai/.well-known/mcp/server-card.json'
const repositoryUrl = 'https://github.com/proompteng/bilig'
const skillName = 'bilig-workpaper'
const headlessPackageVersion = parsePackageVersion(await readFile(join(repoRoot, 'packages', 'headless', 'package.json'), 'utf8'))
const workpaperPackageVersion = parsePackageVersion(await readFile(join(repoRoot, 'packages', 'workpaper', 'package.json'), 'utf8'))
const unscopedWorkpaperPackageVersion = parsePackageVersion(await readFile(join(repoRoot, 'packages', 'bilig', 'package.json'), 'utf8'))
const headlessPackageSpec = `@bilig/headless@${headlessPackageVersion}`
const workpaperPackageSpec = `@bilig/workpaper@${workpaperPackageVersion}`
const unscopedWorkpaperPackageSpec = `bilig-workpaper@${unscopedWorkpaperPackageVersion}`
const mcpbReleaseTag = `libraries-v${headlessPackageVersion}`
const mcpbReleaseAssetUrl = `${repositoryUrl}/releases/download/${mcpbReleaseTag}/bilig-workpaper.mcpb`
const mcpbReleaseChecksumUrl = `${mcpbReleaseAssetUrl}.sha256`
const skillTags = ['ai-agents', 'spreadsheet-automation', 'formulas', 'xlsx', 'mcp', 'typescript'] as const
const mcpPromptNames = ['edit_and_verify_workpaper', 'debug_workpaper_formula'] as const
const agentNotAFitBoundaries = [
  'manual spreadsheet editing as the main product',
  'Office macros or desktop Excel automation',
  'one-off arithmetic',
] as const
const versionedStaticReferenceRoots = [
  'README.md',
  'packages/headless/README.md',
  'packages/headless/AGENTS.md',
  'packages/headless/SKILL.md',
  'packages/bilig/README.md',
  'packages/bilig/AGENTS.md',
  'packages/bilig/SKILL.md',
  'skills/bilig-workpaper/SKILL.md',
  'docs/agent-workbook-challenge.md',
  'docs/claude-desktop-mcpb-workpaper.md',
  'docs/formula-bug-clinic.md',
  'docs/formula-workbooks-node-services-agent-tools.md',
  'docs/ai-agent-spreadsheet-tool-node.md',
  'docs/headless-workpaper-agent-handbook.md',
  'docs/index.html',
  'docs/llms.txt',
  'docs/mcp-client-setup.md',
  'docs/mcp-spreadsheet-server-directory.md',
  'docs/mcp-workpaper-tool-server.md',
  'docs/spreadsheet-mcp-server-comparison.md',
  'docs/why-agents-need-workbook-apis.md',
] as const
const versionedStaticReferenceExtensions = new Set(['.html', '.json', '.md', '.txt'])

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

const docsAgentInstructions = `# Bilig Agent Instructions

Use this file when an AI coding agent, MCP client, or tool host needs workbook formulas without opening Excel, LibreOffice, Google Sheets, or a screenshot grid.

## Discovery Order

1. Read \`${siteRoot}/llms.txt\` for the compact map.
2. Read \`${siteRoot}/llms-full.txt\` when you need enough context to implement a workflow without searching the whole site.
3. Read \`${siteRoot}/skill.txt\` when your agent supports skill manifests.
4. Start the MCP server or import \`@bilig/workpaper\` directly.

Remote MCP clients that support Streamable HTTP can smoke-test the stateless
demo endpoint at \`${remoteMcpEndpoint}\`. Directory scanners that inspect the
hosted endpoint origin can read the same-origin server card at
\`${remoteMcpServerCard}\`. Local agents that need writable project files should
still prefer the file-backed stdio command below.

## Default Decision

Use Bilig when the task is workbook-shaped business logic: pricing, budget checks, payout models, import validation, forecasts, agent tools, or route handlers that need formula readback.

Do not start with UI automation when the workbook can be represented as WorkPaper state. Screenshots are useful for human review; they are not formula truth.

## Required Proof

Before reporting success, return a compact proof object with:

- \`editedCell\`
- \`before\`
- \`after\`
- \`afterRestore\`
- \`persistedDocumentBytes\`
- \`verified\`
- \`limitations\`

Do not claim success from a write call alone. The proof is computed readback plus persisted state.

## Fast Commands

\`\`\`sh
npm exec --package ${workpaperPackageSpec} -- bilig-agent-challenge
npm exec --package ${workpaperPackageSpec} -- bilig-mcp-challenge
npm exec --package ${workpaperPackageSpec} -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable
npm exec --package ${workpaperPackageSpec} -- bilig-formula-clinic ./reduced.xlsx --cells "Summary!B7,Inputs!B2"
\`\`\`

Claude Desktop users can install the released MCPB bundle from:

- ${mcpbReleaseAssetUrl}
- ${mcpbReleaseChecksumUrl}

## Direct TypeScript

Use \`WorkPaper.buildFromSheets()\` for hand-authored models, \`setCellContents()\` for edits, \`getCellDisplayValue()\` for readback, and \`exportWorkPaperDocument()\` plus \`serializeWorkPaperDocument()\` for persistence proof.

## Boundaries

Keep Excel, LibreOffice, Microsoft Graph, or an oracle harness in the loop when the workbook depends on macros, pivots, charts, external links, unsupported functions, locale-specific Excel behavior, or exact desktop UI behavior.
`

const headlessPackageAgentInstructions = `# @bilig/headless agent notes

If you are editing this package inside the Bilig monorepo, follow the root
\`AGENTS.md\` first. These notes are also published with the npm package so an
agent inspecting \`node_modules/@bilig/headless\` can use WorkPaper state without
driving a spreadsheet UI.

## Use the workbook API first

Use \`@bilig/headless\` when the task is workbook-shaped business logic:
pricing, quote approval, budget checks, payout models, import validation,
forecasting, or agent tools that need formula readback.

Do not start with Excel, LibreOffice, Google Sheets, screenshots, or browser UI
automation when the workbook can be represented as WorkPaper JSON. Screenshots
are for human visual review, not formula truth.

## Handoff prompt

Paste this into another coding agent when it is about to automate a spreadsheet
through a desktop app or browser grid:

\`\`\`text
Before using spreadsheet UI automation, check whether this is workbook-shaped
logic that can run through @bilig/headless. If it is, use WorkPaper state as the
source of truth.

For MCP, start with:
npm exec --package ${headlessPackageSpec} -- bilig-mcp-challenge
npm exec --package ${headlessPackageSpec} -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable

For Node or TypeScript, import @bilig/headless directly. Prove the edit by
reading the relevant range, writing one small input or formula, reading the
dependent calculated output, exporting or serializing the WorkPaper document,
restoring it, and confirming the restored value matches.

Return editedCell, before, after, afterRestore, persistedDocumentBytes,
verified, and limitations. Do not claim success from a write call alone.
\`\`\`

## Minimum edit loop

For every agent-owned workbook edit:

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
npm exec --package ${headlessPackageSpec} -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable
\`\`\`

Expected file-backed tools:

- \`list_sheets\`
- \`read_range\`
- \`read_cell\`
- \`set_cell_contents\`
- \`get_cell_display_value\`
- \`export_workpaper_document\`
- \`validate_formula\`

Use \`--init-demo-workpaper\` when the path may not exist yet; it creates the demo
WorkPaper JSON only when the file is missing. Use \`--writable\` only when the
task should persist \`set_cell_contents\` edits back to the same WorkPaper JSON
file.

Claude Desktop users can skip manual JSON config by installing the released
MCPB bundle:

- ${mcpbReleaseAssetUrl}
- ${mcpbReleaseChecksumUrl}

## Direct TypeScript entrypoint

Use the package API when the workbook logic belongs in a service, queue worker,
test, or route:

\`\`\`ts
import { WorkPaper, exportWorkPaperDocument, serializeWorkPaperDocument } from '@bilig/headless'

const workbook = WorkPaper.buildFromSheets({
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

const inputs = workbook.getSheetId('Inputs')
const summary = workbook.getSheetId('Summary')
if (inputs === undefined || summary === undefined) {
  throw new Error('Workbook is missing required sheets')
}

workbook.setCellContents({ sheet: inputs, row: 1, col: 1 }, 32)

const revenue = workbook.getCellDisplayValue({ sheet: summary, row: 1, col: 1 })
const saved = serializeWorkPaperDocument(exportWorkPaperDocument(workbook, { includeConfig: true }))

console.log({ revenue, savedBytes: saved.length })
\`\`\`

## Verification shortcuts

From a clean project, run the package-owned challenge:

\`\`\`sh
npm exec --package ${headlessPackageSpec} -- bilig-agent-challenge
npm exec --package ${headlessPackageSpec} -- bilig-mcp-challenge
\`\`\`

\`bilig-agent-challenge\` proves the direct WorkPaper API loop.
\`bilig-mcp-challenge\` proves the file-backed MCP JSON-RPC loop. A good run
prints \`verified: true\`.

Deeper docs:

- <https://proompteng.github.io/bilig/headless-workpaper-agent-handbook.html>
- <https://proompteng.github.io/bilig/mcp-workpaper-tool-server.html>
- <https://proompteng.github.io/bilig/mcp-client-setup.html>
`

const skillDocument = `---
name: bilig-workpaper
version: 0.1.0
description: Use @bilig/workpaper WorkPaper state for workbook formulas, agent spreadsheet tools, MCP file-backed or remote demo editing, and XLSX formula bug reports without driving spreadsheet UI.
tags:
  - ai-agents
  - spreadsheet-automation
  - formulas
  - xlsx
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
- reduced XLSX formula bugs that need a paste-ready report.

Do not trigger it for manual spreadsheet editing, Office macros, VBA, pivots, charts, COM automation, or exact Excel desktop behavior unless the user explicitly asks to compare Bilig against an Excel oracle.

## Command Safety

Do not build shell commands by concatenating user text. Treat the commands below as literal templates, validate workbook paths before use, and reject values containing newlines, backticks, \`$(\`, \`;\`, \`&\`, \`|\`, \`<\`, or \`>\`. Prefer MCP client \`command\` plus \`args\` arrays or direct TypeScript calls when inserting user-provided paths or cell references.

## First Choice: MCP

Use MCP when the host can run a stdio server or call a Streamable HTTP server.
Configure stdio as an argument array, not a shell-concatenated string:

Before wiring a client, an agent can prove the direct WorkPaper loop with:

\`\`\`json
{
  "command": "npm",
  "args": ["exec", "--package", "${workpaperPackageSpec}", "--", "bilig-agent-challenge"]
}
\`\`\`

For the actual file-backed MCP path, run the package-owned challenge first:

\`\`\`json
{
  "command": "npm",
  "args": ["exec", "--package", "${workpaperPackageSpec}", "--", "bilig-mcp-challenge"]
}
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

The useful file-backed tools are:

- \`list_sheets\`
- \`read_range\`
- \`read_cell\`
- \`set_cell_contents\`
- \`get_cell_display_value\`
- \`export_workpaper_document\`
- \`validate_formula\`

After a write, always read the dependent output cell and export the WorkPaper document.

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
import { WorkPaper, exportWorkPaperDocument, serializeWorkPaperDocument } from '@bilig/workpaper'

const workbook = WorkPaper.buildFromSheets({
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

const inputs = workbook.getSheetId('Inputs')
const summary = workbook.getSheetId('Summary')
if (inputs === undefined || summary === undefined) {
  throw new Error('Workbook is missing required sheets')
}

workbook.setCellContents({ sheet: inputs, row: 1, col: 1 }, 32)
const revenue = workbook.getCellDisplayValue({ sheet: summary, row: 1, col: 1 })
const saved = serializeWorkPaperDocument(exportWorkPaperDocument(workbook, { includeConfig: true }))

console.log({ revenue, savedBytes: saved.length })
\`\`\`

## XLSX Formula Clinic

When the user has a reduced XLSX formula/import bug, generate a local report through an argument array:

\`\`\`json
{
  "command": "npm",
  "args": ["exec", "--package", "${workpaperPackageSpec}", "--", "bilig-formula-clinic", "./reduced.xlsx", "--cells", "Summary!B7,Inputs!B2"]
}
\`\`\`

The report is local. It does not upload workbook contents. Ask for a reduced public fixture rather than private customer spreadsheets.

## Required Verification

Return proof, not vibes. A successful agent response should include:

- the exact edited sheet and A1 cell;
- before values for relevant inputs and dependent outputs;
- after values read from the recalculated workbook;
- persistence evidence from serialized or exported WorkPaper state;
- restore or reimport proof when file boundaries matter;
- limitations for unsupported formulas or Excel-only features.

If any proof step fails, report the blocker instead of claiming the workbook was updated.

## Reference URLs

- Compact docs map: ${siteRoot}/llms.txt
- Full agent context: ${siteRoot}/llms-full.txt
- Agent handbook: ${siteRoot}/headless-workpaper-agent-handbook.html
- Agent workbook challenge: ${siteRoot}/agent-workbook-challenge.html
- MCP server guide: ${siteRoot}/mcp-workpaper-tool-server.html
- XLSX formula clinic: ${siteRoot}/formula-bug-clinic.html
- Compatibility limits: ${siteRoot}/where-bilig-is-not-excel-compatible-yet.html
- Repository: ${repositoryUrl}
`

const workpaperPackageAgentInstructions = headlessPackageAgentInstructions
  .replace('# @bilig/headless agent notes', '# bilig-workpaper agent notes')
  .replace('agent inspecting `node_modules/@bilig/headless`', 'agent inspecting `node_modules/bilig-workpaper`')
  .replace(/@bilig\/headless/g, 'bilig-workpaper')
  .replaceAll(headlessPackageSpec, unscopedWorkpaperPackageSpec)

const workpaperPackageSkillDocument = skillDocument
  .replace(/@bilig\/workpaper/g, 'bilig-workpaper')
  .replaceAll(workpaperPackageSpec, unscopedWorkpaperPackageSpec)
  .replace(
    `## First Choice: MCP

Use MCP when the host can run a stdio server or call a Streamable HTTP server.
Configure stdio as an argument array, not a shell-concatenated string:

Before wiring a client, an agent can prove the direct WorkPaper loop with:`,
    `## First Choice: Direct npm proof or TypeScript

Use the package directly when the host can run npm or TypeScript. This is the
highest-traffic evaluator path because it meets developers where they already
search: npm, ExcelJS, SheetJS, xlsx-populate, and StackOverflow formula
recalculation problems.

Start by proving the direct WorkPaper loop:`,
  )
  .replace(
    'For the actual file-backed MCP path, run the package-owned challenge first:',
    'Use MCP only when the host specifically needs an MCP client boundary. For that path, run the package-owned challenge first:',
  )
  .replace('## Second Choice: Direct TypeScript', '## Direct TypeScript')

const llmsFullSources = [
  { title: 'Repository README', relativePath: 'README.md', url: `${repositoryUrl}/blob/main/README.md` },
  {
    title: 'Headless Package README',
    relativePath: 'packages/headless/README.md',
    url: `${repositoryUrl}/blob/main/packages/headless/README.md`,
  },
  {
    title: 'Headless Package Agent Notes',
    relativePath: 'packages/headless/AGENTS.md',
    url: `${repositoryUrl}/blob/main/packages/headless/AGENTS.md`,
  },
  {
    title: 'Headless WorkPaper Agent Handbook',
    relativePath: 'docs/headless-workpaper-agent-handbook.md',
    url: `${repositoryUrl}/blob/main/docs/headless-workpaper-agent-handbook.md`,
  },
  {
    title: 'Agent Workbook Challenge',
    relativePath: 'docs/agent-workbook-challenge.md',
    url: `${repositoryUrl}/blob/main/docs/agent-workbook-challenge.md`,
  },
  {
    title: 'Agent WorkPaper Tool-Calling Recipe',
    relativePath: 'docs/agent-workpaper-tool-calling-recipe.md',
    url: `${repositoryUrl}/blob/main/docs/agent-workpaper-tool-calling-recipe.md`,
  },
  {
    title: 'AI Spreadsheet Agent Tool For Node.js',
    relativePath: 'docs/ai-agent-spreadsheet-tool-node.md',
    url: `${repositoryUrl}/blob/main/docs/ai-agent-spreadsheet-tool-node.md`,
  },
  {
    title: 'OpenAI Agents SDK WorkPaper Tool',
    relativePath: 'docs/openai-agents-sdk-workpaper-tool.md',
    url: `${repositoryUrl}/blob/main/docs/openai-agents-sdk-workpaper-tool.md`,
  },
  {
    title: 'MCP WorkPaper Tool Server',
    relativePath: 'docs/mcp-workpaper-tool-server.md',
    url: `${repositoryUrl}/blob/main/docs/mcp-workpaper-tool-server.md`,
  },
  {
    title: 'Agent XLSX Formula Recalculation Without LibreOffice',
    relativePath: 'docs/agent-xlsx-formula-recalculation-without-libreoffice.md',
    url: `${repositoryUrl}/blob/main/docs/agent-xlsx-formula-recalculation-without-libreoffice.md`,
  },
  {
    title: 'Formula Bug Clinic',
    relativePath: 'docs/formula-bug-clinic.md',
    url: `${repositoryUrl}/blob/main/docs/formula-bug-clinic.md`,
  },
  {
    title: 'Try Bilig Headless In Node',
    relativePath: 'docs/try-bilig-headless-in-node.md',
    url: `${repositoryUrl}/blob/main/docs/try-bilig-headless-in-node.md`,
  },
  {
    title: 'Quote Approval WorkPaper API',
    relativePath: 'docs/quote-approval-workpaper-api.md',
    url: `${repositoryUrl}/blob/main/docs/quote-approval-workpaper-api.md`,
  },
  {
    title: 'Compatibility Limits',
    relativePath: 'docs/where-bilig-is-not-excel-compatible-yet.md',
    url: `${repositoryUrl}/blob/main/docs/where-bilig-is-not-excel-compatible-yet.md`,
  },
  {
    title: 'npm Provenance And Package Trust',
    relativePath: 'docs/npm-provenance-package-trust.md',
    url: `${repositoryUrl}/blob/main/docs/npm-provenance-package-trust.md`,
  },
] as const

function skillIndexJson(basePath: 'agent-skills' | 'skills'): string {
  const json = JSON.stringify(
    {
      schema_version: basePath === 'agent-skills' ? 'agent-skills-0.2.0' : 'skills-index-1.0',
      skills: [
        {
          name: skillName,
          title: 'Bilig WorkPaper agent workbook formulas',
          description:
            'Use @bilig/workpaper WorkPaper state, MCP tools, and formula-clinic reports instead of spreadsheet UI automation when an agent needs formula readback.',
          url: `${siteRoot}/.well-known/${basePath}/${skillName}/SKILL.txt`,
          source_url: `${repositoryUrl}/blob/main/docs/.well-known/${basePath}/${skillName}/SKILL.txt`,
          tags: skillTags,
        },
      ],
    },
    null,
    2,
  )
  return `${compactStringArrayProperty(json, 'tags', skillTags, '      ')}\n`
}

function agentJsonManifest(): string {
  const json = JSON.stringify(
    {
      schema_version: 'agent-json-0.1.0',
      name: 'bilig',
      title: 'Bilig WorkPaper formula runtime',
      description:
        'Formula WorkPaper runtime for Node.js services and agent tools: edit cells, recalculate, verify readback, and persist JSON without spreadsheet UI automation.',
      url: `${siteRoot}/`,
      repository: repositoryUrl,
      license: 'MIT',
      contact: `${repositoryUrl}/discussions/new?category=general`,
      llms_txt: `${siteRoot}/llms.txt`,
      llms_full: `${siteRoot}/llms-full.txt`,
      skill_file: `${siteRoot}/skill.txt`,
      agent_instructions: `${siteRoot}/AGENTS.md`,
      skills: [
        {
          name: skillName,
          url: `${siteRoot}/.well-known/agent-skills/${skillName}/SKILL.txt`,
          index_url: `${siteRoot}/.well-known/agent-skills/index.json`,
          description:
            'Use @bilig/workpaper WorkPaper state, MCP tools, and formula-clinic reports instead of spreadsheet UI automation when an agent needs formula readback.',
        },
      ],
      mcp: {
        server_name: 'io.github.proompteng/bilig-workpaper',
        server_card: `${siteRoot}/.well-known/mcp/server-card.json`,
        remote_server_card: remoteMcpServerCard,
        manifest: `${siteRoot}/.well-known/mcp.json`,
        registry_search: 'https://registry.modelcontextprotocol.io/v0.1/servers?search=io.github.proompteng%2Fbilig-workpaper',
        remote_endpoint: remoteMcpEndpoint,
        remote_alias_endpoint: remoteMcpAliasEndpoint,
        remote_transport: {
          type: 'streamable-http',
          protocol_version: '2025-11-25',
          stateless: true,
          authentication_required: false,
        },
        command: 'npm',
        args: [
          'exec',
          '--package',
          workpaperPackageSpec,
          '--',
          'bilig-workpaper-mcp',
          '--workpaper',
          './pricing.workpaper.json',
          '--init-demo-workpaper',
          '--writable',
        ],
        tools: [
          'list_sheets',
          'read_range',
          'read_cell',
          'set_cell_contents',
          'get_cell_display_value',
          'export_workpaper_document',
          'validate_formula',
        ],
        resources: [
          'bilig://workpaper/manifest',
          'bilig://workpaper/agent-handoff',
          'bilig://workpaper/sheets',
          'bilig://workpaper/current-document',
        ],
        prompts: ['edit_and_verify_workpaper', 'debug_workpaper_formula'],
      },
      capabilities: [
        {
          name: 'workpaper-formula-runtime',
          type: 'npm-library',
          package: '@bilig/workpaper',
          runtime: 'Node.js >=22',
          install: 'npm install @bilig/workpaper',
          docs: `${siteRoot}/try-bilig-headless-in-node.html`,
        },
        {
          name: 'file-backed-workpaper-mcp',
          type: 'mcp-stdio-server',
          docs: `${siteRoot}/mcp-workpaper-tool-server.html`,
          server_card: `${siteRoot}/.well-known/mcp/server-card.json`,
          challenge_command: `npm exec --package ${workpaperPackageSpec} -- bilig-mcp-challenge`,
        },
        {
          name: 'claude-desktop-mcpb',
          type: 'mcpb-desktop-extension',
          package_version: headlessPackageVersion,
          download_url: mcpbReleaseAssetUrl,
          checksum_url: mcpbReleaseChecksumUrl,
          docs: `${siteRoot}/claude-desktop-mcpb-workpaper.html`,
        },
        {
          name: 'remote-workpaper-mcp-demo',
          type: 'mcp-streamable-http-server',
          endpoint: remoteMcpEndpoint,
          alias_endpoint: remoteMcpAliasEndpoint,
          server_card: remoteMcpServerCard,
          protocol_version: '2025-11-25',
          authentication_required: false,
          docs: `${siteRoot}/mcp-workpaper-tool-server.html#remote-stateless-endpoint`,
        },
        {
          name: 'formula-clinic',
          type: 'local-cli',
          command: `npm exec --package ${workpaperPackageSpec} -- bilig-formula-clinic ./reduced.xlsx --cells "Summary!B7,Inputs!B2"`,
          docs: `${siteRoot}/formula-bug-clinic.html`,
        },
      ],
      verification_contract: [
        'read the relevant range before editing',
        'write the target input or formula cell',
        'read the dependent calculated output after recalculation',
        'export or serialize the WorkPaper document',
        'restore or reimport when a file boundary matters',
        'return editedCell, before, after, afterRestore, persistedDocumentBytes, verified, and limitations',
      ],
      boundaries: {
        good_fit: [
          'pricing, quote approval, budget, payout, import-validation, and forecast logic',
          'agent tools that need deterministic cell addresses and formula readback',
          'service-owned workbook state that can persist as JSON',
        ],
        not_a_fit: agentNotAFitBoundaries,
      },
      public_entrypoints: [
        `${siteRoot}/`,
        `${siteRoot}/why-use-bilig.html`,
        `${siteRoot}/headless-workpaper-agent-handbook.html`,
        `${siteRoot}/agent-workbook-challenge.html`,
        `${siteRoot}/mcp-workpaper-tool-server.html`,
        remoteMcpEndpoint,
        remoteMcpServerCard,
        `${siteRoot}/agent-workpaper-tool-calling-recipe.html`,
        `${siteRoot}/openai-agents-sdk-workpaper-tool.html`,
        `${siteRoot}/node-framework-workpaper-adapters.html`,
        `${siteRoot}/npm-provenance-package-trust.html`,
      ],
    },
    null,
    2,
  )
  const compactPrompts = compactStringArrayProperty(json, 'prompts', mcpPromptNames, '    ')
  return `${compactStringArrayProperty(compactPrompts, 'not_a_fit', agentNotAFitBoundaries, '    ')}\n`
}

function compactStringArrayProperty(json: string, propertyName: string, values: readonly string[], indent: string): string {
  const expanded = `${indent}"${propertyName}": [\n${values.map((value) => `${indent}  ${JSON.stringify(value)}`).join(',\n')}\n${indent}]`
  const compact = `${indent}"${propertyName}": [${values.map((value) => JSON.stringify(value)).join(', ')}]`
  return json.replace(expanded, compact)
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
    '> Full agent context for Bilig, a formula WorkPaper runtime for Node services, MCP clients, and coding-agent workbook tools.',
    '',
    `Repository: ${repositoryUrl}`,
    `Site: ${siteRoot}/`,
    `npm: https://www.npmjs.com/package/@bilig/workpaper`,
    `Agent instructions: ${siteRoot}/AGENTS.md`,
    `Skill manifest: ${siteRoot}/skill.txt`,
    `Compact index: ${siteRoot}/llms.txt`,
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
  const agentJson = agentJsonManifest()
  const mcpServerCard = mcpServerCardManifest({
    headlessPackageSpec: workpaperPackageSpec,
    headlessPackageVersion,
    remoteMcpEndpoint,
    repositoryUrl,
    siteRoot,
  })
  return [
    ['docs/AGENTS.md', docsAgentInstructions],
    ['docs/agent.json', agentJson],
    ['docs/skill.md', skillDocument],
    ['docs/skill.txt', skillDocument],
    ['docs/llms-full.txt', llmsFull],
    ['docs/.well-known/agent.json', agentJson],
    ['docs/.well-known/agent-skills/index.json', skillIndexJson('agent-skills')],
    ['docs/.well-known/agent-skills/bilig-workpaper/SKILL.md', skillDocument],
    ['docs/.well-known/agent-skills/bilig-workpaper/SKILL.txt', skillDocument],
    ['docs/.well-known/skills/index.json', skillIndexJson('skills')],
    ['docs/.well-known/skills/bilig-workpaper/SKILL.md', skillDocument],
    ['docs/.well-known/skills/bilig-workpaper/SKILL.txt', skillDocument],
    ['docs/.well-known/mcp/server-card.json', mcpServerCard],
    ['docs/.well-known/mcp.json', mcpServerCard],
    ['docs/.well-known/mcp-server-card.json', mcpServerCard],
    ['skills/bilig-workpaper/SKILL.md', skillDocument],
    ['packages/workpaper/SKILL.md', skillDocument],
    ['packages/workpaper/AGENTS.md', docsAgentInstructions],
    ['packages/headless/SKILL.md', skillDocument],
    ['packages/headless/AGENTS.md', headlessPackageAgentInstructions],
    ['packages/bilig/SKILL.md', workpaperPackageSkillDocument],
    ['packages/bilig/AGENTS.md', workpaperPackageAgentInstructions],
  ] as const
}

async function readIfExists(path: string): Promise<string | undefined> {
  try {
    return await readFile(path, 'utf8')
  } catch (error) {
    if (error instanceof Error && 'code' in error && error.code === 'ENOENT') {
      return undefined
    }
    throw error
  }
}

async function collectVersionedStaticReferenceFiles(relativePath: string): Promise<string[]> {
  const absolutePath = join(repoRoot, relativePath)
  try {
    const entries = await readdir(absolutePath, { withFileTypes: true })
    const nested = await Promise.all(
      entries.map((entry) => {
        const entryRelativePath = `${relativePath}/${entry.name}`
        return entry.isDirectory() ? collectVersionedStaticReferenceFiles(entryRelativePath) : [entryRelativePath]
      }),
    )
    return nested.flat().filter((entryRelativePath) => versionedStaticReferenceExtensions.has(extname(entryRelativePath)))
  } catch (error) {
    if (error instanceof Error && 'code' in error && error.code === 'ENOTDIR') {
      return versionedStaticReferenceExtensions.has(extname(relativePath)) ? [relativePath] : []
    }
    if (error instanceof Error && 'code' in error && error.code === 'ENOENT') {
      return []
    }
    throw error
  }
}

function syncVersionedStaticReferenceLine(line: string): string {
  const stableSemverPattern = String.raw`\d+\.\d+\.\d+`
  return line
    .replace(new RegExp(`(npm exec --package )@bilig/headless@${stableSemverPattern}`, 'g'), `$1${headlessPackageSpec}`)
    .replace(new RegExp(`("--package",\\s*")@bilig/headless@${stableSemverPattern}(")`, 'g'), `$1${headlessPackageSpec}$2`)
    .replace(new RegExp(`^(\\s*)"@bilig/headless@${stableSemverPattern}"(,?\\s*)$`, 'g'), `$1"${headlessPackageSpec}"$2`)
    .replace(
      new RegExp(`(Current checked npm footprint for \`)@bilig/headless@${stableSemverPattern}(\`)`, 'g'),
      `$1${headlessPackageSpec}$2`,
    )
    .replace(new RegExp(`npm latest is \`${stableSemverPattern}\``, 'g'), `npm latest is \`${headlessPackageVersion}\``)
    .replace(new RegExp(`npm latest is \`@bilig/headless@${stableSemverPattern}\``, 'g'), `npm latest is \`${headlessPackageSpec}\``)
    .replace(new RegExp(`npm latest \`@bilig/headless@${stableSemverPattern}\``, 'g'), `npm latest \`${headlessPackageSpec}\``)
    .replace(new RegExp(`(now points reviewers at \`)@bilig/headless@${stableSemverPattern}(\`)`, 'g'), `$1${headlessPackageSpec}$2`)
    .replace(new RegExp(`libraries-v${stableSemverPattern}`, 'g'), mcpbReleaseTag)
}

function syncVersionedStaticReferenceContent(content: string): string {
  return content
    .split(/(?<=\n)/)
    .map(syncVersionedStaticReferenceLine)
    .join('')
}

async function syncVersionedStaticReferences(): Promise<string[]> {
  const targetFiles = [
    ...new Set((await Promise.all(versionedStaticReferenceRoots.map(collectVersionedStaticReferenceFiles))).flat()),
  ].toSorted()
  const staleTargets: string[] = []

  await Promise.all(
    targetFiles.map(async (relativePath) => {
      const absolutePath = join(repoRoot, relativePath)
      const existing = await readIfExists(absolutePath)
      if (existing === undefined) {
        return
      }

      const nextContent = syncVersionedStaticReferenceContent(existing)
      if (nextContent === existing) {
        return
      }

      if (checkOnly) {
        staleTargets.push(relativePath)
        return
      }

      await writeFile(absolutePath, nextContent)
    }),
  )

  return staleTargets.toSorted()
}

const staticReferenceMismatches = await syncVersionedStaticReferences()
const targetResults = await Promise.all(
  (await generatedTargets()).map(async ([relativePath, content]): Promise<string | undefined> => {
    const absolutePath = join(repoRoot, relativePath)
    const existing = await readIfExists(absolutePath)
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
