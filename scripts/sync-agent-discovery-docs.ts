import { mkdir, readFile, writeFile } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const repoRoot = join(dirname(fileURLToPath(import.meta.url)), '..')
const siteRoot = 'https://proompteng.github.io/bilig'
const repositoryUrl = 'https://github.com/proompteng/bilig'
const skillName = 'bilig-workpaper'
const skillTags = ['ai-agents', 'spreadsheet-automation', 'formulas', 'xlsx', 'mcp', 'typescript'] as const
const agentNotAFitBoundaries = [
  'manual spreadsheet editing as the main product',
  'Office macros or desktop Excel automation',
  'one-off arithmetic',
] as const

const checkOnly = process.argv.includes('--check')

const docsAgentInstructions = `# Bilig Agent Instructions

Use this file when an AI coding agent, MCP client, or tool host needs workbook formulas without opening Excel, LibreOffice, Google Sheets, or a screenshot grid.

## Discovery Order

1. Read \`${siteRoot}/llms.txt\` for the compact map.
2. Read \`${siteRoot}/llms-full.txt\` when you need enough context to implement a workflow without searching the whole site.
3. Read \`${siteRoot}/skill.txt\` when your agent supports skill manifests.
4. Start the MCP server or import \`@bilig/headless\` directly.

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
npm exec --package @bilig/headless -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable
npm exec --package @bilig/headless -- bilig-formula-clinic ./reduced.xlsx --cells "Summary!B7,Inputs!B2"
\`\`\`

## Direct TypeScript

Use \`WorkPaper.buildFromSheets()\` for hand-authored models, \`setCellContents()\` for edits, \`getCellDisplayValue()\` for readback, and \`exportWorkPaperDocument()\` plus \`serializeWorkPaperDocument()\` for persistence proof.

## Boundaries

Keep Excel, LibreOffice, Microsoft Graph, or an oracle harness in the loop when the workbook depends on macros, pivots, charts, external links, unsupported functions, locale-specific Excel behavior, or exact desktop UI behavior.
`

const skillDocument = `---
name: bilig-workpaper
version: 0.1.0
description: Use @bilig/headless WorkPaper state for workbook formulas, agent spreadsheet tools, MCP file-backed editing, and XLSX formula bug reports without driving spreadsheet UI.
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
- MCP clients that can run a stdio server;
- reduced XLSX formula bugs that need a paste-ready report.

Do not trigger it for manual spreadsheet editing, Office macros, VBA, pivots, charts, COM automation, or exact Excel desktop behavior unless the user explicitly asks to compare Bilig against an Excel oracle.

## First Choice: MCP

Use MCP when the host can run a stdio server:

\`\`\`sh
npm exec --package @bilig/headless -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable
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

## Second Choice: Direct TypeScript

Use \`@bilig/headless\` directly when workbook logic belongs in a service, queue worker, test, or route:

\`\`\`ts
import { WorkPaper, exportWorkPaperDocument, serializeWorkPaperDocument } from "@bilig/headless";

const workbook = WorkPaper.buildFromSheets({
  Inputs: [
    ["Metric", "Value"],
    ["Customers", 20],
    ["Average revenue", 1200],
  ],
  Summary: [
    ["Metric", "Value"],
    ["Revenue", "=Inputs!B2*Inputs!B3"],
  ],
});

const inputs = workbook.getSheetId("Inputs");
const summary = workbook.getSheetId("Summary");
if (inputs === undefined || summary === undefined) {
  throw new Error("Workbook is missing required sheets");
}

workbook.setCellContents({ sheet: inputs, row: 1, col: 1 }, 32);
const revenue = workbook.getCellDisplayValue({ sheet: summary, row: 1, col: 1 });
const saved = serializeWorkPaperDocument(exportWorkPaperDocument(workbook, { includeConfig: true }));

console.log({ revenue, savedBytes: saved.length });
\`\`\`

## XLSX Formula Clinic

When the user has a reduced XLSX formula/import bug, generate a local report:

\`\`\`sh
npm exec --package @bilig/headless -- bilig-formula-clinic ./reduced.xlsx --cells "Summary!B7,Inputs!B2"
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
- MCP server guide: ${siteRoot}/mcp-workpaper-tool-server.html
- XLSX formula clinic: ${siteRoot}/formula-bug-clinic.html
- Compatibility limits: ${siteRoot}/where-bilig-is-not-excel-compatible-yet.html
- Repository: ${repositoryUrl}
`

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
    title: 'Agent WorkPaper Tool-Calling Recipe',
    relativePath: 'docs/agent-workpaper-tool-calling-recipe.md',
    url: `${repositoryUrl}/blob/main/docs/agent-workpaper-tool-calling-recipe.md`,
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
            'Use @bilig/headless WorkPaper state, MCP tools, and formula-clinic reports instead of spreadsheet UI automation when an agent needs formula readback.',
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
            'Use @bilig/headless WorkPaper state, MCP tools, and formula-clinic reports instead of spreadsheet UI automation when an agent needs formula readback.',
        },
      ],
      mcp: {
        server_name: 'io.github.proompteng/bilig-workpaper',
        server_card: `${siteRoot}/.well-known/mcp/server-card.json`,
        manifest: `${siteRoot}/.well-known/mcp.json`,
        registry_search: 'https://registry.modelcontextprotocol.io/v0.1/servers?search=io.github.proompteng%2Fbilig-workpaper',
        command: 'npm',
        args: [
          'exec',
          '--package',
          '@bilig/headless',
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
      },
      capabilities: [
        {
          name: 'workpaper-formula-runtime',
          type: 'npm-library',
          package: '@bilig/headless',
          runtime: 'Node.js >=22',
          install: 'npm install @bilig/headless',
          docs: `${siteRoot}/try-bilig-headless-in-node.html`,
        },
        {
          name: 'file-backed-workpaper-mcp',
          type: 'mcp-stdio-server',
          docs: `${siteRoot}/mcp-workpaper-tool-server.html`,
          server_card: `${siteRoot}/.well-known/mcp/server-card.json`,
        },
        {
          name: 'formula-clinic',
          type: 'local-cli',
          command: 'npm exec --package @bilig/headless -- bilig-formula-clinic ./reduced.xlsx --cells "Summary!B7,Inputs!B2"',
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
        `${siteRoot}/mcp-workpaper-tool-server.html`,
        `${siteRoot}/agent-workpaper-tool-calling-recipe.html`,
        `${siteRoot}/node-framework-workpaper-adapters.html`,
        `${siteRoot}/npm-provenance-package-trust.html`,
      ],
    },
    null,
    2,
  )
  return `${compactStringArrayProperty(json, 'not_a_fit', agentNotAFitBoundaries, '    ')}\n`
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
    `npm: https://www.npmjs.com/package/@bilig/headless`,
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
      const content = await readFile(join(repoRoot, source.relativePath), 'utf8')
      return ['', '---', '', `## ${source.title}`, '', `Source: ${source.url}`, '', stripFrontmatter(content)]
    }),
  )

  sourceSections.forEach((section) => sections.push(...section))

  return `${sections.join('\n')}\n`
}

async function generatedTargets(): Promise<ReadonlyArray<readonly [string, string]>> {
  const llmsFull = await buildLlmsFull()
  const agentJson = agentJsonManifest()
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
    ['packages/headless/SKILL.md', skillDocument],
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

const mismatchedTargets = targetResults.filter((target): target is string => target !== undefined)

if (mismatchedTargets.length > 0) {
  console.error(`Agent discovery docs are stale:\n${mismatchedTargets.map((target) => `- ${target}`).join('\n')}`)
  console.error('Run `pnpm agent:discovery:generate`.')
  process.exitCode = 1
}
