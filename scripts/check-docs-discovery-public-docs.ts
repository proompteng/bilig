import { readFile } from 'node:fs/promises'
import { join } from 'node:path'
import {
  requireDocumentNotIncludes,
  requireDocumentsInclude,
  requireDocumentsNotInclude,
  requireIncludes,
} from './check-docs-discovery-core.ts'
import { getBenchmarkDiscoveryEvidence } from './check-docs-discovery-benchmark-evidence.ts'

export async function requireSharedPublicDocsDiscovery(args: {
  readonly docsRoot: string
  readonly readme: string
  readonly headlessReadme: string
  readonly contributing: string
  readonly newContributorGuide: string
  readonly starterIssues: string
  readonly llms: string
  readonly index: string
  readonly issueTemplateConfig: string
  readonly issueTemplateRoot: string
  readonly featureRequestTemplate: string
  readonly ideasDiscussionTemplate: string
  readonly qaDiscussionTemplate: string
  readonly showAndTellDiscussionTemplate: string
  readonly generalDiscussionTemplate: string
  readonly excelImportReadme: string
  readonly publicApi: string
}): Promise<void> {
  const benchmarkEvidence = getBenchmarkDiscoveryEvidence()

  requireDocumentsInclude(
    [
      { path: 'README.md', content: args.readme },
      { path: 'packages/headless/README.md', content: args.headlessReadme },
      { path: 'CONTRIBUTING.md', content: args.contributing },
      { path: 'docs/new-contributor-guide.md', content: args.newContributorGuide },
      { path: 'docs/starter-issues.md', content: args.starterIssues },
      { path: 'docs/llms.txt', content: args.llms },
    ],
    ['https://github.com/proompteng/bilig/issues?q=is%3Aissue%20state%3Aopen%20label%3Afirst-timers-only'],
  )

  const primaryPublicDocs = [
    { path: 'README.md', content: args.readme },
    { path: 'packages/headless/README.md', content: args.headlessReadme },
  ] as const
  requireDocumentsInclude(primaryPublicDocs, [
    '## Choose An Evaluation Path',
    'If you are evaluating...',
    '90-second Node quickstart',
    'Quote approval WorkPaper API',
    'XLSX formula recalculation example',
    'MCP spreadsheet tool server',
    'npm provenance',
    'adoption blocker form',
    'submit a workbook fixture',
    '## TypeScript API Shape',
    'WorkPaper.buildFromSheets({',
    "['Revenue', '=Inputs!B2*Inputs!B3']",
    'workbook.setCellContents({ sheet: inputs, row: 1, col: 1 }, 32)',
    'workbook.getCellDisplayValue({ sheet: summary, row: 1, col: 1 })',
    'serializeWorkPaperDocument(',
    'exportWorkPaperDocument(workbook, { includeConfig: true })',
    '## Proof You Can Reproduce',
    'https://github.com/proompteng/bilig/stargazers',
    'above edits one input',
    'verifies the dependent formula result.',
    'pnpm workpaper:bench:competitive:check',
    benchmarkEvidence.p95HoldoutWorkload,
    benchmarkEvidence.p95HoldoutRatio,
    'compatibility limits',
    'Excel oracle harness',
    'stale cached formula values',
    'https://github.com/proompteng/bilig/discussions/307',
    'https://github.com/proompteng/bilig/discussions/308',
    'SECURITY.md',
    'SUPPORT.md',
    'production-adoption-checklist-headless-workpaper',
  ])
  requireDocumentsNotInclude(primaryPublicDocs, [
    '## Current Public Proof',
    'Latest checked-in snapshot',
    '`12` forks',
    '15,592` npm downloads in the',
    '`10` GitHub Discussions',
    'repository views.',
  ])

  requireDocumentsInclude(
    [
      { path: 'README.md', content: args.readme },
      { path: 'packages/headless/README.md', content: args.headlessReadme },
      { path: 'docs/index.html', content: args.index },
      { path: 'docs/community-launch-pack.md', content: await readFile(join(args.docsRoot, 'community-launch-pack.md'), 'utf8') },
      { path: 'docs/llms.txt', content: args.llms },
      { path: '.github/ISSUE_TEMPLATE/config.yml', content: args.issueTemplateConfig },
      { path: '.github/ISSUE_TEMPLATE.md', content: args.issueTemplateRoot },
      { path: '.github/ISSUE_TEMPLATE/feature_request.yml', content: args.featureRequestTemplate },
      { path: '.github/DISCUSSION_TEMPLATE/ideas.yml', content: args.ideasDiscussionTemplate },
      { path: '.github/DISCUSSION_TEMPLATE/q-a.yml', content: args.qaDiscussionTemplate },
      { path: '.github/DISCUSSION_TEMPLATE/show-and-tell.yml', content: args.showAndTellDiscussionTemplate },
    ],
    ['workbook-automation-examples-node'],
  )

  requireDocumentsInclude(
    [
      { path: 'README.md', content: args.readme },
      { path: 'packages/headless/README.md', content: args.headlessReadme },
      { path: 'docs/index.html', content: args.index },
      { path: 'docs/llms.txt', content: args.llms },
    ],
    [
      'https://github.com/proompteng/bilig/discussions/new?category=general',
      'https://github.com/proompteng/bilig/issues/new?template=workbook_fixture.yml',
      'https://github.com/proompteng/bilig/discussions/414',
      'adoption blocker',
      'submit a workbook fixture',
    ],
  )
  requireDocumentsInclude(
    [
      { path: 'README.md', content: args.readme },
      { path: 'packages/headless/README.md', content: args.headlessReadme },
      { path: 'docs/index.html', content: args.index },
      { path: 'docs/llms.txt', content: args.llms },
    ],
    ['https://github.com/proompteng/bilig/subscription', 'release'],
  )
  requireDocumentsInclude(
    await Promise.all(
      [
        'try-bilig-headless-in-node.md',
        'quote-approval-workpaper-api.md',
        'workbook-automation-examples-node.md',
        'vercel-ai-sdk-langchain-spreadsheet-tool.md',
        'mcp-workpaper-tool-server.md',
        'evaluate-excel-formulas-in-node-typescript.md',
        'google-sheets-api-alternative-node-workpaper.md',
        'headless-spreadsheet-engine-node-services-agents.md',
        'node-spreadsheet-formula-engine.md',
        'server-side-spreadsheet-automation-node.md',
      ].map(async (name) => ({
        path: `docs/${name}`,
        content: await readFile(join(args.docsRoot, name), 'utf8'),
      })),
    ),
    ['https://github.com/proompteng/bilig/discussions/new?category=general', 'adoption blocker'],
  )
  requireDocumentsInclude(
    [{ path: '.github/DISCUSSION_TEMPLATE/general.yml', content: args.generalDiscussionTemplate }],
    ['adoption blocker', 'What proof would unblock you?'],
  )

  const issueTemplateDocs = [
    { path: '.github/ISSUE_TEMPLATE/config.yml', content: args.issueTemplateConfig },
    { path: '.github/ISSUE_TEMPLATE.md', content: args.issueTemplateRoot },
  ] as const
  requireDocumentsInclude(issueTemplateDocs, ['https://github.com/proompteng/bilig/discussions/157'])
  requireDocumentsNotInclude(issueTemplateDocs, ['https://github.com/proompteng/bilig/discussions/115'])

  requireDocumentsInclude(
    [
      { path: 'README.md', content: args.readme },
      { path: 'packages/headless/README.md', content: args.headlessReadme },
      { path: 'docs/index.html', content: args.index },
      { path: 'docs/llms.txt', content: args.llms },
    ],
    [
      'node-spreadsheet-formula-engine',
      'server-side-spreadsheet-automation-node',
      'google-sheets-api-alternative-node-workpaper',
      'production-adoption-checklist-headless-workpaper',
      'examples/serverless-workpaper-api',
      'quote-approval-api',
      'node-framework-workpaper-adapters',
      'submit-workbook-fixture',
      'mcp-spreadsheet-server-directory',
    ],
  )

  requireDocumentsInclude(
    [
      { path: 'README.md', content: args.readme },
      { path: 'packages/headless/README.md', content: args.headlessReadme },
      { path: 'docs/llms.txt', content: args.llms },
    ],
    [
      'examples/headless-workpaper#invoice-totals',
      'examples/headless-workpaper#agent-framework-adapters',
      'examples/headless-workpaper#mcp-tool-server-shape',
      'agent:framework-adapters',
      'agent:mcp-tools',
      'agent:mcp-stdio',
      'npm exec --package @bilig/headless@',
      'https://registry.modelcontextprotocol.io/v0.1/servers?search=io.github.proompteng%2Fbilig-workpaper',
      'vercel-ai-sdk-langchain-spreadsheet-tool',
      'mcp-workpaper-tool-server',
      'mcp-spreadsheet-server-directory',
      'mcp-client-setup',
      'claude-desktop-mcpb-workpaper',
      'examples/headless-workpaper#budget-variance-alerts',
      'examples/headless-workpaper#fulfillment-capacity-plan',
      'examples/headless-workpaper#quote-approval-threshold',
      'examples/headless-workpaper#subscription-mrr-forecast',
    ],
  )

  requireDocumentsInclude(
    [
      { path: 'README.md', content: args.readme },
      { path: 'packages/headless/README.md', content: args.headlessReadme },
      { path: 'docs/llms.txt', content: args.llms },
    ],
    ['docs/javascript-spreadsheet-library-headless-node.md', 'docs/sheetjs-exceljs-alternative-formula-workbook-api.md'],
  )
  requireIncludes(args.llms, 'https://proompteng.github.io/bilig/sheetjs-exceljs-alternative-formula-workbook-api.html', 'docs/llms.txt')
  requireIncludes(args.llms, 'includes a TypeScript WorkPaper runtime check', 'docs/llms.txt')

  requireDocumentsInclude(
    [
      { path: 'README.md', content: args.readme },
      { path: 'packages/headless/README.md', content: args.headlessReadme },
      { path: 'docs/llms.txt', content: args.llms },
      {
        path: 'docs/what-workpaper-benchmark-proves.md',
        content: await readFile(join(args.docsRoot, 'what-workpaper-benchmark-proves.md'), 'utf8'),
      },
    ],
    ['workpaper-benchmark-card.png'],
  )

  requireDocumentsInclude(
    [
      { path: 'packages/headless/README.md', content: args.headlessReadme },
      { path: 'packages/excel-import/README.md', content: args.excelImportReadme },
      { path: 'docs/public-api.md', content: args.publicApi },
    ],
    ['@bilig/headless/xlsx', "import { exportXlsx, importXlsx } from '@bilig/headless/xlsx'", 'workbook.exportSnapshot()'],
  )

  requireDocumentNotIncludes({ path: 'packages/headless/README.md', content: args.headlessReadme }, [
    '](../../docs/',
    '](../../examples/',
    '](../../LICENSE)',
  ])
}
