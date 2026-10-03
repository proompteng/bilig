import { readFile } from 'node:fs/promises'
import { join } from 'node:path'

import { requireFile, requireIncludes, requireNotIncludes } from './check-docs-discovery-core.ts'
import { loadDocsDiscoveryContext } from './check-docs-discovery-context.ts'
import { requireDocsLocalHtmlLinksHaveSources } from './check-docs-discovery-local-links.ts'
import { requireSitemapPublishedSources } from './check-docs-discovery-sitemap.ts'

const context = await loadDocsDiscoveryContext()
const {
  repoRoot,
  docsRoot,
  expectedSitemapUrls,
  sourceFilesByUrl,
  readme,
  index,
  llms,
  llmsFull,
  sitemap,
  headlessReadme,
  workbookCompatibilityReport,
  xlsxFormulaRecalculationNode,
  formulaWorkbooksProof,
  showHnFormulaWorkbooksProof,
  workbookAutomationExamplesDoc,
  serverSideSpreadsheetAutomationNode,
  nodeFrameworkWorkpaperAdaptersDoc,
} = context

requireDocsLocalHtmlLinksHaveSources(docsRoot)
requireSitemapPublishedSources({
  expectedSitemapUrls,
  sitemap,
  siteRoot: context.siteRoot,
  sourceFilesByUrl,
})

await Promise.all(
  [
    'examples/headless-workpaper/README.md',
    'examples/recalc-bridge-workflows/README.md',
    'examples/serverless-workpaper-api/README.md',
    'examples/xlsx-cache-doctor-ci/README.md',
    'examples/xlsx-recalculation-node/README.md',
    'integrations/n8n-nodes-workpaper/README.md',
  ].map((retainedPath) => requireFile(join(repoRoot, retainedPath))),
)

for (const removedNeedle of [
  'product-hunt',
  'community-growth',
  'community-launch',
  'airbyte-workpaper-validation',
  'airflow-workpaper-dag',
  'dagster-workpaper-asset',
  'directus-workpaper-flow-operation',
  'huggingface-workpaper-space',
  'inngest-workpaper-step',
  'kestra-workpaper-flow',
  'langchain-mcp-workpaper-toolnode',
  'langgraph-workpaper-tool-state',
  'mastra-workpaper-tool',
  'meltano-workpaper-utility',
  'prefect-workpaper-flow',
  'temporal-workpaper-activity',
  'triggerdev-workpaper-task',
  'windmill-workpaper-script',
  'pipedream-workpaper-formula-readback',
  'workbook-agent-model',
] as const) {
  for (const [path, content] of [
    ['README.md', readme],
    ['docs/index.html', index],
    ['docs/llms.txt', llms],
    ['docs/llms-full.txt', llmsFull],
    ['packages/workpaper/README.md', headlessReadme],
  ] as const) {
    requireNotIncludes(content, removedNeedle, path)
  }
}

for (const [path, content, needles] of [
  [
    'README.md',
    readme,
    [
      'examples/headless-workpaper',
      'examples/serverless-workpaper-api',
      'examples/xlsx-recalculation-node',
      'examples/recalc-bridge-workflows',
    ],
  ],
  [
    'packages/workpaper/README.md',
    headlessReadme,
    ['examples/headless-workpaper', 'examples/serverless-workpaper-api', 'examples/xlsx-recalculation-node'],
  ],
  ['docs/workbook-compatibility-report.md', workbookCompatibilityReport, ['workbook-compatibility-report', '@bilig/xlsx-formula-recalc']],
  ['docs/xlsx-formula-recalculation-node.md', xlsxFormulaRecalculationNode, ['xlsx-recalc', '@bilig/xlsx-formula-recalc']],
  ['docs/formula-workbooks-node-services-agent-tools.md', formulaWorkbooksProof, ['npm install @bilig/workpaper', 'bilig-workpaper-mcp']],
  [
    'docs/show-hn-formula-workbooks-node-services.md',
    showHnFormulaWorkbooksProof,
    ['bilig-evaluate --door workpaper-service --json', 'bilig-evaluate --door agent-mcp --json'],
  ],
  [
    'docs/workbook-automation-examples-node.md',
    workbookAutomationExamplesDoc,
    ['examples/headless-workpaper', 'invoice-totals', 'subscription-mrr'],
  ],
  ['docs/server-side-spreadsheet-automation-node.md', serverSideSpreadsheetAutomationNode, ['@bilig/workpaper', 'calculated cells']],
  [
    'docs/node-framework-workpaper-adapters.md',
    nodeFrameworkWorkpaperAdaptersDoc,
    ['examples/serverless-workpaper-api', 'Hono', 'Next.js'],
  ],
] as const) {
  for (const needle of needles) {
    requireIncludes(content, needle, path)
  }
}

const n8nReadme = await readFile(join(repoRoot, 'integrations', 'n8n-nodes-workpaper', 'README.md'), 'utf8')
requireIncludes(n8nReadme, '@bilig/n8n-nodes-workpaper', 'integrations/n8n-nodes-workpaper/README.md')
requireNotIncludes(n8nReadme, 'examples/n8n-workpaper-formula-readback', 'integrations/n8n-nodes-workpaper/README.md')
