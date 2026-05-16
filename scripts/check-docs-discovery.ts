import { readFile } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { agentFrameworkDocRequirements, agentFrameworkLlmsRequiredLinks } from './check-docs-discovery-agent-pages.ts'
import { requireAiSdkDiscovery } from './check-docs-discovery-ai-sdk.ts'
import {
  requireDocumentedScriptsExist,
  requireFile,
  requireIncludes,
  requireNoUnsupportedGoogleSheetsTenXClaims,
  requireNotIncludes,
  requirePackageKeywords,
  requirePublishedSource,
} from './check-docs-discovery-core.ts'
import { requireSitemapPublishedSources } from './check-docs-discovery-sitemap.ts'
import { requireHomepageDiscovery } from './check-docs-discovery-homepage.ts'
import { productHuntLaunchAssetFiles, requireGrowthSurfaceDiscovery } from './check-docs-discovery-launch-kit.ts'
import { llmsExternalSurfaceLinks } from './check-docs-discovery-growth-links.ts'
import { requireNpmEvalDiscovery } from './check-docs-discovery-npm-eval.ts'
import { requireOpenAiResponsesDiscovery } from './check-docs-discovery-openai-responses.ts'
import { requireServerlessWorkPaperApiDiscovery } from './check-docs-discovery-serverless.ts'
import { docsSiteSources } from './check-docs-discovery-site-sources.ts'
import { requireStarterIssueDiscovery } from './check-docs-discovery-starter-issues.ts'
import { requireTypeScriptFirstPublicSnippets } from './check-docs-discovery-typescript-snippets.ts'
import { requireXlsxCorpusVerifierDiscovery } from './check-docs-discovery-xlsx-verifier.ts'
import { requireSharedPublicDocsDiscovery } from './check-docs-discovery-public-docs.ts'

const repoRoot = join(dirname(fileURLToPath(import.meta.url)), '..')
const docsRoot = join(repoRoot, 'docs')
const siteRoot = 'https://proompteng.github.io/bilig/'

const expectedSitemapUrls = docsSiteSources.map(([urlPath]) => `${siteRoot}${urlPath}`)
const sourceFilesByUrl = new Map<string, string>(docsSiteSources.map(([urlPath, sourceFile]) => [`${siteRoot}${urlPath}`, sourceFile]))

const [
  readme,
  contributing,
  rootPackageJson,
  index,
  siteCss,
  productCss,
  robots,
  sitemap,
  llms,
  communityLaunchPack,
  productHuntLaunchKit,
  starterIssues,
  newContributorGuide,
  headlessPackageJson,
  headlessReadme,
  excelImportReadme,
  publicApi,
  issueTemplateConfig,
  issueTemplateRoot,
  featureRequestTemplate,
  ideasDiscussionTemplate,
  qaDiscussionTemplate,
  showAndTellDiscussionTemplate,
  pullRequestTemplate,
  dominanceScorecard,
] = await Promise.all([
  readFile(join(repoRoot, 'README.md'), 'utf8'),
  readFile(join(repoRoot, 'CONTRIBUTING.md'), 'utf8'),
  readFile(join(repoRoot, 'package.json'), 'utf8'),
  readFile(join(docsRoot, 'index.html'), 'utf8'),
  readFile(join(docsRoot, 'assets', 'site.css'), 'utf8'),
  readFile(join(docsRoot, 'assets', 'product-demo.css'), 'utf8'),
  readFile(join(docsRoot, 'robots.txt'), 'utf8'),
  readFile(join(docsRoot, 'sitemap.xml'), 'utf8'),
  readFile(join(docsRoot, 'llms.txt'), 'utf8'),
  readFile(join(docsRoot, 'community-launch-pack.md'), 'utf8'),
  readFile(join(docsRoot, 'product-hunt-launch-kit.md'), 'utf8'),
  readFile(join(docsRoot, 'starter-issues.md'), 'utf8'),
  readFile(join(docsRoot, 'new-contributor-guide.md'), 'utf8'),
  readFile(join(repoRoot, 'packages', 'headless', 'package.json'), 'utf8'),
  readFile(join(repoRoot, 'packages', 'headless', 'README.md'), 'utf8'),
  readFile(join(repoRoot, 'packages', 'excel-import', 'README.md'), 'utf8'),
  readFile(join(docsRoot, 'public-api.md'), 'utf8'),
  readFile(join(repoRoot, '.github', 'ISSUE_TEMPLATE', 'config.yml'), 'utf8'),
  readFile(join(repoRoot, '.github', 'ISSUE_TEMPLATE.md'), 'utf8'),
  readFile(join(repoRoot, '.github', 'ISSUE_TEMPLATE', 'feature_request.yml'), 'utf8'),
  readFile(join(repoRoot, '.github', 'DISCUSSION_TEMPLATE', 'ideas.yml'), 'utf8'),
  readFile(join(repoRoot, '.github', 'DISCUSSION_TEMPLATE', 'q-a.yml'), 'utf8'),
  readFile(join(repoRoot, '.github', 'DISCUSSION_TEMPLATE', 'show-and-tell.yml'), 'utf8'),
  readFile(join(repoRoot, '.github', 'PULL_REQUEST_TEMPLATE.md'), 'utf8'),
  readFile(join(repoRoot, 'packages', 'benchmarks', 'baselines', 'bilig-dominance-scorecard.json'), 'utf8'),
])

const [headlessSpreadsheetEngineComparison, sheetjsExceljsAlternativeFormulaWorkbookApi, hyperformulaAlternativeHeadlessWorkpaper] =
  await Promise.all([
    readFile(join(docsRoot, 'headless-spreadsheet-engine-comparison.md'), 'utf8'),
    readFile(join(docsRoot, 'sheetjs-exceljs-alternative-formula-workbook-api.md'), 'utf8'),
    readFile(join(docsRoot, 'hyperformula-alternative-headless-workpaper.md'), 'utf8'),
  ])
const googleSheetsApiBoundaryDoc = await readFile(join(docsRoot, 'google-sheets-api-alternative-node-workpaper.md'), 'utf8')

requireHomepageDiscovery(index, siteCss, productCss)
await requireTypeScriptFirstPublicSnippets(repoRoot)
requireNoUnsupportedGoogleSheetsTenXClaims(dominanceScorecard, {
  'README.md': readme,
  'docs/index.html': index,
  'docs/google-sheets-api-alternative-node-workpaper.md': googleSheetsApiBoundaryDoc,
  'packages/headless/README.md': headlessReadme,
})
requirePackageKeywords(
  headlessPackageJson,
  [
    'agent-tools',
    'excel',
    'formula-engine',
    'headless-spreadsheet',
    'hyperformula',
    'json-persistence',
    'mcp',
    'node',
    'spreadsheet-formulas',
    'typescript',
    'workbook-api',
    'workpaper',
    'xlsx',
  ],
  'packages/headless/package.json',
)
requireIncludes(index, '"downloadUrl": "https://www.npmjs.com/package/@bilig/headless"', 'docs/index.html')
requireIncludes(index, '"applicationCategory": "DeveloperApplication"', 'docs/index.html')
requireIncludes(index, '"@type": "FAQPage"', 'docs/index.html')
for (const required of [
  './why-agents-need-workbook-apis.html',
  './stop-driving-spreadsheets-with-screenshots.html',
  './agent-workpaper-tool-calling-recipe.html',
  './vercel-ai-sdk-langchain-spreadsheet-tool.html',
  './mcp-workpaper-tool-server.html',
  './mcp-spreadsheet-server-directory.html',
  './mcp-client-setup.html',
  './claude-desktop-mcpb-workpaper.html',
  './agent-spreadsheet-tool-call-loop.html',
  './node-service-workpaper-recipe.html',
  './server-side-spreadsheet-automation-node.html',
  './google-sheets-api-alternative-node-workpaper.html',
  './node-spreadsheet-formula-engine.html',
  './evaluate-excel-formulas-in-node-typescript.html',
  './try-bilig-headless-in-node.html',
  './serverless-workpaper-api-route.html',
  './node-framework-workpaper-adapters.html',
  './persisting-formula-backed-workpaper-documents-in-node.html',
  'examples/serverless-workpaper-api#persistence-adapters',
  './building-a-revenue-model-with-headless-workpaper.html',
  './headless-spreadsheet-engine-comparison.html',
  './javascript-spreadsheet-library-headless-node.html',
  './hyperformula-alternative-headless-workpaper.html',
  'https://github.com/proompteng/bilig/stargazers',
]) {
  requireIncludes(index, required, 'docs/index.html')
}

requireIncludes(robots, 'User-agent: *', 'docs/robots.txt')
requireIncludes(robots, 'Allow: /', 'docs/robots.txt')
requireIncludes(robots, `Sitemap: ${siteRoot}sitemap.xml`, 'docs/robots.txt')

const { actualSitemapUrls, sourceFilesToVerify } = requireSitemapPublishedSources({
  expectedSitemapUrls,
  sitemap,
  siteRoot,
  sourceFilesByUrl,
})

await Promise.all(sourceFilesToVerify.map((sourceFile) => requirePublishedSource(join(docsRoot, sourceFile))))
await Promise.all(
  ['README.md', 'package.json', 'quote-approval-api.ts', 'route.ts', 'smoke.ts'].map((sourceFile) =>
    requireFile(join(repoRoot, 'examples', 'serverless-workpaper-api', sourceFile)),
  ),
)
await requireFile(join(repoRoot, 'scripts', 'build-workpaper-mcpb.ts'))
await Promise.all(
  [
    'bilig-hero-workbook-api.png',
    'bilig-hero-workbook-api.svg',
    'bilig-hero-ambient.png',
    'hero-scene.js',
    'github-social-preview.png',
    'workpaper-benchmark-card.png',
    ...productHuntLaunchAssetFiles,
  ].map((sourceFile) => requireFile(join(docsRoot, 'assets', sourceFile))),
)
await Promise.all(
  [
    'fonts.css',
    'product-demo.css',
    'fonts/LICENSE.txt',
    'fonts/README.md',
    'fonts/ibm-plex-mono-400.woff2',
    'fonts/ibm-plex-mono-500.woff2',
    'fonts/ibm-plex-mono-600.woff2',
    'fonts/ibm-plex-sans-400.woff2',
    'fonts/ibm-plex-sans-500.woff2',
    'fonts/ibm-plex-sans-600.woff2',
    'fonts/ibm-plex-sans-700.woff2',
    'fonts/ibm-plex-sans-condensed-600.woff2',
    'fonts/ibm-plex-sans-condensed-700.woff2',
  ].map((sourceFile) => requireFile(join(docsRoot, 'assets', sourceFile))),
)

for (const required of [
  'repository: https://github.com/proompteng/bilig',
  'npm package: https://www.npmjs.com/package/@bilig/headless',
  'npm run agent:tool-call',
  'npm run agent:framework-adapters',
  'npm run agent:verify',
  'https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#json-records-input',
  'https://proompteng.github.io/bilig/why-agents-need-workbook-apis.html',
  'https://proompteng.github.io/bilig/stop-driving-spreadsheets-with-screenshots.html',
  'https://proompteng.github.io/bilig/try-bilig-headless-in-node.html',
  'https://proompteng.github.io/bilig/vercel-ai-sdk-langchain-spreadsheet-tool.html',
  'https://proompteng.github.io/bilig/mcp-workpaper-tool-server.html',
  'https://proompteng.github.io/bilig/mcp-spreadsheet-server-directory.html',
  'https://proompteng.github.io/bilig/mcp-client-setup.html',
  'https://proompteng.github.io/bilig/claude-desktop-mcpb-workpaper.html',
  'https://github.com/proompteng/bilig/blob/main/docs/claude-desktop-mcpb-workpaper.md',
  'https://proompteng.github.io/bilig/agent-workpaper-tool-calling-recipe.html',
  'https://proompteng.github.io/bilig/agent-spreadsheet-tool-call-loop.html',
  'https://proompteng.github.io/bilig/node-service-workpaper-recipe.html',
  'https://proompteng.github.io/bilig/server-side-spreadsheet-automation-node.html',
  'https://proompteng.github.io/bilig/google-sheets-api-alternative-node-workpaper.html',
  'https://proompteng.github.io/bilig/serverless-workpaper-api-route.html',
  'https://proompteng.github.io/bilig/node-framework-workpaper-adapters.html',
  'https://proompteng.github.io/bilig/workbook-automation-examples-node.html',
  'https://github.com/proompteng/bilig/blob/main/docs/workbook-automation-examples-node.md',
  'https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#invoice-totals',
  'https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#budget-variance-alerts',
  'https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#fulfillment-capacity-plan',
  'https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#quote-approval-threshold',
  'https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#subscription-mrr-forecast',
  'https://github.com/proompteng/bilig/tree/main/examples/serverless-workpaper-api',
  'https://github.com/proompteng/bilig/tree/main/examples/serverless-workpaper-api#framework-adapters',
  'https://github.com/proompteng/bilig/discussions',
  'https://github.com/proompteng/bilig/discussions/157',
  'https://github.com/proompteng/bilig/discussions/167',
  'https://github.com/proompteng/bilig/discussions/230',
  'https://github.com/proompteng/bilig/discussions/270',
  'https://github.com/proompteng/bilig/discussions/307',
  'https://github.com/proompteng/bilig/discussions/308',
  'https://github.com/proompteng/bilig/discussions/335',
  'https://github.com/proompteng/bilig/discussions/115',
  'https://proompteng.github.io/bilig/node-spreadsheet-formula-engine.html',
  'https://proompteng.github.io/bilig/evaluate-excel-formulas-in-node-typescript.html',
  'https://github.com/proompteng/bilig/blob/main/docs/node-spreadsheet-formula-engine.md',
  'https://github.com/proompteng/bilig/blob/main/docs/stop-driving-spreadsheets-with-screenshots.md',
  'https://github.com/proompteng/bilig/blob/main/docs/evaluate-excel-formulas-in-node-typescript.md',
  'https://github.com/proompteng/bilig/blob/main/docs/server-side-spreadsheet-automation-node.md',
  'https://github.com/proompteng/bilig/blob/main/docs/google-sheets-api-alternative-node-workpaper.md',
  'https://github.com/proompteng/bilig/blob/main/docs/node-service-workpaper-recipe.md',
  'https://github.com/proompteng/bilig/blob/main/docs/serverless-workpaper-api-route.md',
  'https://github.com/proompteng/bilig/blob/main/docs/node-framework-workpaper-adapters.md',
  'https://github.com/proompteng/bilig/blob/main/docs/csv-shaped-workpaper-input-recipe.md',
  'https://github.com/proompteng/bilig/blob/main/docs/unsupported-formula-troubleshooting-recipe.md',
  'https://github.com/proompteng/bilig/blob/main/docs/agent-workpaper-tool-calling-recipe.md',
  'https://github.com/proompteng/bilig/blob/main/docs/vercel-ai-sdk-langchain-spreadsheet-tool.md',
  'https://github.com/proompteng/bilig/blob/main/docs/mcp-workpaper-tool-server.md',
  'https://github.com/proompteng/bilig/blob/main/docs/mcp-spreadsheet-server-directory.md',
  'https://github.com/proompteng/bilig/blob/main/docs/mcp-client-setup.md',
  'pnpm mcpb:workpaper:build',
  'https://github.com/proompteng/bilig/blob/main/examples/headless-workpaper/mcp-tool-server.ts',
  'https://github.com/proompteng/bilig/blob/main/examples/headless-workpaper/mcp-stdio-server.ts',
  'https://github.com/proompteng/bilig/blob/main/examples/headless-workpaper/agent-framework-adapters.ts',
  'https://github.com/proompteng/bilig/blob/main/docs/agent-spreadsheet-tool-call-loop.md',
  'https://github.com/proompteng/bilig/blob/main/docs/local-workpaper-benchmark-walkthrough.md',
  'https://github.com/proompteng/bilig/blob/main/docs/what-workpaper-benchmark-proves.md',
  'https://github.com/proompteng/bilig/blob/main/docs/hyperformula-alternative-headless-workpaper.md',
  'https://github.com/proompteng/bilig/blob/main/docs/headless-spreadsheet-engine-comparison.md',
  'https://github.com/proompteng/bilig/blob/main/docs/javascript-spreadsheet-library-headless-node.md',
  'https://github.com/proompteng/bilig/blob/main/docs/sheetjs-exceljs-alternative-formula-workbook-api.md',
  'https://github.com/proompteng/bilig/blob/main/docs/where-bilig-is-not-excel-compatible-yet.md',
  'https://github.com/proompteng/bilig/blob/main/docs/formula-edge-sumifs-paired-criteria-fixture.md',
  'https://github.com/proompteng/bilig/blob/main/docs/formula-edge-groupby-spill-fixture.md',
  'https://github.com/proompteng/bilig/blob/main/docs/new-contributor-guide.md',
  'https://github.com/proompteng/bilig/blob/main/docs/starter-issues.md',
  'https://github.com/proompteng/bilig/issues?q=is%3Aissue%20state%3Aopen%20label%3Afirst-timers-only',
]) {
  requireIncludes(llms, required, 'docs/llms.txt')
}
for (const required of agentFrameworkLlmsRequiredLinks) {
  requireIncludes(llms, required, 'docs/llms.txt')
}

await requireSharedPublicDocsDiscovery({
  docsRoot,
  readme,
  headlessReadme,
  contributing,
  newContributorGuide,
  starterIssues,
  llms,
  index,
  issueTemplateConfig,
  issueTemplateRoot,
  featureRequestTemplate,
  ideasDiscussionTemplate,
  qaDiscussionTemplate,
  showAndTellDiscussionTemplate,
  excelImportReadme,
  publicApi,
})

requireIncludes(readme, 'acceptance commands for first patches.', 'README.md')
requireIncludes(headlessReadme, '## Stay Connected', 'packages/headless/README.md')
requireIncludes(headlessReadme, '## More Guides', 'packages/headless/README.md')
requireIncludes(headlessReadme, 'Pick a scoped first patch:', 'packages/headless/README.md')
requireIncludes(headlessReadme, 'When the sanity check passes, these are the next useful pages.', 'packages/headless/README.md')

requireIncludes(newContributorGuide, '## First-Time Command Checklist', 'docs/new-contributor-guide.md')
requireIncludes(newContributorGuide, 'pnpm docs:discovery:check', 'docs/new-contributor-guide.md')
requireIncludes(newContributorGuide, 'pnpm format:check', 'docs/new-contributor-guide.md')
requireIncludes(newContributorGuide, 'pnpm lint', 'docs/new-contributor-guide.md')
requireIncludes(newContributorGuide, 'first-time contributor review happens on GitHub.', 'docs/new-contributor-guide.md')
requireIncludes(contributing, 'pull requests on GitHub are welcome; maintainers', 'CONTRIBUTING.md')
requireIncludes(starterIssues, 'new-contributor-guide.md#first-time-command-checklist', 'docs/starter-issues.md')
requireIncludes(starterIssues, 'https://github.com/proompteng/bilig/blob/main/CONTRIBUTING.md', 'docs/starter-issues.md')
requireIncludes(starterIssues, 'Current starter queue as of May 16, 2026:', 'docs/starter-issues.md')
requireIncludes(starterIssues, '15 open `good first issue` issues.', 'docs/starter-issues.md')
requireIncludes(starterIssues, '15 open `first-timers-only` issues.', 'docs/starter-issues.md')
requireIncludes(starterIssues, '15 open `help wanted` issues.', 'docs/starter-issues.md')
requireIncludes(starterIssues, '9 starter issues are code or test tasks.', 'docs/starter-issues.md')
requireIncludes(starterIssues, '6 starter issues are focused docs or integration transcript tasks.', 'docs/starter-issues.md')
requireIncludes(starterIssues, '0 starter issues are currently under active review.', 'docs/starter-issues.md')
requireIncludes(starterIssues, '## Start Here This Week', 'docs/starter-issues.md')
requireIncludes(starterIssues, 'adds the most familiar Node service entry point.', 'docs/starter-issues.md')
requireIncludes(starterIssues, 'connects the WorkPaper proof loop to a common TypeScript agent stack.', 'docs/starter-issues.md')
requireIncludes(starterIssues, '## Code And Test Starters', 'docs/starter-issues.md')
requireIncludes(starterIssues, '## Integration Docs Starters', 'docs/starter-issues.md')
requireIncludes(starterIssues, '#360: test(headless): cover display-value readback after JSON restore', 'docs/starter-issues.md')
requireIncludes(starterIssues, '#361: test(headless): cover range readback after an input edit', 'docs/starter-issues.md')
requireIncludes(
  starterIssues,
  '#362: test(examples): guard the headless README command index against missing scripts',
  'docs/starter-issues.md',
)
requireIncludes(starterIssues, '#363: test(examples): add invalid-request proof to the HTTP JSON summary smoke', 'docs/starter-issues.md')
requireIncludes(starterIssues, '#366: test(headless): cover changed named expressions after WorkPaper restore', 'docs/starter-issues.md')
requireIncludes(starterIssues, '#367: test(headless): cover dense sheet range read with sparse values', 'docs/starter-issues.md')
requireIncludes(starterIssues, '#368: test(headless): cover two-column formula tiling in fill ranges', 'docs/starter-issues.md')
requireIncludes(starterIssues, '#369: test(headless): cover tab-indented formula prefix detection', 'docs/starter-issues.md')
requireIncludes(starterIssues, '#371: test(examples): add deterministic markdown-report output test', 'docs/starter-issues.md')
requireIncludes(starterIssues, '#273: docs(examples): add Express WorkPaper route smoke', 'docs/starter-issues.md')
requireIncludes(starterIssues, '#283: docs(mcp): add Cursor MCP config for the WorkPaper stdio server', 'docs/starter-issues.md')
requireIncludes(
  starterIssues,
  '#285: docs(mcp): add MCP Inspector smoke-test transcript for the WorkPaper server',
  'docs/starter-issues.md',
)
requireIncludes(starterIssues, '#300: docs(examples): add tRPC WorkPaper procedure smoke', 'docs/starter-issues.md')
requireIncludes(starterIssues, '#334: docs(agent): add OpenAI Responses streaming tool-call transcript', 'docs/starter-issues.md')
requireIncludes(starterIssues, '#358: docs(agent): add AI SDK onStepFinish WorkPaper transcript', 'docs/starter-issues.md')
requireIncludes(starterIssues, 'Add `help wanted` only when an external contributor can make progress', 'docs/starter-issues.md')
requireNotIncludes(starterIssues, '115 open `first-timers-only` issues.', 'docs/starter-issues.md')
requireNotIncludes(starterIssues, '#370: test(examples): add malformed CSV fixture check to the csv-shaped smoke', 'docs/starter-issues.md')
requireNotIncludes(starterIssues, 'https://github.com/proompteng/bilig/pull/373', 'docs/starter-issues.md')
requireNotIncludes(starterIssues, 'https://github.com/proompteng/bilig/pull/271', 'docs/starter-issues.md')
requireNotIncludes(starterIssues, 'https://github.com/proompteng/bilig/pull/291', 'docs/starter-issues.md')
requireNotIncludes(starterIssues, 'https://github.com/proompteng/bilig/pull/295', 'docs/starter-issues.md')
requireNotIncludes(starterIssues, 'https://github.com/proompteng/bilig/pull/251', 'docs/starter-issues.md')
requireNotIncludes(starterIssues, 'https://github.com/proompteng/bilig/pull/349', 'docs/starter-issues.md')
requireNotIncludes(starterIssues, 'https://github.com/proompteng/bilig/pull/351', 'docs/starter-issues.md')
requireNotIncludes(starterIssues, 'https://github.com/proompteng/bilig/pull/352', 'docs/starter-issues.md')
requireNotIncludes(starterIssues, 'https://github.com/proompteng/bilig/pull/353', 'docs/starter-issues.md')
requireNotIncludes(starterIssues, 'https://github.com/proompteng/bilig/pull/365', 'docs/starter-issues.md')
requireIncludes(contributing, 'new-contributor-guide.md#first-time-command-checklist', 'CONTRIBUTING.md')
requireIncludes(llms, 'first-patch list capped at 15 scoped issues.', 'docs/llms.txt')
requireIncludes(llms, 'https://github.com/proompteng/bilig/issues/273', 'docs/llms.txt')
requireIncludes(llms, 'https://github.com/proompteng/bilig/issues/283', 'docs/llms.txt')
requireIncludes(llms, 'https://github.com/proompteng/bilig/issues/285', 'docs/llms.txt')
requireIncludes(llms, 'https://github.com/proompteng/bilig/issues/300', 'docs/llms.txt')
requireIncludes(llms, 'https://github.com/proompteng/bilig/issues/334', 'docs/llms.txt')
requireIncludes(llms, 'https://github.com/proompteng/bilig/issues/358', 'docs/llms.txt')
requireNotIncludes(llms, 'https://github.com/proompteng/bilig/issues/272', 'docs/llms.txt')
requireNotIncludes(llms, 'https://github.com/proompteng/bilig/issues/277', 'docs/llms.txt')
requireNotIncludes(llms, 'https://github.com/proompteng/bilig/issues/281', 'docs/llms.txt')
requireIncludes(
  await readFile(join(docsRoot, 'evaluate-excel-formulas-in-node-typescript.md'), 'utf8'),
  'npx tsx eval-node-formulas.ts',
  'docs/evaluate-excel-formulas-in-node-typescript.md',
)
requireIncludes(
  await readFile(join(docsRoot, 'server-side-spreadsheet-automation-node.md'), 'utf8'),
  'npx tsx eval.ts',
  'docs/server-side-spreadsheet-automation-node.md',
)
for (const required of [
  'title: Google Sheets API alternative for local Node workbook execution',
  'That is the boundary. `bilig` is not trying to replace Google Sheets.',
  'npm install @bilig/headless',
  'npx tsx eval.ts',
  '"verified": true',
  'https://developers.google.com/workspace/sheets/api/guides/concepts',
  'https://developers.google.com/workspace/sheets/api/guides/values',
] as const) {
  requireIncludes(googleSheetsApiBoundaryDoc, required, 'docs/google-sheets-api-alternative-node-workpaper.md')
}
requireIncludes(readme, 'Google Sheets API boundary', 'README.md')
requireIncludes(headlessReadme, 'Google Sheets API boundary', 'packages/headless/README.md')
requireIncludes(index, './google-sheets-api-alternative-node-workpaper.html', 'docs/index.html')
requireIncludes(llms, 'https://proompteng.github.io/bilig/google-sheets-api-alternative-node-workpaper.html', 'docs/llms.txt')
requireIncludes(llms, 'https://github.com/proompteng/bilig/blob/main/docs/google-sheets-api-alternative-node-workpaper.md', 'docs/llms.txt')

requireXlsxCorpusVerifierDiscovery(await readFile(join(docsRoot, 'xlsx-corpus-verifier-walkthrough.md'), 'utf8'))
requireIncludes(index, './xlsx-corpus-verifier-walkthrough.html', 'docs/index.html')
requireIncludes(llms, 'https://proompteng.github.io/bilig/xlsx-corpus-verifier-walkthrough.html', 'docs/llms.txt')

const [
  whyAgentsDoc,
  agentToolCallingDoc,
  aiSdkLangChainDoc,
  mcpWorkPaperToolServerDoc,
  mcpSpreadsheetServerDirectoryDoc,
  mcpClientSetupDoc,
  claudeDesktopMcpbDoc,
  agentToolCallLoopDoc,
] = await Promise.all([
  readFile(join(docsRoot, 'why-agents-need-workbook-apis.md'), 'utf8'),
  readFile(join(docsRoot, 'agent-workpaper-tool-calling-recipe.md'), 'utf8'),
  readFile(join(docsRoot, 'vercel-ai-sdk-langchain-spreadsheet-tool.md'), 'utf8'),
  readFile(join(docsRoot, 'mcp-workpaper-tool-server.md'), 'utf8'),
  readFile(join(docsRoot, 'mcp-spreadsheet-server-directory.md'), 'utf8'),
  readFile(join(docsRoot, 'mcp-client-setup.md'), 'utf8'),
  readFile(join(docsRoot, 'claude-desktop-mcpb-workpaper.md'), 'utf8'),
  readFile(join(docsRoot, 'agent-spreadsheet-tool-call-loop.md'), 'utf8'),
])
requireIncludes(
  whyAgentsDoc,
  'description: Why coding agents should edit workbook formulas through a Node.js WorkPaper API',
  'docs/why-agents-need-workbook-apis.md',
)
requireIncludes(
  agentToolCallingDoc,
  'description: Wrap @bilig/headless workbook reads, writes, formula readback, and persistence as deterministic Node.js tools',
  'docs/agent-workpaper-tool-calling-recipe.md',
)
requireIncludes(agentToolCallingDoc, 'OpenAI Responses API Tool Wrapper', 'docs/agent-workpaper-tool-calling-recipe.md')
requireIncludes(
  agentToolCallingDoc,
  'https://developers.openai.com/api/docs/guides/function-calling',
  'docs/agent-workpaper-tool-calling-recipe.md',
)
requireIncludes(agentToolCallingDoc, 'function_call_output', 'docs/agent-workpaper-tool-calling-recipe.md')
requireIncludes(agentToolCallingDoc, 'npm run agent:framework-adapters', 'docs/agent-workpaper-tool-calling-recipe.md')
requireIncludes(
  aiSdkLangChainDoc,
  'description: Wrap @bilig/headless WorkPaper reads, verified edits, formula contracts, and persistence checks as AI SDK, LangChain, Mastra, LlamaIndex.TS, LangGraph.js, CopilotKit, and Cloudflare Agents tools',
  'docs/vercel-ai-sdk-langchain-spreadsheet-tool.md',
)
requireIncludes(aiSdkLangChainDoc, 'npm run agent:framework-adapters', 'docs/vercel-ai-sdk-langchain-spreadsheet-tool.md')
requireIncludes(aiSdkLangChainDoc, 'Mastra `createTool()`', 'docs/vercel-ai-sdk-langchain-spreadsheet-tool.md')
requireIncludes(aiSdkLangChainDoc, 'LlamaIndex.TS tools', 'docs/vercel-ai-sdk-langchain-spreadsheet-tool.md')
requireIncludes(aiSdkLangChainDoc, 'LangGraph.js `ToolNode`', 'docs/vercel-ai-sdk-langchain-spreadsheet-tool.md')
requireIncludes(aiSdkLangChainDoc, 'CopilotKit `useCopilotAction`', 'docs/vercel-ai-sdk-langchain-spreadsheet-tool.md')
requireIncludes(aiSdkLangChainDoc, 'Cloudflare Agents API and agent tools', 'docs/vercel-ai-sdk-langchain-spreadsheet-tool.md')
const agentFrameworkDocs = await Promise.all(
  agentFrameworkDocRequirements.map(async ({ path, includes }) => ({
    path,
    includes,
    content: await readFile(join(repoRoot, path), 'utf8'),
  })),
)
for (const { path, includes, content } of agentFrameworkDocs) {
  for (const required of includes) {
    requireIncludes(content, required, path)
  }
}
requireIncludes(
  mcpWorkPaperToolServerDoc,
  'description: Expose @bilig/headless workbook reads, verified edits, formula contracts, and persistence checks through MCP-style tools/list and tools/call handlers',
  'docs/mcp-workpaper-tool-server.md',
)
requireIncludes(mcpWorkPaperToolServerDoc, 'npm run agent:mcp-tools', 'docs/mcp-workpaper-tool-server.md')
requireIncludes(mcpWorkPaperToolServerDoc, 'npm run --silent agent:mcp-stdio', 'docs/mcp-workpaper-tool-server.md')
requireIncludes(mcpWorkPaperToolServerDoc, 'npm exec --package @bilig/headless -- bilig-workpaper-mcp', 'docs/mcp-workpaper-tool-server.md')
requireIncludes(
  mcpWorkPaperToolServerDoc,
  'npm exec --package @bilig/headless -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --writable',
  'docs/mcp-workpaper-tool-server.md',
)
requireIncludes(mcpWorkPaperToolServerDoc, '`list_sheets`, `read_range`, `read_cell`', 'docs/mcp-workpaper-tool-server.md')
requireIncludes(
  mcpWorkPaperToolServerDoc,
  'WorkPaper JSON back to the same file after `set_cell_contents`',
  'docs/mcp-workpaper-tool-server.md',
)
requireIncludes(mcpWorkPaperToolServerDoc, 'io.github.proompteng/bilig-workpaper', 'docs/mcp-workpaper-tool-server.md')
requireIncludes(
  mcpWorkPaperToolServerDoc,
  'https://registry.modelcontextprotocol.io/v0.1/servers?search=io.github.proompteng%2Fbilig-workpaper',
  'docs/mcp-workpaper-tool-server.md',
)
requireIncludes(mcpWorkPaperToolServerDoc, 'tools/list', 'docs/mcp-workpaper-tool-server.md')
requireIncludes(mcpWorkPaperToolServerDoc, 'tools/call', 'docs/mcp-workpaper-tool-server.md')
requireIncludes(mcpWorkPaperToolServerDoc, 'MCP tool annotations', 'docs/mcp-workpaper-tool-server.md')
requireIncludes(
  mcpWorkPaperToolServerDoc,
  '`read_workpaper_summary` is read-only, idempotent, and closed-world',
  'docs/mcp-workpaper-tool-server.md',
)
requireIncludes(
  mcpWorkPaperToolServerDoc,
  '`set_workpaper_input_cell` mutates the local WorkPaper state, is idempotent',
  'docs/mcp-workpaper-tool-server.md',
)
requireIncludes(
  mcpWorkPaperToolServerDoc,
  'https://modelcontextprotocol.io/specification/2025-06-18/server/tools',
  'docs/mcp-workpaper-tool-server.md',
)
requireIncludes(mcpWorkPaperToolServerDoc, 'https://github.com/proompteng/bilig/discussions/230', 'docs/mcp-workpaper-tool-server.md')
requireIncludes(agentToolCallingDoc, 'https://github.com/proompteng/bilig/discussions/335', 'docs/agent-workpaper-tool-calling-recipe.md')
requireIncludes(mcpWorkPaperToolServerDoc, 'mcp-client-setup.md', 'docs/mcp-workpaper-tool-server.md')
for (const required of [
  'description: Live directory and install status for the Bilig WorkPaper MCP server',
  'npm exec --package @bilig/headless -- bilig-workpaper-mcp',
  'io.github.proompteng/bilig-workpaper',
  'https://registry.modelcontextprotocol.io/v0.1/servers?search=io.github.proompteng%2Fbilig-workpaper',
  'https://glama.ai/mcp/servers/proompteng/bilig',
  'https://github.com/chatmcp/mcpso/issues/2295',
  'https://github.com/cline/mcp-marketplace/issues/1557',
  'https://mcpserver.cc/en?q=bilig',
  'bcdce4e1-3b05-4be2-b611-2a2abb8baf79',
  'https://agentndx.ai/browse?q=bilig',
  'AgentNDX submission was accepted for review on May 13, 2026',
  'https://github.com/YuzeHao2023/Awesome-MCP-Servers/pull/244',
  'https://github.com/toolsdk-ai/toolsdk-mcp-registry/pull/309',
  'https://github.com/ever-works/awesome-mcp-servers-data/pull/4',
  'https://github.com/jmstfv/mcpserve/pull/19',
  'Not indexed in public search on May 14, 2026',
  'https://www.pulsemcp.com/servers?search=bilig&q=bilig',
  'https://github.com/proompteng/bilig/issues/384',
  'marked `@bilig/headless@0.14.26` as the',
  'https://github.com/proompteng/bilig/actions/runs/25949339848',
  'read_workpaper_summary',
  'set_workpaper_input_cell',
]) {
  requireIncludes(mcpSpreadsheetServerDirectoryDoc, required, 'docs/mcp-spreadsheet-server-directory.md')
}
requireIncludes(
  mcpClientSetupDoc,
  'description: Copy-paste MCP client configuration for running the published @bilig/headless WorkPaper stdio server from Claude, Cursor, VS Code, Cline, and Codex.',
  'docs/mcp-client-setup.md',
)
for (const required of [
  'npm exec --package @bilig/headless -- bilig-workpaper-mcp',
  'pnpm mcpb:workpaper:build',
  'claude-desktop-mcpb-workpaper.md',
  'claude mcp add-json bilig-workpaper',
  '.cursor/mcp.json',
  '.vscode/mcp.json',
  'cline_mcp_settings.json',
  '~/.cline/data/settings/cline_mcp_settings.json',
  '[mcp_servers.bilig-workpaper]',
  'https://code.visualstudio.com/docs/copilot/reference/mcp-configuration',
  'https://docs.cline.bot/mcp/adding-and-configuring-servers',
  'https://platform.openai.com/docs/docs-mcp',
]) {
  requireIncludes(mcpClientSetupDoc, required, 'docs/mcp-client-setup.md')
}
requireIncludes(rootPackageJson, '"mcpb:workpaper:build": "tsx scripts/build-workpaper-mcpb.ts"', 'package.json')
for (const required of [
  'description: Build a Claude Desktop MCPB bundle for the published @bilig/headless WorkPaper MCP server',
  'pnpm mcpb:workpaper:build',
  'BILIG_HEADLESS_VERSION=$(npm view @bilig/headless version)',
  'pnpm mcpb:workpaper:build -- --package-version "$BILIG_HEADLESS_VERSION"',
  'build/mcpb/bilig-workpaper.mcpb',
  'open build/mcpb/bilig-workpaper.mcpb',
  'read_workpaper_summary',
  'set_workpaper_input_cell',
  '"entry_point": "server/index.js"',
  'https://registry.modelcontextprotocol.io/v0.1/servers?search=io.github.proompteng%2Fbilig-workpaper',
]) {
  requireIncludes(claudeDesktopMcpbDoc, required, 'docs/claude-desktop-mcpb-workpaper.md')
}
requireGrowthSurfaceDiscovery(communityLaunchPack, llms, productHuntLaunchKit, requireIncludes)
requireNotIncludes(llms, '## launch and feedback', 'docs/llms.txt')
requireNotIncludes(llms, 'conversion-feedback comment after npm download and clone traffic review', 'docs/llms.txt')
requireNotIncludes(llms, 'published dev article source', 'docs/llms.txt')
for (const removedGrowthLink of llmsExternalSurfaceLinks) {
  requireNotIncludes(llms, removedGrowthLink, 'docs/llms.txt')
}
requireIncludes(
  aiSdkLangChainDoc,
  'https://ai-sdk.dev/docs/ai-sdk-core/tools-and-tool-calling',
  'docs/vercel-ai-sdk-langchain-spreadsheet-tool.md',
)
requireIncludes(
  aiSdkLangChainDoc,
  'https://docs.langchain.com/oss/javascript/langchain/tools',
  'docs/vercel-ai-sdk-langchain-spreadsheet-tool.md',
)
requireIncludes(
  agentToolCallLoopDoc,
  'description: A runnable @bilig/headless loop where an agent writes one workbook input',
  'docs/agent-spreadsheet-tool-call-loop.md',
)
for (const [path, content] of [
  ['docs/why-agents-need-workbook-apis.md', whyAgentsDoc],
  ['docs/agent-workpaper-tool-calling-recipe.md', agentToolCallingDoc],
  ['docs/vercel-ai-sdk-langchain-spreadsheet-tool.md', aiSdkLangChainDoc],
  ['docs/mcp-workpaper-tool-server.md', mcpWorkPaperToolServerDoc],
  ['docs/mcp-spreadsheet-server-directory.md', mcpSpreadsheetServerDirectoryDoc],
  ['docs/mcp-client-setup.md', mcpClientSetupDoc],
  ['docs/claude-desktop-mcpb-workpaper.md', claudeDesktopMcpbDoc],
  ['docs/agent-spreadsheet-tool-call-loop.md', agentToolCallLoopDoc],
  ['docs/workbook-automation-examples-node.md', await readFile(join(docsRoot, 'workbook-automation-examples-node.md'), 'utf8')],
  ['docs/server-side-spreadsheet-automation-node.md', await readFile(join(docsRoot, 'server-side-spreadsheet-automation-node.md'), 'utf8')],
  ['docs/google-sheets-api-alternative-node-workpaper.md', googleSheetsApiBoundaryDoc],
  ['docs/node-framework-workpaper-adapters.md', await readFile(join(docsRoot, 'node-framework-workpaper-adapters.md'), 'utf8')],
  ['docs/dev-to-workbook-apis-post.md', await readFile(join(docsRoot, 'dev-to-workbook-apis-post.md'), 'utf8')],
] as const) {
  requireIncludes(content, 'image: /assets/github-social-preview.png', path)
}

requireIncludes(issueTemplateConfig, 'https://github.com/proompteng/bilig/discussions/213', '.github/ISSUE_TEMPLATE/config.yml')
requireIncludes(
  pullRequestTemplate,
  'For public docs or example work, include the page or discussion that a new',
  '.github/PULL_REQUEST_TEMPLATE.md',
)

for (const required of [
  '## Use-Case Chooser',
  'Formula-backed calculations inside a Node service',
  'Agent writeback that must prove the value after an edit',
  'XLSX parsing, export, styling, images, and workbook-file metadata',
  'Persisting a workbook document as JSON and restoring it later',
  'Embedding a spreadsheet UI that users edit directly',
  '[Node quickstart](try-bilig-headless-in-node.md)',
  '[agent tool-calling recipe](agent-workpaper-tool-calling-recipe.md)',
  '[SheetJS and ExcelJS boundary guide](sheetjs-exceljs-alternative-formula-workbook-api.md)',
  '[HyperFormula alternative notes](hyperformula-alternative-headless-workpaper.md)',
  '[documented Excel gaps](where-bilig-is-not-excel-compatible-yet.md)',
]) {
  requireIncludes(headlessSpreadsheetEngineComparison, required, 'docs/headless-spreadsheet-engine-comparison.md')
}

for (const [path, content] of [
  ['docs/sheetjs-exceljs-alternative-formula-workbook-api.md', sheetjsExceljsAlternativeFormulaWorkbookApi],
  ['docs/hyperformula-alternative-headless-workpaper.md', hyperformulaAlternativeHeadlessWorkpaper],
] as const) {
  requireIncludes(
    content,
    '[headless spreadsheet engine use-case chooser](headless-spreadsheet-engine-comparison.md#use-case-chooser)',
    path,
  )
}

for (const required of [
  'title: SheetJS and ExcelJS alternative for formula-backed workbook APIs',
  'canonical_url: https://proompteng.github.io/bilig/sheetjs-exceljs-alternative-formula-workbook-api.html',
  'Research date: 2026-05-14.',
  '## TypeScript Evaluation Path',
  'npm install -D tsx typescript @types/node',
  'const workbook = WorkPaper.buildFromSheets({',
  'workbook.setCellContents({ sheet: inputs, row: 1, col: 1 }, 40)',
  'verified: before === 36864 && after === 46080 && afterRestore === after',
  'SheetJS Pro has a formula calculator component',
  'ExcelJS can store formulas and supplied results',
] as const) {
  requireIncludes(sheetjsExceljsAlternativeFormulaWorkbookApi, required, 'docs/sheetjs-exceljs-alternative-formula-workbook-api.md')
}

requireIncludes(
  await readFile(join(docsRoot, 'node-spreadsheet-formula-engine.md'), 'utf8'),
  'cat > formula-engine-smoke.ts',
  'docs/node-spreadsheet-formula-engine.md',
)

const discussionDocs = {
  readme: ['README.md', readme],
  headless: ['packages/headless/README.md', headlessReadme],
  agent: ['docs/agent-workpaper-tool-calling-recipe.md', agentToolCallingDoc],
  index: ['docs/index.html', index],
  launch: ['docs/community-launch-pack.md', communityLaunchPack],
  llms: ['docs/llms.txt', llms],
  mcp: ['docs/mcp-workpaper-tool-server.md', mcpWorkPaperToolServerDoc],
} as const

const discussionDocChecks = [
  ['https://github.com/proompteng/bilig/discussions/157', ['readme', 'headless', 'index', 'launch', 'llms']],
  ['https://github.com/proompteng/bilig/discussions/213', ['readme', 'launch', 'llms']],
  ['https://github.com/proompteng/bilig/discussions/230', ['mcp', 'llms']],
  ['https://github.com/proompteng/bilig/discussions/167', ['index', 'launch', 'llms']],
  ['https://github.com/proompteng/bilig/discussions/307', ['readme', 'headless', 'index', 'launch', 'llms']],
  ['https://github.com/proompteng/bilig/discussions/308', ['readme', 'headless', 'launch', 'llms']],
  ['https://github.com/proompteng/bilig/discussions/335', ['readme', 'headless', 'agent', 'launch', 'llms']],
  ['https://github.com/proompteng/bilig/discussions/340', ['readme', 'headless', 'index', 'launch', 'llms']],
  ['https://github.com/proompteng/bilig/discussions/382', ['launch', 'llms']],
] as const

for (const [url, docKeys] of discussionDocChecks) {
  for (const docKey of docKeys) {
    const [path, content] = discussionDocs[docKey]
    requireIncludes(content, url, path)
  }
}

requireStarterIssueDiscovery(starterIssues, llms)

const [headlessExampleReadme, headlessExamplePackage, headlessPackageManifest, headlessServerJson] = await Promise.all([
  readFile(join(repoRoot, 'examples', 'headless-workpaper', 'README.md'), 'utf8'),
  readFile(join(repoRoot, 'examples', 'headless-workpaper', 'package.json'), 'utf8'),
  readFile(join(repoRoot, 'packages', 'headless', 'package.json'), 'utf8'),
  readFile(join(repoRoot, 'packages', 'headless', 'server.json'), 'utf8'),
])
await requireNpmEvalDiscovery(repoRoot, docsRoot, readme, headlessReadme, headlessExampleReadme)
await requireOpenAiResponsesDiscovery({
  repoRoot,
  docsRoot,
  readme,
  headlessReadme,
  index,
  llms,
  agentToolCallingDoc,
  headlessExampleReadme,
  headlessExamplePackage,
})
await requireAiSdkDiscovery({
  repoRoot,
  docsRoot,
  readme,
  headlessReadme,
  index,
  llms,
  agentToolCallingDoc,
  aiSdkLangChainDoc,
  headlessExampleReadme,
  headlessExamplePackage,
})
await requireFile(join(repoRoot, 'examples', 'headless-workpaper', 'agent-framework-adapters.ts'))
await requireFile(join(repoRoot, 'examples', 'headless-workpaper', 'mcp-tool-server.ts'))
await requireFile(join(repoRoot, 'examples', 'headless-workpaper', 'mcp-stdio-server.ts'))
requireDocumentedScriptsExist(headlessExampleReadme, headlessExamplePackage, 'examples/headless-workpaper/README.md')
requireIncludes(headlessExampleReadme, 'npm run invoice-totals', 'examples/headless-workpaper/README.md')
requireIncludes(headlessExampleReadme, '## Invoice Totals', 'examples/headless-workpaper/README.md')
requireIncludes(headlessExampleReadme, 'npm run csv-shaped', 'examples/headless-workpaper/README.md')
requireIncludes(headlessExampleReadme, '## CSV Shaped Input', 'examples/headless-workpaper/README.md')
requireIncludes(headlessReadme, 'npm run csv-shaped', 'packages/headless/README.md')
requireIncludes(headlessReadme, 'CSV shaped input', 'packages/headless/README.md')
requireIncludes(headlessExampleReadme, 'npm run budget-variance', 'examples/headless-workpaper/README.md')
requireIncludes(headlessExampleReadme, '## Budget Variance Alerts', 'examples/headless-workpaper/README.md')
requireIncludes(headlessExampleReadme, 'npm run fulfillment-capacity', 'examples/headless-workpaper/README.md')
requireIncludes(headlessExampleReadme, '## Fulfillment Capacity Plan', 'examples/headless-workpaper/README.md')
requireIncludes(headlessExampleReadme, 'npm run quote-approval', 'examples/headless-workpaper/README.md')
requireIncludes(headlessExampleReadme, '## Quote Approval Threshold', 'examples/headless-workpaper/README.md')
requireIncludes(headlessExampleReadme, 'npm run subscription-mrr', 'examples/headless-workpaper/README.md')
requireIncludes(headlessExampleReadme, '## Subscription MRR Forecast', 'examples/headless-workpaper/README.md')
requireIncludes(headlessExampleReadme, 'npm run agent:framework-adapters', 'examples/headless-workpaper/README.md')
requireIncludes(headlessExampleReadme, '## Agent Framework Adapters', 'examples/headless-workpaper/README.md')
requireIncludes(headlessExampleReadme, 'npm run agent:mcp-tools', 'examples/headless-workpaper/README.md')
requireIncludes(headlessExampleReadme, 'npm run agent:mcp-stdio', 'examples/headless-workpaper/README.md')
requireIncludes(headlessExampleReadme, '## MCP Tool Server Shape', 'examples/headless-workpaper/README.md')
requireIncludes(headlessExampleReadme, '## MCP Stdio Server', 'examples/headless-workpaper/README.md')
requireIncludes(headlessExampleReadme, 'annotations.', 'examples/headless-workpaper/README.md')
requireIncludes(headlessExampleReadme, 'read tool is annotated as read-only', 'examples/headless-workpaper/README.md')
requireIncludes(
  headlessExamplePackage,
  '"agent:framework-adapters": "tsx agent-framework-adapters.ts"',
  'examples/headless-workpaper/package.json',
)
requireIncludes(headlessExamplePackage, '"agent:mcp-tools": "tsx mcp-tool-server.ts"', 'examples/headless-workpaper/package.json')
requireIncludes(headlessExamplePackage, '"agent:mcp-stdio": "tsx mcp-stdio-server.ts"', 'examples/headless-workpaper/package.json')
await requireServerlessWorkPaperApiDiscovery({
  repoRoot,
  docsRoot,
  readme,
  headlessReadme,
  llms,
})
requireIncludes(headlessPackageManifest, '"mcpName": "io.github.proompteng/bilig-workpaper"', 'packages/headless/package.json')
requireIncludes(headlessPackageManifest, '"bilig-workpaper-mcp": "./dist/work-paper-mcp-stdio-bin.js"', 'packages/headless/package.json')
requireIncludes(headlessReadme, 'bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --writable', 'packages/headless/README.md')
requireIncludes(headlessReadme, '`set_cell_contents` edits back to the same file', 'packages/headless/README.md')
requireIncludes(readme, 'bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --writable', 'README.md')
requireIncludes(readme, '`export_workpaper_document`, and `validate_formula`', 'README.md')
requireIncludes(headlessServerJson, '"name": "io.github.proompteng/bilig-workpaper"', 'packages/headless/server.json')
requireIncludes(headlessServerJson, '"identifier": "@bilig/headless"', 'packages/headless/server.json')
requireIncludes(headlessReadme, 'npm run invoice-totals', 'packages/headless/README.md')
requireIncludes(headlessReadme, '#invoice-totals', 'packages/headless/README.md')
requireIncludes(headlessReadme, 'npm run budget-variance', 'packages/headless/README.md')
requireIncludes(headlessReadme, '#budget-variance-alerts', 'packages/headless/README.md')
requireIncludes(headlessReadme, 'npm run fulfillment-capacity', 'packages/headless/README.md')
requireIncludes(headlessReadme, '#fulfillment-capacity-plan', 'packages/headless/README.md')
requireIncludes(headlessReadme, 'npm run quote-approval', 'packages/headless/README.md')
requireIncludes(headlessReadme, '#quote-approval-threshold', 'packages/headless/README.md')
requireIncludes(headlessReadme, 'npm run subscription-mrr', 'packages/headless/README.md')
requireIncludes(headlessReadme, '#subscription-mrr-forecast', 'packages/headless/README.md')

for (const required of [
  '## Clean npm Sanity Check',
  'mkdir bilig-headless-sanity',
  'npx tsx sanity.ts',
  'curl -fsSLo sanity.ts https://proompteng.github.io/bilig/npm-eval.ts',
  'afterRestore',
  'matching `after`/`afterRestore` values are the check.',
]) {
  requireIncludes(headlessReadme, required, 'packages/headless/README.md')
}

console.log(
  JSON.stringify(
    {
      ok: true,
      sitemapUrlCount: actualSitemapUrls.length,
      robots: 'ok',
      llms: 'ok',
    },
    null,
    2,
  ),
)
