import { readFile, stat } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const repoRoot = join(dirname(fileURLToPath(import.meta.url)), '..')
const docsRoot = join(repoRoot, 'docs')
const siteRoot = 'https://proompteng.github.io/bilig/'

const expectedSitemapUrls = [
  siteRoot,
  `${siteRoot}why-agents-need-workbook-apis.html`,
  `${siteRoot}agent-workpaper-tool-calling-recipe.html`,
  `${siteRoot}vercel-ai-sdk-langchain-spreadsheet-tool.html`,
  `${siteRoot}mcp-workpaper-tool-server.html`,
  `${siteRoot}agent-spreadsheet-tool-call-loop.html`,
  `${siteRoot}node-service-workpaper-recipe.html`,
  `${siteRoot}node-spreadsheet-formula-engine.html`,
  `${siteRoot}evaluate-excel-formulas-in-node-typescript.html`,
  `${siteRoot}try-bilig-headless-in-node.html`,
  `${siteRoot}workbook-automation-examples-node.html`,
  `${siteRoot}serverless-workpaper-api-route.html`,
  `${siteRoot}csv-shaped-workpaper-input-recipe.html`,
  `${siteRoot}unsupported-formula-troubleshooting-recipe.html`,
  `${siteRoot}local-workpaper-benchmark-walkthrough.html`,
  `${siteRoot}building-a-revenue-model-with-headless-workpaper.html`,
  `${siteRoot}persisting-formula-backed-workpaper-documents-in-node.html`,
  `${siteRoot}what-workpaper-benchmark-proves.html`,
  `${siteRoot}hyperformula-alternative-headless-workpaper.html`,
  `${siteRoot}headless-spreadsheet-engine-comparison.html`,
  `${siteRoot}sheetjs-exceljs-alternative-formula-workbook-api.html`,
  `${siteRoot}where-bilig-is-not-excel-compatible-yet.html`,
  `${siteRoot}xlsx-corpus-verifier-walkthrough.html`,
  `${siteRoot}formula-edge-xlookup-exact-fixture.html`,
  `${siteRoot}formula-edge-sumifs-paired-criteria-fixture.html`,
  `${siteRoot}formula-edge-groupby-spill-fixture.html`,
  `${siteRoot}starter-issues.html`,
  `${siteRoot}community-launch-pack.html`,
  `${siteRoot}community-growth-snapshot.html`,
  `${siteRoot}llms.txt`,
] as const

const sourceFilesByUrl = new Map<string, string>([
  [siteRoot, 'index.html'],
  [`${siteRoot}why-agents-need-workbook-apis.html`, 'why-agents-need-workbook-apis.md'],
  [`${siteRoot}agent-workpaper-tool-calling-recipe.html`, 'agent-workpaper-tool-calling-recipe.md'],
  [`${siteRoot}vercel-ai-sdk-langchain-spreadsheet-tool.html`, 'vercel-ai-sdk-langchain-spreadsheet-tool.md'],
  [`${siteRoot}mcp-workpaper-tool-server.html`, 'mcp-workpaper-tool-server.md'],
  [`${siteRoot}agent-spreadsheet-tool-call-loop.html`, 'agent-spreadsheet-tool-call-loop.md'],
  [`${siteRoot}node-service-workpaper-recipe.html`, 'node-service-workpaper-recipe.md'],
  [`${siteRoot}node-spreadsheet-formula-engine.html`, 'node-spreadsheet-formula-engine.md'],
  [`${siteRoot}evaluate-excel-formulas-in-node-typescript.html`, 'evaluate-excel-formulas-in-node-typescript.md'],
  [`${siteRoot}try-bilig-headless-in-node.html`, 'try-bilig-headless-in-node.md'],
  [`${siteRoot}workbook-automation-examples-node.html`, 'workbook-automation-examples-node.md'],
  [`${siteRoot}serverless-workpaper-api-route.html`, 'serverless-workpaper-api-route.md'],
  [`${siteRoot}csv-shaped-workpaper-input-recipe.html`, 'csv-shaped-workpaper-input-recipe.md'],
  [`${siteRoot}unsupported-formula-troubleshooting-recipe.html`, 'unsupported-formula-troubleshooting-recipe.md'],
  [`${siteRoot}local-workpaper-benchmark-walkthrough.html`, 'local-workpaper-benchmark-walkthrough.md'],
  [`${siteRoot}building-a-revenue-model-with-headless-workpaper.html`, 'building-a-revenue-model-with-headless-workpaper.md'],
  [`${siteRoot}persisting-formula-backed-workpaper-documents-in-node.html`, 'persisting-formula-backed-workpaper-documents-in-node.md'],
  [`${siteRoot}what-workpaper-benchmark-proves.html`, 'what-workpaper-benchmark-proves.md'],
  [`${siteRoot}hyperformula-alternative-headless-workpaper.html`, 'hyperformula-alternative-headless-workpaper.md'],
  [`${siteRoot}headless-spreadsheet-engine-comparison.html`, 'headless-spreadsheet-engine-comparison.md'],
  [`${siteRoot}sheetjs-exceljs-alternative-formula-workbook-api.html`, 'sheetjs-exceljs-alternative-formula-workbook-api.md'],
  [`${siteRoot}where-bilig-is-not-excel-compatible-yet.html`, 'where-bilig-is-not-excel-compatible-yet.md'],
  [`${siteRoot}xlsx-corpus-verifier-walkthrough.html`, 'xlsx-corpus-verifier-walkthrough.md'],
  [`${siteRoot}formula-edge-xlookup-exact-fixture.html`, 'formula-edge-xlookup-exact-fixture.md'],
  [`${siteRoot}formula-edge-sumifs-paired-criteria-fixture.html`, 'formula-edge-sumifs-paired-criteria-fixture.md'],
  [`${siteRoot}formula-edge-groupby-spill-fixture.html`, 'formula-edge-groupby-spill-fixture.md'],
  [`${siteRoot}starter-issues.html`, 'starter-issues.md'],
  [`${siteRoot}community-launch-pack.html`, 'community-launch-pack.md'],
  [`${siteRoot}community-growth-snapshot.html`, 'community-growth-snapshot.md'],
  [`${siteRoot}llms.txt`, 'llms.txt'],
])

function requireIncludes(haystack: string, needle: string, context: string): void {
  if (!haystack.includes(needle)) {
    throw new Error(`${context} is missing ${needle}`)
  }
}

function requireNotIncludes(haystack: string, needle: string, context: string): void {
  if (haystack.includes(needle)) {
    throw new Error(`${context} must not include ${needle}`)
  }
}

async function requireFile(path: string): Promise<void> {
  const info = await stat(path)
  if (!info.isFile()) {
    throw new Error(`${path} is not a file`)
  }
}

function getFrontMatter(content: string): string | undefined {
  if (!content.startsWith('---\n')) {
    return undefined
  }

  const end = content.indexOf('\n---', 4)
  if (end === -1) {
    return undefined
  }

  return content.slice(4, end)
}

async function requirePublishedSource(path: string): Promise<void> {
  await requireFile(path)

  if (!path.endsWith('.md')) {
    return
  }

  const frontMatter = getFrontMatter(await readFile(path, 'utf8'))
  if (frontMatter !== undefined && /^published:\s*false\s*$/m.test(frontMatter)) {
    throw new Error(`${path} is listed in the sitemap but has published: false`)
  }
}

function extractSitemapUrls(sitemap: string): string[] {
  return [...sitemap.matchAll(/<loc>([^<]+)<\/loc>/g)].map((match) => match[1] ?? '')
}

function extractNpmRunScripts(readme: string): string[] {
  const scripts = new Set<string>()

  // Match the command form used throughout the headless example README,
  // including optional npm flags such as `npm run --silent script-name`.
  for (const match of readme.matchAll(/\bnpm\s+run(?:\s+--[\w-]+)*\s+([\w:-]+)/g)) {
    const script = match[1]
    if (script !== undefined) {
      scripts.add(script)
    }
  }

  return [...scripts].toSorted()
}

function getPackageScripts(packageJson: string, context: string): Record<string, unknown> {
  const manifest: unknown = JSON.parse(packageJson)

  if (typeof manifest !== 'object' || manifest === null || !('scripts' in manifest)) {
    throw new Error(`${context} is missing a scripts object`)
  }

  const { scripts } = manifest
  if (typeof scripts !== 'object' || scripts === null || Array.isArray(scripts)) {
    throw new Error(`${context} scripts must be an object`)
  }

  return scripts
}

function requireDocumentedScriptsExist(readme: string, packageJson: string, context: string): void {
  const scripts = getPackageScripts(packageJson, 'examples/headless-workpaper/package.json')

  for (const documentedScript of extractNpmRunScripts(readme)) {
    if (!(documentedScript in scripts)) {
      throw new Error(`${context} documents missing package.json script: npm run ${documentedScript}`)
    }
  }
}

const [
  readme,
  contributing,
  index,
  siteCss,
  robots,
  sitemap,
  llms,
  starterIssues,
  newContributorGuide,
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
] = await Promise.all([
  readFile(join(repoRoot, 'README.md'), 'utf8'),
  readFile(join(repoRoot, 'CONTRIBUTING.md'), 'utf8'),
  readFile(join(docsRoot, 'index.html'), 'utf8'),
  readFile(join(docsRoot, 'assets', 'site.css'), 'utf8'),
  readFile(join(docsRoot, 'robots.txt'), 'utf8'),
  readFile(join(docsRoot, 'sitemap.xml'), 'utf8'),
  readFile(join(docsRoot, 'llms.txt'), 'utf8'),
  readFile(join(docsRoot, 'starter-issues.md'), 'utf8'),
  readFile(join(docsRoot, 'new-contributor-guide.md'), 'utf8'),
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
])

requireIncludes(index, '<link rel="canonical" href="https://proompteng.github.io/bilig/" />', 'docs/index.html')
requireIncludes(
  index,
  '<link rel="sitemap" type="application/xml" href="https://proompteng.github.io/bilig/sitemap.xml" />',
  'docs/index.html',
)
requireIncludes(
  index,
  '<link rel="alternate" type="text/plain" href="https://proompteng.github.io/bilig/llms.txt" title="llms.txt" />',
  'docs/index.html',
)
requireIncludes(index, '"@type": "SoftwareSourceCode"', 'docs/index.html')
requireIncludes(index, '"codeRepository": "https://github.com/proompteng/bilig"', 'docs/index.html')
requireIncludes(index, '<title>bilig - Headless Spreadsheet Engine for Node.js Services and Agents</title>', 'docs/index.html')
requireIncludes(index, '<meta name="robots" content="index, follow, max-image-preview:large" />', 'docs/index.html')
requireIncludes(index, '<link rel="icon" type="image/svg+xml" href="./assets/favicon.svg" />', 'docs/index.html')
requireIncludes(index, '<link rel="stylesheet" href="./assets/site.css?v=2026-05-13-6" />', 'docs/index.html')
requireIncludes(index, '<link rel="stylesheet" href="./assets/product-demo.css?v=2026-05-13-1" />', 'docs/index.html')
requireNotIncludes(index, './assets/fonts.css', 'docs/index.html')
requireNotIncludes(index, 'bilig-hero-workbook-api.png?v=2026-05-08-2', 'docs/index.html')
requireNotIncludes(siteCss, 'bilig-hero-workbook-api.png?v=2026-05-08-2', 'docs/assets/site.css')
requireIncludes(index, 'Revenue.workpaper', 'docs/index.html')
requireIncludes(index, 'Build a workbook in Node, change inputs through code', 'docs/index.html')
requireIncludes(index, '<strong>55 small starter issues are open.</strong>', 'docs/index.html')
requireIncludes(index, '<strong>0.13.19</strong>', 'docs/index.html')
requireIncludes(index, '<span>Open first-timer issues</span>', 'docs/index.html')
requireIncludes(index, '<strong>55</strong>', 'docs/index.html')
requireNotIncludes(index, '<strong>40 starter tasks</strong>', 'docs/index.html')
requireNotIncludes(index, '<strong>0.13.9</strong>', 'docs/index.html')
requireIncludes(index, '"downloadUrl": "https://www.npmjs.com/package/@bilig/headless"', 'docs/index.html')
requireIncludes(index, '"applicationCategory": "DeveloperApplication"', 'docs/index.html')
requireIncludes(index, '"@type": "FAQPage"', 'docs/index.html')
for (const required of [
  './why-agents-need-workbook-apis.html',
  './agent-workpaper-tool-calling-recipe.html',
  './vercel-ai-sdk-langchain-spreadsheet-tool.html',
  './mcp-workpaper-tool-server.html',
  './agent-spreadsheet-tool-call-loop.html',
  './node-service-workpaper-recipe.html',
  './node-spreadsheet-formula-engine.html',
  './evaluate-excel-formulas-in-node-typescript.html',
  './try-bilig-headless-in-node.html',
  './serverless-workpaper-api-route.html',
  './persisting-formula-backed-workpaper-documents-in-node.html',
  './building-a-revenue-model-with-headless-workpaper.html',
  './headless-spreadsheet-engine-comparison.html',
  './hyperformula-alternative-headless-workpaper.html',
]) {
  requireIncludes(index, required, 'docs/index.html')
}

requireIncludes(robots, 'User-agent: *', 'docs/robots.txt')
requireIncludes(robots, 'Allow: /', 'docs/robots.txt')
requireIncludes(robots, `Sitemap: ${siteRoot}sitemap.xml`, 'docs/robots.txt')

const actualSitemapUrls = extractSitemapUrls(sitemap)
if (actualSitemapUrls.length !== expectedSitemapUrls.length) {
  throw new Error(`sitemap has ${String(actualSitemapUrls.length)} urls, expected ${String(expectedSitemapUrls.length)}`)
}

const sourceFilesToVerify: string[] = []

for (const expectedUrl of expectedSitemapUrls) {
  if (!actualSitemapUrls.includes(expectedUrl)) {
    throw new Error(`sitemap is missing ${expectedUrl}`)
  }

  const sourceFile = sourceFilesByUrl.get(expectedUrl)
  if (sourceFile === undefined) {
    throw new Error(`no source file mapping for ${expectedUrl}`)
  }
  sourceFilesToVerify.push(sourceFile)
}

await Promise.all(sourceFilesToVerify.map((sourceFile) => requirePublishedSource(join(docsRoot, sourceFile))))
await Promise.all(
  ['README.md', 'package.json', 'route.ts', 'smoke.ts'].map((sourceFile) =>
    requireFile(join(repoRoot, 'examples', 'serverless-workpaper-api', sourceFile)),
  ),
)
await Promise.all(
  ['github-social-preview.png', 'workpaper-benchmark-card.png'].map((sourceFile) => requireFile(join(docsRoot, 'assets', sourceFile))),
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

for (const url of actualSitemapUrls) {
  if (!url.startsWith(siteRoot)) {
    throw new Error(`sitemap url is outside ${siteRoot}: ${url}`)
  }
}

for (const required of [
  'repository: https://github.com/proompteng/bilig',
  'npm package: https://www.npmjs.com/package/@bilig/headless',
  'npm run agent:tool-call',
  'npm run agent:framework-adapters',
  'npm run agent:verify',
  'https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#json-records-input',
  'https://proompteng.github.io/bilig/why-agents-need-workbook-apis.html',
  'https://proompteng.github.io/bilig/try-bilig-headless-in-node.html',
  'https://proompteng.github.io/bilig/vercel-ai-sdk-langchain-spreadsheet-tool.html',
  'https://proompteng.github.io/bilig/mcp-workpaper-tool-server.html',
  'https://proompteng.github.io/bilig/agent-workpaper-tool-calling-recipe.html',
  'https://proompteng.github.io/bilig/agent-spreadsheet-tool-call-loop.html',
  'https://proompteng.github.io/bilig/node-service-workpaper-recipe.html',
  'https://proompteng.github.io/bilig/serverless-workpaper-api-route.html',
  'https://proompteng.github.io/bilig/workbook-automation-examples-node.html',
  'https://github.com/proompteng/bilig/blob/main/docs/workbook-automation-examples-node.md',
  'https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#invoice-totals',
  'https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#budget-variance-alerts',
  'https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#fulfillment-capacity-plan',
  'https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#quote-approval-threshold',
  'https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#subscription-mrr-forecast',
  'https://github.com/proompteng/bilig/tree/main/examples/serverless-workpaper-api',
  'https://github.com/proompteng/bilig/discussions',
  'https://github.com/proompteng/bilig/discussions/157',
  'https://github.com/proompteng/bilig/discussions/167',
  'https://github.com/proompteng/bilig/discussions/230',
  'https://github.com/proompteng/bilig/discussions/270',
  'https://github.com/proompteng/bilig/discussions/115',
  'https://github.com/proompteng/bilig/blob/main/docs/dev-to-workbook-apis-post.md',
  'https://proompteng.github.io/bilig/node-spreadsheet-formula-engine.html',
  'https://proompteng.github.io/bilig/evaluate-excel-formulas-in-node-typescript.html',
  'https://github.com/proompteng/bilig/blob/main/docs/node-spreadsheet-formula-engine.md',
  'https://github.com/proompteng/bilig/blob/main/docs/evaluate-excel-formulas-in-node-typescript.md',
  'https://github.com/proompteng/bilig/blob/main/docs/node-service-workpaper-recipe.md',
  'https://github.com/proompteng/bilig/blob/main/docs/serverless-workpaper-api-route.md',
  'https://github.com/proompteng/bilig/blob/main/docs/csv-shaped-workpaper-input-recipe.md',
  'https://github.com/proompteng/bilig/blob/main/docs/unsupported-formula-troubleshooting-recipe.md',
  'https://github.com/proompteng/bilig/blob/main/docs/agent-workpaper-tool-calling-recipe.md',
  'https://github.com/proompteng/bilig/blob/main/docs/vercel-ai-sdk-langchain-spreadsheet-tool.md',
  'https://github.com/proompteng/bilig/blob/main/docs/mcp-workpaper-tool-server.md',
  'https://github.com/proompteng/bilig/blob/main/examples/headless-workpaper/mcp-tool-server.ts',
  'https://github.com/proompteng/bilig/blob/main/examples/headless-workpaper/mcp-stdio-server.ts',
  'https://github.com/proompteng/bilig/blob/main/examples/headless-workpaper/agent-framework-adapters.ts',
  'https://github.com/proompteng/bilig/blob/main/docs/agent-spreadsheet-tool-call-loop.md',
  'https://github.com/proompteng/bilig/blob/main/docs/local-workpaper-benchmark-walkthrough.md',
  'https://github.com/proompteng/bilig/blob/main/docs/what-workpaper-benchmark-proves.md',
  'https://github.com/proompteng/bilig/blob/main/docs/hyperformula-alternative-headless-workpaper.md',
  'https://github.com/proompteng/bilig/blob/main/docs/headless-spreadsheet-engine-comparison.md',
  'https://github.com/proompteng/bilig/blob/main/docs/sheetjs-exceljs-alternative-formula-workbook-api.md',
  'https://github.com/proompteng/bilig/blob/main/docs/where-bilig-is-not-excel-compatible-yet.md',
  'https://github.com/proompteng/bilig/blob/main/docs/formula-edge-sumifs-paired-criteria-fixture.md',
  'https://github.com/proompteng/bilig/blob/main/docs/formula-edge-groupby-spill-fixture.md',
  'https://github.com/proompteng/bilig/blob/main/docs/new-contributor-guide.md',
  'https://github.com/proompteng/bilig/blob/main/docs/starter-issues.md',
  'https://github.com/proompteng/bilig/blob/main/docs/community-launch-pack.md',
  'https://github.com/proompteng/bilig/blob/main/docs/community-growth-snapshot.md',
  'https://github.com/proompteng/bilig/issues?q=is%3Aissue%20state%3Aopen%20label%3Afirst-timers-only',
]) {
  requireIncludes(llms, required, 'docs/llms.txt')
}

for (const [path, content] of [
  ['README.md', readme],
  ['packages/headless/README.md', headlessReadme],
  ['CONTRIBUTING.md', contributing],
  ['docs/new-contributor-guide.md', newContributorGuide],
  ['docs/starter-issues.md', starterIssues],
  ['docs/llms.txt', llms],
] as const) {
  requireIncludes(content, 'https://github.com/proompteng/bilig/issues?q=is%3Aissue%20state%3Aopen%20label%3Afirst-timers-only', path)
}

for (const [path, content] of [
  ['README.md', readme],
  ['packages/headless/README.md', headlessReadme],
] as const) {
  requireIncludes(content, '## Current Public Proof', path)
  requireIncludes(content, 'https://proompteng.github.io/bilig/community-growth-snapshot.html', path)
  requireIncludes(content, 'https://github.com/proompteng/bilig/stargazers', path)
  requireIncludes(content, '`10` forks', path)
  requireIncludes(content, '15,592` npm downloads in the', path)
  requireIncludes(content, '`55` open', path)
  requireIncludes(content, '`good first issue` tickets', path)
}

requireIncludes(newContributorGuide, '## First-Time Command Checklist', 'docs/new-contributor-guide.md')
requireIncludes(newContributorGuide, 'pnpm docs:discovery:check', 'docs/new-contributor-guide.md')
requireIncludes(newContributorGuide, 'pnpm format:check', 'docs/new-contributor-guide.md')
requireIncludes(newContributorGuide, 'pnpm lint', 'docs/new-contributor-guide.md')
requireIncludes(starterIssues, 'new-contributor-guide.md#first-time-command-checklist', 'docs/starter-issues.md')
requireIncludes(starterIssues, 'https://github.com/proompteng/bilig/blob/main/CONTRIBUTING.md', 'docs/starter-issues.md')
requireIncludes(starterIssues, '55 open `first-timers-only` issues.', 'docs/starter-issues.md')
requireIncludes(starterIssues, '49 issues are generally available for a new contributor to claim.', 'docs/starter-issues.md')
requireIncludes(starterIssues, '### npm Smoke Test Improvements', 'docs/starter-issues.md')
requireIncludes(starterIssues, 'https://github.com/proompteng/bilig/issues/265', 'docs/starter-issues.md')
requireIncludes(starterIssues, 'https://github.com/proompteng/bilig/issues/269', 'docs/starter-issues.md')
requireIncludes(starterIssues, '#269: docs(headless): add package-manager variants for the smoke test', 'docs/starter-issues.md')
requireIncludes(starterIssues, '#272: docs(examples): add NestJS WorkPaper controller smoke', 'docs/starter-issues.md')
requireIncludes(starterIssues, '#277: docs(examples): add Bun.serve WorkPaper route smoke', 'docs/starter-issues.md')
requireIncludes(starterIssues, '#278: docs(examples): add SvelteKit WorkPaper route smoke', 'docs/starter-issues.md')
requireIncludes(starterIssues, '#281: docs(examples): add Cloudflare D1 WorkPaper persistence smoke', 'docs/starter-issues.md')
requireIncludes(starterIssues, '#266: docs(try): add a Docker smoke test for clean Node 24 runs', 'docs/starter-issues.md')
requireIncludes(starterIssues, 'https://github.com/proompteng/bilig/pull/271', 'docs/starter-issues.md')
requireNotIncludes(starterIssues, 'https://github.com/proompteng/bilig/pull/251', 'docs/starter-issues.md')
requireIncludes(contributing, 'new-contributor-guide.md#first-time-command-checklist', 'CONTRIBUTING.md')
requireIncludes(llms, '55 open first-timers-only issues, 49 generally available, 6 already in review', 'docs/llms.txt')
requireIncludes(llms, 'https://github.com/proompteng/bilig/issues/272', 'docs/llms.txt')
requireIncludes(llms, 'https://github.com/proompteng/bilig/issues/277', 'docs/llms.txt')
requireIncludes(llms, 'https://github.com/proompteng/bilig/issues/281', 'docs/llms.txt')
requireIncludes(
  await readFile(join(docsRoot, 'evaluate-excel-formulas-in-node-typescript.md'), 'utf8'),
  'npx tsx eval-node-formulas.ts',
  'docs/evaluate-excel-formulas-in-node-typescript.md',
)

for (const [path, content] of [
  ['README.md', readme],
  ['packages/headless/README.md', headlessReadme],
  ['docs/index.html', index],
  ['docs/community-launch-pack.md', await readFile(join(docsRoot, 'community-launch-pack.md'), 'utf8')],
  ['docs/llms.txt', llms],
  ['.github/ISSUE_TEMPLATE/config.yml', issueTemplateConfig],
  ['.github/ISSUE_TEMPLATE.md', issueTemplateRoot],
  ['.github/ISSUE_TEMPLATE/feature_request.yml', featureRequestTemplate],
  ['.github/DISCUSSION_TEMPLATE/ideas.yml', ideasDiscussionTemplate],
  ['.github/DISCUSSION_TEMPLATE/q-a.yml', qaDiscussionTemplate],
  ['.github/DISCUSSION_TEMPLATE/show-and-tell.yml', showAndTellDiscussionTemplate],
] as const) {
  requireIncludes(content, 'workbook-automation-examples-node', path)
}

const [whyAgentsDoc, agentToolCallingDoc, aiSdkLangChainDoc, mcpWorkPaperToolServerDoc, agentToolCallLoopDoc] = await Promise.all([
  readFile(join(docsRoot, 'why-agents-need-workbook-apis.md'), 'utf8'),
  readFile(join(docsRoot, 'agent-workpaper-tool-calling-recipe.md'), 'utf8'),
  readFile(join(docsRoot, 'vercel-ai-sdk-langchain-spreadsheet-tool.md'), 'utf8'),
  readFile(join(docsRoot, 'mcp-workpaper-tool-server.md'), 'utf8'),
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
requireIncludes(agentToolCallingDoc, 'npm run agent:framework-adapters', 'docs/agent-workpaper-tool-calling-recipe.md')
requireIncludes(
  aiSdkLangChainDoc,
  'description: Wrap @bilig/headless WorkPaper reads, verified edits, formula contracts, and persistence checks as Vercel AI SDK and LangChain-style tools',
  'docs/vercel-ai-sdk-langchain-spreadsheet-tool.md',
)
requireIncludes(aiSdkLangChainDoc, 'npm run agent:framework-adapters', 'docs/vercel-ai-sdk-langchain-spreadsheet-tool.md')
requireIncludes(
  mcpWorkPaperToolServerDoc,
  'description: Expose @bilig/headless workbook reads, verified edits, formula contracts, and persistence checks through MCP-style tools/list and tools/call handlers',
  'docs/mcp-workpaper-tool-server.md',
)
requireIncludes(mcpWorkPaperToolServerDoc, 'npm run agent:mcp-tools', 'docs/mcp-workpaper-tool-server.md')
requireIncludes(mcpWorkPaperToolServerDoc, 'npm run --silent agent:mcp-stdio', 'docs/mcp-workpaper-tool-server.md')
requireIncludes(mcpWorkPaperToolServerDoc, 'npm exec --package @bilig/headless -- bilig-workpaper-mcp', 'docs/mcp-workpaper-tool-server.md')
requireIncludes(mcpWorkPaperToolServerDoc, 'io.github.proompteng/bilig-workpaper', 'docs/mcp-workpaper-tool-server.md')
requireIncludes(
  mcpWorkPaperToolServerDoc,
  'https://registry.modelcontextprotocol.io/v0.1/servers?search=io.github.proompteng%2Fbilig-workpaper',
  'docs/mcp-workpaper-tool-server.md',
)
requireIncludes(mcpWorkPaperToolServerDoc, 'tools/list', 'docs/mcp-workpaper-tool-server.md')
requireIncludes(mcpWorkPaperToolServerDoc, 'tools/call', 'docs/mcp-workpaper-tool-server.md')
requireIncludes(
  mcpWorkPaperToolServerDoc,
  'https://modelcontextprotocol.io/specification/2025-06-18/server/tools',
  'docs/mcp-workpaper-tool-server.md',
)
requireIncludes(mcpWorkPaperToolServerDoc, 'https://github.com/proompteng/bilig/discussions/230', 'docs/mcp-workpaper-tool-server.md')
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
  ['docs/agent-spreadsheet-tool-call-loop.md', agentToolCallLoopDoc],
  ['docs/workbook-automation-examples-node.md', await readFile(join(docsRoot, 'workbook-automation-examples-node.md'), 'utf8')],
  ['docs/dev-to-workbook-apis-post.md', await readFile(join(docsRoot, 'dev-to-workbook-apis-post.md'), 'utf8')],
] as const) {
  requireIncludes(content, 'image: /assets/github-social-preview.png', path)
}

for (const [path, content] of [
  ['.github/ISSUE_TEMPLATE/config.yml', issueTemplateConfig],
  ['.github/ISSUE_TEMPLATE.md', issueTemplateRoot],
] as const) {
  requireIncludes(content, 'https://github.com/proompteng/bilig/discussions/157', path)
  requireNotIncludes(content, 'https://github.com/proompteng/bilig/discussions/115', path)
}

requireIncludes(issueTemplateConfig, 'https://github.com/proompteng/bilig/discussions/213', '.github/ISSUE_TEMPLATE/config.yml')
requireIncludes(
  pullRequestTemplate,
  'For public docs or example work, include the page or discussion that a new',
  '.github/PULL_REQUEST_TEMPLATE.md',
)

for (const [path, content] of [
  ['README.md', readme],
  ['packages/headless/README.md', headlessReadme],
  ['docs/index.html', index],
  ['docs/llms.txt', llms],
] as const) {
  requireIncludes(content, 'node-spreadsheet-formula-engine', path)
  requireIncludes(content, 'examples/serverless-workpaper-api', path)
}

for (const [path, content] of [
  ['README.md', readme],
  ['packages/headless/README.md', headlessReadme],
  ['docs/llms.txt', llms],
] as const) {
  requireIncludes(content, 'examples/headless-workpaper#invoice-totals', path)
  requireIncludes(content, 'examples/headless-workpaper#agent-framework-adapters', path)
  requireIncludes(content, 'examples/headless-workpaper#mcp-tool-server-shape', path)
  requireIncludes(content, 'npm run agent:framework-adapters', path)
  requireIncludes(content, 'npm run agent:mcp-tools', path)
  requireIncludes(content, 'npm run agent:mcp-stdio', path)
  requireIncludes(content, 'npm exec --package @bilig/headless -- bilig-workpaper-mcp', path)
  requireIncludes(content, 'https://registry.modelcontextprotocol.io/v0.1/servers?search=io.github.proompteng%2Fbilig-workpaper', path)
  requireIncludes(content, 'vercel-ai-sdk-langchain-spreadsheet-tool', path)
  requireIncludes(content, 'mcp-workpaper-tool-server', path)
  requireIncludes(content, 'examples/headless-workpaper#budget-variance-alerts', path)
  requireIncludes(content, 'examples/headless-workpaper#fulfillment-capacity-plan', path)
  requireIncludes(content, 'examples/headless-workpaper#quote-approval-threshold', path)
  requireIncludes(content, 'examples/headless-workpaper#subscription-mrr-forecast', path)
}

for (const [path, content] of [
  ['README.md', readme],
  ['packages/headless/README.md', headlessReadme],
  ['docs/llms.txt', llms],
] as const) {
  requireIncludes(content, 'docs/sheetjs-exceljs-alternative-formula-workbook-api.md', path)
}

for (const [path, content] of [
  ['README.md', readme],
  ['packages/headless/README.md', headlessReadme],
  ['docs/index.html', index],
  ['docs/llms.txt', llms],
  ['docs/what-workpaper-benchmark-proves.md', await readFile(join(docsRoot, 'what-workpaper-benchmark-proves.md'), 'utf8')],
] as const) {
  requireIncludes(content, 'workpaper-benchmark-card.png', path)
}

for (const [path, content] of [
  ['README.md', readme],
  ['packages/headless/README.md', headlessReadme],
  ['docs/index.html', index],
  ['docs/community-launch-pack.md', await readFile(join(docsRoot, 'community-launch-pack.md'), 'utf8')],
  ['docs/llms.txt', llms],
] as const) {
  requireIncludes(content, 'https://github.com/proompteng/bilig/discussions/157', path)
}

for (const [path, content] of [
  ['README.md', readme],
  ['docs/community-launch-pack.md', await readFile(join(docsRoot, 'community-launch-pack.md'), 'utf8')],
  ['docs/llms.txt', llms],
] as const) {
  requireIncludes(content, 'https://github.com/proompteng/bilig/discussions/213', path)
}

for (const [path, content] of [
  ['docs/mcp-workpaper-tool-server.md', mcpWorkPaperToolServerDoc],
  ['docs/llms.txt', llms],
] as const) {
  requireIncludes(content, 'https://github.com/proompteng/bilig/discussions/230', path)
}

for (const [path, content] of [
  ['docs/index.html', index],
  ['docs/community-launch-pack.md', await readFile(join(docsRoot, 'community-launch-pack.md'), 'utf8')],
  ['docs/llms.txt', llms],
] as const) {
  requireIncludes(content, 'https://github.com/proompteng/bilig/discussions/167', path)
}

const currentStarterIssueNumbers = [
  134, 153, 154, 155, 156, 158, 159, 162, 163, 193, 194, 195, 196, 197, 198, 207, 208, 209, 210, 211, 212, 217, 218, 219, 220, 221, 222,
  223, 231, 233, 247, 248, 249, 250, 255, 256, 257, 258, 259, 260, 265, 266, 267, 268, 269, 272, 273, 274, 275, 276, 277, 278, 279, 280,
  281,
]

for (const required of currentStarterIssueNumbers.map((issueNumber) => `https://github.com/proompteng/bilig/issues/${issueNumber}`)) {
  requireIncludes(starterIssues, required, 'docs/starter-issues.md')
  requireIncludes(llms, required, 'docs/llms.txt')
}

for (const closedIssue of [
  '137',
  '138',
  '141',
  '142',
  '143',
  '144',
  '145',
  '146',
  '147',
  '148',
  '149',
  '150',
  '151',
  '152',
  '224',
  '199',
  '200',
  '201',
  '202',
  '203',
  '204',
  '205',
  '228',
  '229',
  '246',
  '160',
  '161',
  '164',
  '165',
  '166',
  '168',
  '169',
  '170',
  '171',
  '172',
  '173',
  '174',
  '175',
  '176',
  '178',
  '179',
  '180',
  '181',
  '182',
  '183',
  '184',
  '185',
  '186',
  '187',
  '188',
  '189',
  '190',
  '191',
  '192',
  '227',
]) {
  if (starterIssues.includes(`https://github.com/proompteng/bilig/issues/${closedIssue}`)) {
    throw new Error(`docs/starter-issues.md still links to closed starter issue #${closedIssue}`)
  }

  if (llms.includes(`https://github.com/proompteng/bilig/issues/${closedIssue}`)) {
    throw new Error(`docs/llms.txt still links to closed starter issue #${closedIssue}`)
  }
}

const publicDocs = [
  ['packages/headless/README.md', headlessReadme],
  ['packages/excel-import/README.md', excelImportReadme],
  ['docs/public-api.md', publicApi],
] as const

for (const [path, content] of publicDocs) {
  for (const blockedSnippet of ['pnpm add @bilig/headless @bilig/excel-import', 'pnpm add @bilig/excel-import']) {
    if (content.includes(blockedSnippet)) {
      throw new Error(`${path} points users at unpublished npm package command: ${blockedSnippet}`)
    }
  }
}

for (const blockedLink of ['](../../docs/', '](../../examples/', '](../../LICENSE)']) {
  requireNotIncludes(headlessReadme, blockedLink, 'packages/headless/README.md')
}

const [headlessExampleReadme, headlessExamplePackage, headlessPackageManifest, headlessServerJson] = await Promise.all([
  readFile(join(repoRoot, 'examples', 'headless-workpaper', 'README.md'), 'utf8'),
  readFile(join(repoRoot, 'examples', 'headless-workpaper', 'package.json'), 'utf8'),
  readFile(join(repoRoot, 'packages', 'headless', 'package.json'), 'utf8'),
  readFile(join(repoRoot, 'packages', 'headless', 'server.json'), 'utf8'),
])
await requireFile(join(repoRoot, 'examples', 'headless-workpaper', 'agent-framework-adapters.ts'))
await requireFile(join(repoRoot, 'examples', 'headless-workpaper', 'mcp-tool-server.ts'))
await requireFile(join(repoRoot, 'examples', 'headless-workpaper', 'mcp-stdio-server.ts'))
requireDocumentedScriptsExist(headlessExampleReadme, headlessExamplePackage, 'examples/headless-workpaper/README.md')
requireIncludes(headlessExampleReadme, 'npm run invoice-totals', 'examples/headless-workpaper/README.md')
requireIncludes(headlessExampleReadme, '## Invoice Totals', 'examples/headless-workpaper/README.md')
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
requireIncludes(
  headlessExamplePackage,
  '"agent:framework-adapters": "tsx agent-framework-adapters.ts"',
  'examples/headless-workpaper/package.json',
)
requireIncludes(headlessExamplePackage, '"agent:mcp-tools": "tsx mcp-tool-server.ts"', 'examples/headless-workpaper/package.json')
requireIncludes(headlessExamplePackage, '"agent:mcp-stdio": "tsx mcp-stdio-server.ts"', 'examples/headless-workpaper/package.json')
requireIncludes(headlessPackageManifest, '"mcpName": "io.github.proompteng/bilig-workpaper"', 'packages/headless/package.json')
requireIncludes(headlessPackageManifest, '"bilig-workpaper-mcp": "./dist/work-paper-mcp-stdio-bin.js"', 'packages/headless/package.json')
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
  "import { WorkPaper } from '@bilig/headless'",
  'console.log({ revenue: (cell as NumericCell).value, verified: true });',
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
