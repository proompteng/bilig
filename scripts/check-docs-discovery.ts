import { readFile, stat } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { agentFrameworkDocRequirements, agentFrameworkLlmsRequiredLinks } from './check-docs-discovery-agent-pages.ts'
import { communityLaunchPackRequiredLinks, llmsExternalSurfaceLinks } from './check-docs-discovery-growth-links.ts'
import { requireHomepageDiscovery } from './check-docs-discovery-homepage.ts'
import { docsSiteSources } from './check-docs-discovery-site-sources.ts'

const repoRoot = join(dirname(fileURLToPath(import.meta.url)), '..')
const docsRoot = join(repoRoot, 'docs')
const siteRoot = 'https://proompteng.github.io/bilig/'

const expectedSitemapUrls = docsSiteSources.map(([urlPath]) => `${siteRoot}${urlPath}`)
const sourceFilesByUrl = new Map<string, string>(docsSiteSources.map(([urlPath, sourceFile]) => [`${siteRoot}${urlPath}`, sourceFile]))

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

function requirePackageKeywords(packageJson: string, requiredKeywords: readonly string[], context: string): void {
  const manifest: unknown = JSON.parse(packageJson)

  if (typeof manifest !== 'object' || manifest === null || !('keywords' in manifest)) {
    throw new Error(`${context} is missing a keywords array`)
  }

  const { keywords } = manifest
  if (!Array.isArray(keywords) || !keywords.every((keyword) => typeof keyword === 'string')) {
    throw new Error(`${context} keywords must be an array of strings`)
  }

  for (const requiredKeyword of requiredKeywords) {
    if (!keywords.includes(requiredKeyword)) {
      throw new Error(`${context} is missing discovery keyword: ${requiredKeyword}`)
    }
  }
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
  rootPackageJson,
  index,
  siteCss,
  robots,
  sitemap,
  llms,
  communityLaunchPack,
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
] = await Promise.all([
  readFile(join(repoRoot, 'README.md'), 'utf8'),
  readFile(join(repoRoot, 'CONTRIBUTING.md'), 'utf8'),
  readFile(join(repoRoot, 'package.json'), 'utf8'),
  readFile(join(docsRoot, 'index.html'), 'utf8'),
  readFile(join(docsRoot, 'assets', 'site.css'), 'utf8'),
  readFile(join(docsRoot, 'robots.txt'), 'utf8'),
  readFile(join(docsRoot, 'sitemap.xml'), 'utf8'),
  readFile(join(docsRoot, 'llms.txt'), 'utf8'),
  readFile(join(docsRoot, 'community-launch-pack.md'), 'utf8'),
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
])

const [headlessSpreadsheetEngineComparison, sheetjsExceljsAlternativeFormulaWorkbookApi, hyperformulaAlternativeHeadlessWorkpaper] =
  await Promise.all([
    readFile(join(docsRoot, 'headless-spreadsheet-engine-comparison.md'), 'utf8'),
    readFile(join(docsRoot, 'sheetjs-exceljs-alternative-formula-workbook-api.md'), 'utf8'),
    readFile(join(docsRoot, 'hyperformula-alternative-headless-workpaper.md'), 'utf8'),
  ])

requireHomepageDiscovery(index, siteCss)
requirePackageKeywords(
  headlessPackageJson,
  ['calculation', 'compute', 'excel', 'headless-spreadsheet', 'node-spreadsheet', 'sheet', 'spreadsheet-engine', 'worksheet'],
  'packages/headless/package.json',
)
requireIncludes(index, '"downloadUrl": "https://www.npmjs.com/package/@bilig/headless"', 'docs/index.html')
requireIncludes(index, '"applicationCategory": "DeveloperApplication"', 'docs/index.html')
requireIncludes(index, '"@type": "FAQPage"', 'docs/index.html')
for (const required of [
  './why-agents-need-workbook-apis.html',
  './agent-workpaper-tool-calling-recipe.html',
  './vercel-ai-sdk-langchain-spreadsheet-tool.html',
  './mcp-workpaper-tool-server.html',
  './mcp-spreadsheet-server-directory.html',
  './mcp-client-setup.html',
  './claude-desktop-mcpb-workpaper.html',
  './agent-spreadsheet-tool-call-loop.html',
  './node-service-workpaper-recipe.html',
  './server-side-spreadsheet-automation-node.html',
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
await requireFile(join(repoRoot, 'scripts', 'build-workpaper-mcpb.ts'))
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
  'https://proompteng.github.io/bilig/mcp-spreadsheet-server-directory.html',
  'https://proompteng.github.io/bilig/mcp-client-setup.html',
  'https://proompteng.github.io/bilig/claude-desktop-mcpb-workpaper.html',
  'https://github.com/proompteng/bilig/blob/main/docs/claude-desktop-mcpb-workpaper.md',
  'https://proompteng.github.io/bilig/agent-workpaper-tool-calling-recipe.html',
  'https://proompteng.github.io/bilig/agent-spreadsheet-tool-call-loop.html',
  'https://proompteng.github.io/bilig/node-service-workpaper-recipe.html',
  'https://proompteng.github.io/bilig/server-side-spreadsheet-automation-node.html',
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
  'https://github.com/proompteng/bilig/discussions/230#discussioncomment-16907632',
  'https://github.com/proompteng/bilig/discussions/115',
  'https://github.com/proompteng/bilig/blob/main/docs/dev-to-workbook-apis-post.md',
  'https://proompteng.github.io/bilig/node-spreadsheet-formula-engine.html',
  'https://proompteng.github.io/bilig/evaluate-excel-formulas-in-node-typescript.html',
  'https://github.com/proompteng/bilig/blob/main/docs/node-spreadsheet-formula-engine.md',
  'https://github.com/proompteng/bilig/blob/main/docs/evaluate-excel-formulas-in-node-typescript.md',
  'https://github.com/proompteng/bilig/blob/main/docs/server-side-spreadsheet-automation-node.md',
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
  'https://github.com/proompteng/bilig/blob/main/docs/community-launch-pack.md',
  'https://github.com/proompteng/bilig/blob/main/docs/community-growth-snapshot.md',
  'https://github.com/proompteng/bilig/issues?q=is%3Aissue%20state%3Aopen%20label%3Afirst-timers-only',
]) {
  requireIncludes(llms, required, 'docs/llms.txt')
}
for (const required of agentFrameworkLlmsRequiredLinks) {
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
  requireIncludes(content, '## Proof You Can Reproduce', path)
  requireIncludes(content, 'https://proompteng.github.io/bilig/community-growth-snapshot.html', path)
  requireIncludes(content, 'https://github.com/proompteng/bilig/stargazers', path)
  requireIncludes(content, 'edits one input, restores the', path)
  requireIncludes(content, 'saved JSON document, and verifies the dependent formula result.', path)
  requireIncludes(content, 'pnpm workpaper:bench:competitive:check', path)
  requireIncludes(content, 'lookup-approximate-duplicates', path)
  requireIncludes(content, '1.043x', path)
  requireIncludes(content, 'compatibility limits', path)
  requireIncludes(content, 'stars, npm downloads, starter issues, Discussions, traffic, and clones.', path)
  requireIncludes(content, 'https://github.com/proompteng/bilig/discussions/307', path)
  requireIncludes(content, 'https://github.com/proompteng/bilig/discussions/308', path)
  requireNotIncludes(content, '## Current Public Proof', path)
  requireNotIncludes(content, 'Latest checked-in snapshot', path)
  requireNotIncludes(content, '`12` forks', path)
  requireNotIncludes(content, '15,592` npm downloads in the', path)
  requireNotIncludes(content, '`10` GitHub Discussions', path)
  requireNotIncludes(content, 'repository views.', path)
}

requireIncludes(newContributorGuide, '## First-Time Command Checklist', 'docs/new-contributor-guide.md')
requireIncludes(newContributorGuide, 'pnpm docs:discovery:check', 'docs/new-contributor-guide.md')
requireIncludes(newContributorGuide, 'pnpm format:check', 'docs/new-contributor-guide.md')
requireIncludes(newContributorGuide, 'pnpm lint', 'docs/new-contributor-guide.md')
requireIncludes(starterIssues, 'new-contributor-guide.md#first-time-command-checklist', 'docs/starter-issues.md')
requireIncludes(starterIssues, 'https://github.com/proompteng/bilig/blob/main/CONTRIBUTING.md', 'docs/starter-issues.md')
requireIncludes(starterIssues, '89 open `first-timers-only` issues.', 'docs/starter-issues.md')
requireIncludes(starterIssues, '89 issues are generally available for a new contributor to claim.', 'docs/starter-issues.md')
requireIncludes(starterIssues, '0 issues already have active pull requests.', 'docs/starter-issues.md')
requireIncludes(starterIssues, '### npm Smoke Test Improvements', 'docs/starter-issues.md')
requireIncludes(starterIssues, '### JavaScript Library Comparison Starters', 'docs/starter-issues.md')
requireIncludes(starterIssues, 'https://github.com/proompteng/bilig/issues/265', 'docs/starter-issues.md')
requireIncludes(starterIssues, 'https://github.com/proompteng/bilig/issues/269', 'docs/starter-issues.md')
requireIncludes(starterIssues, '#269: docs(headless): add package-manager variants for the smoke test', 'docs/starter-issues.md')
requireIncludes(starterIssues, '#272: docs(examples): add NestJS WorkPaper controller smoke', 'docs/starter-issues.md')
requireIncludes(starterIssues, '#277: docs(examples): add Bun.serve WorkPaper route smoke', 'docs/starter-issues.md')
requireIncludes(starterIssues, '#278: docs(examples): add SvelteKit WorkPaper route smoke', 'docs/starter-issues.md')
requireIncludes(starterIssues, '#281: docs(examples): add Cloudflare D1 WorkPaper persistence smoke', 'docs/starter-issues.md')
requireIncludes(starterIssues, '#286: docs(examples): add Nuxt Nitro WorkPaper route smoke', 'docs/starter-issues.md')
requireIncludes(starterIssues, '#292: docs(mcp): add Zed MCP config for WorkPaper', 'docs/starter-issues.md')
requireIncludes(starterIssues, '#293: docs(mcp): add Continue MCP config for WorkPaper', 'docs/starter-issues.md')
requireIncludes(starterIssues, '#296: docs(mcpb): add Windows install notes for the Claude Desktop bundle', 'docs/starter-issues.md')
requireIncludes(starterIssues, '#297: docs(mcpb): add a Claude Desktop MCPB troubleshooting table', 'docs/starter-issues.md')
requireIncludes(starterIssues, '#298: docs(mcpb): add a copy-paste verification transcript for the bundle server', 'docs/starter-issues.md')
requireIncludes(starterIssues, '#323: docs(agent): add Mastra WorkPaper verification transcript', 'docs/starter-issues.md')
requireIncludes(starterIssues, '#324: docs(agent): add LlamaIndex.TS WorkPaper verification transcript', 'docs/starter-issues.md')
requireIncludes(starterIssues, '#325: docs(agent): add LangGraph.js ToolNode state handoff note', 'docs/starter-issues.md')
requireIncludes(starterIssues, '#326: docs(agent): add CopilotKit WorkPaper action UI result note', 'docs/starter-issues.md')
requireIncludes(starterIssues, '#327: docs(agent): add Cloudflare Agents WorkPaper state persistence note', 'docs/starter-issues.md')
requireIncludes(starterIssues, '#328: docs(site): add social preview checklist to the contributor guide', 'docs/starter-issues.md')
requireIncludes(starterIssues, '#329: docs(site): add OpenGraph cache-bust note to the launch pack', 'docs/starter-issues.md')
requireIncludes(starterIssues, '#299: docs(examples): add AdonisJS WorkPaper controller smoke', 'docs/starter-issues.md')
requireIncludes(starterIssues, '#300: docs(examples): add tRPC WorkPaper procedure smoke', 'docs/starter-issues.md')
requireIncludes(starterIssues, '#301: docs(storage): add Drizzle WorkPaper JSON persistence recipe', 'docs/starter-issues.md')
requireIncludes(starterIssues, '#302: docs(storage): add Kysely WorkPaper JSON persistence recipe', 'docs/starter-issues.md')
requireIncludes(starterIssues, '#303: docs(storage): add Upstash Redis WorkPaper persistence recipe', 'docs/starter-issues.md')
requireIncludes(starterIssues, '#304: docs(storage): add Neon Postgres WorkPaper persistence recipe', 'docs/starter-issues.md')
requireIncludes(starterIssues, '#305: docs(storage): add Cloudflare R2 WorkPaper persistence recipe', 'docs/starter-issues.md')
requireIncludes(starterIssues, '#306: docs(storage): add AWS S3 WorkPaper persistence recipe', 'docs/starter-issues.md')
requireIncludes(starterIssues, '#330: docs(examples): add Koa WorkPaper route smoke', 'docs/starter-issues.md')
requireIncludes(starterIssues, '#331: docs(examples): add Elysia WorkPaper route smoke', 'docs/starter-issues.md')
requireIncludes(starterIssues, '#332: docs(storage): add MongoDB WorkPaper JSON persistence recipe', 'docs/starter-issues.md')
requireIncludes(starterIssues, '#333: docs(comparison): add Univer spreadsheet UI boundary note', 'docs/starter-issues.md')
requireIncludes(starterIssues, '#257: docs(examples): add a runnable Hono WorkPaper route smoke', 'docs/starter-issues.md')
requireIncludes(starterIssues, '#258: docs(examples): add Cloudflare KV WorkPaper persistence snippet', 'docs/starter-issues.md')
requireIncludes(starterIssues, '#259: docs(service): add Prisma-backed WorkPaper JSON persistence recipe', 'docs/starter-issues.md')
requireIncludes(starterIssues, '#260: docs(examples): add Fastify WorkPaper route smoke snippet', 'docs/starter-issues.md')
requireNotIncludes(starterIssues, 'https://github.com/proompteng/bilig/pull/271', 'docs/starter-issues.md')
requireNotIncludes(starterIssues, 'https://github.com/proompteng/bilig/pull/291', 'docs/starter-issues.md')
requireNotIncludes(starterIssues, 'https://github.com/proompteng/bilig/pull/295', 'docs/starter-issues.md')
requireNotIncludes(starterIssues, 'https://github.com/proompteng/bilig/pull/251', 'docs/starter-issues.md')
requireIncludes(contributing, 'new-contributor-guide.md#first-time-command-checklist', 'CONTRIBUTING.md')
requireIncludes(llms, 'see docs/starter-issues.md for the maintained', 'docs/llms.txt')
requireIncludes(llms, 'https://github.com/proompteng/bilig/issues/272', 'docs/llms.txt')
requireIncludes(llms, 'https://github.com/proompteng/bilig/issues/277', 'docs/llms.txt')
requireIncludes(llms, 'https://github.com/proompteng/bilig/issues/281', 'docs/llms.txt')
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
requireIncludes(mcpWorkPaperToolServerDoc, 'mcp-client-setup.md', 'docs/mcp-workpaper-tool-server.md')
for (const required of [
  'description: Live directory and install status for the Bilig WorkPaper MCP server',
  'npm exec --package @bilig/headless -- bilig-workpaper-mcp',
  'io.github.proompteng/bilig-workpaper',
  'https://registry.modelcontextprotocol.io/v0.1/servers?search=io.github.proompteng%2Fbilig-workpaper',
  'https://glama.ai/mcp/servers/proompteng/bilig',
  'https://github.com/chatmcp/mcpso/issues/2295',
  'https://github.com/cline/mcp-marketplace/issues/1557',
  'Not indexed yet as of May 13, 2026',
  'https://www.pulsemcp.com/servers?search=bilig&q=bilig',
  'marked `@bilig/headless@0.14.0` as the',
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
for (const required of communityLaunchPackRequiredLinks) {
  requireIncludes(communityLaunchPack, required, 'docs/community-launch-pack.md')
}
for (const required of llmsExternalSurfaceLinks) {
  requireIncludes(llms, required, 'docs/llms.txt')
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
  ['docs/node-framework-workpaper-adapters.md', await readFile(join(docsRoot, 'node-framework-workpaper-adapters.md'), 'utf8')],
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
  requireIncludes(content, 'server-side-spreadsheet-automation-node', path)
  requireIncludes(content, 'examples/serverless-workpaper-api', path)
  requireIncludes(content, 'node-framework-workpaper-adapters', path)
  requireIncludes(content, 'mcp-spreadsheet-server-directory', path)
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
  requireIncludes(content, 'mcp-spreadsheet-server-directory', path)
  requireIncludes(content, 'mcp-client-setup', path)
  requireIncludes(content, 'claude-desktop-mcpb-workpaper', path)
  requireIncludes(content, 'examples/headless-workpaper#budget-variance-alerts', path)
  requireIncludes(content, 'examples/headless-workpaper#fulfillment-capacity-plan', path)
  requireIncludes(content, 'examples/headless-workpaper#quote-approval-threshold', path)
  requireIncludes(content, 'examples/headless-workpaper#subscription-mrr-forecast', path)
}

for (const required of [
  '## Use-Case Chooser',
  'Formula-backed calculations inside a Node service',
  'Agent writeback that must prove the value after an edit',
  'XLSX parsing, export, styling, images, and workbook-file metadata',
  'Persisting a workbook document as JSON and restoring it later',
  'Embedding a spreadsheet UI that users edit directly',
  '[npm smoke test](try-bilig-headless-in-node.md)',
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

for (const [path, content] of [
  ['README.md', readme],
  ['packages/headless/README.md', headlessReadme],
  ['docs/llms.txt', llms],
] as const) {
  requireIncludes(content, 'docs/javascript-spreadsheet-library-headless-node.md', path)
  requireIncludes(content, 'docs/sheetjs-exceljs-alternative-formula-workbook-api.md', path)
}

for (const [path, content] of [
  ['README.md', readme],
  ['packages/headless/README.md', headlessReadme],
  ['docs/llms.txt', llms],
  ['docs/what-workpaper-benchmark-proves.md', await readFile(join(docsRoot, 'what-workpaper-benchmark-proves.md'), 'utf8')],
] as const) {
  requireIncludes(content, 'workpaper-benchmark-card.png', path)
}

const discussionDocs = {
  readme: ['README.md', readme],
  headless: ['packages/headless/README.md', headlessReadme],
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
] as const

for (const [url, docKeys] of discussionDocChecks) {
  for (const docKey of docKeys) {
    const [path, content] = discussionDocs[docKey]
    requireIncludes(content, url, path)
  }
}

const currentStarterIssueNumbers = [
  134, 153, 155, 156, 158, 159, 162, 163, 193, 194, 195, 196, 197, 198, 207, 208, 209, 210, 211, 212, 217, 218, 219, 220, 221, 222, 223,
  233, 247, 248, 249, 250, 255, 256, 257, 258, 259, 260, 265, 267, 268, 269, 272, 273, 274, 275, 277, 278, 279, 280, 281, 283, 284, 285,
  286, 287, 288, 289, 290, 292, 293, 296, 297, 298, 299, 300, 301, 302, 303, 304, 305, 306, 309, 310, 311, 312, 313, 314, 323, 324, 325,
  326, 327, 328, 329, 330, 331, 332, 333,
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
  '154',
  '224',
  '231',
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
  '266',
  '282',
  '294',
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
  '276',
  '227',
  '315',
  '316',
  '317',
  '318',
  '319',
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
requireIncludes(headlessExampleReadme, 'annotations.', 'examples/headless-workpaper/README.md')
requireIncludes(headlessExampleReadme, 'read tool is annotated as read-only', 'examples/headless-workpaper/README.md')
requireIncludes(
  headlessExamplePackage,
  '"agent:framework-adapters": "tsx agent-framework-adapters.ts"',
  'examples/headless-workpaper/package.json',
)
requireIncludes(headlessExamplePackage, '"agent:mcp-tools": "tsx mcp-tool-server.ts"', 'examples/headless-workpaper/package.json')
requireIncludes(headlessExamplePackage, '"agent:mcp-stdio": "tsx mcp-stdio-server.ts"', 'examples/headless-workpaper/package.json')
await requireFile(join(repoRoot, 'examples', 'serverless-workpaper-api', 'framework-adapters.ts'))
await requireFile(join(repoRoot, 'examples', 'serverless-workpaper-api', 'next-route-handler.ts'))
await requireFile(join(repoRoot, 'examples', 'serverless-workpaper-api', 'persistence-adapters.ts'))
const [serverlessExampleReadme, serverlessExamplePackage] = await Promise.all([
  readFile(join(repoRoot, 'examples', 'serverless-workpaper-api', 'README.md'), 'utf8'),
  readFile(join(repoRoot, 'examples', 'serverless-workpaper-api', 'package.json'), 'utf8'),
])
const serverlessWorkPaperApiRouteDoc = await readFile(join(docsRoot, 'serverless-workpaper-api-route.md'), 'utf8')
const persistenceDoc = await readFile(join(docsRoot, 'persisting-formula-backed-workpaper-documents-in-node.md'), 'utf8')
requireIncludes(serverlessExampleReadme, 'npm run next-route-handler', 'examples/serverless-workpaper-api/README.md')
requireIncludes(serverlessExampleReadme, '## Next.js App Router Smoke', 'examples/serverless-workpaper-api/README.md')
requireIncludes(serverlessWorkPaperApiRouteDoc, 'npm run next-route-handler', 'docs/serverless-workpaper-api-route.md')
requireIncludes(serverlessExampleReadme, 'npm run framework-adapters', 'examples/serverless-workpaper-api/README.md')
requireIncludes(serverlessExampleReadme, '## Framework Adapters', 'examples/serverless-workpaper-api/README.md')
requireIncludes(serverlessExampleReadme, 'npm run persistence-adapters', 'examples/serverless-workpaper-api/README.md')
requireIncludes(serverlessExampleReadme, '## Persistence Adapters', 'examples/serverless-workpaper-api/README.md')
requireIncludes(serverlessExampleReadme, 'Postgres JSONB', 'examples/serverless-workpaper-api/README.md')
requireIncludes(
  persistenceDoc,
  'examples/serverless-workpaper-api/persistence-adapters.ts',
  'docs/persisting-formula-backed-workpaper-documents-in-node.md',
)
requireIncludes(persistenceDoc, 'npm run persistence-adapters', 'docs/persisting-formula-backed-workpaper-documents-in-node.md')
requireIncludes(persistenceDoc, 'Postgres JSONB', 'docs/persisting-formula-backed-workpaper-documents-in-node.md')
requireIncludes(persistenceDoc, 'Redis or string-KV adapter', 'docs/persisting-formula-backed-workpaper-documents-in-node.md')
requireIncludes(
  serverlessExamplePackage,
  '"next-route-handler": "tsx next-route-handler.ts"',
  'examples/serverless-workpaper-api/package.json',
)
requireIncludes(
  serverlessExamplePackage,
  '"framework-adapters": "tsx framework-adapters.ts"',
  'examples/serverless-workpaper-api/package.json',
)
requireIncludes(
  serverlessExamplePackage,
  '"persistence-adapters": "tsx persistence-adapters.ts"',
  'examples/serverless-workpaper-api/package.json',
)
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
requireIncludes(headlessReadme, 'npm run next-route-handler', 'packages/headless/README.md')
requireIncludes(headlessReadme, 'npm run framework-adapters', 'packages/headless/README.md')
requireIncludes(headlessReadme, 'npm run persistence-adapters', 'packages/headless/README.md')
requireIncludes(headlessReadme, '#persistence-adapters', 'packages/headless/README.md')
requireIncludes(headlessReadme, 'node-framework-workpaper-adapters.html', 'packages/headless/README.md')

for (const required of [
  '## Clean npm Sanity Check',
  'mkdir bilig-headless-sanity',
  'npx tsx sanity.ts',
  'createWorkPaperFromDocument',
  'serializeWorkPaperDocument',
  'workbook.setCellContents({ sheet: revenue, row: 1, col: 1 }, 32);',
  'console.log({ before, after, sheets: restored.getSheetNames(), bytes: saved.length, verified });',
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
