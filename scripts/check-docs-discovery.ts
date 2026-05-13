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
  `${siteRoot}agent-spreadsheet-tool-call-loop.html`,
  `${siteRoot}node-service-workpaper-recipe.html`,
  `${siteRoot}node-spreadsheet-formula-engine.html`,
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
  `${siteRoot}llms.txt`,
] as const

const sourceFilesByUrl = new Map<string, string>([
  [siteRoot, 'index.html'],
  [`${siteRoot}why-agents-need-workbook-apis.html`, 'why-agents-need-workbook-apis.md'],
  [`${siteRoot}agent-workpaper-tool-calling-recipe.html`, 'agent-workpaper-tool-calling-recipe.md'],
  [`${siteRoot}agent-spreadsheet-tool-call-loop.html`, 'agent-spreadsheet-tool-call-loop.md'],
  [`${siteRoot}node-service-workpaper-recipe.html`, 'node-service-workpaper-recipe.md'],
  [`${siteRoot}node-spreadsheet-formula-engine.html`, 'node-spreadsheet-formula-engine.md'],
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

const [
  readme,
  contributing,
  index,
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
requireIncludes(index, 'bilig-hero-workbook-api.png?v=2026-05-08-2', 'docs/index.html')
requireIncludes(index, '"downloadUrl": "https://www.npmjs.com/package/@bilig/headless"', 'docs/index.html')
requireIncludes(index, '"applicationCategory": "DeveloperApplication"', 'docs/index.html')
requireIncludes(index, '"@type": "FAQPage"', 'docs/index.html')
for (const required of [
  './why-agents-need-workbook-apis.html',
  './agent-workpaper-tool-calling-recipe.html',
  './agent-spreadsheet-tool-call-loop.html',
  './node-service-workpaper-recipe.html',
  './node-spreadsheet-formula-engine.html',
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
  ['README.md', 'package.json', 'route.mjs', 'smoke.mjs'].map((sourceFile) =>
    requireFile(join(repoRoot, 'examples', 'serverless-workpaper-api', sourceFile)),
  ),
)
await Promise.all(
  ['github-social-preview.png', 'workpaper-benchmark-card.png'].map((sourceFile) => requireFile(join(docsRoot, 'assets', sourceFile))),
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
  'npm run agent:verify',
  'https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#json-records-input',
  'https://proompteng.github.io/bilig/why-agents-need-workbook-apis.html',
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
  'https://github.com/proompteng/bilig/discussions/115',
  'https://github.com/proompteng/bilig/blob/main/docs/dev-to-workbook-apis-post.md',
  'https://proompteng.github.io/bilig/node-spreadsheet-formula-engine.html',
  'https://github.com/proompteng/bilig/blob/main/docs/node-spreadsheet-formula-engine.md',
  'https://github.com/proompteng/bilig/blob/main/docs/node-service-workpaper-recipe.md',
  'https://github.com/proompteng/bilig/blob/main/docs/serverless-workpaper-api-route.md',
  'https://github.com/proompteng/bilig/blob/main/docs/csv-shaped-workpaper-input-recipe.md',
  'https://github.com/proompteng/bilig/blob/main/docs/unsupported-formula-troubleshooting-recipe.md',
  'https://github.com/proompteng/bilig/blob/main/docs/agent-workpaper-tool-calling-recipe.md',
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

requireIncludes(newContributorGuide, '## First-Time Command Checklist', 'docs/new-contributor-guide.md')
requireIncludes(newContributorGuide, 'pnpm docs:discovery:check', 'docs/new-contributor-guide.md')
requireIncludes(newContributorGuide, 'pnpm format:check', 'docs/new-contributor-guide.md')
requireIncludes(newContributorGuide, 'pnpm lint', 'docs/new-contributor-guide.md')
requireIncludes(starterIssues, 'new-contributor-guide.md#first-time-command-checklist', 'docs/starter-issues.md')
requireIncludes(contributing, 'new-contributor-guide.md#first-time-command-checklist', 'CONTRIBUTING.md')

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

const [whyAgentsDoc, agentToolCallingDoc, agentToolCallLoopDoc] = await Promise.all([
  readFile(join(docsRoot, 'why-agents-need-workbook-apis.md'), 'utf8'),
  readFile(join(docsRoot, 'agent-workpaper-tool-calling-recipe.md'), 'utf8'),
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
requireIncludes(
  agentToolCallLoopDoc,
  'description: A runnable @bilig/headless loop where an agent writes one workbook input',
  'docs/agent-spreadsheet-tool-call-loop.md',
)
for (const [path, content] of [
  ['docs/why-agents-need-workbook-apis.md', whyAgentsDoc],
  ['docs/agent-workpaper-tool-calling-recipe.md', agentToolCallingDoc],
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
  ['docs/index.html', index],
  ['docs/community-launch-pack.md', await readFile(join(docsRoot, 'community-launch-pack.md'), 'utf8')],
  ['docs/llms.txt', llms],
] as const) {
  requireIncludes(content, 'https://github.com/proompteng/bilig/discussions/167', path)
}

for (const required of [
  'https://github.com/proompteng/bilig/issues/134',
  'https://github.com/proompteng/bilig/issues/153',
  'https://github.com/proompteng/bilig/issues/154',
  'https://github.com/proompteng/bilig/issues/155',
  'https://github.com/proompteng/bilig/issues/156',
  'https://github.com/proompteng/bilig/issues/158',
  'https://github.com/proompteng/bilig/issues/159',
  'https://github.com/proompteng/bilig/issues/162',
  'https://github.com/proompteng/bilig/issues/163',
  'https://github.com/proompteng/bilig/issues/193',
  'https://github.com/proompteng/bilig/issues/194',
  'https://github.com/proompteng/bilig/issues/195',
  'https://github.com/proompteng/bilig/issues/196',
  'https://github.com/proompteng/bilig/issues/197',
  'https://github.com/proompteng/bilig/issues/198',
  'https://github.com/proompteng/bilig/issues/199',
  'https://github.com/proompteng/bilig/issues/207',
  'https://github.com/proompteng/bilig/issues/208',
  'https://github.com/proompteng/bilig/issues/209',
  'https://github.com/proompteng/bilig/issues/210',
  'https://github.com/proompteng/bilig/issues/211',
  'https://github.com/proompteng/bilig/issues/212',
  'https://github.com/proompteng/bilig/issues/217',
  'https://github.com/proompteng/bilig/issues/218',
  'https://github.com/proompteng/bilig/issues/219',
  'https://github.com/proompteng/bilig/issues/220',
  'https://github.com/proompteng/bilig/issues/221',
  'https://github.com/proompteng/bilig/issues/222',
  'https://github.com/proompteng/bilig/issues/223',
  'https://github.com/proompteng/bilig/issues/224',
]) {
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
  '200',
  '201',
  '202',
  '203',
  '204',
  '205',
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

const headlessExampleReadme = await readFile(join(repoRoot, 'examples', 'headless-workpaper', 'README.md'), 'utf8')
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
  'node --input-type=module',
  "import { WorkPaper } from '@bilig/headless'",
  'console.log({ revenue: cell.value, verified: true })',
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
