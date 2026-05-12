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
requireIncludes(index, '<title>bilig - Headless Spreadsheet Engine for Node.js and Agents</title>', 'docs/index.html')
requireIncludes(index, '<meta name="robots" content="index, follow, max-image-preview:large" />', 'docs/index.html')
requireIncludes(index, '<link rel="icon" type="image/svg+xml" href="./assets/favicon.svg" />', 'docs/index.html')
requireIncludes(index, 'bilig-hero-workbook-api.png?v=2026-05-08-2', 'docs/index.html')
requireIncludes(index, '"downloadUrl": "https://www.npmjs.com/package/@bilig/headless"', 'docs/index.html')
requireIncludes(index, '"applicationCategory": "DeveloperApplication"', 'docs/index.html')
requireIncludes(index, '"@type": "FAQPage"', 'docs/index.html')

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
  'https://github.com/proompteng/bilig/tree/main/examples/serverless-workpaper-api',
  'https://github.com/proompteng/bilig/discussions',
  'https://github.com/proompteng/bilig/discussions/157',
  'https://github.com/proompteng/bilig/discussions/167',
  'https://github.com/proompteng/bilig/discussions/115',
  'https://github.com/proompteng/bilig/blob/main/docs/dev-to-workbook-apis-post.md',
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

for (const [path, content] of [
  ['README.md', readme],
  ['packages/headless/README.md', headlessReadme],
  ['docs/index.html', index],
  ['docs/llms.txt', llms],
] as const) {
  requireIncludes(content, 'examples/serverless-workpaper-api', path)
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
  ['docs/community-launch-pack.md', await readFile(join(docsRoot, 'community-launch-pack.md'), 'utf8')],
  ['docs/llms.txt', llms],
] as const) {
  requireIncludes(content, 'https://github.com/proompteng/bilig/discussions/157', path)
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
  'https://github.com/proompteng/bilig/issues/138',
  'https://github.com/proompteng/bilig/issues/141',
  'https://github.com/proompteng/bilig/issues/142',
  'https://github.com/proompteng/bilig/issues/143',
  'https://github.com/proompteng/bilig/issues/144',
  'https://github.com/proompteng/bilig/issues/145',
  'https://github.com/proompteng/bilig/issues/146',
  'https://github.com/proompteng/bilig/issues/147',
  'https://github.com/proompteng/bilig/issues/148',
  'https://github.com/proompteng/bilig/issues/149',
  'https://github.com/proompteng/bilig/issues/150',
  'https://github.com/proompteng/bilig/issues/151',
  'https://github.com/proompteng/bilig/issues/152',
  'https://github.com/proompteng/bilig/issues/153',
  'https://github.com/proompteng/bilig/issues/154',
  'https://github.com/proompteng/bilig/issues/155',
  'https://github.com/proompteng/bilig/issues/156',
  'https://github.com/proompteng/bilig/issues/158',
  'https://github.com/proompteng/bilig/issues/159',
  'https://github.com/proompteng/bilig/issues/160',
  'https://github.com/proompteng/bilig/issues/162',
  'https://github.com/proompteng/bilig/issues/163',
  'https://github.com/proompteng/bilig/issues/171',
  'https://github.com/proompteng/bilig/issues/172',
  'https://github.com/proompteng/bilig/issues/173',
  'https://github.com/proompteng/bilig/issues/174',
]) {
  requireIncludes(starterIssues, required, 'docs/starter-issues.md')
  requireIncludes(llms, required, 'docs/llms.txt')
}

for (const closedIssue of ['137', '161', '164', '165', '166', '168', '169', '170']) {
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
