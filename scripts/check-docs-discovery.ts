import { readFile } from 'node:fs/promises'
import { join } from 'node:path'
import { agentFrameworkDocRequirements, agentFrameworkLlmsRequiredLinks } from './check-docs-discovery-agent-pages.ts'
import {
  requireFile,
  requireIncludes,
  requireNoUnsupportedGoogleSheetsTenXClaims,
  requireNotIncludes,
  requirePackageKeywords,
  requirePublishedSource,
} from './check-docs-discovery-core.ts'
import { loadDocsDiscoveryContext } from './check-docs-discovery-context.ts'
import { requireSitemapPublishedSources } from './check-docs-discovery-sitemap.ts'
import { requireHomepageDiscovery } from './check-docs-discovery-homepage.ts'
import { productHuntLaunchAssetFiles, requireGrowthSurfaceDiscovery } from './check-docs-discovery-launch-kit.ts'
import { llmsExternalSurfaceLinks } from './check-docs-discovery-growth-links.ts'
import { requireFormulaProofDiscovery } from './check-docs-discovery-proof-pages.ts'
import { requireStarterIssueDiscovery } from './check-docs-discovery-starter-issues.ts'
import { requireTypeScriptFirstPublicSnippets } from './check-docs-discovery-typescript-snippets.ts'
import { requireXlsxCorpusVerifierDiscovery } from './check-docs-discovery-xlsx-verifier.ts'
import { requireXlsxCalcAlternativeDiscovery } from './check-docs-discovery-xlsx-calc.ts'
import { requireSharedPublicDocsDiscovery } from './check-docs-discovery-public-docs.ts'
import { requireHeadlessExampleDiscovery } from './check-docs-discovery-headless-examples.ts'
import { homepageRequiredLinks, llmsRequiredLinks } from './check-docs-discovery-public-link-manifest.ts'

const {
  repoRoot,
  docsRoot,
  siteRoot,
  expectedSitemapUrls,
  sourceFilesByUrl,
  benchmarkEvidence,
  headlessPackageVersion,
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
  headlessExamplePackageJson,
  headlessReadme,
  headlessAgentNotes,
  excelImportReadme,
  dockerfile,
  publicApi,
  issueTemplateConfig,
  issueTemplateRoot,
  featureRequestTemplate,
  ideasDiscussionTemplate,
  qaDiscussionTemplate,
  showAndTellDiscussionTemplate,
  generalDiscussionTemplate,
  pullRequestTemplate,
  dominanceScorecard,
  headlessSpreadsheetEngineComparison,
  sheetjsExceljsAlternativeFormulaWorkbookApi,
  hyperformulaAlternativeHeadlessWorkpaper,
  xlsxFormulaRecalculationNode,
  agentXlsxFormulaRecalculationWithoutLibreOffice,
  staleXlsxFormulaCacheNode,
  microsoftGraphExcelRecalculationNode,
  formulaWorkbooksProof,
  showHnFormulaWorkbooksProof,
  googleSheetsApiBoundaryDoc,
  npmProvenancePackageTrustDoc,
  xlsxCorpusVerifierWalkthrough,
  whyAgentsDoc,
  headlessWorkpaperAgentHandbook,
  agentToolCallingDoc,
  aiSdkLangChainDoc,
  mcpWorkPaperToolServerDoc,
  mcpSpreadsheetServerDirectoryDoc,
  mcpClientSetupDoc,
  claudeDesktopMcpbDoc,
  agentToolCallLoopDoc,
  mcpServerCard,
  mcpServerCardMcpJson,
  mcpServerCardLegacyJson,
  workbookAutomationExamplesDoc,
  serverSideSpreadsheetAutomationNode,
  nodeFrameworkWorkpaperAdaptersDoc,
  devToWorkbookApisPost,
  evaluateExcelFormulasInNodeTypescript,
  nodeSpreadsheetFormulaEngine,
} = await loadDocsDiscoveryContext()

const headlessSpreadsheetEngineNodeServicesAgents = await readFile(
  join(docsRoot, 'headless-spreadsheet-engine-node-services-agents.md'),
  'utf8',
)

requireHomepageDiscovery(index, siteCss, productCss)
await requireXlsxCalcAlternativeDiscovery(docsRoot)
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
    'excel-formulas',
    'formula-recalculation',
    'formula-engine',
    'headless-spreadsheet',
    'hyperformula',
    'mcp',
    'mcp-server',
    'node',
    'spreadsheet-automation',
    'spreadsheet-engine',
    'spreadsheet-formulas',
    'spreadsheet-mcp',
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
for (const required of homepageRequiredLinks) {
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

for (const required of llmsRequiredLinks) {
  requireIncludes(llms, required, 'docs/llms.txt')
}
for (const required of agentFrameworkLlmsRequiredLinks) {
  requireIncludes(llms, required, 'docs/llms.txt')
}

requireFormulaProofDiscovery({
  benchmarkEvidence,
  communityLaunchPack,
  formulaWorkbooksProof,
  headlessReadme,
  index,
  llms,
  readme,
  requireIncludes,
  showHnFormulaWorkbooksProof,
})

for (const required of [
  'title: Fix stale XLSX formula values in Node.js',
  'An `.xlsx` can store both the formula text',
  'Run a formula runtime before reading',
  '`@bilig/headless` when the service can own the workbook state locally',
  'https://github.com/proompteng/bilig/stargazers',
] as const) {
  requireIncludes(staleXlsxFormulaCacheNode, required, 'docs/stale-xlsx-formula-cache-node.md')
}
requireIncludes(index, './stale-xlsx-formula-cache-node.html', 'docs/index.html')
requireIncludes(readme, 'docs/stale-xlsx-formula-cache-node.md', 'README.md')
requireIncludes(headlessReadme, 'docs/stale-xlsx-formula-cache-node.md', 'packages/headless/README.md')
requireIncludes(llms, 'https://proompteng.github.io/bilig/stale-xlsx-formula-cache-node.html', 'docs/llms.txt')
requireIncludes(llms, 'https://github.com/proompteng/bilig/blob/main/docs/stale-xlsx-formula-cache-node.md', 'docs/llms.txt')

for (const required of [
  'title: Microsoft Graph Excel recalculation vs local Node WorkPaper',
  'POST /me/drive/items/{id}/workbook/application/calculate',
  'Files.ReadWrite',
  'application permissions are not supported for that API',
  'Use `@bilig/headless` when the workbook is service-owned state',
  'https://learn.microsoft.com/en-us/graph/api/workbookapplication-calculate',
  'https://github.com/proompteng/bilig/stargazers',
] as const) {
  requireIncludes(microsoftGraphExcelRecalculationNode, required, 'docs/microsoft-graph-excel-recalculation-node.md')
}
requireIncludes(index, './microsoft-graph-excel-recalculation-node.html', 'docs/index.html')
requireIncludes(readme, 'docs/microsoft-graph-excel-recalculation-node.md', 'README.md')
requireIncludes(headlessReadme, 'docs/microsoft-graph-excel-recalculation-node.md', 'packages/headless/README.md')
requireIncludes(llms, 'https://proompteng.github.io/bilig/microsoft-graph-excel-recalculation-node.html', 'docs/llms.txt')
requireIncludes(llms, 'https://github.com/proompteng/bilig/blob/main/docs/microsoft-graph-excel-recalculation-node.md', 'docs/llms.txt')

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
  generalDiscussionTemplate,
  excelImportReadme,
  publicApi,
})

requireIncludes(readme, 'acceptance commands for first patches.', 'README.md')
requireIncludes(readme, 'docs/why-use-bilig.md', 'README.md')
requireIncludes(readme, 'The published package also carries `AGENTS.md`', 'README.md')
requireIncludes(readme, 'agent handoff prompt', 'README.md')
requireIncludes(index, './headless-workpaper-agent-handbook.html">Agent handoff prompt', 'docs/index.html')
requireIncludes(llms, '## agent handoff prompt', 'docs/llms.txt')
requireIncludes(llms, 'Do not claim success from a write call alone.', 'docs/llms.txt')
requireIncludes(llms, 'pnpm --dir bilig/examples/headless-workpaper install --ignore-workspace', 'docs/llms.txt')
requireIncludes(llms, 'pnpm --dir bilig/examples/headless-workpaper run agent:framework-adapters', 'docs/llms.txt')
requireIncludes(llms, 'pnpm --dir examples/headless-workpaper run agent:mcp-tools', 'docs/llms.txt')
requireNotIncludes(llms, 'cd bilig/examples/headless-workpaper', 'docs/llms.txt')
requireNotIncludes(llms, '\nnpm start\n', 'docs/llms.txt')
requireIncludes(headlessReadme, 'https://proompteng.github.io/bilig/why-use-bilig.html', 'packages/headless/README.md')
requireIncludes(headlessReadme, 'The npm tarball also includes `AGENTS.md`', 'packages/headless/README.md')
requireIncludes(headlessPackageJson, '"AGENTS.md"', 'packages/headless/package.json')
requireIncludes(headlessAgentNotes, '## Handoff prompt', 'packages/headless/AGENTS.md')
requireIncludes(headlessAgentNotes, 'Do not claim success from a write call alone.', 'packages/headless/AGENTS.md')
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
  evaluateExcelFormulasInNodeTypescript,
  'npx tsx eval-node-formulas.ts',
  'docs/evaluate-excel-formulas-in-node-typescript.md',
)
requireIncludes(serverSideSpreadsheetAutomationNode, 'npx tsx eval.ts', 'docs/server-side-spreadsheet-automation-node.md')
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

for (const required of [
  'title: Verify npm provenance for @bilig/headless',
  'npm view @bilig/headless@latest version dist.attestations dist.signatures --json',
  'npm audit signatures',
  'dist.attestations.provenance.predicateType',
  'npm publish ... --provenance',
  'https://docs.npmjs.com/trusted-publishers/',
  'https://docs.npmjs.com/viewing-package-provenance/',
  'https://scorecard.dev/',
  'official OpenSSF Scorecard action',
  'uploaded as SARIF to GitHub code',
] as const) {
  requireIncludes(npmProvenancePackageTrustDoc, required, 'docs/npm-provenance-package-trust.md')
}
requireIncludes(readme, `@bilig/headless@${headlessPackageVersion}`, 'README.md')
requireIncludes(readme, 'npm view @bilig/headless@latest version dist.attestations dist.signatures --json', 'README.md')
requireIncludes(readme, 'npm provenance and package trust', 'README.md')
requireIncludes(readme, 'https://api.scorecard.dev/projects/github.com/proompteng/bilig/badge', 'README.md')
requireIncludes(readme, 'uploaded to GitHub code scanning on every `main` update', 'README.md')
requireIncludes(headlessReadme, `@bilig/headless@${headlessPackageVersion}`, 'packages/headless/README.md')
requireIncludes(
  headlessReadme,
  'npm view @bilig/headless@latest version dist.attestations dist.signatures --json',
  'packages/headless/README.md',
)
requireIncludes(headlessReadme, 'npm provenance and package trust guide', 'packages/headless/README.md')
requireIncludes(headlessReadme, 'https://api.scorecard.dev/projects/github.com/proompteng/bilig/badge', 'packages/headless/README.md')
requireIncludes(headlessReadme, 'uploaded to GitHub code scanning on every `main` update', 'packages/headless/README.md')
requireIncludes(readme, 'examples/xlsx-recalculation-node', 'README.md')
requireIncludes(readme, 'docs/xlsx-formula-recalculation-node.md', 'README.md')
requireIncludes(readme, 'docs/agent-xlsx-formula-recalculation-without-libreoffice.md', 'README.md')
requireIncludes(readme, 'docs/excel-file-calculation-engine-node.md', 'README.md')
requireIncludes(readme, 'docs/exceljs-shared-formula-recalculation-node.md', 'README.md')
requireIncludes(headlessReadme, 'examples/xlsx-recalculation-node', 'packages/headless/README.md')
requireIncludes(headlessReadme, 'docs/xlsx-formula-recalculation-node.md', 'packages/headless/README.md')
requireIncludes(
  headlessReadme,
  'https://proompteng.github.io/bilig/agent-xlsx-formula-recalculation-without-libreoffice.html',
  'packages/headless/README.md',
)
requireIncludes(headlessReadme, 'docs/excel-file-calculation-engine-node.md', 'packages/headless/README.md')
requireIncludes(headlessReadme, 'docs/exceljs-shared-formula-recalculation-node.md', 'packages/headless/README.md')
requireIncludes(index, 'examples/xlsx-recalculation-node', 'docs/index.html')
requireIncludes(index, './xlsx-formula-recalculation-node.html', 'docs/index.html')
requireIncludes(index, './xlsx-recalculation-proof.html', 'docs/index.html')
requireIncludes(index, './agent-xlsx-formula-recalculation-without-libreoffice.html', 'docs/index.html')
requireIncludes(index, './excel-file-calculation-engine-node.html', 'docs/index.html')
requireIncludes(index, './exceljs-shared-formula-recalculation-node.html', 'docs/index.html')
requireIncludes(index, './xlsx-template-formula-recalculation-node.html', 'docs/index.html')
requireIncludes(index, './xlsx-populate-formula-result-node.html', 'docs/index.html')
requireIncludes(llms, 'https://github.com/proompteng/bilig/tree/main/examples/xlsx-recalculation-node', 'docs/llms.txt')
requireIncludes(llms, 'https://proompteng.github.io/bilig/xlsx-formula-recalculation-node.html', 'docs/llms.txt')
requireIncludes(llms, 'https://github.com/proompteng/bilig/blob/main/docs/xlsx-formula-recalculation-node.md', 'docs/llms.txt')
requireIncludes(llms, 'https://proompteng.github.io/bilig/xlsx-recalculation-proof.html', 'docs/llms.txt')
requireIncludes(llms, 'https://proompteng.github.io/bilig/xlsx-recalculation-proof.ts', 'docs/llms.txt')
requireIncludes(llms, 'creates an XLSX workbook, edits inputs, recalculates formulas in Node.js', 'docs/llms.txt')
for (const required of [
  'title: Agent XLSX formula recalculation without LibreOffice',
  'canonical_url: https://proompteng.github.io/bilig/agent-xlsx-formula-recalculation-without-libreoffice.html',
  'curl -fsSLO https://proompteng.github.io/bilig/xlsx-recalculation-proof.ts',
  '"formulasSurvivedXlsxRoundTrip": true',
  'verified: true',
  '[MCP spreadsheet tool server](mcp-workpaper-tool-server.md)',
]) {
  requireIncludes(agentXlsxFormulaRecalculationWithoutLibreOffice, required, 'docs/agent-xlsx-formula-recalculation-without-libreoffice.md')
}
requireIncludes(llms, 'https://proompteng.github.io/bilig/agent-xlsx-formula-recalculation-without-libreoffice.html', 'docs/llms.txt')
requireIncludes(
  llms,
  'https://github.com/proompteng/bilig/blob/main/docs/agent-xlsx-formula-recalculation-without-libreoffice.md',
  'docs/llms.txt',
)
requireIncludes(llms, 'gives spreadsheet agents a Node.js tool contract', 'docs/llms.txt')
requireIncludes(llms, 'https://proompteng.github.io/bilig/excel-file-calculation-engine-node.html', 'docs/llms.txt')
requireIncludes(llms, 'https://github.com/proompteng/bilig/blob/main/docs/excel-file-calculation-engine-node.md', 'docs/llms.txt')
requireIncludes(llms, 'covers backend routes that write request inputs into an XLSX workbook', 'docs/llms.txt')
requireIncludes(llms, 'https://proompteng.github.io/bilig/exceljs-shared-formula-recalculation-node.html', 'docs/llms.txt')
requireIncludes(llms, 'https://github.com/proompteng/bilig/blob/main/docs/exceljs-shared-formula-recalculation-node.md', 'docs/llms.txt')
requireIncludes(llms, 'documents the XLSX shared-formula expansion path', 'docs/llms.txt')
requireIncludes(llms, 'https://proompteng.github.io/bilig/xlsx-template-formula-recalculation-node.html', 'docs/llms.txt')
requireIncludes(llms, 'https://github.com/proompteng/bilig/blob/main/docs/xlsx-template-formula-recalculation-node.md', 'docs/llms.txt')
requireIncludes(llms, 'template substitution -> formula runtime -> verified readback', 'docs/llms.txt')
requireIncludes(llms, 'https://proompteng.github.io/bilig/xlsx-populate-formula-result-node.html', 'docs/llms.txt')
requireIncludes(llms, 'https://github.com/proompteng/bilig/blob/main/docs/xlsx-populate-formula-result-node.md', 'docs/llms.txt')
requireIncludes(llms, 'separates formula serialization from recalculation', 'docs/llms.txt')
requireIncludes(index, './npm-provenance-package-trust.html', 'docs/index.html')
requireIncludes(llms, 'https://proompteng.github.io/bilig/npm-provenance-package-trust.html', 'docs/llms.txt')
requireIncludes(llms, 'https://github.com/proompteng/bilig/blob/main/docs/npm-provenance-package-trust.md', 'docs/llms.txt')
await requireFile(join(repoRoot, '.github', 'workflows', 'scorecard.yml'))

requireXlsxCorpusVerifierDiscovery(xlsxCorpusVerifierWalkthrough)
requireIncludes(index, './xlsx-corpus-verifier-walkthrough.html', 'docs/index.html')
requireIncludes(llms, 'https://proompteng.github.io/bilig/xlsx-corpus-verifier-walkthrough.html', 'docs/llms.txt')

const jekyllConfig = await readFile(join(docsRoot, '_config.yml'), 'utf8')
requireIncludes(jekyllConfig, 'include:', 'docs/_config.yml')
requireIncludes(jekyllConfig, '  - .well-known', 'docs/_config.yml')
if (mcpServerCardMcpJson !== mcpServerCard) {
  throw new Error('docs/.well-known/mcp.json must match docs/.well-known/mcp/server-card.json')
}
if (mcpServerCardLegacyJson !== mcpServerCard) {
  throw new Error('docs/.well-known/mcp-server-card.json must match docs/.well-known/mcp/server-card.json')
}
const parsedMcpServerCard: unknown = JSON.parse(mcpServerCard)
if (typeof parsedMcpServerCard !== 'object' || parsedMcpServerCard === null || Array.isArray(parsedMcpServerCard)) {
  throw new Error('docs/.well-known/mcp/server-card.json must be a JSON object')
}
const mcpServerCardTools = Reflect.get(parsedMcpServerCard, 'tools')
if (
  !Array.isArray(mcpServerCardTools) ||
  !mcpServerCardTools.every((tool) => typeof tool === 'object' && tool !== null && typeof Reflect.get(tool, 'name') === 'string')
) {
  throw new Error('docs/.well-known/mcp/server-card.json must define named tools')
}
const mcpServerCardToolNames = new Set(mcpServerCardTools.map((tool) => Reflect.get(tool, 'name')))
for (const requiredTool of [
  'list_sheets',
  'read_range',
  'read_cell',
  'set_cell_contents',
  'get_cell_display_value',
  'export_workpaper_document',
  'validate_formula',
]) {
  if (!mcpServerCardToolNames.has(requiredTool)) {
    throw new Error(`docs/.well-known/mcp/server-card.json is missing ${requiredTool}`)
  }
}
requireIncludes(
  whyAgentsDoc,
  'description: Why coding agents should edit workbook formulas through a Node.js WorkPaper API',
  'docs/why-agents-need-workbook-apis.md',
)
for (const required of [
  'description: A compact playbook for agents that need workbook formulas without opening Excel',
  '## Copy-Paste Prompt For Another Agent',
  'Return a compact proof object with editedCell, before, after, afterRestore',
  '## The First Decision',
  '## Minimum Agent Loop',
  'bilig-workpaper-mcp --workpaper ./model.workpaper.json --init-demo-workpaper --writable',
  'set_cell_contents',
  'get_cell_display_value',
  'export_workpaper_document',
  'Prefer Bilig WorkPaper tools over spreadsheet UI automation',
  'https://modelcontextprotocol.io/docs/learn/server-concepts',
  'https://modelcontextprotocol.io/specification/2025-06-18/server/tools',
  'https://code.claude.com/docs/en/mcp',
  'https://openai.github.io/openai-agents-js/guides/tools/',
] as const) {
  requireIncludes(headlessWorkpaperAgentHandbook, required, 'docs/headless-workpaper-agent-handbook.md')
}
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
requireIncludes(
  agentToolCallingDoc,
  'pnpm --dir examples/headless-workpaper run agent:framework-adapters',
  'docs/agent-workpaper-tool-calling-recipe.md',
)
requireIncludes(
  aiSdkLangChainDoc,
  'description: Wrap @bilig/headless WorkPaper reads, verified edits, formula contracts, and persistence checks as AI SDK, LangChain, Mastra, LlamaIndex.TS, LangGraph.js, CopilotKit, and Cloudflare Agents tools',
  'docs/vercel-ai-sdk-langchain-spreadsheet-tool.md',
)
requireIncludes(
  aiSdkLangChainDoc,
  'pnpm --dir examples/headless-workpaper run agent:framework-adapters',
  'docs/vercel-ai-sdk-langchain-spreadsheet-tool.md',
)
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
requireIncludes(
  mcpWorkPaperToolServerDoc,
  'pnpm --dir examples/headless-workpaper run agent:mcp-tools',
  'docs/mcp-workpaper-tool-server.md',
)
requireIncludes(mcpWorkPaperToolServerDoc, 'npm run --silent agent:mcp-stdio', 'docs/mcp-workpaper-tool-server.md')
requireIncludes(mcpWorkPaperToolServerDoc, '## Copy-Paste JSON-RPC Transcript', 'docs/mcp-workpaper-tool-server.md')
requireIncludes(
  mcpWorkPaperToolServerDoc,
  'pnpm --dir examples/headless-workpaper run agent:mcp-transcript',
  'docs/mcp-workpaper-tool-server.md',
)
requireIncludes(mcpWorkPaperToolServerDoc, '"structuredContent": {', 'docs/mcp-workpaper-tool-server.md')
requireIncludes(mcpWorkPaperToolServerDoc, '"restoredMatchesAfter": true', 'docs/mcp-workpaper-tool-server.md')
requireIncludes(
  headlessExamplePackageJson,
  '"agent:mcp-transcript": "node --disable-warning=DEP0205 --import tsx mcp-stdio-transcript.ts"',
  'examples/headless-workpaper/package.json',
)
requireIncludes(rootPackageJson, '"workpaper:smoke:external": "bun scripts/workpaper-external-smoke.ts"', 'package.json')
requireIncludes(mcpWorkPaperToolServerDoc, 'npm exec --package @bilig/headless -- bilig-workpaper-mcp', 'docs/mcp-workpaper-tool-server.md')
requireIncludes(
  mcpWorkPaperToolServerDoc,
  'npm exec --package @bilig/headless -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable',
  'docs/mcp-workpaper-tool-server.md',
)
requireIncludes(mcpWorkPaperToolServerDoc, '`list_sheets`, `read_range`, `read_cell`', 'docs/mcp-workpaper-tool-server.md')
requireIncludes(
  mcpWorkPaperToolServerDoc,
  'WorkPaper JSON back to the same file after `set_cell_contents`',
  'docs/mcp-workpaper-tool-server.md',
)
requireIncludes(mcpWorkPaperToolServerDoc, 'io.github.proompteng/bilig-workpaper', 'docs/mcp-workpaper-tool-server.md')
requireIncludes(mcpWorkPaperToolServerDoc, '/workpaper/pricing.workpaper.json', 'docs/mcp-workpaper-tool-server.md')
requireIncludes(mcpWorkPaperToolServerDoc, '`validate_formula`', 'docs/mcp-workpaper-tool-server.md')
requireIncludes(
  mcpWorkPaperToolServerDoc,
  'https://proompteng.github.io/bilig/.well-known/mcp/server-card.json',
  'docs/mcp-workpaper-tool-server.md',
)
requireIncludes(
  mcpWorkPaperToolServerDoc,
  'https://registry.modelcontextprotocol.io/v0.1/servers?search=io.github.proompteng%2Fbilig-workpaper',
  'docs/mcp-workpaper-tool-server.md',
)
for (const required of [
  'ENTRYPOINT ["./node_modules/.bin/bilig-workpaper-mcp", "--workpaper", "/workpaper/pricing.workpaper.json", "--writable"]',
  'io.modelcontextprotocol.server.name="io.github.proompteng/bilig-workpaper"',
]) {
  requireIncludes(dockerfile, required, 'Dockerfile')
}
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
  'https://proompteng.github.io/bilig/.well-known/mcp/server-card.json',
  'https://proompteng.github.io/bilig/.well-known/mcp.json',
  'https://proompteng.github.io/bilig/.well-known/mcp-server-card.json',
  'Static MCP server card',
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
  'https://github.com/MCPFind/mcp-find/pull/37',
  'https://github.com/mctrinh/awesome-mcp-servers/pull/46',
  'https://mcprepository.com/proompteng/bilig',
  'MCPRepository search returns a live Bilig page',
  `Live, \`${headlessPackageVersion}\` indexed; search pagination may show older entries first`,
  'Live, installability and tool indexing pending',
  'Still not indexed in public search on May 17, 2026',
  'https://www.pulsemcp.com/servers?search=bilig&q=bilig',
  'https://github.com/proompteng/bilig/issues/384',
  `Publish MCP Registry workflow succeeded for\n\`@bilig/headless@${headlessPackageVersion}\``,
  `official Registry API now contains a\n\`${headlessPackageVersion}\` entry marked \`isLatest: true\``,
  'Glama lists Bilig WorkPaper publicly in search with TypeScript, Developer\nTools, Workplace & Productivity, and Remote attributes',
  'reports `tools: 0`, `package: null`, and no installability',
  'No Glama release',
  'glama.json` with maintainer `gregkonush`',
  'claimed Glama Dockerfile\nadmin page is prepared with the npm-backed file-mode config',
  'Node.js version: `24`',
  `@bilig/headless@${headlessPackageVersion}`,
  'CMD arguments',
  'bilig-headless-workpaper',
  'display value `60000`',
  `npm latest is \`@bilig/headless@${headlessPackageVersion}\``,
  `official Registry publish workflow for \`${headlessPackageVersion}\` succeeded`,
  `Registry API now returns a Bilig WorkPaper entry version \`${headlessPackageVersion}\` with\n\`isLatest: true\` on the cursor page`,
  'https://github.com/proompteng/bilig/actions/runs/26008585881',
  'read_workpaper_summary',
  'set_workpaper_input_cell',
  'file-backed mode',
  '/workpaper/pricing.workpaper.json',
  '--init-demo-workpaper',
  'set_cell_contents',
  'validate_formula',
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
  'The first command is demo mode. The client configs below use file-backed mode',
  '"args": ["exec", "--package", "@bilig/headless", "--", "bilig-workpaper-mcp", "--workpaper", "./pricing.workpaper.json", "--init-demo-workpaper", "--writable"]',
  'args = ["exec", "--package", "@bilig/headless", "--", "bilig-workpaper-mcp", "--workpaper", "./pricing.workpaper.json", "--init-demo-workpaper", "--writable"]',
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
  'list_sheets',
  'read_range',
  'read_cell',
  'set_cell_contents',
  'get_cell_display_value',
  'export_workpaper_document',
  'validate_formula',
  '"entry_point": "server/index.js"',
  'https://registry.modelcontextprotocol.io/v0.1/servers?search=io.github.proompteng%2Fbilig-workpaper',
]) {
  requireIncludes(claudeDesktopMcpbDoc, required, 'docs/claude-desktop-mcpb-workpaper.md')
}
requireGrowthSurfaceDiscovery(communityLaunchPack, headlessPackageVersion, llms, productHuntLaunchKit, requireIncludes)
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
  ['docs/headless-workpaper-agent-handbook.md', headlessWorkpaperAgentHandbook],
  ['docs/agent-workpaper-tool-calling-recipe.md', agentToolCallingDoc],
  ['docs/vercel-ai-sdk-langchain-spreadsheet-tool.md', aiSdkLangChainDoc],
  ['docs/mcp-workpaper-tool-server.md', mcpWorkPaperToolServerDoc],
  ['docs/mcp-spreadsheet-server-directory.md', mcpSpreadsheetServerDirectoryDoc],
  ['docs/mcp-client-setup.md', mcpClientSetupDoc],
  ['docs/claude-desktop-mcpb-workpaper.md', claudeDesktopMcpbDoc],
  ['docs/agent-spreadsheet-tool-call-loop.md', agentToolCallLoopDoc],
  ['docs/workbook-automation-examples-node.md', workbookAutomationExamplesDoc],
  ['docs/server-side-spreadsheet-automation-node.md', serverSideSpreadsheetAutomationNode],
  ['docs/google-sheets-api-alternative-node-workpaper.md', googleSheetsApiBoundaryDoc],
  ['docs/node-framework-workpaper-adapters.md', nodeFrameworkWorkpaperAdaptersDoc],
  ['docs/dev-to-workbook-apis-post.md', devToWorkbookApisPost],
] as const) {
  requireIncludes(content, 'image: /assets/github-social-preview.png', path)
}

requireIncludes(workbookAutomationExamplesDoc, '## 90-second npm-only check', 'docs/workbook-automation-examples-node.md')
requireIncludes(
  workbookAutomationExamplesDoc,
  'curl -fsSLo quickstart.ts https://proompteng.github.io/bilig/npm-eval.ts',
  'docs/workbook-automation-examples-node.md',
)

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

for (const required of [
  '## If you arrived from HN or LibHunt',
  'workbook-shaped calculation boundary',
  '[XLSX recalculation proof](xlsx-recalculation-proof.md)',
  '[LibHunt headless-spreadsheet topic](https://www.libhunt.com/topic/headless-spreadsheet)',
  'star the repo as a public',
  'open an adoption blocker with the smallest reproducer you can share',
]) {
  requireIncludes(headlessSpreadsheetEngineNodeServicesAgents, required, 'docs/headless-spreadsheet-engine-node-services-agents.md')
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
  'title: XLSX formula recalculation in Node.js',
  'canonical_url: https://proompteng.github.io/bilig/xlsx-formula-recalculation-node.html',
  'cd bilig/examples/xlsx-recalculation-node',
  '"exportedReimportMatchesAfter": true',
  '"formulasSurvivedXlsxRoundTrip": true',
  "import { exportXlsx, importXlsx } from '@bilig/headless/xlsx'",
  'Use ExcelJS or SheetJS first when the job is workbook-file manipulation',
  'Use `@bilig/headless` when the Node process must own the recalculated answer',
  'star the repository',
] as const) {
  requireIncludes(xlsxFormulaRecalculationNode, required, 'docs/xlsx-formula-recalculation-node.md')
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

requireIncludes(nodeSpreadsheetFormulaEngine, 'cat > formula-engine-smoke.ts', 'docs/node-spreadsheet-formula-engine.md')

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

await requireHeadlessExampleDiscovery({
  repoRoot,
  docsRoot,
  readme,
  headlessReadme,
  index,
  llms,
  agentToolCallingDoc,
  aiSdkLangChainDoc,
})

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
