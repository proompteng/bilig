import { readFile } from 'node:fs/promises'
import { join } from 'node:path'
import { agentFrameworkLlmsRequiredLinks } from './check-docs-discovery-agent-pages.ts'
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
import { productHuntLaunchAssetFiles } from './check-docs-discovery-launch-kit.ts'
import { requireFormulaProofDiscovery } from './check-docs-discovery-proof-pages.ts'
import { requireTypeScriptFirstPublicSnippets } from './check-docs-discovery-typescript-snippets.ts'
import { requireXlsxCorpusVerifierDiscovery } from './check-docs-discovery-xlsx-verifier.ts'
import { requireXlsxCalcAlternativeDiscovery } from './check-docs-discovery-xlsx-calc.ts'
import { requireSharedPublicDocsDiscovery } from './check-docs-discovery-public-docs.ts'
import { homepageRequiredLinks, llmsRequiredLinks } from './check-docs-discovery-public-link-manifest.ts'
import { requireAgentPublicSurfaceDiscovery } from './check-docs-discovery-agent-surfaces.ts'

const docsDiscoveryContext = await loadDocsDiscoveryContext()
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
  index,
  siteCss,
  productCss,
  robots,
  sitemap,
  llms,
  llmsFull,
  agentJson,
  agentJsonRoot,
  docsAgentNotes,
  docsSkill,
  agentSkillsIndex,
  legacySkillsIndex,
  communityLaunchPack,
  starterIssues,
  newContributorGuide,
  headlessPackageJson,
  headlessReadme,
  headlessAgentNotes,
  headlessSkillNotes,
  excelImportReadme,
  publicApi,
  issueTemplateConfig,
  issueTemplateRoot,
  featureRequestTemplate,
  ideasDiscussionTemplate,
  qaDiscussionTemplate,
  showAndTellDiscussionTemplate,
  generalDiscussionTemplate,
  dominanceScorecard,
  agentXlsxFormulaRecalculationWithoutLibreOffice,
  staleXlsxFormulaCacheNode,
  microsoftGraphExcelRecalculationNode,
  formulaWorkbooksProof,
  showHnFormulaWorkbooksProof,
  googleSheetsApiBoundaryDoc,
  npmProvenancePackageTrustDoc,
  xlsxCorpusVerifierWalkthrough,
  serverSideSpreadsheetAutomationNode,
  evaluateExcelFormulasInNodeTypescript,
} = docsDiscoveryContext
const headlessPackageSpec = `@bilig/headless@${headlessPackageVersion}`
const mcpbReleaseAssetUrl = `https://github.com/proompteng/bilig/releases/download/libraries-v${headlessPackageVersion}/bilig-workpaper.mcpb`
const mcpbReleaseChecksumUrl = `${mcpbReleaseAssetUrl}.sha256`

const headlessSpreadsheetEngineNodeServicesAgents = await readFile(
  join(docsRoot, 'headless-spreadsheet-engine-node-services-agents.md'),
  'utf8',
)
const spreadsheetMcpServerComparison = await readFile(join(docsRoot, 'spreadsheet-mcp-server-comparison.md'), 'utf8')
const rootSkillNotes = await readFile(join(repoRoot, 'skills', 'bilig-workpaper', 'SKILL.md'), 'utf8')

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
requireIncludes(
  index,
  '<link rel="alternate" type="application/json" href="https://proompteng.github.io/bilig/.well-known/agent.json" title="agent.json" />',
  'docs/index.html',
)
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
requireIncludes(llms, 'https://proompteng.github.io/bilig/AGENTS.md', 'docs/llms.txt')
requireIncludes(llms, 'https://proompteng.github.io/bilig/.well-known/agent.json', 'docs/llms.txt')
requireIncludes(llms, 'https://proompteng.github.io/bilig/agent.json', 'docs/llms.txt')
requireIncludes(llms, 'https://proompteng.github.io/bilig/skill.txt', 'docs/llms.txt')
requireIncludes(llms, 'https://proompteng.github.io/bilig/llms-full.txt', 'docs/llms.txt')
requireIncludes(llms, 'https://proompteng.github.io/bilig/.well-known/agent-skills/index.json', 'docs/llms.txt')
requireIncludes(llms, 'https://proompteng.github.io/bilig/.well-known/skills/index.json', 'docs/llms.txt')
requireIncludes(readme, 'docs/.well-known/agent.json', 'README.md')
requireIncludes(headlessReadme, 'https://proompteng.github.io/bilig/.well-known/agent.json', 'packages/headless/README.md')
requireIncludes(llms, 'Do not claim success from a write call alone.', 'docs/llms.txt')
requireIncludes(llms, 'pnpm --dir bilig/examples/headless-workpaper install --ignore-workspace', 'docs/llms.txt')
requireIncludes(llms, 'pnpm --dir bilig/examples/headless-workpaper run agent:framework-adapters', 'docs/llms.txt')
requireIncludes(llms, 'pnpm --dir examples/headless-workpaper run agent:mcp-tools', 'docs/llms.txt')
requireNotIncludes(llms, 'cd bilig/examples/headless-workpaper', 'docs/llms.txt')
requireNotIncludes(llms, '\nnpm start\n', 'docs/llms.txt')
requireIncludes(headlessReadme, 'https://proompteng.github.io/bilig/why-use-bilig.html', 'packages/headless/README.md')
requireIncludes(headlessReadme, 'The npm tarball also includes `AGENTS.md`', 'packages/headless/README.md')
requireIncludes(headlessReadme, 'SKILL.md', 'packages/headless/README.md')
requireIncludes(headlessPackageJson, '"AGENTS.md"', 'packages/headless/package.json')
requireIncludes(headlessPackageJson, '"SKILL.md"', 'packages/headless/package.json')
requireIncludes(headlessAgentNotes, '## Handoff prompt', 'packages/headless/AGENTS.md')
requireIncludes(headlessAgentNotes, 'Do not claim success from a write call alone.', 'packages/headless/AGENTS.md')
requireIncludes(headlessAgentNotes, `npm exec --package ${headlessPackageSpec} -- bilig-mcp-challenge`, 'packages/headless/AGENTS.md')
requireIncludes(
  headlessAgentNotes,
  `npm exec --package ${headlessPackageSpec} -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable`,
  'packages/headless/AGENTS.md',
)
requireIncludes(headlessSkillNotes, 'name: bilig-workpaper', 'packages/headless/SKILL.md')
requireIncludes(headlessSkillNotes, '"bilig-formula-clinic", "./reduced.xlsx"', 'packages/headless/SKILL.md')
requireIncludes(headlessSkillNotes, 'Do not trigger it for manual spreadsheet editing', 'packages/headless/SKILL.md')
requireIncludes(headlessSkillNotes, '## Command Safety', 'packages/headless/SKILL.md')
requireIncludes(headlessSkillNotes, 'argument array, not a shell-concatenated string', 'packages/headless/SKILL.md')
requireNotIncludes(headlessSkillNotes, 'allowed-tools:', 'packages/headless/SKILL.md')
requireNotIncludes(headlessSkillNotes, 'argument-hint:', 'packages/headless/SKILL.md')
requireIncludes(docsAgentNotes, '## Discovery Order', 'docs/AGENTS.md')
requireIncludes(docsAgentNotes, 'Do not claim success from a write call alone.', 'docs/AGENTS.md')
requireIncludes(docsSkill, 'name: bilig-workpaper', 'docs/skill.md')
requireIncludes(docsSkill, '## Required Verification', 'docs/skill.md')
requireIncludes(docsSkill, '## Command Safety', 'docs/skill.md')
requireNotIncludes(docsSkill, 'allowed-tools:', 'docs/skill.md')
requireNotIncludes(docsSkill, 'argument-hint:', 'docs/skill.md')
requireIncludes(rootSkillNotes, '## Command Safety', 'skills/bilig-workpaper/SKILL.md')
requireIncludes(rootSkillNotes, 'argument array, not a shell-concatenated string', 'skills/bilig-workpaper/SKILL.md')
requireNotIncludes(rootSkillNotes, 'allowed-tools:', 'skills/bilig-workpaper/SKILL.md')
requireNotIncludes(rootSkillNotes, 'argument-hint:', 'skills/bilig-workpaper/SKILL.md')
if (agentJsonRoot !== agentJson) {
  throw new Error('docs/agent.json must match docs/.well-known/agent.json')
}
const parsedAgentJson: unknown = JSON.parse(agentJson)
if (typeof parsedAgentJson !== 'object' || parsedAgentJson === null || Array.isArray(parsedAgentJson)) {
  throw new Error('docs/.well-known/agent.json must be a JSON object')
}
for (const [fieldName, expectedValue] of [
  ['name', 'bilig'],
  ['repository', 'https://github.com/proompteng/bilig'],
  ['llms_txt', 'https://proompteng.github.io/bilig/llms.txt'],
  ['llms_full', 'https://proompteng.github.io/bilig/llms-full.txt'],
  ['skill_file', 'https://proompteng.github.io/bilig/skill.txt'],
  ['agent_instructions', 'https://proompteng.github.io/bilig/AGENTS.md'],
] as const) {
  if (Reflect.get(parsedAgentJson, fieldName) !== expectedValue) {
    throw new Error(`docs/.well-known/agent.json ${fieldName} must be ${expectedValue}`)
  }
}
const parsedAgentJsonMcp = Reflect.get(parsedAgentJson, 'mcp')
if (typeof parsedAgentJsonMcp !== 'object' || parsedAgentJsonMcp === null || Array.isArray(parsedAgentJsonMcp)) {
  throw new Error('docs/.well-known/agent.json must define an mcp object')
}
if (Reflect.get(parsedAgentJsonMcp, 'server_card') !== 'https://proompteng.github.io/bilig/.well-known/mcp/server-card.json') {
  throw new Error('docs/.well-known/agent.json must point at the MCP server card')
}
if (Reflect.get(parsedAgentJsonMcp, 'remote_endpoint') !== 'https://bilig.proompteng.ai/mcp') {
  throw new Error('docs/.well-known/agent.json must advertise the hosted MCP endpoint')
}
const agentJsonMcpRemoteTransport = Reflect.get(parsedAgentJsonMcp, 'remote_transport')
if (
  typeof agentJsonMcpRemoteTransport !== 'object' ||
  agentJsonMcpRemoteTransport === null ||
  Reflect.get(agentJsonMcpRemoteTransport, 'type') !== 'streamable-http' ||
  Reflect.get(agentJsonMcpRemoteTransport, 'protocol_version') !== '2025-11-25'
) {
  throw new Error('docs/.well-known/agent.json must advertise the hosted Streamable HTTP MCP transport')
}
const agentJsonMcpTools = Reflect.get(parsedAgentJsonMcp, 'tools')
if (!Array.isArray(agentJsonMcpTools) || !agentJsonMcpTools.every((tool) => typeof tool === 'string')) {
  throw new Error('docs/.well-known/agent.json mcp.tools must be a string array')
}
for (const requiredTool of [
  'list_sheets',
  'set_cell_contents',
  'get_cell_display_value',
  'export_workpaper_document',
  'validate_formula',
]) {
  if (!agentJsonMcpTools.includes(requiredTool)) {
    throw new Error(`docs/.well-known/agent.json mcp.tools is missing ${requiredTool}`)
  }
}
const agentJsonMcpResources = Reflect.get(parsedAgentJsonMcp, 'resources')
if (!Array.isArray(agentJsonMcpResources) || !agentJsonMcpResources.every((resource) => typeof resource === 'string')) {
  throw new Error('docs/.well-known/agent.json mcp.resources must be a string array')
}
for (const requiredResource of ['bilig://workpaper/manifest', 'bilig://workpaper/agent-handoff', 'bilig://workpaper/current-document']) {
  if (!agentJsonMcpResources.includes(requiredResource)) {
    throw new Error(`docs/.well-known/agent.json mcp.resources is missing ${requiredResource}`)
  }
}
const agentJsonMcpPrompts = Reflect.get(parsedAgentJsonMcp, 'prompts')
if (!Array.isArray(agentJsonMcpPrompts) || !agentJsonMcpPrompts.every((prompt) => typeof prompt === 'string')) {
  throw new Error('docs/.well-known/agent.json mcp.prompts must be a string array')
}
for (const requiredPrompt of ['edit_and_verify_workpaper', 'debug_workpaper_formula']) {
  if (!agentJsonMcpPrompts.includes(requiredPrompt)) {
    throw new Error(`docs/.well-known/agent.json mcp.prompts is missing ${requiredPrompt}`)
  }
}
const agentJsonCapabilities = Reflect.get(parsedAgentJson, 'capabilities')
if (
  !Array.isArray(agentJsonCapabilities) ||
  !agentJsonCapabilities.some(
    (capability) =>
      typeof capability === 'object' &&
      capability !== null &&
      Reflect.get(capability, 'name') === 'file-backed-workpaper-mcp' &&
      Reflect.get(capability, 'server_card') === 'https://proompteng.github.io/bilig/.well-known/mcp/server-card.json',
  )
) {
  throw new Error('docs/.well-known/agent.json must advertise the file-backed MCP capability')
}
if (
  !agentJsonCapabilities.some(
    (capability) =>
      typeof capability === 'object' &&
      capability !== null &&
      Reflect.get(capability, 'name') === 'remote-workpaper-mcp-demo' &&
      Reflect.get(capability, 'endpoint') === 'https://bilig.proompteng.ai/mcp',
  )
) {
  throw new Error('docs/.well-known/agent.json must advertise the remote MCP demo capability')
}
if (
  !agentJsonCapabilities.some(
    (capability) =>
      typeof capability === 'object' &&
      capability !== null &&
      Reflect.get(capability, 'name') === 'claude-desktop-mcpb' &&
      Reflect.get(capability, 'type') === 'mcpb-desktop-extension' &&
      Reflect.get(capability, 'package_version') === headlessPackageVersion &&
      Reflect.get(capability, 'download_url') === mcpbReleaseAssetUrl &&
      Reflect.get(capability, 'checksum_url') === mcpbReleaseChecksumUrl,
  )
) {
  throw new Error('docs/.well-known/agent.json must advertise the Claude Desktop MCPB release asset')
}
requireIncludes(
  agentSkillsIndex,
  'https://proompteng.github.io/bilig/.well-known/agent-skills/bilig-workpaper/SKILL.txt',
  'docs/.well-known/agent-skills/index.json',
)
requireIncludes(
  legacySkillsIndex,
  'https://proompteng.github.io/bilig/.well-known/skills/bilig-workpaper/SKILL.txt',
  'docs/.well-known/skills/index.json',
)
requireIncludes(llmsFull, '## Generated Skill Manifest', 'docs/llms-full.txt')
requireIncludes(llmsFull, '## Headless WorkPaper Agent Handbook', 'docs/llms-full.txt')
requireIncludes(llmsFull, `npm exec --package ${headlessPackageSpec} -- bilig-mcp-challenge`, 'docs/llms-full.txt')
requireIncludes(llmsFull, `npm exec --package ${headlessPackageSpec} -- bilig-workpaper-mcp`, 'docs/llms-full.txt')
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

await requireAgentPublicSurfaceDiscovery({
  context: docsDiscoveryContext,
  headlessSpreadsheetEngineNodeServicesAgents,
  spreadsheetMcpServerComparison,
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
