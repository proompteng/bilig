import { readFile } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

import { getBenchmarkDiscoveryEvidence } from './check-docs-discovery-benchmark-evidence.ts'
import { docsSiteSources } from './check-docs-discovery-site-sources.ts'

export interface DocsDiscoveryContext {
  readonly repoRoot: string
  readonly docsRoot: string
  readonly siteRoot: string
  readonly expectedSitemapUrls: readonly string[]
  readonly sourceFilesByUrl: ReadonlyMap<string, string>
  readonly benchmarkEvidence: ReturnType<typeof getBenchmarkDiscoveryEvidence>
  readonly headlessPackageVersion: string
  readonly readme: string
  readonly contributing: string
  readonly rootPackageJson: string
  readonly index: string
  readonly siteCss: string
  readonly productCss: string
  readonly robots: string
  readonly sitemap: string
  readonly llms: string
  readonly llmsFull: string
  readonly agentJson: string
  readonly agentJsonRoot: string
  readonly docsAgentNotes: string
  readonly docsSkill: string
  readonly agentSkillsIndex: string
  readonly legacySkillsIndex: string
  readonly communityLaunchPack: string
  readonly productHuntLaunchKit: string
  readonly starterIssues: string
  readonly newContributorGuide: string
  readonly headlessPackageJson: string
  readonly headlessExamplePackageJson: string
  readonly headlessReadme: string
  readonly headlessAgentNotes: string
  readonly headlessSkillNotes: string
  readonly excelImportReadme: string
  readonly dockerfile: string
  readonly publicApi: string
  readonly issueTemplateConfig: string
  readonly issueTemplateRoot: string
  readonly featureRequestTemplate: string
  readonly ideasDiscussionTemplate: string
  readonly qaDiscussionTemplate: string
  readonly showAndTellDiscussionTemplate: string
  readonly generalDiscussionTemplate: string
  readonly pullRequestTemplate: string
  readonly dominanceScorecard: string
  readonly headlessSpreadsheetEngineComparison: string
  readonly sheetjsExceljsAlternativeFormulaWorkbookApi: string
  readonly hyperformulaAlternativeHeadlessWorkpaper: string
  readonly xlsxFormulaRecalculationNode: string
  readonly agentXlsxFormulaRecalculationWithoutLibreOffice: string
  readonly staleXlsxFormulaCacheNode: string
  readonly microsoftGraphExcelRecalculationNode: string
  readonly formulaWorkbooksProof: string
  readonly showHnFormulaWorkbooksProof: string
  readonly googleSheetsApiBoundaryDoc: string
  readonly npmProvenancePackageTrustDoc: string
  readonly xlsxCorpusVerifierWalkthrough: string
  readonly whyAgentsDoc: string
  readonly headlessWorkpaperAgentHandbook: string
  readonly agentToolCallingDoc: string
  readonly aiSdkLangChainDoc: string
  readonly mcpWorkPaperToolServerDoc: string
  readonly mcpSpreadsheetServerDirectoryDoc: string
  readonly mcpClientSetupDoc: string
  readonly claudeDesktopMcpbDoc: string
  readonly agentToolCallLoopDoc: string
  readonly mcpServerCard: string
  readonly mcpServerCardMcpJson: string
  readonly mcpServerCardLegacyJson: string
  readonly workbookAutomationExamplesDoc: string
  readonly serverSideSpreadsheetAutomationNode: string
  readonly nodeFrameworkWorkpaperAdaptersDoc: string
  readonly devToWorkbookApisPost: string
  readonly evaluateExcelFormulasInNodeTypescript: string
  readonly nodeSpreadsheetFormulaEngine: string
}

export async function loadDocsDiscoveryContext(): Promise<DocsDiscoveryContext> {
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
    llmsFull,
    agentJson,
    agentJsonRoot,
    docsAgentNotes,
    docsSkill,
    agentSkillsIndex,
    legacySkillsIndex,
    communityLaunchPack,
    productHuntLaunchKit,
    starterIssues,
    newContributorGuide,
    headlessPackageJson,
    headlessExamplePackageJson,
    headlessReadme,
    headlessAgentNotes,
    headlessSkillNotes,
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
    readFile(join(docsRoot, 'llms-full.txt'), 'utf8'),
    readFile(join(docsRoot, '.well-known', 'agent.json'), 'utf8'),
    readFile(join(docsRoot, 'agent.json'), 'utf8'),
    readFile(join(docsRoot, 'AGENTS.md'), 'utf8'),
    readFile(join(docsRoot, 'skill.md'), 'utf8'),
    readFile(join(docsRoot, '.well-known', 'agent-skills', 'index.json'), 'utf8'),
    readFile(join(docsRoot, '.well-known', 'skills', 'index.json'), 'utf8'),
    readFile(join(docsRoot, 'community-launch-pack.md'), 'utf8'),
    readFile(join(docsRoot, 'product-hunt-launch-kit.md'), 'utf8'),
    readFile(join(docsRoot, 'starter-issues.md'), 'utf8'),
    readFile(join(docsRoot, 'new-contributor-guide.md'), 'utf8'),
    readFile(join(repoRoot, 'packages', 'headless', 'package.json'), 'utf8'),
    readFile(join(repoRoot, 'examples', 'headless-workpaper', 'package.json'), 'utf8'),
    readFile(join(repoRoot, 'packages', 'headless', 'README.md'), 'utf8'),
    readFile(join(repoRoot, 'packages', 'headless', 'AGENTS.md'), 'utf8'),
    readFile(join(repoRoot, 'packages', 'headless', 'SKILL.md'), 'utf8'),
    readFile(join(repoRoot, 'packages', 'excel-import', 'README.md'), 'utf8'),
    readFile(join(repoRoot, 'Dockerfile'), 'utf8'),
    readFile(join(docsRoot, 'public-api.md'), 'utf8'),
    readFile(join(repoRoot, '.github', 'ISSUE_TEMPLATE', 'config.yml'), 'utf8'),
    readFile(join(repoRoot, '.github', 'ISSUE_TEMPLATE.md'), 'utf8'),
    readFile(join(repoRoot, '.github', 'ISSUE_TEMPLATE', 'feature_request.yml'), 'utf8'),
    readFile(join(repoRoot, '.github', 'DISCUSSION_TEMPLATE', 'ideas.yml'), 'utf8'),
    readFile(join(repoRoot, '.github', 'DISCUSSION_TEMPLATE', 'q-a.yml'), 'utf8'),
    readFile(join(repoRoot, '.github', 'DISCUSSION_TEMPLATE', 'show-and-tell.yml'), 'utf8'),
    readFile(join(repoRoot, '.github', 'DISCUSSION_TEMPLATE', 'general.yml'), 'utf8'),
    readFile(join(repoRoot, '.github', 'PULL_REQUEST_TEMPLATE.md'), 'utf8'),
    readFile(join(repoRoot, 'packages', 'benchmarks', 'baselines', 'bilig-dominance-scorecard.json'), 'utf8'),
    readFile(join(docsRoot, 'headless-spreadsheet-engine-comparison.md'), 'utf8'),
    readFile(join(docsRoot, 'sheetjs-exceljs-alternative-formula-workbook-api.md'), 'utf8'),
    readFile(join(docsRoot, 'hyperformula-alternative-headless-workpaper.md'), 'utf8'),
    readFile(join(docsRoot, 'xlsx-formula-recalculation-node.md'), 'utf8'),
    readFile(join(docsRoot, 'agent-xlsx-formula-recalculation-without-libreoffice.md'), 'utf8'),
    readFile(join(docsRoot, 'stale-xlsx-formula-cache-node.md'), 'utf8'),
    readFile(join(docsRoot, 'microsoft-graph-excel-recalculation-node.md'), 'utf8'),
    readFile(join(docsRoot, 'formula-workbooks-node-services-agent-tools.md'), 'utf8'),
    readFile(join(docsRoot, 'show-hn-formula-workbooks-node-services.md'), 'utf8'),
    readFile(join(docsRoot, 'google-sheets-api-alternative-node-workpaper.md'), 'utf8'),
    readFile(join(docsRoot, 'npm-provenance-package-trust.md'), 'utf8'),
    readFile(join(docsRoot, 'xlsx-corpus-verifier-walkthrough.md'), 'utf8'),
    readFile(join(docsRoot, 'why-agents-need-workbook-apis.md'), 'utf8'),
    readFile(join(docsRoot, 'headless-workpaper-agent-handbook.md'), 'utf8'),
    readFile(join(docsRoot, 'agent-workpaper-tool-calling-recipe.md'), 'utf8'),
    readFile(join(docsRoot, 'vercel-ai-sdk-langchain-spreadsheet-tool.md'), 'utf8'),
    readFile(join(docsRoot, 'mcp-workpaper-tool-server.md'), 'utf8'),
    readFile(join(docsRoot, 'mcp-spreadsheet-server-directory.md'), 'utf8'),
    readFile(join(docsRoot, 'mcp-client-setup.md'), 'utf8'),
    readFile(join(docsRoot, 'claude-desktop-mcpb-workpaper.md'), 'utf8'),
    readFile(join(docsRoot, 'agent-spreadsheet-tool-call-loop.md'), 'utf8'),
    readFile(join(docsRoot, '.well-known', 'mcp', 'server-card.json'), 'utf8'),
    readFile(join(docsRoot, '.well-known', 'mcp.json'), 'utf8'),
    readFile(join(docsRoot, '.well-known', 'mcp-server-card.json'), 'utf8'),
    readFile(join(docsRoot, 'workbook-automation-examples-node.md'), 'utf8'),
    readFile(join(docsRoot, 'server-side-spreadsheet-automation-node.md'), 'utf8'),
    readFile(join(docsRoot, 'node-framework-workpaper-adapters.md'), 'utf8'),
    readFile(join(docsRoot, 'dev-to-workbook-apis-post.md'), 'utf8'),
    readFile(join(docsRoot, 'evaluate-excel-formulas-in-node-typescript.md'), 'utf8'),
    readFile(join(docsRoot, 'node-spreadsheet-formula-engine.md'), 'utf8'),
  ])

  return {
    repoRoot,
    docsRoot,
    siteRoot,
    expectedSitemapUrls,
    sourceFilesByUrl,
    benchmarkEvidence: getBenchmarkDiscoveryEvidence(),
    headlessPackageVersion: parseHeadlessPackageVersion(headlessPackageJson),
    readme,
    contributing,
    rootPackageJson,
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
    productHuntLaunchKit,
    starterIssues,
    newContributorGuide,
    headlessPackageJson,
    headlessExamplePackageJson,
    headlessReadme,
    headlessAgentNotes,
    headlessSkillNotes,
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
  }
}

export function parseHeadlessPackageVersion(packageJson: string): string {
  const parsedPackage: unknown = JSON.parse(packageJson)
  const version =
    typeof parsedPackage === 'object' && parsedPackage !== null && !Array.isArray(parsedPackage)
      ? Reflect.get(parsedPackage, 'version')
      : undefined
  if (typeof version !== 'string') {
    throw new Error('packages/headless/package.json is missing a string version')
  }
  return version
}
