import { readFile } from 'node:fs/promises'
import { join } from 'node:path'

import { agentFrameworkDocRequirements } from './check-docs-discovery-agent-pages.ts'
import { requireIncludes, requireNotIncludes } from './check-docs-discovery-core.ts'
import type { DocsDiscoveryContext } from './check-docs-discovery-context.ts'
import { llmsExternalSurfaceLinks } from './check-docs-discovery-growth-links.ts'
import { requireHeadlessExampleDiscovery } from './check-docs-discovery-headless-examples.ts'
import { requireGrowthSurfaceDiscovery } from './check-docs-discovery-launch-kit.ts'
import { requireStarterIssueDiscovery } from './check-docs-discovery-starter-issues.ts'

export async function requireAgentPublicSurfaceDiscovery(input: {
  readonly context: DocsDiscoveryContext
  readonly headlessSpreadsheetEngineNodeServicesAgents: string
  readonly spreadsheetMcpServerComparison: string
}): Promise<void> {
  const {
    repoRoot,
    docsRoot,
    rootPackageJson,
    index,
    llms,
    mcpServerCard,
    mcpServerCardMcpJson,
    mcpServerCardLegacyJson,
    communityLaunchPack,
    productHuntLaunchKit,
    starterIssues,
    headlessPackageVersion,
    readme,
    headlessReadme,
    issueTemplateConfig,
    pullRequestTemplate,
    dockerfile,
    headlessExamplePackageJson,
    headlessSpreadsheetEngineComparison,
    sheetjsExceljsAlternativeFormulaWorkbookApi,
    hyperformulaAlternativeHeadlessWorkpaper,
    xlsxFormulaRecalculationNode,
    googleSheetsApiBoundaryDoc,
    whyAgentsDoc,
    headlessWorkpaperAgentHandbook,
    agentToolCallingDoc,
    aiSdkLangChainDoc,
    mcpWorkPaperToolServerDoc,
    mcpSpreadsheetServerDirectoryDoc,
    mcpClientSetupDoc,
    claudeDesktopMcpbDoc,
    agentToolCallLoopDoc,
    workbookAutomationExamplesDoc,
    serverSideSpreadsheetAutomationNode,
    nodeFrameworkWorkpaperAdaptersDoc,
    devToWorkbookApisPost,
    nodeSpreadsheetFormulaEngine,
  } = input.context
  const { headlessSpreadsheetEngineNodeServicesAgents, spreadsheetMcpServerComparison } = input
  const headlessPackageSpec = `@bilig/headless@${headlessPackageVersion}`
  const mcpbReleaseAssetUrl = `https://github.com/proompteng/bilig/releases/download/libraries-v${headlessPackageVersion}/bilig-workpaper.mcpb`
  const mcpbReleaseChecksumUrl = `${mcpbReleaseAssetUrl}.sha256`
  const officialRegistryLatestMarkedVersion = '0.27.0'

  const jekyllConfig = await readFile(join(docsRoot, '_config.yml'), 'utf8')
  const openAiAgentsSdkDoc = await readFile(join(docsRoot, 'openai-agents-sdk-workpaper-tool.md'), 'utf8')
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
  if (Reflect.get(parsedMcpServerCard, 'protocolVersion') !== '2025-11-25') {
    throw new Error('docs/.well-known/mcp/server-card.json must advertise the latest MCP protocol version')
  }
  const mcpServerInfo = Reflect.get(parsedMcpServerCard, 'serverInfo')
  if (typeof mcpServerInfo !== 'object' || mcpServerInfo === null || Reflect.get(mcpServerInfo, 'version') !== headlessPackageVersion) {
    throw new Error('docs/.well-known/mcp/server-card.json serverInfo.version must match @bilig/headless')
  }
  const mcpServerCardTransports = Reflect.get(parsedMcpServerCard, 'transports')
  if (
    !Array.isArray(mcpServerCardTransports) ||
    !mcpServerCardTransports.some(
      (transport) =>
        typeof transport === 'object' &&
        transport !== null &&
        Reflect.get(transport, 'type') === 'streamable-http' &&
        Reflect.get(transport, 'url') === 'https://bilig.proompteng.ai/mcp' &&
        Reflect.get(transport, 'stateless') === true,
    )
  ) {
    throw new Error('docs/.well-known/mcp/server-card.json must advertise the hosted stateless Streamable HTTP endpoint')
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
  for (const tool of mcpServerCardTools) {
    requireMcpToolDiscoveryContract(tool)
  }
  const mcpServerCardCapabilities = Reflect.get(parsedMcpServerCard, 'capabilities')
  if (
    typeof mcpServerCardCapabilities !== 'object' ||
    mcpServerCardCapabilities === null ||
    Reflect.get(mcpServerCardCapabilities, 'resources') !== true ||
    Reflect.get(mcpServerCardCapabilities, 'prompts') !== true
  ) {
    throw new Error('docs/.well-known/mcp/server-card.json must advertise resources and prompts')
  }
  const mcpServerCardResources = Reflect.get(parsedMcpServerCard, 'resources')
  if (
    !Array.isArray(mcpServerCardResources) ||
    !mcpServerCardResources.every(
      (resource) => typeof resource === 'object' && resource !== null && typeof Reflect.get(resource, 'uri') === 'string',
    )
  ) {
    throw new Error('docs/.well-known/mcp/server-card.json must define resource URIs')
  }
  const mcpServerCardResourceUris = new Set(mcpServerCardResources.map((resource) => Reflect.get(resource, 'uri')))
  for (const requiredResource of [
    'bilig://workpaper/manifest',
    'bilig://workpaper/agent-handoff',
    'bilig://workpaper/sheets',
    'bilig://workpaper/current-document',
  ]) {
    if (!mcpServerCardResourceUris.has(requiredResource)) {
      throw new Error(`docs/.well-known/mcp/server-card.json is missing ${requiredResource}`)
    }
  }
  const mcpServerCardPrompts = Reflect.get(parsedMcpServerCard, 'prompts')
  if (
    !Array.isArray(mcpServerCardPrompts) ||
    !mcpServerCardPrompts.every(
      (prompt) => typeof prompt === 'object' && prompt !== null && typeof Reflect.get(prompt, 'name') === 'string',
    )
  ) {
    throw new Error('docs/.well-known/mcp/server-card.json must define named prompts')
  }
  const mcpServerCardPromptNames = new Set(mcpServerCardPrompts.map((prompt) => Reflect.get(prompt, 'name')))
  for (const requiredPrompt of ['edit_and_verify_workpaper', 'debug_workpaper_formula']) {
    if (!mcpServerCardPromptNames.has(requiredPrompt)) {
      throw new Error(`docs/.well-known/mcp/server-card.json is missing ${requiredPrompt}`)
    }
  }
  requireIncludes(
    whyAgentsDoc,
    'description: Why coding agents should edit workbook formulas through a Node.js WorkPaper API',
    'docs/why-agents-need-workbook-apis.md',
  )
  for (const required of [
    '## MCP In 30 Seconds',
    'bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable',
    'set_cell_contents',
    'export_workpaper_document',
    '"editedCell": "Inputs!B3"',
    '"restoredMatchesAfter": true',
  ]) {
    requireIncludes(whyAgentsDoc, required, 'docs/why-agents-need-workbook-apis.md')
  }
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
    'https://modelcontextprotocol.io/specification/2025-11-25/server/tools',
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
  requireIncludes(
    agentToolCallingDoc,
    'pnpm --dir examples/headless-workpaper run agent:openai-agents-sdk',
    'docs/agent-workpaper-tool-calling-recipe.md',
  )
  requireIncludes(agentToolCallingDoc, 'openai-agents-sdk-workpaper-tool.md', 'docs/agent-workpaper-tool-calling-recipe.md')
  requireIncludes(agentToolCallingDoc, 'function_call_output', 'docs/agent-workpaper-tool-calling-recipe.md')
  for (const [path, content] of [
    ['README.md', readme],
    ['packages/headless/README.md', headlessReadme],
    ['docs/index.html', index],
    ['docs/llms.txt', llms],
    ['docs/agent-workpaper-tool-calling-recipe.md', agentToolCallingDoc],
    ['examples/headless-workpaper/package.json', headlessExamplePackageJson],
    ['examples/headless-workpaper/README.md', await readFile(join(repoRoot, 'examples', 'headless-workpaper', 'README.md'), 'utf8')],
  ] as const) {
    requireIncludes(content, 'agent:openai-agents-sdk', path)
  }
  for (const required of [
    'title: OpenAI Agents SDK WorkPaper tools',
    'description: Wrap @bilig/headless workbook reads and verified edits as OpenAI Agents SDK function tools.',
    'image: /assets/github-social-preview.png',
    'pnpm --dir examples/headless-workpaper run agent:openai-agents-sdk',
    'examples/headless-workpaper/openai-agents-sdk-tool-smoke.ts',
    'https://openai.github.io/openai-agents-js/guides/tools/',
    'OpenAI Agents SDK Agent -> tool() -> invokeFunctionTool()',
    'restoredMatchesAfter',
  ] as const) {
    requireIncludes(openAiAgentsSdkDoc, required, 'docs/openai-agents-sdk-workpaper-tool.md')
  }
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
    'description: Expose @bilig/headless workbook reads, verified edits, formula contracts, persistence checks, resources, and prompts through MCP.',
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
  requireIncludes(
    mcpWorkPaperToolServerDoc,
    `npm exec --package ${headlessPackageSpec} -- bilig-mcp-challenge`,
    'docs/mcp-workpaper-tool-server.md',
  )
  requireIncludes(
    mcpWorkPaperToolServerDoc,
    `npm exec --package ${headlessPackageSpec} -- bilig-workpaper-mcp`,
    'docs/mcp-workpaper-tool-server.md',
  )
  requireIncludes(
    mcpWorkPaperToolServerDoc,
    `npm exec --package ${headlessPackageSpec} -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable`,
    'docs/mcp-workpaper-tool-server.md',
  )
  requireIncludes(mcpWorkPaperToolServerDoc, '`list_sheets`, `read_range`, `read_cell`', 'docs/mcp-workpaper-tool-server.md')
  requireIncludes(
    mcpWorkPaperToolServerDoc,
    'WorkPaper JSON back to the same file after `set_cell_contents`',
    'docs/mcp-workpaper-tool-server.md',
  )
  requireIncludes(mcpWorkPaperToolServerDoc, 'resources/list', 'docs/mcp-workpaper-tool-server.md')
  requireIncludes(mcpWorkPaperToolServerDoc, 'prompts/list', 'docs/mcp-workpaper-tool-server.md')
  requireIncludes(mcpWorkPaperToolServerDoc, 'bilig://workpaper/agent-handoff', 'docs/mcp-workpaper-tool-server.md')
  requireIncludes(mcpWorkPaperToolServerDoc, 'edit_and_verify_workpaper', 'docs/mcp-workpaper-tool-server.md')
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
    'https://bilig.proompteng.ai/.well-known/mcp/server-card.json',
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
    'https://modelcontextprotocol.io/specification/2025-11-25/server/tools',
    'docs/mcp-workpaper-tool-server.md',
  )
  requireIncludes(mcpWorkPaperToolServerDoc, 'https://github.com/proompteng/bilig/discussions/230', 'docs/mcp-workpaper-tool-server.md')
  requireIncludes(agentToolCallingDoc, 'https://github.com/proompteng/bilig/discussions/335', 'docs/agent-workpaper-tool-calling-recipe.md')
  requireIncludes(mcpWorkPaperToolServerDoc, 'mcp-client-setup.md', 'docs/mcp-workpaper-tool-server.md')
  for (const required of [
    '## Named Public Alternatives',
    'https://github.com/henilcalagiya/google-sheets-mcp',
    'https://github.com/dream-num/univer-mcp',
    'https://github.com/GRID-is/claude-mcp',
    'A file library can preserve formulas without recalculating fresh results in Node',
    'Do not pitch Bilig as "another Google\nSheets MCP server"',
    'A long-running SheetJS issue asks\nwhether a formula value can be refreshed after changing an input cell',
    'ExcelJS discussion describes JSON-driven workbook edits where shared formulas',
  ]) {
    requireIncludes(spreadsheetMcpServerComparison, required, 'docs/spreadsheet-mcp-server-comparison.md')
  }
  for (const required of [
    'description: Live directory and install status for the Bilig WorkPaper MCP server',
    `npm exec --package ${headlessPackageSpec} -- bilig-workpaper-mcp`,
    'io.github.proompteng/bilig-workpaper',
    'https://registry.modelcontextprotocol.io/v0.1/servers?search=io.github.proompteng%2Fbilig-workpaper',
    'https://glama.ai/mcp/servers/proompteng/bilig',
    'https://bilig.proompteng.ai/.well-known/mcp/server-card.json',
    'https://proompteng.github.io/bilig/.well-known/mcp/server-card.json',
    'https://proompteng.github.io/bilig/.well-known/mcp.json',
    'https://proompteng.github.io/bilig/.well-known/mcp-server-card.json',
    'https://smithery.ai/servers/gkonushev/bilig-workpaper',
    'npx -y smithery mcp add gkonushev/bilig-workpaper',
    "npx -y smithery tool call bilig-workpaper list_sheets '{}'",
    'Static MCP server card',
    'https://github.com/chatmcp/mcpso/issues/2295',
    'https://github.com/cline/mcp-marketplace/issues/1557',
    'https://github.com/docker/mcp-registry/pull/3606',
    'https://github.com/aaif-goose/goose/pull/9315',
    'Docker MCP Registry             | Submitted for maintainer review; source commit and readme refreshed on May 19',
    'Goose MCP catalog               | Closed by maintainer while Goose pauses new MCP server additions',
    'mcp.so                          | Submitted for maintainer review; issue body refreshed on May 19',
    'Cline MCP Marketplace           | Submitted for maintainer review; issue body refreshed on May 19',
    'The Docker MCP Registry pull request was refreshed on May 19, 2026',
    '`a1ecdd52cda3d54e0254afce129a9012c5027826`',
    `the PR body\nnow points reviewers at \`@bilig/headless@${headlessPackageVersion}\` and \`libraries-v${headlessPackageVersion}\``,
    'The Goose MCP catalog pull request was closed on May 19, 2026',
    'Do not resubmit there until maintainers reopen that path',
    'The mcp.so and Cline MCP Marketplace submissions were refreshed on May 19, 2026',
    'by editing the existing issue bodies, not by adding more comments',
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
    'Live; `smithery mcp add` smoke connected and listed demo workbook sheets',
    `Live but latest marker lags npm; \`${officialRegistryLatestMarkedVersion}\` is latest-marked while npm latest is \`${headlessPackageVersion}\``,
    'Live with `Try in Browser`; seven tools indexed with A-grade TDQS',
    'Live in PulseMCP-backed lookup as `Bilig WorkPaper`',
    'https://www.pulsemcp.com/servers?search=bilig&q=bilig',
    'https://github.com/proompteng/bilig/issues/384',
    'Live verification on May 19, 2026 found Bilig WorkPaper\nthrough a PulseMCP-backed lookup query',
    'Smithery lists Bilig WorkPaper as `gkonushev/bilig-workpaper`',
    'Live verification on May 19, 2026 returned a connected Smithery server',
    'Glama lists Bilig WorkPaper publicly with TypeScript, Developer Tools',
    'file-backed tools',
    'A-grade Tool Definition Quality',
    "Glama's source crawl, hosted smoke build, and JSON\nAPI can refresh on different cadences",
    `npm latest is \`@bilig/headless@${headlessPackageVersion}\``,
    `official Registry latest-marked entry is\n\`io.github.proompteng/bilig-workpaper@${officialRegistryLatestMarkedVersion}\``,
    `does not yet match npm latest \`@bilig/headless@${headlessPackageVersion}\``,
    'remote `https://bilig.proompteng.ai/mcp`',
    'limit=100',
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
    'description: Remote MCP smoke endpoint plus copy-paste local stdio configuration for Bilig WorkPaper in Claude, Cursor, VS Code, Cline, and Codex.',
    'docs/mcp-client-setup.md',
  )
  for (const required of [
    '## Remote smoke in 30 seconds',
    '## Smithery install',
    'https://smithery.ai/servers/gkonushev/bilig-workpaper',
    'npx -y smithery mcp add gkonushev/bilig-workpaper',
    "npx -y smithery tool call bilig-workpaper list_sheets '{}'",
    'https://bilig.proompteng.ai/mcp',
    'https://bilig.proompteng.ai/.well-known/mcp/server-card.json',
    'mcp-protocol-version: 2025-11-25',
    '## Persistent file-backed stdio server',
    `npm exec --package ${headlessPackageSpec} -- bilig-workpaper-mcp`,
    'The first command is demo mode. The client configs below use file-backed mode',
    `"args": ["exec", "--package", "${headlessPackageSpec}", "--", "bilig-workpaper-mcp", "--workpaper", "./pricing.workpaper.json", "--init-demo-workpaper", "--writable"]`,
    `args = ["exec", "--package", "${headlessPackageSpec}", "--", "bilig-workpaper-mcp", "--workpaper", "./pricing.workpaper.json", "--init-demo-workpaper", "--writable"]`,
    'pnpm mcpb:workpaper:build',
    mcpbReleaseAssetUrl,
    mcpbReleaseChecksumUrl,
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
    'description: Download or reproduce the Claude Desktop MCPB bundle for the published @bilig/headless WorkPaper MCP server',
    mcpbReleaseAssetUrl,
    mcpbReleaseChecksumUrl,
    'Open the downloaded `.mcpb` file with Claude Desktop',
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
    '[Bilig WorkPaper MCPB privacy policy](workpaper-mcpb-privacy.md)',
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
    ['docs/openai-agents-sdk-workpaper-tool.md', openAiAgentsSdkDoc],
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
}

function requireMcpToolDiscoveryContract(tool: unknown): void {
  if (typeof tool !== 'object' || tool === null || Array.isArray(tool)) {
    throw new Error('docs/.well-known/mcp/server-card.json tools must be JSON objects')
  }

  const name = Reflect.get(tool, 'name')
  if (typeof name !== 'string' || name.length === 0) {
    throw new Error('docs/.well-known/mcp/server-card.json tools must have names')
  }
  if (typeof Reflect.get(tool, 'description') !== 'string' || Reflect.get(tool, 'description').length === 0) {
    throw new Error(`docs/.well-known/mcp/server-card.json tool ${name} must have a description`)
  }
  if (typeof Reflect.get(tool, 'title') !== 'string' || Reflect.get(tool, 'title').length === 0) {
    throw new Error(`docs/.well-known/mcp/server-card.json tool ${name} must have a title`)
  }

  const inputSchema = Reflect.get(tool, 'inputSchema')
  if (typeof inputSchema !== 'object' || inputSchema === null || Array.isArray(inputSchema)) {
    throw new Error(`docs/.well-known/mcp/server-card.json tool ${name} must have an inputSchema`)
  }
  requireSchemaPropertyDescriptions(name, 'inputSchema', inputSchema)

  const outputSchema = Reflect.get(tool, 'outputSchema')
  if (typeof outputSchema !== 'object' || outputSchema === null || Array.isArray(outputSchema)) {
    throw new Error(`docs/.well-known/mcp/server-card.json tool ${name} must have an outputSchema`)
  }
  requireSchemaPropertyDescriptions(name, 'outputSchema', outputSchema)

  const annotations = Reflect.get(tool, 'annotations')
  if (typeof annotations !== 'object' || annotations === null || Array.isArray(annotations)) {
    throw new Error(`docs/.well-known/mcp/server-card.json tool ${name} must have annotations`)
  }
  for (const requiredAnnotation of ['title', 'readOnlyHint', 'destructiveHint', 'idempotentHint', 'openWorldHint'] as const) {
    if (!Reflect.has(annotations, requiredAnnotation)) {
      throw new Error(`docs/.well-known/mcp/server-card.json tool ${name} annotations must include ${requiredAnnotation}`)
    }
  }
  if (Reflect.get(annotations, 'title') !== Reflect.get(tool, 'title')) {
    throw new Error(`docs/.well-known/mcp/server-card.json tool ${name} annotation title must match tool title`)
  }
}

function requireSchemaPropertyDescriptions(toolName: string, schemaName: 'inputSchema' | 'outputSchema', schema: object): void {
  const properties = Reflect.get(schema, 'properties')
  if (properties === undefined) {
    return
  }
  if (typeof properties !== 'object' || properties === null || Array.isArray(properties)) {
    throw new Error(`docs/.well-known/mcp/server-card.json tool ${toolName} ${schemaName}.properties must be an object`)
  }

  for (const [propertyName, propertySchema] of Object.entries(properties)) {
    requireSchemaPropertyDescription(toolName, `${schemaName}.${propertyName}`, propertySchema)
  }
}

function requireSchemaPropertyDescription(toolName: string, propertyPath: string, propertySchema: unknown): void {
  if (typeof propertySchema !== 'object' || propertySchema === null || Array.isArray(propertySchema)) {
    throw new Error(`docs/.well-known/mcp/server-card.json tool ${toolName} ${propertyPath} must be a JSON schema object`)
  }
  if (typeof Reflect.get(propertySchema, 'description') !== 'string' || Reflect.get(propertySchema, 'description').length === 0) {
    throw new Error(`docs/.well-known/mcp/server-card.json tool ${toolName} ${propertyPath} must have a description`)
  }

  const nestedProperties = Reflect.get(propertySchema, 'properties')
  if (nestedProperties !== undefined) {
    if (typeof nestedProperties !== 'object' || nestedProperties === null || Array.isArray(nestedProperties)) {
      throw new Error(`docs/.well-known/mcp/server-card.json tool ${toolName} ${propertyPath}.properties must be an object`)
    }
    for (const [nestedName, nestedSchema] of Object.entries(nestedProperties)) {
      requireSchemaPropertyDescription(toolName, `${propertyPath}.${nestedName}`, nestedSchema)
    }
  }

  const items = Reflect.get(propertySchema, 'items')
  if (items !== undefined) {
    if (typeof items !== 'object' || items === null || Array.isArray(items)) {
      throw new Error(`docs/.well-known/mcp/server-card.json tool ${toolName} ${propertyPath}.items must be a JSON schema object`)
    }
    const itemProperties = Reflect.get(items, 'properties')
    if (itemProperties !== undefined) {
      if (typeof itemProperties !== 'object' || itemProperties === null || Array.isArray(itemProperties)) {
        throw new Error(`docs/.well-known/mcp/server-card.json tool ${toolName} ${propertyPath}.items.properties must be an object`)
      }
      for (const [itemPropertyName, itemPropertySchema] of Object.entries(itemProperties)) {
        requireSchemaPropertyDescription(toolName, `${propertyPath}.items.${itemPropertyName}`, itemPropertySchema)
      }
    }
  }
}
