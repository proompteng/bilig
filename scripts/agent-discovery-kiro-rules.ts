import { type AgentIdeRuleInput, commandSafetyStandard, workbookProofStandard } from './agent-discovery-ide-rules.ts'
import { buildFileBackedMcpServerConfig } from './agent-discovery-mcp-configs.ts'

export function buildKiroWorkpaperSteering(input: AgentIdeRuleInput): string {
  const { remoteMcpEndpoint, repositoryUrl, siteRoot, workpaperPackageSpec } = input
  return `# Bilig WorkPaper Steering

Use this Kiro steering file when a task mentions spreadsheet-shaped business
logic: pricing, quote approvals, payout checks, budgets, imports, forecasts,
workbook-file formula diagnostics, or formula readback after changing cells.

Kiro also reads the workspace \`AGENTS.md\`. This steering file is narrower:
it tells Kiro when to prefer Bilig WorkPaper and the project-local MCP server
before trying Excel, LibreOffice, Google Sheets, browser grids, screenshots, or
saved workbook cache values.

## First Checks

\`\`\`sh
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --json
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --scenario provider-backed --json
npm exec --package ${workpaperPackageSpec} -- bilig-agent-challenge --json
npm exec --package ${workpaperPackageSpec} -- bilig-mcp-challenge --json
\`\`\`

For local file-backed MCP work, use the project server in
\`.kiro/settings/mcp.json\`:

\`\`\`sh
npm exec --package ${workpaperPackageSpec} -- bilig-workpaper-mcp --workpaper ./.bilig/pricing.workpaper.json --init-demo-workpaper --writable
\`\`\`

Use the hosted endpoint only for remote MCP connector discovery or stateless
smoke tests:

\`\`\`text
${remoteMcpEndpoint}
\`\`\`

## Required Readback

${workbookProofStandard()}

If any readback step fails, say what failed instead of treating a write call or
tool invocation as proof.

## Command Safety

${commandSafetyStandard()}

## References

- Kiro steering: .kiro/steering/bilig-workpaper.md
- Kiro project MCP config: .kiro/settings/mcp.json
- Docs map: ${siteRoot}/llms.txt
- Agent rule chooser: ${siteRoot}/agent-rule-chooser.html
- Agent handbook: ${siteRoot}/headless-workpaper-agent-handbook.html
- Repository: ${repositoryUrl}
`
}

export function buildKiroMcpConfig(input: AgentIdeRuleInput): string {
  return buildFileBackedMcpServerConfig({
    serverKey: 'bilig-workpaper',
    workpaperPackageSpec: input.workpaperPackageSpec,
    workpaperPath: './.bilig/pricing.workpaper.json',
  })
}
