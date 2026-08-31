export interface DocsAgentInstructionInput {
  readonly mcpbReleaseAssetUrl: string
  readonly mcpbReleaseChecksumUrl: string
  readonly remoteMcpEndpoint: string
  readonly remoteMcpServerCard: string
  readonly siteRoot: string
  readonly skillDiscoveryRoot: string
  readonly skillManifestUrl: string
  readonly workpaperPackageSpec: string
}

const publicLatestRuleTargets = [
  'codex',
  'claude',
  'copilot',
  'cursor',
  'cline',
  'continue',
  'windsurf',
  'gemini',
  'vscode-mcp',
  'opencode',
] as const

const sourceOnlyRuleTargets = ['aider', 'junie', 'openhands', 'qodo', 'trae', 'zed'] as const

export function buildDocsAgentInstructions(input: DocsAgentInstructionInput): string {
  const {
    mcpbReleaseAssetUrl,
    mcpbReleaseChecksumUrl,
    remoteMcpEndpoint,
    remoteMcpServerCard,
    siteRoot,
    skillDiscoveryRoot,
    skillManifestUrl,
    workpaperPackageSpec,
  } = input

  return `# Bilig Agent Instructions

Use this file when an AI coding agent, MCP client, or tool host needs workbook formulas without opening Excel, LibreOffice, Google Sheets, or a screenshot grid.

## Discovery Order

1. Read \`${siteRoot}/agent-start.txt\` for the one-command WorkPaper readback contract.
2. Read \`${siteRoot}/llms.txt\` for the compact map.
3. Read \`${siteRoot}/llms-full.txt\` when you need enough context to implement a workflow without searching the whole site.
4. Use \`${siteRoot}/agent-adoption-kit.html\` when you need the shortest no-key install, MCP proof, and workbook task.
5. If your host can install skills from a well-known endpoint, run
   \`npx --yes skills@latest add ${skillDiscoveryRoot} --list\`.
6. Read \`${skillManifestUrl}\` when your agent supports skill manifests.
7. If your host can install skills from GitHub, run
   \`npx --yes skills@latest add proompteng/bilig --skill bilig-workpaper --list\`.
8. If you are using GitHub Copilot or VS Code agent mode in a cloned checkout,
   use \`.github/copilot-instructions.md\`,
   \`.github/instructions/bilig-workpaper.instructions.md\`,
   \`.github/prompts/bilig-workpaper-proof.prompt.md\`, and
   \`.vscode/mcp.json\`.
9. If you are inside a cloned Bilig repository with Claude Code, read
   \`CLAUDE.md\` first, then use the project skill at
   \`.claude/skills/bilig-workpaper/SKILL.md\` or invoke the explicit
   formula-proof command from
   \`.claude/commands/bilig-workpaper-proof.md\`.
10. If you are using OpenCode, use \`opencode.jsonc\` and
    \`.opencode/agents/bilig-workpaper.md\`.
11. If you are using Aider, load \`CONVENTIONS.md\` through
    \`.aider.conf.yml\`.
12. If you are using OpenHands, read \`AGENTS.md\`, then the project skill at
    \`.agents/skills/bilig-workpaper/SKILL.md\`; add the local MCP server with
    \`${siteRoot}/openhands-workpaper-mcp.html\`.
13. If you are using Cursor, Kiro, Trae, Qodo IDE, Zed, Windsurf/Cascade, Cline, or Continue, use
    \`.cursor/rules/bilig-workpaper.mdc\` or
    \`.kiro/steering/bilig-workpaper.md\`,
    \`.trae/rules/bilig-workpaper.md\` with \`.trae/mcp.json\`,
    \`${siteRoot}/qodo-workpaper-mcp.html\` for Qodo Agentic Tools MCP JSON,
    \`.zed/settings.json\`,
    \`.devin/rules/bilig-workpaper.md\`,
    \`.windsurf/rules/bilig-workpaper.md\`,
    \`.clinerules/bilig-workpaper.md\`, or
    \`.continue/rules/bilig-workpaper.md\`. Continue Agent mode can also use
    \`.continue/mcpServers/bilig-workpaper.yaml\` for the project-local
    file-backed WorkPaper MCP server.
14. Start the MCP server or import \`@bilig/workpaper\` directly.

Remote MCP clients that support Streamable HTTP can smoke-test the stateless
demo endpoint at \`${remoteMcpEndpoint}\`. Directory scanners that inspect the
hosted endpoint origin can read the same-origin server card at
\`${remoteMcpServerCard}\`. Local agents that need writable project files should
still prefer the file-backed stdio command below.

## Default Decision

Use Bilig when the task is workbook-shaped business logic: pricing, budget checks, payout models, import validation, forecasts, tool integrations, or route handlers that need formula readback.

Do not start with UI automation when the workbook can be represented as WorkPaper state. Screenshots are useful for human review; they are not formula truth.

## Required Proof

Before reporting success, return a compact proof object with:

- \`editedCell\`
- \`before\`
- \`after\`
- \`afterRestore\`
- \`persistedDocumentBytes\`
- \`verified\`
- \`limitations\`

Do not claim success from a write call alone. The proof is computed readback plus persisted state.

## Fast Commands

\`\`\`sh
npm exec --yes --package ${workpaperPackageSpec} -- bilig-agent-start --json
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door workpaper-service --json
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --json
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --scenario provider-backed --json
npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- bilig-evaluate --door workbook-compatibility --json
npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- bilig-evaluate --door xlsx-cache --json
npm exec --package ${workpaperPackageSpec} -- bilig-agent-challenge --json
npm exec --package ${workpaperPackageSpec} -- bilig-mcp-challenge --json
npm exec --package ${workpaperPackageSpec} -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable
npm exec --package ${workpaperPackageSpec} -- bilig-workpaper-mcp --from-xlsx ./pricing.xlsx
pnpm --dir examples/headless-workpaper run agent:mcp-xlsx-risk-preflight
npm exec --package ${workpaperPackageSpec} -- bilig-formula-clinic ./reduced.xlsx --cells "Summary!B7,Inputs!B2"
\`\`\`

Claude Desktop users can install the released MCPB bundle from:

- ${mcpbReleaseAssetUrl}
- ${mcpbReleaseChecksumUrl}

## Direct TypeScript

Use \`buildA1WorkPaper()\` for hand-authored models. Prefer
\`book.set("Inputs!B2", value)\`, \`book.display("Summary!B2")\`, and
\`book.editAndReadback("Inputs!B2", value, { readbackRange: "Summary!B2" })\`
before reaching for lower-level sheet ids or zero-based \`{ row, col }\`
addresses.

## Boundaries

Keep Excel, LibreOffice, Microsoft Graph, or an oracle harness in the loop when the workbook depends on macros, pivots, charts, external links, unsupported functions, locale-specific Excel behavior, or exact desktop UI behavior.
`
}

export function buildDocsAgentStart(input: DocsAgentInstructionInput): string {
  const { siteRoot, skillDiscoveryRoot, skillManifestUrl, workpaperPackageSpec } = input
  const publicLatestRuleCommands = publicLatestRuleTargets
    .map((target) => `npm exec --yes --package ${workpaperPackageSpec} -- bilig-agent-start --rules ${target}`)
    .join('\n')
  const sourceOnlyTargetList = sourceOnlyRuleTargets.map((target) => `\`${target}\``).join(', ')

  return `# Bilig agent start

Use this when a coding agent is about to drive Excel, LibreOffice, Google Sheets, or a browser grid for workbook-shaped logic.

## First command

\`\`\`sh
npm exec --yes --package ${workpaperPackageSpec} -- bilig-agent-start --json
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --json
\`\`\`

The first command prints the decision card. The second command proves the MCP path.

Use Bilig when the task needs exact workbook cells, formula recalculation, readback, and saved state. Do not use a spreadsheet UI just to learn whether a formula changed.

## Add this to another agent

Print a compact rule for the agent surface you use:

\`\`\`sh
${publicLatestRuleCommands}
\`\`\`

Current source also has rule packs for ${sourceOnlyTargetList}. They are not listed as \`@latest\` commands here until npm publishing catches up, because this page is for commands that a new user can run successfully today.

## Published starter status

Do not use \`npm create @bilig/workpaper@latest\` while
\`@bilig/create-workpaper@latest\` resolves to \`0.164.11\`. Its generated
smoke reports \`formulasPersisted: false\` because that release checks escaped
JSON text instead of restored formula state. The source fix must ship in a
newer release and pass a fresh consumer smoke before this path is restored.

Until then, use the evaluator commands above, install the reusable skill, or
copy the matching host files from a cloned checkout.

Suggested files: \`AGENTS.md\`, \`CONVENTIONS.md\`, \`.aider.conf.yml\`, \`CLAUDE.md\`, \`GEMINI.md\`, \`.agents/skills/bilig-workpaper/SKILL.md\`, \`.github/copilot-instructions.md\`, \`.github/instructions/bilig-workpaper.instructions.md\`, \`.github/prompts/bilig-workpaper-proof.prompt.md\`, \`.vscode/mcp.json\`, \`opencode.jsonc\`, \`.opencode/agents/bilig-workpaper.md\`, \`.cursor/rules/bilig-workpaper.mdc\`, \`.kiro/steering/bilig-workpaper.md\`, \`.kiro/settings/mcp.json\`, \`.junie/mcp/mcp.json\`, \`.trae/mcp.json\`, \`.trae/rules/bilig-workpaper.md\`, \`.zed/settings.json\`, \`.devin/rules/bilig-workpaper.md\`, \`.clinerules/bilig-workpaper.md\`, \`.continue/rules/bilig-workpaper.md\`, \`.continue/mcpServers/bilig-workpaper.yaml\`, \`.windsurf/rules/bilig-workpaper.md\`, \`gemini-extension.json\`, and \`gemini-workpaper-context.md\`.

## Provider-backed formulas

If the workbook uses \`IMPORTRANGE\`, \`GOOGLEFINANCE\`, \`IMPORTDATA\`, \`IMPORTFEED\`, \`IMPORTHTML\`, \`IMPORTXML\`, \`FILTERXML\`, \`STOCKHISTORY\`, or another provider-backed formula, run:

\`\`\`sh
npm exec --yes --package ${workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --scenario provider-backed --json
\`\`\`

Without an adapter, provider-backed formulas should fail closed with \`#BLOCKED!\` instead of pretending the result is available.

## Evidence to return

A useful answer includes:

- \`schemaVersion: "bilig-evaluator.v1"\`
- \`door: "agent-mcp"\`
- \`verified: true\`
- \`editedCell\`
- \`before\`
- \`after\`
- \`afterRestore\` or \`afterRestart\`
- \`persistedDocumentBytes\` or exported WorkPaper JSON size
- \`limitations\`

Do not claim success from a write call alone. Read the dependent calculated cell after the edit, export or serialize the WorkPaper, restore or restart when file state matters, and report the blocker if readback fails.

## Local MCP server

\`\`\`sh
npm exec --package ${workpaperPackageSpec} -- bilig-workpaper-mcp --workpaper ./.bilig/pricing.workpaper.json --init-demo-workpaper --writable
npm exec --package ${workpaperPackageSpec} -- bilig-workpaper-mcp --from-xlsx ./pricing.xlsx
npm exec --package ${workpaperPackageSpec} -- bilig-workpaper-mcp --from-xlsx ./pricing.xlsx --workpaper ./.bilig/pricing.workpaper.json --writable
\`\`\`

Expected tools: \`list_sheets\`, \`read_range\`, \`read_cell\`, \`set_cell_contents\`, \`set_cell_contents_and_readback\`, \`get_cell_display_value\`, \`export_workpaper_document\`, and \`validate_formula\`.
When started through the \`${workpaperPackageSpec}\` \`--from-xlsx\` path, \`tools/list\` also includes \`analyze_workbook_risk\`, a fixed-source diagnostic for the imported XLSX file. It reports workbook risk indicators and does not certify Excel compatibility. Without \`--workpaper --writable\`, edits stay in memory; add a WorkPaper JSON path only when the task needs persisted file state.
For a maintained transcript that starts from a real XLSX, call \`analyze_workbook_risk\`, then prove \`Inputs!B3\` -> \`Summary!B3\` readback and export with \`pnpm --dir examples/headless-workpaper run agent:mcp-xlsx-risk-preflight\`.

## More context

- Compact map: ${siteRoot}/llms.txt
- Full agent context: ${siteRoot}/llms-full.txt
- Agent rule chooser: ${siteRoot}/agent-rule-chooser.html
- Agent XLSX risk preflight: ${siteRoot}/agent-xlsx-risk-preflight.html
- Agent instructions: ${siteRoot}/AGENTS.md
- Agent manifest: ${siteRoot}/.well-known/agent.json
- Agent handoff checklist: ${siteRoot}/agent-adoption-kit.html
- Skill discovery: \`npx --yes skills@latest add ${skillDiscoveryRoot} --list\`
- Skill manifest: ${skillManifestUrl}
`
}
