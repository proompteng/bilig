import { readFile } from 'node:fs/promises'
import { join } from 'node:path'

import { requireIncludes, requireNotIncludes } from './check-docs-discovery-core.ts'

export async function requireAgentRuleChooserDiscovery(input: {
  readonly docsRoot: string
  readonly index: string
  readonly llms: string
  readonly llmsFull: string
  readonly readme: string
}): Promise<void> {
  const { docsRoot, index, llms, llmsFull, readme } = input
  const [agentRuleChooser, agentStart] = await Promise.all([
    readFile(join(docsRoot, 'agent-rule-chooser.md'), 'utf8'),
    readFile(join(docsRoot, 'agent-start.txt'), 'utf8'),
  ])

  for (const required of [
    'title: Coding agent rule chooser for Bilig WorkPaper',
    'description: Pick the Bilig instruction, rule, prompt, or MCP config for Codex',
    'npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json',
    'schemaVersion: "bilig-evaluator.v1"',
    'door: "agent-mcp"',
    '| Codex | `AGENTS.md`',
    '| Claude Code | `CLAUDE.md`, then `.claude/skills/bilig-workpaper/SKILL.md`',
    '| GitHub Copilot | `.github/copilot-instructions.md`',
    '| VS Code agent mode | `.github/copilot-instructions.md`',
    '| Cursor | `.cursor/rules/bilig-workpaper.mdc`',
    '| Kiro | `.kiro/steering/bilig-workpaper.md`',
    '`.kiro/settings/mcp.json` defines the project-local file-backed WorkPaper MCP server',
    '| Roo Code | `.roo/rules/bilig-workpaper.md`',
    '`.roo/mcp.json` defines the project-local file-backed WorkPaper MCP server',
    '| Trae | `.trae/rules/bilig-workpaper.md`',
    '`.trae/mcp.json` defines the project-local file-backed WorkPaper MCP server after Project MCP is enabled',
    '[Trae WorkPaper MCP setup](trae-workpaper-mcp.md)',
    '| Qodo IDE | `AGENTS.md` plus the [Qodo WorkPaper MCP setup](qodo-workpaper-mcp.md)',
    'Paste the `bilig-workpaper` JSON into Qodo Agentic Tools MCP settings',
    '[Qodo WorkPaper MCP setup](qodo-workpaper-mcp.md)',
    '| Zed | `.zed/settings.json`, root `AGENTS.md`, and `.agents/skills/bilig-workpaper/SKILL.md`',
    '`.zed/settings.json` defines the project-local `context_servers.bilig-workpaper` MCP server',
    '| JetBrains Junie | `AGENTS.md` in the repo root',
    '`.junie/mcp/mcp.json` defines the file-backed WorkPaper MCP server',
    '| OpenHands | `AGENTS.md`, then `.agents/skills/bilig-workpaper/SKILL.md`',
    'openhands mcp add bilig-workpaper --transport stdio npm -- exec --yes --package @bilig/workpaper@latest -- bilig-workpaper-mcp --workpaper ./.bilig/pricing.workpaper.json --init-demo-workpaper --writable',
    'https://docs.openhands.dev/openhands/usage/cli/mcp-servers',
    'https://docs.openhands.dev/overview/skills',
    '| OpenCode | `opencode.jsonc`, then `.opencode/agents/bilig-workpaper.md`',
    'https://opencode.ai/docs/config/',
    'https://opencode.ai/docs/mcp-servers/',
    'https://opencode.ai/docs/agents/',
    '[OpenCode WorkPaper MCP setup](opencode-workpaper-mcp.md)',
    '| Aider | `CONVENTIONS.md`, loaded by `.aider.conf.yml`; public page: [Aider WorkPaper conventions](aider-workpaper-conventions.md).',
    '[Aider WorkPaper conventions](aider-workpaper-conventions.md)',
    'https://aider.chat/docs/usage/conventions.html',
    'https://aider.chat/docs/config/aider_conf.html',
    '| Windsurf/Cascade | `.devin/rules/bilig-workpaper.md`',
    '| Cline | `.clinerules/bilig-workpaper.md`',
    '| Continue | `.continue/rules/bilig-workpaper.md`',
    '`.continue/mcpServers/bilig-workpaper.yaml` defines the project-local file-backed WorkPaper MCP server',
    '| Gemini CLI | `gemini-extension.json` plus `gemini-workpaper-context.md`',
    'The published starter overlay is release-pending.',
    'Until a newer release passes a fresh consumer smoke',
    '`.vscode/mcp.json` uses the VS Code `servers` shape.',
    'Kiro reads workspace steering from `.kiro/steering/`',
    'project MCP servers',
    'from `.kiro/settings/mcp.json`',
    'Roo Code reads workspace rules from `.roo/rules/`',
    'Trae reads project rules from `.trae/rules/`',
    'Project MCP servers from',
    '`.trae/mcp.json`; enable Project MCP',
    'Zed reads project context servers from `.zed/settings.json`',
    '`AGENTS.md` and `.agents/skills/bilig-workpaper/SKILL.md` as project context',
    'Junie project MCP config lives at `.junie/mcp/mcp.json`',
    'Aider loads `CONVENTIONS.md` through `.aider.conf.yml`',
    'Cascade/Devin docs currently prefer `.devin/rules`',
    'mirror remains for compatible Windsurf/Cascade installs',
    'https://docs.windsurf.com/windsurf/cascade/memories',
    'https://docs.cline.bot/customization/cline-rules',
    'https://docs.continue.dev/customize/rules',
    'https://docs.continue.dev/customize/deep-dives/mcp',
    'https://kiro.dev/docs/steering/',
    'https://kiro.dev/docs/mcp/configuration/',
    'https://roocodeinc.github.io/Roo-Code/features/custom-instructions',
    'https://roocodeinc.github.io/Roo-Code/features/mcp/using-mcp-in-roo/',
    'https://docs.trae.ai/ide/model-context-protocol',
    'https://docs.trae.ai/ide/add-mcp-servers',
    'https://docs.trae.ai/ide/rules',
    'https://docs.trae.ai/ide/skills',
    'https://docs.qodo.ai/qodo-documentation/qodo-ide/tools-mcps/agentic-tools-mcps',
    'https://docs.qodo.ai/qodo-documentation/qodo-review/configuration/qodo-merge-configuration',
    'https://zed.dev/docs/ai/mcp',
    'https://zed.dev/docs/ai/rules',
    'https://zed.dev/docs/ai/tool-permissions',
    'https://junie.jetbrains.com/docs/junie-plugin-mcp-settings.html',
    'https://junie.jetbrains.com/docs/guidelines-and-memory.html',
    '[Agent WorkPaper evaluator matrix](agent-proof-matrix.md)',
  ] as const) {
    requireIncludes(agentRuleChooser, required, 'docs/agent-rule-chooser.md')
  }

  requireNotIncludes(agentRuleChooser, 'No code changes are required', 'docs/agent-rule-chooser.md')

  for (const [path, content] of [
    ['README.md', readme],
    ['docs/index.html', index],
    ['docs/llms.txt', llms],
  ] as const) {
    requireIncludes(content, 'agent-rule-chooser', path)
  }

  requireIncludes(llmsFull, '## Host Rule Chooser', 'docs/llms-full.txt')
  requireIncludes(llmsFull, 'Source: https://github.com/proompteng/bilig/blob/main/docs/agent-rule-chooser.md', 'docs/llms-full.txt')
  requireIncludes(agentStart, 'Agent rule chooser: https://proompteng.github.io/bilig/agent-rule-chooser.html', 'docs/agent-start.txt')
}
