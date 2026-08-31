import { readFile } from 'node:fs/promises'
import { join } from 'node:path'

import { requireIncludes, requireMatches, requireNotIncludes, requirePackageKeywords } from './check-docs-discovery-core.ts'

export async function requireAgentInstructionDiscovery(input: {
  readonly repoRoot: string
  readonly docsRoot: string
  readonly headlessPackageVersion: string
}): Promise<void> {
  const { repoRoot, docsRoot, headlessPackageVersion } = input
  const headlessPackageSpec = `@bilig/headless@${headlessPackageVersion}`
  const rawHostedSkillManifestUrl = 'https://bilig.proompteng.ai/.well-known/agent-skills/bilig-workpaper/SKILL.txt'
  const [
    docsAgentNotes,
    docsAgentStart,
    wellKnownAgentStart,
    docsSkill,
    claudeProjectMemory,
    claudeProjectSkillNotes,
    claudeProjectCommandNotes,
    cursorProjectRuleNotes,
    devinProjectRuleNotes,
    windsurfProjectRuleNotes,
    clineProjectRuleNotes,
    continueProjectRuleNotes,
    continueMcpConfig,
    traeMcpConfig,
    traeProjectRuleNotes,
    aiderConventions,
    aiderConfig,
    openCodeAgentNotes,
    copilotInstructions,
    copilotWorkpaperInstructions,
    copilotPrompt,
    claudeCodeMcpConfig,
    cursorMcpConfig,
    kiroMcpConfig,
    kiroSteeringNotes,
    junieMcpConfig,
    vscodeMcpConfig,
    openCodeMcpConfig,
    reusableMcpConfig,
    rootSkillNotes,
    workpaperPackageJson,
    workpaperPackageReadme,
    workpaperPackageAgentNotes,
    workpaperPackageSkillNotes,
    headlessAgentNotes,
    headlessSkillNotes,
  ] = await Promise.all([
    readFile(join(docsRoot, 'AGENTS.md'), 'utf8'),
    readFile(join(docsRoot, 'agent-start.txt'), 'utf8'),
    readFile(join(docsRoot, '.well-known', 'agent-start.txt'), 'utf8'),
    readFile(join(docsRoot, 'skill.md'), 'utf8'),
    readFile(join(repoRoot, 'CLAUDE.md'), 'utf8'),
    readFile(join(repoRoot, '.claude', 'skills', 'bilig-workpaper', 'SKILL.md'), 'utf8'),
    readFile(join(repoRoot, '.claude', 'commands', 'bilig-workpaper-proof.md'), 'utf8'),
    readFile(join(repoRoot, '.cursor', 'rules', 'bilig-workpaper.mdc'), 'utf8'),
    readFile(join(repoRoot, '.devin', 'rules', 'bilig-workpaper.md'), 'utf8'),
    readFile(join(repoRoot, '.windsurf', 'rules', 'bilig-workpaper.md'), 'utf8'),
    readFile(join(repoRoot, '.clinerules', 'bilig-workpaper.md'), 'utf8'),
    readFile(join(repoRoot, '.continue', 'rules', 'bilig-workpaper.md'), 'utf8'),
    readFile(join(repoRoot, '.continue', 'mcpServers', 'bilig-workpaper.yaml'), 'utf8'),
    readFile(join(repoRoot, '.trae', 'mcp.json'), 'utf8'),
    readFile(join(repoRoot, '.trae', 'rules', 'bilig-workpaper.md'), 'utf8'),
    readFile(join(repoRoot, 'CONVENTIONS.md'), 'utf8'),
    readFile(join(repoRoot, '.aider.conf.yml'), 'utf8'),
    readFile(join(repoRoot, '.opencode', 'agents', 'bilig-workpaper.md'), 'utf8'),
    readFile(join(repoRoot, '.github', 'copilot-instructions.md'), 'utf8'),
    readFile(join(repoRoot, '.github', 'instructions', 'bilig-workpaper.instructions.md'), 'utf8'),
    readFile(join(repoRoot, '.github', 'prompts', 'bilig-workpaper-proof.prompt.md'), 'utf8'),
    readFile(join(repoRoot, '.mcp.json'), 'utf8'),
    readFile(join(repoRoot, '.cursor', 'mcp.json'), 'utf8'),
    readFile(join(repoRoot, '.kiro', 'settings', 'mcp.json'), 'utf8'),
    readFile(join(repoRoot, '.kiro', 'steering', 'bilig-workpaper.md'), 'utf8'),
    readFile(join(repoRoot, '.junie', 'mcp', 'mcp.json'), 'utf8'),
    readFile(join(repoRoot, '.vscode', 'mcp.json'), 'utf8'),
    readFile(join(repoRoot, 'opencode.jsonc'), 'utf8'),
    readFile(join(repoRoot, 'mcp', 'bilig-workpaper.mcp.json'), 'utf8'),
    readFile(join(repoRoot, 'skills', 'bilig-workpaper', 'SKILL.md'), 'utf8'),
    readFile(join(repoRoot, 'packages', 'bilig', 'package.json'), 'utf8'),
    readFile(join(repoRoot, 'packages', 'bilig', 'README.md'), 'utf8'),
    readFile(join(repoRoot, 'packages', 'bilig', 'AGENTS.md'), 'utf8'),
    readFile(join(repoRoot, 'packages', 'bilig', 'SKILL.md'), 'utf8'),
    readFile(join(repoRoot, 'packages', 'headless', 'AGENTS.md'), 'utf8'),
    readFile(join(repoRoot, 'packages', 'headless', 'SKILL.md'), 'utf8'),
  ])

  requireIncludes(workpaperPackageJson, '"bilig-agent-challenge": "./bin/bilig-agent-challenge.js"', 'packages/bilig/package.json')
  requireIncludes(workpaperPackageJson, '"bilig-workpaper-mcp": "./bin/bilig-workpaper-mcp.js"', 'packages/bilig/package.json')
  requireIncludes(workpaperPackageJson, '"AGENTS.md"', 'packages/bilig/package.json')
  requireIncludes(workpaperPackageJson, '"SKILL.md"', 'packages/bilig/package.json')
  requirePackageKeywords(
    workpaperPackageJson,
    [
      'bilig-workpaper',
      'formula-engine',
      'formula-recalculation',
      'formula-workbook',
      'mcp',
      'mcp-server',
      'mcp-tools',
      'model-context-protocol',
      'node-services',
      'server-side-formula-engine',
      'server-side-formulas',
      'tool-integration',
      'workbook-api',
      'workbook-formulas',
      'workbook-runtime',
      'workpaper',
      'workpaper-json',
      'workpaper-runtime',
    ],
    'packages/bilig/package.json',
  )
  for (const staleKeyword of [
    'excel-formulas',
    'headless-spreadsheet',
    'server-side-spreadsheet',
    'spreadsheet-automation',
    'xlsx',
  ] as const) {
    requireNotIncludes(workpaperPackageJson, `"${staleKeyword}"`, 'packages/bilig/package.json')
  }
  requireIncludes(workpaperPackageReadme, 'The npm tarball includes `AGENTS.md`, `SKILL.md`', 'packages/bilig/README.md')
  requireIncludes(workpaperPackageReadme, 'verified formula readback', 'packages/bilig/README.md')
  requireIncludes(workpaperPackageReadme, 'npm exec --package bilig-workpaper -- bilig-agent-challenge', 'packages/bilig/README.md')
  requireIncludes(workpaperPackageAgentNotes, 'node_modules/bilig-workpaper', 'packages/bilig/AGENTS.md')
  requireIncludes(workpaperPackageAgentNotes, 'npm exec --package bilig-workpaper@', 'packages/bilig/AGENTS.md')
  requireIncludes(workpaperPackageSkillNotes, 'Use bilig-workpaper WorkPaper state', 'packages/bilig/SKILL.md')
  requireIncludes(workpaperPackageSkillNotes, '## First Check: Agent Evaluator', 'packages/bilig/SKILL.md')
  requireIncludes(workpaperPackageSkillNotes, '"bilig-evaluate", "--door", "agent-mcp", "--json"', 'packages/bilig/SKILL.md')
  requireIncludes(workpaperPackageSkillNotes, 'reduced formula/import bugs that need a local report', 'packages/bilig/SKILL.md')
  requireIncludes(workpaperPackageSkillNotes, '## First Choice: MCP', 'packages/bilig/SKILL.md')
  requireIncludes(workpaperPackageSkillNotes, '"--package", "bilig-workpaper@', 'packages/bilig/SKILL.md')

  requireIncludes(headlessAgentNotes, '## Handoff prompt', 'packages/headless/AGENTS.md')
  requireIncludes(headlessAgentNotes, 'Do not claim success from a write call alone.', 'packages/headless/AGENTS.md')
  requireIncludes(
    headlessAgentNotes,
    `npm exec --package ${headlessPackageSpec} -- bilig-mcp-challenge --json`,
    'packages/headless/AGENTS.md',
  )
  requireIncludes(
    headlessAgentNotes,
    `npm exec --package ${headlessPackageSpec} -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable`,
    'packages/headless/AGENTS.md',
  )
  requireIncludes(headlessSkillNotes, 'name: bilig-workpaper', 'packages/headless/SKILL.md')
  requireIncludes(headlessSkillNotes, '"bilig-formula-clinic"', 'packages/headless/SKILL.md')
  requireIncludes(headlessSkillNotes, '"./reduced.xlsx"', 'packages/headless/SKILL.md')
  requireIncludes(headlessSkillNotes, 'Do not trigger it for manual spreadsheet editing', 'packages/headless/SKILL.md')
  requireIncludes(headlessSkillNotes, '## Command Safety', 'packages/headless/SKILL.md')
  requireIncludes(headlessSkillNotes, 'argument array, not a shell-concatenated string', 'packages/headless/SKILL.md')
  requireNotIncludes(headlessSkillNotes, 'allowed-tools:', 'packages/headless/SKILL.md')
  requireNotIncludes(headlessSkillNotes, 'argument-hint:', 'packages/headless/SKILL.md')

  requireIncludes(docsAgentNotes, '## Discovery Order', 'docs/AGENTS.md')
  requireNotIncludes(docsAgentNotes, '13. Start the MCP server or import `@bilig/workpaper` directly.', 'docs/AGENTS.md')
  requireIncludes(docsAgentNotes, '14. Start the MCP server or import `@bilig/workpaper` directly.', 'docs/AGENTS.md')
  requireIncludes(docsAgentNotes, 'Do not claim success from a write call alone.', 'docs/AGENTS.md')
  requireIncludes(docsAgentNotes, 'read\n   `CLAUDE.md` first', 'docs/AGENTS.md')
  requireIncludes(docsAgentNotes, '.claude/skills/bilig-workpaper/SKILL.md', 'docs/AGENTS.md')
  requireIncludes(docsAgentNotes, '.claude/commands/bilig-workpaper-proof.md', 'docs/AGENTS.md')
  requireIncludes(docsAgentNotes, '.cursor/rules/bilig-workpaper.mdc', 'docs/AGENTS.md')
  requireIncludes(docsAgentNotes, '.devin/rules/bilig-workpaper.md', 'docs/AGENTS.md')
  requireIncludes(docsAgentNotes, '.windsurf/rules/bilig-workpaper.md', 'docs/AGENTS.md')
  requireIncludes(docsAgentNotes, '.clinerules/bilig-workpaper.md', 'docs/AGENTS.md')
  requireIncludes(docsAgentNotes, '.continue/rules/bilig-workpaper.md', 'docs/AGENTS.md')
  requireIncludes(docsAgentNotes, '.continue/mcpServers/bilig-workpaper.yaml', 'docs/AGENTS.md')
  requireIncludes(docsAgentNotes, '.trae/rules/bilig-workpaper.md', 'docs/AGENTS.md')
  requireIncludes(docsAgentNotes, '.trae/mcp.json', 'docs/AGENTS.md')
  requireIncludes(docsAgentNotes, 'CONVENTIONS.md', 'docs/AGENTS.md')
  requireIncludes(docsAgentNotes, '.aider.conf.yml', 'docs/AGENTS.md')
  requireIncludes(docsAgentNotes, '.agents/skills/bilig-workpaper/SKILL.md', 'docs/AGENTS.md')
  requireIncludes(docsAgentNotes, 'openhands-workpaper-mcp.html', 'docs/AGENTS.md')
  requireNotIncludes(docsAgentNotes, 'examples/goose-workpaper-mcp/recipe.yaml', 'docs/AGENTS.md')
  requireNotIncludes(docsAgentNotes, 'goose-workpaper-mcp.html', 'docs/AGENTS.md')
  requireIncludes(docsAgentNotes, 'opencode.jsonc', 'docs/AGENTS.md')
  requireIncludes(docsAgentNotes, '.opencode/agents/bilig-workpaper.md', 'docs/AGENTS.md')
  requireIncludes(docsAgentNotes, '.github/copilot-instructions.md', 'docs/AGENTS.md')
  requireIncludes(docsAgentNotes, '.github/instructions/bilig-workpaper.instructions.md', 'docs/AGENTS.md')
  requireIncludes(docsAgentNotes, '.github/prompts/bilig-workpaper-proof.prompt.md', 'docs/AGENTS.md')
  requireIncludes(docsAgentNotes, '.vscode/mcp.json', 'docs/AGENTS.md')
  requireIncludes(docsAgentNotes, rawHostedSkillManifestUrl, 'docs/AGENTS.md')
  requireNotIncludes(docsAgentNotes, 'https://proompteng.github.io/bilig/skill.txt', 'docs/AGENTS.md')
  requireIncludes(
    docsAgentNotes,
    'npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- bilig-evaluate --door workbook-compatibility --json',
    'docs/AGENTS.md',
  )
  requireIncludes(docsAgentNotes, 'npm exec --package @bilig/workpaper@latest -- bilig-agent-challenge --json', 'docs/AGENTS.md')
  requireIncludes(docsAgentNotes, 'npm exec --package @bilig/workpaper@latest -- bilig-mcp-challenge --json', 'docs/AGENTS.md')
  requireIncludes(aiderConventions, '# Aider Bilig WorkPaper Conventions', 'CONVENTIONS.md')
  requireIncludes(aiderConventions, 'bilig-evaluate --door agent-mcp --json', 'CONVENTIONS.md')
  requireIncludes(aiderConventions, 'bilig-workpaper-mcp --from-xlsx ./pricing.xlsx', 'CONVENTIONS.md')
  requireIncludes(aiderConventions, 'Do not claim success from a write call alone.', 'CONVENTIONS.md')
  requireIncludes(aiderConfig, 'read:', '.aider.conf.yml')
  requireIncludes(aiderConfig, '- CONVENTIONS.md', '.aider.conf.yml')
  if (docsAgentStart !== wellKnownAgentStart) {
    throw new Error('docs/agent-start.txt must match docs/.well-known/agent-start.txt')
  }
  requireIncludes(docsAgentStart, '# Bilig agent start', 'docs/agent-start.txt')
  requireIncludes(docsAgentStart, 'npm exec --yes --package @bilig/workpaper@latest -- bilig-agent-start --json', 'docs/agent-start.txt')
  requireIncludes(
    docsAgentStart,
    'npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json',
    'docs/agent-start.txt',
  )
  requireIncludes(docsAgentStart, 'The first command prints the decision card.', 'docs/agent-start.txt')
  requireIncludes(
    docsAgentStart,
    'npm exec --yes --package @bilig/workpaper@latest -- bilig-agent-start --rules codex',
    'docs/agent-start.txt',
  )
  requireIncludes(
    docsAgentStart,
    'npm exec --yes --package @bilig/workpaper@latest -- bilig-agent-start --rules cursor',
    'docs/agent-start.txt',
  )
  requireIncludes(
    docsAgentStart,
    'npm exec --yes --package @bilig/workpaper@latest -- bilig-agent-start --rules cline',
    'docs/agent-start.txt',
  )
  requireIncludes(
    docsAgentStart,
    'npm exec --yes --package @bilig/workpaper@latest -- bilig-agent-start --rules continue',
    'docs/agent-start.txt',
  )
  requireIncludes(
    docsAgentStart,
    'npm exec --yes --package @bilig/workpaper@latest -- bilig-agent-start --rules windsurf',
    'docs/agent-start.txt',
  )
  requireIncludes(
    docsAgentStart,
    'npm exec --yes --package @bilig/workpaper@latest -- bilig-agent-start --rules gemini',
    'docs/agent-start.txt',
  )
  requireIncludes(
    docsAgentStart,
    'npm exec --yes --package @bilig/workpaper@latest -- bilig-agent-start --rules vscode-mcp',
    'docs/agent-start.txt',
  )
  requireIncludes(
    docsAgentStart,
    'npm exec --yes --package @bilig/workpaper@latest -- bilig-agent-start --rules opencode',
    'docs/agent-start.txt',
  )
  for (const releasePendingTarget of ['aider', 'junie', 'openhands', 'qodo', 'trae', 'zed']) {
    requireNotIncludes(
      docsAgentStart,
      `npm exec --yes --package @bilig/workpaper@latest -- bilig-agent-start --rules ${releasePendingTarget}`,
      'docs/agent-start.txt',
    )
  }
  requireIncludes(
    docsAgentStart,
    'Current source also has rule packs for `aider`, `junie`, `openhands`, `qodo`, `trae`, `zed`.',
    'docs/agent-start.txt',
  )
  requireIncludes(docsAgentStart, 'not listed as `@latest` commands here until npm publishing catches up', 'docs/agent-start.txt')
  requireIncludes(docsAgentStart, '.github/copilot-instructions.md', 'docs/agent-start.txt')
  requireIncludes(docsAgentStart, '.github/instructions/bilig-workpaper.instructions.md', 'docs/agent-start.txt')
  requireIncludes(docsAgentStart, '.github/prompts/bilig-workpaper-proof.prompt.md', 'docs/agent-start.txt')
  requireIncludes(docsAgentStart, 'CLAUDE.md', 'docs/agent-start.txt')
  requireIncludes(docsAgentStart, 'CONVENTIONS.md', 'docs/agent-start.txt')
  requireIncludes(docsAgentStart, '.aider.conf.yml', 'docs/agent-start.txt')
  requireIncludes(docsAgentStart, '.agents/skills/bilig-workpaper/SKILL.md', 'docs/agent-start.txt')
  requireNotIncludes(docsAgentStart, 'examples/goose-workpaper-mcp/recipe.yaml', 'docs/agent-start.txt')
  requireIncludes(docsAgentStart, '.cursor/rules/bilig-workpaper.mdc', 'docs/agent-start.txt')
  requireIncludes(docsAgentStart, '.devin/rules/bilig-workpaper.md', 'docs/agent-start.txt')
  requireIncludes(docsAgentStart, '.clinerules/bilig-workpaper.md', 'docs/agent-start.txt')
  requireIncludes(docsAgentStart, '.continue/rules/bilig-workpaper.md', 'docs/agent-start.txt')
  requireIncludes(docsAgentStart, '.continue/mcpServers/bilig-workpaper.yaml', 'docs/agent-start.txt')
  requireIncludes(docsAgentStart, '.windsurf/rules/bilig-workpaper.md', 'docs/agent-start.txt')
  requireIncludes(docsAgentStart, 'opencode.jsonc', 'docs/agent-start.txt')
  requireIncludes(docsAgentStart, '.opencode/agents/bilig-workpaper.md', 'docs/agent-start.txt')
  requireIncludes(docsAgentStart, 'GEMINI.md', 'docs/agent-start.txt')
  requireIncludes(docsAgentStart, 'gemini-extension.json', 'docs/agent-start.txt')
  requireIncludes(docsAgentStart, 'gemini-workpaper-context.md', 'docs/agent-start.txt')
  requireIncludes(docsAgentStart, '.vscode/mcp.json', 'docs/agent-start.txt')
  requireIncludes(docsAgentStart, '.trae/mcp.json', 'docs/agent-start.txt')
  requireIncludes(docsAgentStart, '.trae/rules/bilig-workpaper.md', 'docs/agent-start.txt')
  requireIncludes(docsAgentStart, 'The source fix must ship in a newer release', 'docs/agent-start.txt')
  requireNotIncludes(docsAgentStart, 'npm create @bilig/workpaper@latest . -- --add-agent', 'docs/agent-start.txt')
  requireIncludes(
    docsAgentStart,
    'npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --scenario provider-backed --json',
    'docs/agent-start.txt',
  )
  requireIncludes(docsAgentStart, 'GOOGLEFINANCE', 'docs/agent-start.txt')
  requireIncludes(docsAgentStart, 'IMPORTXML', 'docs/agent-start.txt')
  requireIncludes(docsAgentStart, '#BLOCKED!', 'docs/agent-start.txt')
  requireIncludes(docsAgentStart, 'schemaVersion: "bilig-evaluator.v1"', 'docs/agent-start.txt')
  requireIncludes(docsAgentStart, 'verified: true', 'docs/agent-start.txt')
  requireIncludes(docsAgentStart, 'afterRestore', 'docs/agent-start.txt')
  requireIncludes(docsAgentStart, 'afterRestart', 'docs/agent-start.txt')
  requireIncludes(docsAgentStart, 'persistedDocumentBytes', 'docs/agent-start.txt')
  requireIncludes(docsAgentStart, 'Do not claim success from a write call alone.', 'docs/agent-start.txt')
  requireIncludes(
    docsAgentStart,
    'npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp --workpaper ./.bilig/pricing.workpaper.json --init-demo-workpaper --writable',
    'docs/agent-start.txt',
  )
  requireIncludes(docsAgentStart, 'set_cell_contents_and_readback', 'docs/agent-start.txt')
  requireIncludes(docsAgentStart, 'https://proompteng.github.io/bilig/llms.txt', 'docs/agent-start.txt')
  requireIncludes(docsAgentStart, 'https://proompteng.github.io/bilig/agent-rule-chooser.html', 'docs/agent-start.txt')
  requireIncludes(docsAgentStart, 'https://proompteng.github.io/bilig/.well-known/agent.json', 'docs/agent-start.txt')
  requireIncludes(docsSkill, 'name: bilig-workpaper', 'docs/skill.md')
  requireIncludes(docsSkill, '## Required Verification', 'docs/skill.md')
  requireIncludes(docsSkill, '## Command Safety', 'docs/skill.md')
  requireIncludes(docsSkill, '## First Check: Agent Evaluator', 'docs/skill.md')
  requireIncludes(docsSkill, '"bilig-evaluate", "--door", "agent-mcp", "--json"', 'docs/skill.md')
  requireIncludes(docsSkill, 'npx --yes skills@latest add https://bilig.proompteng.ai --list', 'docs/skill.md')
  requireIncludes(docsSkill, 'npx --yes skills@latest add proompteng/bilig --skill bilig-workpaper --list', 'docs/skill.md')
  requireIncludes(docsSkill, 'returned `tools` array as the source', 'docs/skill.md')
  requireIncludes(docsSkill, 'set_cell_contents_and_readback', 'docs/skill.md')
  requireIncludes(docsSkill, 'currently published package', 'docs/skill.md')
  requireIncludes(docsSkill, '"bilig-agent-challenge", "--json"', 'docs/skill.md')
  requireIncludes(docsSkill, '"bilig-mcp-challenge", "--json"', 'docs/skill.md')
  requireNotIncludes(docsSkill, 'allowed-tools:', 'docs/skill.md')
  requireNotIncludes(docsSkill, 'argument-hint:', 'docs/skill.md')

  requireIncludes(claudeProjectMemory, '# Claude Code Project Instructions', 'CLAUDE.md')
  requireIncludes(claudeProjectMemory, 'Read `AGENTS.md` first', 'CLAUDE.md')
  requireIncludes(claudeProjectMemory, 'npm exec --yes --package @bilig/workpaper@latest -- bilig-agent-start --json', 'CLAUDE.md')
  requireIncludes(
    claudeProjectMemory,
    'npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json',
    'CLAUDE.md',
  )
  requireIncludes(claudeProjectMemory, '.mcp.json', 'CLAUDE.md')
  requireIncludes(claudeProjectMemory, '.claude/skills/bilig-workpaper/SKILL.md', 'CLAUDE.md')
  requireIncludes(claudeProjectMemory, '.claude/commands/bilig-workpaper-proof.md', 'CLAUDE.md')
  requireIncludes(claudeProjectMemory, 'Do not claim success from a write call alone.', 'CLAUDE.md')
  requireIncludes(claudeProjectMemory, 'https://proompteng.github.io/bilig/agent-rule-chooser.html', 'CLAUDE.md')

  requireIncludes(claudeProjectSkillNotes, 'name: bilig-workpaper', '.claude/skills/bilig-workpaper/SKILL.md')
  requireIncludes(claudeProjectSkillNotes, '## Command Safety', '.claude/skills/bilig-workpaper/SKILL.md')
  requireIncludes(claudeProjectSkillNotes, '"bilig-evaluate", "--door", "agent-mcp", "--json"', '.claude/skills/bilig-workpaper/SKILL.md')
  requireIncludes(claudeProjectSkillNotes, '"bilig-mcp-challenge", "--json"', '.claude/skills/bilig-workpaper/SKILL.md')
  requireIncludes(claudeProjectSkillNotes, 'Return readback, not a write-only claim.', '.claude/skills/bilig-workpaper/SKILL.md')
  requireNotIncludes(claudeProjectSkillNotes, 'allowed-tools:', '.claude/skills/bilig-workpaper/SKILL.md')
  requireNotIncludes(claudeProjectSkillNotes, 'argument-hint:', '.claude/skills/bilig-workpaper/SKILL.md')
  requireIncludes(
    claudeProjectCommandNotes,
    'description: Verify workbook formula edits with Bilig WorkPaper',
    '.claude/commands/bilig-workpaper-proof.md',
  )
  requireIncludes(claudeProjectCommandNotes, '$ARGUMENTS', '.claude/commands/bilig-workpaper-proof.md')
  requireIncludes(
    claudeProjectCommandNotes,
    'npm exec --package @bilig/workpaper@latest -- bilig-mcp-challenge --json',
    '.claude/commands/bilig-workpaper-proof.md',
  )
  requireIncludes(
    claudeProjectCommandNotes,
    'npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp --from-xlsx ./pricing.xlsx',
    '.claude/commands/bilig-workpaper-proof.md',
  )
  requireIncludes(
    claudeProjectCommandNotes,
    'npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp --from-xlsx ./pricing.xlsx --workpaper ./.bilig/pricing.workpaper.json --writable',
    '.claude/commands/bilig-workpaper-proof.md',
  )
  requireIncludes(claudeProjectCommandNotes, 'Do not drive', '.claude/commands/bilig-workpaper-proof.md')
  requireIncludes(claudeProjectCommandNotes, 'do not claim success from a write call alone.', '.claude/commands/bilig-workpaper-proof.md')

  requireIncludes(cursorProjectRuleNotes, 'description: Use Bilig WorkPaper', '.cursor/rules/bilig-workpaper.mdc')
  requireIncludes(cursorProjectRuleNotes, 'alwaysApply: false', '.cursor/rules/bilig-workpaper.mdc')
  requireIncludes(
    cursorProjectRuleNotes,
    'npm exec --package @bilig/workpaper@latest -- bilig-agent-challenge --json',
    '.cursor/rules/bilig-workpaper.mdc',
  )
  requireIncludes(cursorProjectRuleNotes, 'Do not claim success from a write call alone.', '.cursor/rules/bilig-workpaper.mdc')
  requireIncludes(cursorProjectRuleNotes, 'https://proompteng.github.io/bilig/llms.txt', '.cursor/rules/bilig-workpaper.mdc')

  if (devinProjectRuleNotes !== windsurfProjectRuleNotes) {
    throw new Error('.devin/rules/bilig-workpaper.md must match .windsurf/rules/bilig-workpaper.md')
  }
  requireIncludes(devinProjectRuleNotes, 'trigger: model_decision', '.devin/rules/bilig-workpaper.md')
  requireIncludes(devinProjectRuleNotes, 'Windsurf/Cascade agent needs spreadsheet-shaped business', '.devin/rules/bilig-workpaper.md')
  requireIncludes(
    devinProjectRuleNotes,
    'npm exec --package @bilig/workpaper@latest -- bilig-mcp-challenge --json',
    '.devin/rules/bilig-workpaper.md',
  )
  requireIncludes(devinProjectRuleNotes, 'If any readback step fails', '.devin/rules/bilig-workpaper.md')
  requireIncludes(devinProjectRuleNotes, 'https://proompteng.github.io/bilig/llms.txt', '.devin/rules/bilig-workpaper.md')

  requireIncludes(windsurfProjectRuleNotes, 'trigger: model_decision', '.windsurf/rules/bilig-workpaper.md')
  requireIncludes(
    windsurfProjectRuleNotes,
    'Windsurf/Cascade agent needs spreadsheet-shaped business',
    '.windsurf/rules/bilig-workpaper.md',
  )
  requireIncludes(
    windsurfProjectRuleNotes,
    'npm exec --package @bilig/workpaper@latest -- bilig-mcp-challenge --json',
    '.windsurf/rules/bilig-workpaper.md',
  )
  requireIncludes(windsurfProjectRuleNotes, 'If any readback step fails', '.windsurf/rules/bilig-workpaper.md')
  requireIncludes(windsurfProjectRuleNotes, 'https://proompteng.github.io/bilig/llms.txt', '.windsurf/rules/bilig-workpaper.md')

  requireIncludes(clineProjectRuleNotes, 'Cline can read this workspace rule', '.clinerules/bilig-workpaper.md')
  requireIncludes(
    clineProjectRuleNotes,
    'npm exec --package @bilig/workpaper@latest -- bilig-agent-challenge --json',
    '.clinerules/bilig-workpaper.md',
  )
  requireIncludes(clineProjectRuleNotes, 'Do not claim success from a write call alone.', '.clinerules/bilig-workpaper.md')
  requireIncludes(
    clineProjectRuleNotes,
    'https://proompteng.github.io/bilig/mcp-workpaper-tool-server.html',
    '.clinerules/bilig-workpaper.md',
  )

  requireIncludes(continueProjectRuleNotes, 'name: Bilig WorkPaper Formula Check', '.continue/rules/bilig-workpaper.md')
  requireIncludes(continueProjectRuleNotes, '.continue/mcpServers/bilig-workpaper.yaml', '.continue/rules/bilig-workpaper.md')
  requireIncludes(
    continueProjectRuleNotes,
    'npm exec --package @bilig/workpaper@latest -- bilig-mcp-challenge --json',
    '.continue/rules/bilig-workpaper.md',
  )
  requireIncludes(continueProjectRuleNotes, 'Do not claim success from a write call alone.', '.continue/rules/bilig-workpaper.md')
  requireIncludes(continueProjectRuleNotes, 'https://proompteng.github.io/bilig/llms-full.txt', '.continue/rules/bilig-workpaper.md')
  for (const required of [
    'name: Bilig WorkPaper MCP',
    'version: 0.0.1',
    'schema: v1',
    'mcpServers:',
    '- name: Bilig WorkPaper File',
    'type: stdio',
    'command: npm',
  ] as const) {
    requireIncludes(continueMcpConfig, required, '.continue/mcpServers/bilig-workpaper.yaml')
  }
  for (const requiredArg of [
    'exec',
    '--yes',
    '--package',
    '@bilig/workpaper@latest',
    '--',
    'bilig-workpaper-mcp',
    '--workpaper',
    './.bilig/pricing.workpaper.json',
    '--init-demo-workpaper',
    '--writable',
  ] as const) {
    requireYamlListValue(continueMcpConfig, requiredArg, '.continue/mcpServers/bilig-workpaper.yaml')
  }
  for (const required of [
    '"mcpServers"',
    '"bilig-workpaper"',
    '"command": "npm"',
    '"--yes"',
    '"--package"',
    '"@bilig/workpaper@latest"',
    '"bilig-workpaper-mcp"',
    '"./.bilig/pricing.workpaper.json"',
    '"--init-demo-workpaper"',
    '"--writable"',
  ] as const) {
    requireIncludes(traeMcpConfig, required, '.trae/mcp.json')
  }
  for (const required of [
    'Trae project rule',
    'Project MCP is enabled',
    '.trae/mcp.json',
    'bilig-evaluate --door agent-mcp --json',
    'Do not claim success from a write call alone.',
    'trae-workpaper-mcp.html',
  ] as const) {
    requireIncludes(traeProjectRuleNotes, required, '.trae/rules/bilig-workpaper.md')
  }

  requireIncludes(openCodeAgentNotes, 'mode: subagent', '.opencode/agents/bilig-workpaper.md')
  requireIncludes(openCodeAgentNotes, 'bilig-workpaper_*', '.opencode/agents/bilig-workpaper.md')
  requireIncludes(
    openCodeAgentNotes,
    'npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json',
    '.opencode/agents/bilig-workpaper.md',
  )
  requireIncludes(openCodeAgentNotes, 'set_cell_contents_and_readback', '.opencode/agents/bilig-workpaper.md')
  requireIncludes(openCodeAgentNotes, 'Do not claim success from a write call alone.', '.opencode/agents/bilig-workpaper.md')
  requireIncludes(
    openCodeAgentNotes,
    'https://proompteng.github.io/bilig/opencode-workpaper-mcp.html',
    '.opencode/agents/bilig-workpaper.md',
  )

  requireIncludes(copilotInstructions, '## Copilot Agent WorkPaper Path', '.github/copilot-instructions.md')
  requireIncludes(copilotInstructions, '.github/instructions/bilig-workpaper.instructions.md', '.github/copilot-instructions.md')
  requireIncludes(copilotInstructions, '.github/prompts/bilig-workpaper-proof.prompt.md', '.github/copilot-instructions.md')
  requireIncludes(copilotInstructions, '.vscode/mcp.json', '.github/copilot-instructions.md')
  requireIncludes(
    copilotInstructions,
    'npm exec --package @bilig/workpaper@latest -- bilig-mcp-challenge --json',
    '.github/copilot-instructions.md',
  )
  requireIncludes(
    copilotInstructions,
    'npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp --from-xlsx ./pricing.xlsx',
    '.github/copilot-instructions.md',
  )
  requireIncludes(
    copilotInstructions,
    'npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp --from-xlsx ./pricing.xlsx --workpaper ./.bilig/pricing.workpaper.json --writable',
    '.github/copilot-instructions.md',
  )
  requireIncludes(copilotInstructions, 'Do not claim success from a write call alone.', '.github/copilot-instructions.md')
  requireIncludes(copilotWorkpaperInstructions, "applyTo: '**/*'", '.github/instructions/bilig-workpaper.instructions.md')
  requireIncludes(copilotWorkpaperInstructions, '# Bilig WorkPaper Formula Proof', '.github/instructions/bilig-workpaper.instructions.md')
  requireIncludes(
    copilotWorkpaperInstructions,
    'npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json',
    '.github/instructions/bilig-workpaper.instructions.md',
  )
  requireIncludes(
    copilotWorkpaperInstructions,
    'npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --scenario provider-backed --json',
    '.github/instructions/bilig-workpaper.instructions.md',
  )
  requireIncludes(
    copilotWorkpaperInstructions,
    'npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp --from-xlsx ./pricing.xlsx',
    '.github/instructions/bilig-workpaper.instructions.md',
  )
  requireIncludes(
    copilotWorkpaperInstructions,
    'npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp --from-xlsx ./pricing.xlsx --workpaper ./.bilig/pricing.workpaper.json --writable',
    '.github/instructions/bilig-workpaper.instructions.md',
  )
  requireIncludes(
    copilotWorkpaperInstructions,
    '.github/prompts/bilig-workpaper-proof.prompt.md',
    '.github/instructions/bilig-workpaper.instructions.md',
  )
  requireIncludes(copilotWorkpaperInstructions, '.vscode/mcp.json', '.github/instructions/bilig-workpaper.instructions.md')
  requireIncludes(
    copilotWorkpaperInstructions,
    'Do not claim success from a write call alone.',
    '.github/instructions/bilig-workpaper.instructions.md',
  )
  requireIncludes(copilotPrompt, 'name: bilig-workpaper-proof', '.github/prompts/bilig-workpaper-proof.prompt.md')
  requireIncludes(copilotPrompt, 'Task: ${input:task:', '.github/prompts/bilig-workpaper-proof.prompt.md')
  requireIncludes(
    copilotPrompt,
    'npm exec --package @bilig/workpaper@latest -- bilig-agent-challenge --json',
    '.github/prompts/bilig-workpaper-proof.prompt.md',
  )
  requireIncludes(
    copilotPrompt,
    'npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp --from-xlsx ./pricing.xlsx',
    '.github/prompts/bilig-workpaper-proof.prompt.md',
  )
  requireIncludes(copilotPrompt, '`biligWorkpaperFile`', '.github/prompts/bilig-workpaper-proof.prompt.md')
  requireIncludes(claudeCodeMcpConfig, '"bilig-workpaper"', '.mcp.json')
  requireIncludes(claudeCodeMcpConfig, '"command": "npm"', '.mcp.json')
  requireIncludes(claudeCodeMcpConfig, '"--package"', '.mcp.json')
  requireIncludes(claudeCodeMcpConfig, '"@bilig/workpaper@latest"', '.mcp.json')
  requireIncludes(claudeCodeMcpConfig, '"--workpaper"', '.mcp.json')
  requireIncludes(claudeCodeMcpConfig, '"./.bilig/pricing.workpaper.json"', '.mcp.json')
  requireIncludes(claudeCodeMcpConfig, '"--init-demo-workpaper"', '.mcp.json')
  requireIncludes(claudeCodeMcpConfig, '"--writable"', '.mcp.json')
  requireIncludes(cursorMcpConfig, '"biligWorkpaperFile"', '.cursor/mcp.json')
  requireIncludes(cursorMcpConfig, '"command": "npm"', '.cursor/mcp.json')
  requireIncludes(cursorMcpConfig, '"@bilig/workpaper@latest"', '.cursor/mcp.json')
  requireIncludes(cursorMcpConfig, '"./.bilig/pricing.workpaper.json"', '.cursor/mcp.json')
  requireIncludes(kiroMcpConfig, '"bilig-workpaper"', '.kiro/settings/mcp.json')
  requireIncludes(kiroMcpConfig, '"command": "npm"', '.kiro/settings/mcp.json')
  requireIncludes(kiroMcpConfig, '"@bilig/workpaper@latest"', '.kiro/settings/mcp.json')
  requireIncludes(kiroMcpConfig, '"./.bilig/pricing.workpaper.json"', '.kiro/settings/mcp.json')
  requireIncludes(kiroSteeringNotes, 'Kiro steering', '.kiro/steering/bilig-workpaper.md')
  requireIncludes(kiroSteeringNotes, 'bilig-evaluate --door agent-mcp --json', '.kiro/steering/bilig-workpaper.md')
  requireIncludes(kiroSteeringNotes, 'Do not claim success from a write call alone.', '.kiro/steering/bilig-workpaper.md')
  requireIncludes(junieMcpConfig, '"biligWorkpaperFile"', '.junie/mcp/mcp.json')
  requireIncludes(junieMcpConfig, '"command": "npm"', '.junie/mcp/mcp.json')
  requireIncludes(junieMcpConfig, '"@bilig/workpaper@latest"', '.junie/mcp/mcp.json')
  requireIncludes(junieMcpConfig, '"./.bilig/pricing.workpaper.json"', '.junie/mcp/mcp.json')
  requireIncludes(vscodeMcpConfig, '"biligWorkpaperDemo"', '.vscode/mcp.json')
  requireIncludes(vscodeMcpConfig, '"biligWorkpaperFile"', '.vscode/mcp.json')
  requireIncludes(vscodeMcpConfig, '"${workspaceFolder}/.bilig/pricing.workpaper.json"', '.vscode/mcp.json')
  requireIncludes(openCodeMcpConfig, '"$schema": "https://opencode.ai/config.json"', 'opencode.jsonc')
  requireIncludes(openCodeMcpConfig, '"instructions"', 'opencode.jsonc')
  requireIncludes(openCodeMcpConfig, '"AGENTS.md"', 'opencode.jsonc')
  requireIncludes(openCodeMcpConfig, '"bilig-workpaper"', 'opencode.jsonc')
  requireIncludes(openCodeMcpConfig, '"type": "local"', 'opencode.jsonc')
  requireIncludes(openCodeMcpConfig, '"command"', 'opencode.jsonc')
  requireIncludes(openCodeMcpConfig, '"@bilig/workpaper@latest"', 'opencode.jsonc')
  requireIncludes(openCodeMcpConfig, '"./.bilig/pricing.workpaper.json"', 'opencode.jsonc')
  requireIncludes(openCodeMcpConfig, '"bilig-workpaper-demo"', 'opencode.jsonc')
  requireIncludes(openCodeMcpConfig, '"url": "https://bilig.proompteng.ai/mcp"', 'opencode.jsonc')
  requireIncludes(openCodeMcpConfig, '"enabled": false', 'opencode.jsonc')
  requireIncludes(reusableMcpConfig, '"bilig-workpaper"', 'mcp/bilig-workpaper.mcp.json')
  requireIncludes(reusableMcpConfig, '"@bilig/workpaper@latest"', 'mcp/bilig-workpaper.mcp.json')
  requireIncludes(reusableMcpConfig, '"./.bilig/pricing.workpaper.json"', 'mcp/bilig-workpaper.mcp.json')

  requireIncludes(rootSkillNotes, '## Command Safety', 'skills/bilig-workpaper/SKILL.md')
  requireIncludes(rootSkillNotes, '## First Check: Agent Evaluator', 'skills/bilig-workpaper/SKILL.md')
  requireIncludes(rootSkillNotes, '"bilig-evaluate", "--door", "agent-mcp", "--json"', 'skills/bilig-workpaper/SKILL.md')
  requireIncludes(rootSkillNotes, 'argument array, not a shell-concatenated string', 'skills/bilig-workpaper/SKILL.md')
  requireIncludes(
    rootSkillNotes,
    'npx --yes skills@latest add proompteng/bilig --skill bilig-workpaper --list',
    'skills/bilig-workpaper/SKILL.md',
  )
  requireIncludes(rootSkillNotes, 'returned `tools` array as the source', 'skills/bilig-workpaper/SKILL.md')
  requireIncludes(rootSkillNotes, 'set_cell_contents_and_readback', 'skills/bilig-workpaper/SKILL.md')
  requireIncludes(rootSkillNotes, 'currently published package', 'skills/bilig-workpaper/SKILL.md')
  requireNotIncludes(rootSkillNotes, 'allowed-tools:', 'skills/bilig-workpaper/SKILL.md')
  requireNotIncludes(rootSkillNotes, 'argument-hint:', 'skills/bilig-workpaper/SKILL.md')
}

function requireYamlListValue(content: string, value: string, context: string): void {
  const escapedValue = value.replace(/[.*+?^${}()|[\]\\]/gu, '\\$&')
  requireMatches(
    content,
    new RegExp(`^\\s*-\\s*(?:"${escapedValue}"|'${escapedValue}'|${escapedValue})\\s*$`, 'mu'),
    `list value ${value}`,
    context,
  )
}
