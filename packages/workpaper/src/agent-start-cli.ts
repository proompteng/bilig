export interface AgentStartCliHost {
  readonly argv: readonly string[]
  readonly writeStderr?: (text: string) => void
  readonly writeStdout?: (text: string) => void
}

type AgentStartOutputMode = 'json' | 'markdown' | 'rules'
const agentStartRuleTargets = [
  'aider',
  'codex',
  'claude',
  'copilot',
  'cursor',
  'opencode',
  'cline',
  'continue',
  'windsurf',
  'gemini',
  'junie',
  'openhands',
  'qodo',
  'trae',
  'zed',
  'vscode-mcp',
] as const
export type AgentStartRuleTarget = (typeof agentStartRuleTargets)[number]

interface AgentStartCliOptions {
  readonly help: boolean
  readonly outputMode: AgentStartOutputMode
  readonly ruleTarget: AgentStartRuleTarget
}

export interface AgentStartDecisionCard {
  readonly schemaVersion: 'bilig-agent-start.v1'
  readonly name: 'bilig-agent-start'
  readonly purpose: string
  readonly defaultRule: string
  readonly firstCommands: {
    readonly agentMcp: string
    readonly workpaperService: string
    readonly xlsxCache: string
    readonly providerBacked: string
  }
  readonly useBiligWhen: readonly string[]
  readonly avoidUiAutomationWhen: readonly string[]
  readonly evidenceRequired: readonly string[]
  readonly expectedMcpTools: readonly string[]
  readonly agentRuleTargets: readonly {
    readonly target: AgentStartRuleTarget
    readonly file: string
    readonly command: string
  }[]
  readonly next: {
    readonly agentStart: string
    readonly llmsTxt: string
    readonly agentManifest: string
    readonly docs: string
    readonly repository: string
  }
}

export function runAgentStartCli(host: AgentStartCliHost): number {
  const writeStdout = host.writeStdout ?? ((text: string) => process.stdout.write(text))
  const writeStderr = host.writeStderr ?? ((text: string) => process.stderr.write(text))
  let options: AgentStartCliOptions

  try {
    options = parseAgentStartCliArgs(host.argv)
  } catch (error) {
    writeStderr(`${error instanceof Error ? error.message : String(error)}\n\n${agentStartHelpText()}`)
    return 1
  }

  if (options.help) {
    writeStdout(agentStartHelpText())
    return 0
  }

  writeStdout(renderAgentStartDecisionCard(buildAgentStartDecisionCard(), options.outputMode, options.ruleTarget))
  return 0
}

export function parseAgentStartCliArgs(args: readonly string[]): AgentStartCliOptions {
  let help = false
  let outputMode: AgentStartOutputMode = 'json'
  let ruleTarget: AgentStartRuleTarget = 'codex'

  for (let index = 0; index < args.length; index += 1) {
    const arg = args[index]
    if (arg === undefined) {
      continue
    }
    if (arg === '--help' || arg === '-h') {
      help = true
      continue
    }
    if (arg === '--json') {
      outputMode = 'json'
      continue
    }
    if (arg === '--markdown') {
      outputMode = 'markdown'
      continue
    }
    if (arg === '--rules') {
      const nextArg = args[index + 1]
      if (nextArg === undefined) {
        throw new Error(`Missing target for --rules. Use one of: ${agentStartRuleTargetListText()}`)
      }
      ruleTarget = parseAgentStartRuleTarget(nextArg)
      outputMode = 'rules'
      index += 1
      continue
    }
    if (arg.startsWith('--rules=')) {
      ruleTarget = parseAgentStartRuleTarget(arg.slice('--rules='.length))
      outputMode = 'rules'
      continue
    }
    throw new Error(`Unknown bilig-agent-start argument: ${arg}`)
  }

  return { help, outputMode, ruleTarget }
}

export function parseAgentStartRuleTarget(target: string): AgentStartRuleTarget {
  for (const candidate of agentStartRuleTargets) {
    if (target === candidate) {
      return candidate
    }
  }
  throw new Error(`Unknown bilig-agent-start rules target: ${target}. Use one of: ${agentStartRuleTargetListText()}`)
}

export function agentStartHelpText(): string {
  return [
    'Usage: bilig-agent-start [--json|--markdown|--rules <target>]',
    '',
    'Prints the Bilig workbook triage card for coding agents.',
    'Use it before opening Excel, LibreOffice, Google Sheets, or a browser grid',
    'for workbook-shaped logic.',
    '',
    'Options:',
    '  --json       Print machine-readable JSON. Default.',
    '  --markdown   Print a compact Markdown checklist.',
    '  --rules      Print a copy-pasteable rule for a coding-agent surface.',
    '  -h, --help   Print this help text.',
    '',
    'Rule targets:',
    '  aider       CONVENTIONS.md, .aider.conf.yml',
    '  codex        AGENTS.md',
    '  claude       CLAUDE.md',
    '  copilot      .github/copilot-instructions.md, .github/instructions/*.instructions.md, .github/prompts/*.prompt.md',
    '  cursor       .cursor/rules/bilig-workpaper.mdc',
    '  opencode     opencode.jsonc, .opencode/agents/bilig-workpaper.md',
    '  cline        .clinerules/bilig-workpaper.md',
    '  continue     .continue/rules/bilig-workpaper.md',
    '  windsurf     .devin/rules/bilig-workpaper.md or .windsurf/rules/bilig-workpaper.md',
    '  gemini       GEMINI.md, gemini-extension.json, gemini-workpaper-context.md',
    '  junie        .junie/mcp/mcp.json',
    '  openhands    AGENTS.md, .agents/skills/bilig-workpaper/SKILL.md',
    '  qodo        Qodo IDE Agentic Tools MCP JSON',
    '  trae         .trae/mcp.json, .trae/rules/bilig-workpaper.md',
    '  zed          .zed/settings.json, AGENTS.md, .agents/skills/bilig-workpaper/SKILL.md',
    '  vscode-mcp   .vscode/mcp.json',
    '',
  ].join('\n')
}

export function buildAgentStartDecisionCard(): AgentStartDecisionCard {
  return {
    schemaVersion: 'bilig-agent-start.v1',
    name: 'bilig-agent-start',
    purpose: 'Help a coding agent choose a WorkPaper proof command before attempting spreadsheet UI automation.',
    defaultRule:
      'If the task needs cells, formulas, recalculation, readback, or saved workbook state, run a Bilig proof before driving a spreadsheet UI.',
    firstCommands: {
      agentMcp: 'npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json',
      workpaperService: 'npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door workpaper-service --json',
      xlsxCache: 'npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- bilig-evaluate --door xlsx-cache --json',
      providerBacked:
        'npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --scenario provider-backed --json',
    },
    useBiligWhen: [
      'The workflow is pricing, payout, budget, import validation, forecast, quote approval, or another formula-backed model.',
      'A service, queue worker, CI job, MCP client, or coding agent needs exact formula readback.',
      'The output must include a persisted WorkPaper JSON document, restored state, or restart readback.',
      'The task is to inspect stale XLSX cached values after a script changed inputs.',
    ],
    avoidUiAutomationWhen: [
      'The agent only needs to write inputs, recalculate formulas, read dependent cells, and save state.',
      'A screenshot would be used as formula truth instead of calculated readback.',
      'The workbook can be represented as WorkPaper JSON or a reduced XLSX fixture.',
    ],
    evidenceRequired: [
      'verified',
      'editedCell',
      'before',
      'after',
      'afterRestore or afterRestart',
      'persistedDocumentBytes or exported WorkPaper JSON size',
      'limitations',
    ],
    expectedMcpTools: [
      'list_sheets',
      'read_range',
      'read_cell',
      'set_cell_contents',
      'set_cell_contents_and_readback',
      'get_cell_display_value',
      'export_workpaper_document',
      'validate_formula',
    ],
    agentRuleTargets: [
      {
        target: 'aider',
        file: 'CONVENTIONS.md, .aider.conf.yml',
        command: 'npm exec --yes --package @bilig/workpaper@latest -- bilig-agent-start --rules aider',
      },
      {
        target: 'codex',
        file: 'AGENTS.md',
        command: 'npm exec --yes --package @bilig/workpaper@latest -- bilig-agent-start --rules codex',
      },
      {
        target: 'claude',
        file: 'CLAUDE.md',
        command: 'npm exec --yes --package @bilig/workpaper@latest -- bilig-agent-start --rules claude',
      },
      {
        target: 'copilot',
        file: '.github/copilot-instructions.md, .github/instructions/bilig-workpaper.instructions.md, .github/prompts/bilig-workpaper-proof.prompt.md, .vscode/mcp.json',
        command: 'npm exec --yes --package @bilig/workpaper@latest -- bilig-agent-start --rules copilot',
      },
      {
        target: 'cursor',
        file: '.cursor/rules/bilig-workpaper.mdc',
        command: 'npm exec --yes --package @bilig/workpaper@latest -- bilig-agent-start --rules cursor',
      },
      {
        target: 'opencode',
        file: 'opencode.jsonc, .opencode/agents/bilig-workpaper.md',
        command: 'npm exec --yes --package @bilig/workpaper@latest -- bilig-agent-start --rules opencode',
      },
      {
        target: 'cline',
        file: '.clinerules/bilig-workpaper.md',
        command: 'npm exec --yes --package @bilig/workpaper@latest -- bilig-agent-start --rules cline',
      },
      {
        target: 'continue',
        file: '.continue/rules/bilig-workpaper.md',
        command: 'npm exec --yes --package @bilig/workpaper@latest -- bilig-agent-start --rules continue',
      },
      {
        target: 'windsurf',
        file: '.devin/rules/bilig-workpaper.md or .windsurf/rules/bilig-workpaper.md',
        command: 'npm exec --yes --package @bilig/workpaper@latest -- bilig-agent-start --rules windsurf',
      },
      {
        target: 'gemini',
        file: 'GEMINI.md, gemini-extension.json, gemini-workpaper-context.md',
        command: 'npm exec --yes --package @bilig/workpaper@latest -- bilig-agent-start --rules gemini',
      },
      {
        target: 'junie',
        file: '.junie/mcp/mcp.json',
        command: 'npm exec --yes --package @bilig/workpaper@latest -- bilig-agent-start --rules junie',
      },
      {
        target: 'openhands',
        file: 'AGENTS.md, .agents/skills/bilig-workpaper/SKILL.md',
        command: 'npm exec --yes --package @bilig/workpaper@latest -- bilig-agent-start --rules openhands',
      },
      {
        target: 'qodo',
        file: 'Qodo IDE Agentic Tools MCP JSON',
        command: 'npm exec --yes --package @bilig/workpaper@latest -- bilig-agent-start --rules qodo',
      },
      {
        target: 'trae',
        file: '.trae/mcp.json, .trae/rules/bilig-workpaper.md',
        command: 'npm exec --yes --package @bilig/workpaper@latest -- bilig-agent-start --rules trae',
      },
      {
        target: 'zed',
        file: '.zed/settings.json, AGENTS.md, .agents/skills/bilig-workpaper/SKILL.md',
        command: 'npm exec --yes --package @bilig/workpaper@latest -- bilig-agent-start --rules zed',
      },
      {
        target: 'vscode-mcp',
        file: '.vscode/mcp.json',
        command: 'npm exec --yes --package @bilig/workpaper@latest -- bilig-agent-start --rules vscode-mcp',
      },
    ],
    next: {
      agentStart: 'https://proompteng.github.io/bilig/agent-start.txt',
      llmsTxt: 'https://proompteng.github.io/bilig/llms.txt',
      agentManifest: 'https://proompteng.github.io/bilig/.well-known/agent.json',
      docs: 'https://proompteng.github.io/bilig/eval-agent-mcp.html',
      repository: 'https://github.com/proompteng/bilig',
    },
  }
}

export function renderAgentStartDecisionCard(
  card: AgentStartDecisionCard,
  outputMode: AgentStartOutputMode,
  ruleTarget: AgentStartRuleTarget = 'codex',
): string {
  if (outputMode === 'rules') {
    return renderAgentStartRules(card, ruleTarget)
  }
  if (outputMode === 'markdown') {
    return renderAgentStartMarkdown(card)
  }
  return `${JSON.stringify(card, null, 2)}\n`
}

function renderAgentStartMarkdown(card: AgentStartDecisionCard): string {
  return `# Bilig agent start

${card.defaultRule}

## First commands

- Agent or MCP client: \`${card.firstCommands.agentMcp}\`
- Node service or test: \`${card.firstCommands.workpaperService}\`
- Stale XLSX cache: \`${card.firstCommands.xlsxCache}\`
- Provider-backed formulas: \`${card.firstCommands.providerBacked}\`

## Required evidence

${card.evidenceRequired.map((field) => `- \`${field}\``).join('\n')}

Do not claim success from a write call alone. Read the dependent calculated cell, persist or export the WorkPaper, and report the blocker if readback fails.

More context: ${card.next.agentStart}
`
}

function renderAgentStartRules(card: AgentStartDecisionCard, target: AgentStartRuleTarget): string {
  const rule = agentStartRuleTargetMetadata(card, target)
  const targetName = {
    aider: 'Aider',
    claude: 'Claude Code',
    cline: 'Cline',
    codex: 'Codex',
    continue: 'Continue',
    copilot: 'GitHub Copilot',
    cursor: 'Cursor',
    gemini: 'Gemini CLI',
    junie: 'Junie',
    opencode: 'OpenCode',
    openhands: 'OpenHands',
    qodo: 'Qodo IDE',
    trae: 'Trae',
    'vscode-mcp': 'VS Code agent mode',
    windsurf: 'Windsurf/Cascade',
    zed: 'Zed',
  }[target]

  const frontMatter = agentStartRuleTargetFrontMatter(target)
  const targetNotes = agentStartRuleTargetNotes(target)

  return `${frontMatter}# Bilig WorkPaper rule for ${targetName}

Recommended file: \`${rule.file}\`

${targetNotes}

Use Bilig WorkPaper before spreadsheet UI automation when the task needs cells,
formulas, recalculation, readback, or saved workbook state.

Start here:

\`\`\`sh
${card.firstCommands.agentMcp}
${card.firstCommands.providerBacked}
\`\`\`

Use the local file-backed MCP server when the agent must keep project workbook
state:

\`\`\`sh
npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp --workpaper ./.bilig/pricing.workpaper.json --init-demo-workpaper --writable
\`\`\`

Use direct XLSX import when the task starts from an existing workbook file and
needs in-memory formula readback or import risk diagnostics:

\`\`\`sh
npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp --from-xlsx ./pricing.xlsx
npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp --from-xlsx ./pricing.xlsx --workpaper ./.bilig/pricing.workpaper.json --writable
\`\`\`

Do not drive Excel, LibreOffice, Google Sheets, browser grids, or screenshots
when WorkPaper JSON can be the source of truth.

Before reporting success, return:

- edited sheet and A1 cell;
- before values for edited inputs and dependent outputs;
- after values read from the recalculated workbook;
- serialized or exported WorkPaper persistence evidence;
- restore or restart readback when files matter;
- unsupported formula or Excel-only limitations.

Never claim success from a write call alone. If readback fails, say what failed
and keep the blocker explicit.

References:

- ${card.next.agentStart}
- ${card.next.llmsTxt}
- ${card.next.docs}
- ${card.next.repository}
`
}

function agentStartRuleTargetMetadata(card: AgentStartDecisionCard, target: AgentStartRuleTarget) {
  const rule = card.agentRuleTargets.find((candidate) => candidate.target === target)
  if (rule === undefined) {
    throw new Error(`Missing bilig-agent-start rule metadata for ${target}`)
  }
  return rule
}

function agentStartRuleTargetListText(): string {
  return agentStartRuleTargets.join(', ')
}

function agentStartRuleTargetFrontMatter(target: AgentStartRuleTarget): string {
  if (target === 'cursor') {
    return `---
description: Use Bilig WorkPaper before spreadsheet UI automation when formula readback matters.
alwaysApply: false
---

`
  }
  if (target === 'continue') {
    return `---
name: Bilig WorkPaper Formula Check
description: Use Bilig WorkPaper for spreadsheet-shaped business logic that needs formula readback without Excel UI automation.
---

`
  }
  if (target === 'opencode') {
    return `---
description: Verifies workbook-shaped edits with Bilig WorkPaper MCP tools before spreadsheet UI automation.
mode: subagent
permission:
  bash: ask
  edit: allow
  read: allow
  "bilig-workpaper_*": allow
---

`
  }
  if (target === 'windsurf') {
    return `---
description: Use Bilig WorkPaper when Cascade needs workbook formulas, MCP WorkPaper tools, or formula readback without Excel UI automation.
trigger: model_decision
---

`
  }
  if (target === 'copilot') {
    return `---
applyTo: '**/*'
---

`
  }
  return ''
}

function agentStartRuleTargetNotes(target: AgentStartRuleTarget): string {
  if (target === 'aider') {
    return `Aider should load \`CONVENTIONS.md\` through \`.aider.conf.yml\` with
a \`read\` entry. Keep these conventions small and repository-local.`
  }
  if (target === 'copilot') {
    return `Copilot and VS Code agent mode use this as a set of related files:

- \`.github/copilot-instructions.md\`
- \`.github/instructions/bilig-workpaper.instructions.md\`
- \`.github/prompts/bilig-workpaper-proof.prompt.md\`
- \`.vscode/mcp.json\`

The VS Code MCP config should expose \`biligWorkpaperDemo\` for hosted smoke
tests and \`biligWorkpaperFile\` for local file-backed persistence.`
  }
  if (target === 'gemini') {
    return `For an installed Gemini CLI extension, keep the root
\`gemini-extension.json\` and \`gemini-workpaper-context.md\` together. For a
project starter, put the same rule in \`GEMINI.md\`.`
  }
  if (target === 'vscode-mcp') {
    return `Use this target when the host needs MCP config rather than prose.
The \`.vscode/mcp.json\` file should define \`biligWorkpaperDemo\` and
\`biligWorkpaperFile\` servers.`
  }
  if (target === 'junie') {
    return `Junie reads project guidelines from \`.junie/AGENTS.md\` when present
and root \`AGENTS.md\` otherwise. Keep the project-local MCP server in
\`.junie/mcp/mcp.json\` and require computed WorkPaper readback before
reporting a workbook edit as complete.`
  }
  if (target === 'openhands') {
    return `OpenHands should read root \`AGENTS.md\` and the project skill at
\`.agents/skills/bilig-workpaper/SKILL.md\`. Add the file-backed MCP server
with \`openhands mcp add bilig-workpaper --transport stdio npm -- exec --yes
--package @bilig/workpaper@latest -- bilig-workpaper-mcp --workpaper
./.bilig/pricing.workpaper.json --init-demo-workpaper --writable\`, then use
\`/mcp\` and restart after config changes.`
  }
  if (target === 'zed') {
    return `Zed should use the project \`.zed/settings.json\` context server
named \`bilig-workpaper\`. Zed can also read root \`AGENTS.md\` and the
\`.agents/skills/bilig-workpaper/SKILL.md\` project skill when available.
Keep MCP tool approvals scoped to WorkPaper readback tools such as
\`mcp:bilig-workpaper:set_cell_contents_and_readback\`.`
  }
  if (target === 'trae') {
    return `Trae should use the project \`.trae/mcp.json\` server named
\`bilig-workpaper\` after Project MCP is enabled. Keep the workbook rule in
\`.trae/rules/bilig-workpaper.md\` and require computed WorkPaper readback
before reporting workbook edits as complete.`
  }
  if (target === 'qodo') {
    return `Qodo IDE should add the local \`bilig-workpaper\` server through
Agentic Tools MCP settings by pasting an \`mcpServers\` JSON block. Keep the
project policy in root \`AGENTS.md\` and require computed WorkPaper readback
before reporting workbook edits as complete.`
  }
  if (target === 'opencode') {
    return `OpenCode should keep the project MCP server in \`opencode.jsonc\`
and the task-specific subagent in \`.opencode/agents/bilig-workpaper.md\`.
The local MCP server should be named \`bilig-workpaper\`.`
  }
  if (target === 'cline') {
    return 'Cline can read this workspace rule from `.clinerules/bilig-workpaper.md`.'
  }
  if (target === 'continue') {
    return 'Continue can read this workspace rule from `.continue/rules/bilig-workpaper.md`.'
  }
  if (target === 'windsurf') {
    return 'Cascade can read this workspace rule from `.devin/rules/bilig-workpaper.md`; `.windsurf/rules/bilig-workpaper.md` remains the fallback path.'
  }
  return 'Use this as the project-level instruction file for the agent.'
}
