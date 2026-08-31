export function withStarterWorkpaperPath(content: string): string {
  return content
    .replaceAll('./.bilig/pricing.workpaper.json', '__WORKPAPER_PATH__')
    .replaceAll('${workspaceFolder}/.bilig/pricing.workpaper.json', '__WORKPAPER_PATH__')
    .replaceAll('./pricing.workpaper.json', '__WORKPAPER_PATH__')
}

export function buildStarterAgentOverlayInstructions(): string {
  return `# Agent Instructions

Use \`@bilig/workpaper\` as the source of truth for workbook logic in this
project. Do not open Excel, LibreOffice, Google Sheets, or a browser grid for
primary formula work unless a human explicitly asks for visual review.

## Verify First

\`\`\`sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --scenario revenue-plan --json
\`\`\`

That command is the package-owned proof gate. A valid run includes
\`verified: true\` and covers \`SUM\`, \`SUMIF\`, \`XLOOKUP\`, \`FILTER\`, a named
expression, persistence, and restart readback. Generated starters also include
\`npm run agent:verify\`.

## Agent Surfaces

This starter includes project instructions for common coding agents:

- Claude Code: \`CLAUDE.md\`, \`.claude/skills/bilig-workpaper/SKILL.md\`,
  and \`.claude/commands/bilig-workpaper-proof.md\`
- Codex and other AGENTS-aware tools: \`AGENTS.md\`
- Gemini CLI: \`GEMINI.md\`
- GitHub Copilot and VS Code agent mode:
  \`.github/copilot-instructions.md\`,
  \`.github/instructions/bilig-workpaper.instructions.md\`,
  \`.github/prompts/bilig-workpaper-proof.prompt.md\`, and \`.vscode/mcp.json\`
- Cursor: \`.cursor/rules/bilig-workpaper.mdc\` and \`.cursor/mcp.json\`
- Kiro: \`.kiro/steering/bilig-workpaper.md\` and \`.kiro/settings/mcp.json\`
- Roo Code: \`.roo/rules/bilig-workpaper.md\` and \`.roo/mcp.json\`
- Trae: \`.trae/rules/bilig-workpaper.md\` and \`.trae/mcp.json\`
- Qodo IDE: paste the \`bilig-workpaper\` JSON from the public Qodo setup guide
  into Qodo Agentic Tools MCP settings; keep root \`AGENTS.md\` as policy
- OpenHands: \`.agents/skills/bilig-workpaper/SKILL.md\`
- OpenCode: \`opencode.jsonc\` and \`.opencode/agents/bilig-workpaper.md\`
- Aider: \`CONVENTIONS.md\` loaded by \`.aider.conf.yml\`
- Cascade/Devin: \`.devin/rules/bilig-workpaper.md\`
- Cline: \`.clinerules/bilig-workpaper.md\`
- Continue: \`.continue/rules/bilig-workpaper.md\` and
  \`.continue/mcpServers/bilig-workpaper.yaml\`
- Windsurf/Cascade fallback: \`.windsurf/rules/bilig-workpaper.md\`

## Preferred Agent Loop

1. Read the relevant sheet, range, or API output before editing.
2. Name the exact sheet and A1 cell target.
3. Validate formulas before writing them.
4. Prefer \`set_cell_contents_and_readback\` for one-call edit plus dependent
   output readback.
5. Otherwise, write one small input or formula change and then read the
   dependent calculated output after recalculation.
6. Export or serialize the WorkPaper document.
7. Report \`editedCell\`, \`before\`, \`after\`, \`afterRestore\` or persistence
   evidence, \`verified\`, and known limitations.

Do not claim success from a write call alone. Success requires computed
readback plus persisted WorkPaper state.

## MCP Server

Start the persistent project-local MCP server with:

\`\`\`sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-workpaper-mcp --workpaper __WORKPAPER_PATH__ --init-demo-workpaper --writable
\`\`\`

It launches:

\`\`\`sh
bilig-workpaper-mcp --workpaper __WORKPAPER_PATH__ --init-demo-workpaper --writable
\`\`\`

Expected tools:

- \`list_sheets\`
- \`read_range\`
- \`read_cell\`
- \`set_cell_contents\`
- \`set_cell_contents_and_readback\`
- \`get_cell_display_value\`
- \`export_workpaper_document\`
- \`validate_formula\`
`
}

export function buildStarterClaudeInstructions(): string {
  return `# Claude Project Instructions

Read \`AGENTS.md\` first. For workbook tasks, prefer the Bilig WorkPaper API or
the project-local MCP server over spreadsheet UI automation.

Before reporting success, run:

\`\`\`sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --scenario revenue-plan --json
\`\`\`

That evaluator checks MCP tool discovery, mutation, recalculated \`SUM\`,
\`SUMIF\`, \`XLOOKUP\`, \`FILTER\`, a named expression, persistence, and restart
readback.

For MCP use, start:

\`\`\`sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-workpaper-mcp --workpaper __WORKPAPER_PATH__ --init-demo-workpaper --writable
\`\`\`

Claude Code can load the project skill at
\`.claude/skills/bilig-workpaper/SKILL.md\`. For an explicit proof contract, use
the project command in \`.claude/commands/bilig-workpaper-proof.md\`.
`
}

export function buildStarterGeminiInstructions(): string {
  return `# Bilig WorkPaper Instructions For Gemini

Use \`@bilig/workpaper\` as the source of truth for workbook logic in this
project. Prefer the WorkPaper API or the project-local MCP server over Excel,
LibreOffice, Google Sheets, browser grids, screenshots, or stale cached XLSX
values.

Before reporting success, run:

\`\`\`sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --scenario revenue-plan --json
\`\`\`

That evaluator checks MCP tool discovery, mutation, recalculated \`SUM\`,
\`SUMIF\`, \`XLOOKUP\`, \`FILTER\`, a named expression, persistence, and restart
readback.

For MCP use, start:

\`\`\`sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-workpaper-mcp --workpaper __WORKPAPER_PATH__ --init-demo-workpaper --writable
\`\`\`

Return proof with:

- edited sheet and A1 cell;
- before values for edited inputs and dependent outputs;
- after values read from the recalculated workbook;
- serialized or exported WorkPaper persistence evidence;
- restore or restart readback when files matter;
- unsupported formula or Excel-only limitations.

Do not claim success from a write call alone. If a proof step fails, say which
step failed and what evidence is missing.
`
}

export function buildStarterOverlayReadme(): string {
  return `# \`__PROJECT_NAME__\`

Formula WorkPaper starter for Node services, MCP clients, and host
integrations, built with \`@bilig/workpaper\`.

\`\`\`sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --scenario revenue-plan --json
\`\`\`

This starter includes \`npm run agent:verify\`, which runs the service smoke, the
basic MCP evaluator, and the richer revenue-plan evaluator:

- \`npm run smoke\`: writes quote inputs through a service-style API handler,
  recalculates formulas, persists WorkPaper JSON, restores it, and checks
  \`verified: true\`.
- \`npm run agent:evaluate\`: runs
  \`bilig-evaluate --door agent-mcp --scenario revenue-plan --json\`,
  discovers MCP tools, edits a WorkPaper cell, reads recalculated \`SUM\`,
  \`SUMIF\`, \`XLOOKUP\`, \`FILTER\`, and named-expression outputs, exports JSON,
  restarts from disk, and checks \`verified: true\`.
- \`npm run agent:evaluate:basic\`: runs the smaller
  \`bilig-evaluate --door agent-mcp --json\` smoke when you only need the
  minimal MCP contract.

Use \`npm run mcp:challenge\` only when you need the lower-level JSON-RPC
diagnostic transcript.

Start the local API:

\`\`\`sh
npm run dev
curl http://localhost:8788/api/quote/approval
curl -X POST http://localhost:8788/api/quote/approval \\
  -H 'content-type: application/json' \\
  -d '{"units":40,"listPrice":1200,"discount":0.05,"unitCost":760,"minimumMargin":0.3}'
\`\`\`

Start the persistent project-local MCP server:

\`\`\`sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-workpaper-mcp --workpaper __WORKPAPER_PATH__ --init-demo-workpaper --writable
\`\`\`

The server owns \`__WORKPAPER_PATH__\`, initializes it when missing, writes
through MCP tools, recalculates formulas, and persists edits back to disk.
Project MCP configs are included for Claude Code (\`.mcp.json\`), Cursor, Kiro,
Roo Code, Trae, Junie, Zed, Continue, and VS Code. Aider loads
\`CONVENTIONS.md\` through \`.aider.conf.yml\`. Claude Code gets the project
skill at \`.claude/skills/bilig-workpaper/SKILL.md\`; OpenHands gets the same
workflow at \`.agents/skills/bilig-workpaper/SKILL.md\`; OpenCode gets
\`opencode.jsonc\` and \`.opencode/agents/bilig-workpaper.md\`. Trae gets
\`.trae/mcp.json\` and \`.trae/rules/bilig-workpaper.md\`. Qodo IDE uses the
same \`mcpServers\` JSON through Agentic Tools MCP settings rather than a
repo-native file. Zed gets \`.zed/settings.json\` with the
\`context_servers.bilig-workpaper\` MCP server. Continue gets
\`.continue/mcpServers/bilig-workpaper.yaml\` for Agent-mode MCP tools. Other
clients can use the same command from \`mcp/bilig-workpaper.mcp.json\`.

Host instruction files are included for Claude Code, Codex/AGENTS-aware tools,
Gemini CLI, GitHub Copilot, VS Code agent mode, Cursor, Kiro, Roo Code, Trae,
Qodo IDE, OpenHands, OpenCode, Zed, Junie, Aider, Cline, Continue,
Cascade/Devin, and Windsurf/Cascade. They all point at the same rule: read
first, edit one cell or formula, recalculate, read the dependent output,
persist or export, and report proof.

The published overlay installer is release-pending while
\`@bilig/create-workpaper@latest\` resolves to \`0.164.11\`; that release's
smoke reports \`formulasPersisted: false\`. Until a newer release passes a
fresh consumer smoke, copy only the host files needed by the current project
instead of running the generator.

For existing repos, the generated MCP configs keep WorkPaper state under
\`./.bilig/pricing.workpaper.json\` so the overlay does not add a noisy root
workbook file.

Host handoff:

\`\`\`text
Use Bilig WorkPaper tools instead of spreadsheet UI automation. Read the
relevant range first, write one precise input or formula change, read the
dependent calculated output, export or serialize the WorkPaper document, and
report editedCell, before, after, persistence evidence, verified, and
limitations. Do not claim success from a write call alone.
\`\`\`

Host handoff checklist: <https://proompteng.github.io/bilig/agent-adoption-kit.html>

Learn more: <https://github.com/proompteng/bilig>
`
}

export function buildStarterOverlayPackageJson(): string {
  return `${JSON.stringify(
    {
      name: '__PROJECT_NAME__',
      private: true,
      type: 'module',
      scripts: {
        dev: 'tsx src/index.ts --serve',
        smoke: 'tsx src/index.ts',
        typecheck: 'tsc --noEmit',
        'agent:verify': 'npm run smoke && npm run agent:evaluate:basic && npm run agent:evaluate',
        'agent:evaluate': 'bilig-evaluate --door agent-mcp --scenario revenue-plan --json',
        'agent:evaluate:basic': 'bilig-evaluate --door agent-mcp --json',
        'mcp:challenge': 'bilig-mcp-challenge --json',
        'mcp:server': 'bilig-workpaper-mcp --workpaper __WORKPAPER_PATH__ --init-demo-workpaper --writable',
      },
      dependencies: {
        '@bilig/workpaper': '__BILIG_WORKPAPER_VERSION__',
      },
      devDependencies: {
        '@types/node': '25.5.0',
        tsx: '4.21.0',
        typescript: '6.0.2',
      },
    },
    null,
    2,
  )}\n`
}
