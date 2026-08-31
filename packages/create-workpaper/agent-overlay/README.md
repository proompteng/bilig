# `__PROJECT_NAME__`

Formula WorkPaper starter for Node services, MCP clients, and host
integrations, built with `@bilig/workpaper`.

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --scenario revenue-plan --json
```

This starter includes `npm run agent:verify`, which runs the service smoke, the
basic MCP evaluator, and the richer revenue-plan evaluator:

- `npm run smoke`: writes quote inputs through a service-style API handler,
  recalculates formulas, persists WorkPaper JSON, restores it, and checks
  `verified: true`.
- `npm run agent:evaluate`: runs
  `bilig-evaluate --door agent-mcp --scenario revenue-plan --json`,
  discovers MCP tools, edits a WorkPaper cell, reads recalculated `SUM`,
  `SUMIF`, `XLOOKUP`, `FILTER`, and named-expression outputs, exports JSON,
  restarts from disk, and checks `verified: true`.
- `npm run agent:evaluate:basic`: runs the smaller
  `bilig-evaluate --door agent-mcp --json` smoke when you only need the
  minimal MCP contract.

Use `npm run mcp:challenge` only when you need the lower-level JSON-RPC
diagnostic transcript.

Start the local API:

```sh
npm run dev
curl http://localhost:8788/api/quote/approval
curl -X POST http://localhost:8788/api/quote/approval \
  -H 'content-type: application/json' \
  -d '{"units":40,"listPrice":1200,"discount":0.05,"unitCost":760,"minimumMargin":0.3}'
```

Start the persistent project-local MCP server:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-workpaper-mcp --workpaper __WORKPAPER_PATH__ --init-demo-workpaper --writable
```

The server owns `__WORKPAPER_PATH__`, initializes it when missing, writes
through MCP tools, recalculates formulas, and persists edits back to disk.
Project MCP configs are included for Claude Code (`.mcp.json`), Cursor, Kiro,
Roo Code, Trae, Junie, Zed, Continue, and VS Code. Aider loads
`CONVENTIONS.md` through `.aider.conf.yml`. Claude Code gets the project
skill at `.claude/skills/bilig-workpaper/SKILL.md`; OpenHands gets the same
workflow at `.agents/skills/bilig-workpaper/SKILL.md`; OpenCode gets
`opencode.jsonc` and `.opencode/agents/bilig-workpaper.md`. Trae gets
`.trae/mcp.json` and `.trae/rules/bilig-workpaper.md`. Qodo IDE uses the
same `mcpServers` JSON through Agentic Tools MCP settings rather than a
repo-native file. Zed gets `.zed/settings.json` with the
`context_servers.bilig-workpaper` MCP server. Continue gets
`.continue/mcpServers/bilig-workpaper.yaml` for Agent-mode MCP tools. Other
clients can use the same command from `mcp/bilig-workpaper.mcp.json`.

Host instruction files are included for Claude Code, Codex/AGENTS-aware tools,
Gemini CLI, GitHub Copilot, VS Code agent mode, Cursor, Kiro, Roo Code, Trae,
Qodo IDE, OpenHands, OpenCode, Zed, Junie, Aider, Cline, Continue,
Cascade/Devin, and Windsurf/Cascade. They all point at the same rule: read
first, edit one cell or formula, recalculate, read the dependent output,
persist or export, and report proof.

The published overlay installer is release-pending while
`@bilig/create-workpaper@latest` resolves to `0.164.11`; that release's
smoke reports `formulasPersisted: false`. Until a newer release passes a
fresh consumer smoke, copy only the host files needed by the current project
instead of running the generator.

For existing repos, the generated MCP configs keep WorkPaper state under
`./.bilig/pricing.workpaper.json` so the overlay does not add a noisy root
workbook file.

Host handoff:

```text
Use Bilig WorkPaper tools instead of spreadsheet UI automation. Read the
relevant range first, write one precise input or formula change, read the
dependent calculated output, export or serialize the WorkPaper document, and
report editedCell, before, after, persistence evidence, verified, and
limitations. Do not claim success from a write call alone.
```

Host handoff checklist: <https://proompteng.github.io/bilig/agent-adoption-kit.html>

Learn more: <https://github.com/proompteng/bilig>
