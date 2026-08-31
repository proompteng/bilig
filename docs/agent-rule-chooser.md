---
title: Coding agent rule chooser for Bilig WorkPaper
published: true
description: Pick the Bilig instruction, rule, prompt, or MCP config for Codex, Claude Code, GitHub Copilot, VS Code, Cursor, Kiro, Roo Code, Trae, Qodo IDE, Zed, JetBrains Junie, OpenHands, OpenCode, Aider, Windsurf, Cline, Continue, and Gemini CLI.
tags: agents, agent-rules, mcp, workbook formulas, coding agents
canonical_url: https://proompteng.github.io/bilig/agent-rule-chooser.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Coding Agent Rule Chooser

Use this page when a coding agent is in a cloned repo and you need to know
which Bilig file it should read before spreadsheet-shaped work.

Run the no-key agent MCP proof first:

```sh
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --json
```

The result must include `schemaVersion: "bilig-evaluator.v1"`,
`door: "agent-mcp"`, `verified: true`, edited cell evidence, formula readback,
exported or persisted WorkPaper state, and restore or restart readback. A write
call alone is not success.

## Quick Choice

| Agent host | Use this Bilig file | Tool hookup | Proof bar |
| --- | --- | --- | --- |
| Codex | `AGENTS.md` in the repo directory chain. Public handoff: `docs/AGENTS.md`. | Optional `.mcp.json` when the Codex environment supports MCP. | Run `bilig-agent-start --json`, then `bilig-evaluate --door agent-mcp --json`. |
| Claude Code | `CLAUDE.md`, then `.claude/skills/bilig-workpaper/SKILL.md` or `.claude/commands/bilig-workpaper-proof.md`. | `.mcp.json` defines the file-backed `bilig-workpaper` stdio server. | Use `/bilig-workpaper-proof <task>` before Excel, LibreOffice, Sheets, browser grids, or screenshots. |
| GitHub Copilot | `.github/copilot-instructions.md` plus `.github/instructions/bilig-workpaper.instructions.md`. | `.github/prompts/bilig-workpaper-proof.prompt.md` for the task prompt, `.vscode/mcp.json` in VS Code. | Copilot should return WorkPaper readback fields, not spreadsheet UI status. |
| VS Code agent mode | `.github/copilot-instructions.md` and `.github/instructions/bilig-workpaper.instructions.md`. | `.vscode/mcp.json` for `biligWorkpaperDemo` and `biligWorkpaperFile`. | Use the workspace MCP config before copying a generic `mcpServers` manifest. |
| Cursor | `.cursor/rules/bilig-workpaper.mdc`. | `.cursor/mcp.json` for local file-backed WorkPaper tools. | Treat `.cursorrules` as legacy; use the project rule and MCP config here. |
| Kiro | `.kiro/steering/bilig-workpaper.md`; Kiro also loads root `AGENTS.md` when present. | `.kiro/settings/mcp.json` defines the project-local file-backed WorkPaper MCP server. | Use Kiro steering and the project MCP server before spreadsheet UI automation. |
| Roo Code | `.roo/rules/bilig-workpaper.md`; Roo also loads root `AGENTS.md` by default. | `.roo/mcp.json` defines the project-local file-backed WorkPaper MCP server. | Use Roo's project rule and MCP server before spreadsheet UI automation. |
| Trae | `.trae/rules/bilig-workpaper.md`; Trae also loads root `AGENTS.md` when present. | `.trae/mcp.json` defines the project-local file-backed WorkPaper MCP server after Project MCP is enabled. | Use Trae's project rule and MCP server before spreadsheet UI automation. |
| Qodo IDE | `AGENTS.md` plus the [Qodo WorkPaper MCP setup](qodo-workpaper-mcp.md). | Paste the `bilig-workpaper` JSON into Qodo Agentic Tools MCP settings. | Use Qodo's local MCP tool before spreadsheet UI automation; do not claim a repo-native `.qodo` config file. |
| Zed | `.zed/settings.json`, root `AGENTS.md`, and `.agents/skills/bilig-workpaper/SKILL.md`. | `.zed/settings.json` defines the project-local `context_servers.bilig-workpaper` MCP server. | Use Zed's context server before spreadsheet UI automation and keep tool permissions scoped to WorkPaper readback. |
| JetBrains Junie | `AGENTS.md` in the repo root; `.junie/AGENTS.md` can add narrower project memory when needed. | `.junie/mcp/mcp.json` defines the file-backed WorkPaper MCP server. | Use Junie MCP tools for workbook readback and require persisted WorkPaper evidence before reporting success. |
| OpenHands | `AGENTS.md`, then `.agents/skills/bilig-workpaper/SKILL.md`. | `openhands mcp add bilig-workpaper --transport stdio npm -- exec --yes --package @bilig/workpaper@latest -- bilig-workpaper-mcp --workpaper ./.bilig/pricing.workpaper.json --init-demo-workpaper --writable`. | Use `/mcp` in the conversation and restart after MCP config changes. |
| OpenCode | `opencode.jsonc`, then `.opencode/agents/bilig-workpaper.md`. | `opencode.jsonc` defines the local `bilig-workpaper` MCP server and a disabled hosted demo server. | Invoke the `@bilig-workpaper` subagent for workbook-shaped tasks and require readback fields. |
| Aider | `CONVENTIONS.md`, loaded by `.aider.conf.yml`; public page: [Aider WorkPaper conventions](aider-workpaper-conventions.md). | Run the local `bilig-workpaper-mcp` command from the conventions when state must persist. | Keep Aider's answer tied to WorkPaper readback, export or restore evidence, and explicit limitations. |
| Windsurf/Cascade | `.devin/rules/bilig-workpaper.md`, with `.windsurf/rules/bilig-workpaper.md` kept as a fallback. | Start with the same `bilig-evaluate --door agent-mcp --json` command, then file-backed MCP if state must persist. | The rule uses `trigger: model_decision`; require computed readback before reporting success. |
| Cline | `.clinerules/bilig-workpaper.md`. | Add MCP through Cline's current MCP settings when direct tool calls are needed. | Cline should use the workspace rule when workbook formulas, cells, or MCP WorkPaper tools appear. |
| Continue | `.continue/rules/bilig-workpaper.md`. | `.continue/mcpServers/bilig-workpaper.yaml` defines the project-local file-backed WorkPaper MCP server. | Use the rule for Agent, Chat, and Edit requests; use the MCP block from Continue Agent mode when the task needs direct workbook tools. |
| Gemini CLI | `gemini-extension.json` plus `gemini-workpaper-context.md`; generated starters also include `GEMINI.md`. | `gemini extensions install https://github.com/proompteng/bilig --ref main`. | The extension starts the `bilig-workpaper` MCP server and injects the WorkPaper proof context. |

## Existing Repo Overlay

The published starter overlay is release-pending. Do not use
`npm create @bilig/workpaper@latest` while `@bilig/create-workpaper@latest`
resolves to `0.164.11`; that release's generated smoke reports
`formulasPersisted: false`. Until a newer release passes a fresh consumer
smoke, copy only the host files named in the table above and use their direct
`npm exec` MCP commands.

## Confusion Guards

- `docs/AGENTS.md` is a public handoff page. Codex reads `AGENTS.md` from the
  cloned repo directory chain.
- Claude Code reads `CLAUDE.md`, not `AGENTS.md`; this repo's project memory
  routes it to the Claude Code skill, slash command, and `.mcp.json`.
- `.vscode/mcp.json` uses the VS Code `servers` shape. `mcp/bilig-workpaper.mcp.json`
  is the reusable `mcpServers` shape for other clients.
- Kiro reads workspace steering from `.kiro/steering/` and project MCP servers
  from `.kiro/settings/mcp.json`; root `AGENTS.md` stays the shared policy.
- Roo Code reads workspace rules from `.roo/rules/` and project MCP servers
  from `.roo/mcp.json`; root `AGENTS.md` stays the shared policy.
- Trae reads project rules from `.trae/rules/` and Project MCP servers from
  `.trae/mcp.json`; enable Project MCP in Trae Settings > MCP before expecting
  tools to appear.
- Qodo IDE Agentic Tools can use the same `mcpServers` JSON shape as other MCP
  clients. Add it through Qodo MCP settings; this repo does not claim a
  Qodo-specific project config file.
- Zed reads project context servers from `.zed/settings.json`. Zed can use
  `AGENTS.md` and `.agents/skills/bilig-workpaper/SKILL.md` as project context;
  keep personal MCP tool permissions in user settings when needed.
- Junie project MCP config lives at `.junie/mcp/mcp.json`; root `AGENTS.md`
  remains the shared project instruction file unless `.junie/AGENTS.md` is
  needed for Junie-only memory.
- Aider loads `CONVENTIONS.md` through `.aider.conf.yml`; keep the file focused
  on WorkPaper proof, not broad repo policy that belongs in `AGENTS.md`.
- Cascade/Devin docs currently prefer `.devin/rules`; the `.windsurf/rules`
  mirror remains for compatible Windsurf/Cascade installs.
- `GEMINI.md` is the normal Gemini CLI context file, but this repo exposes the
  installable Gemini extension path first.

## Official Host Docs Checked

- [Codex AGENTS.md](https://github.com/openai/codex/blob/main/docs/agents_md.md)
- [Claude Code memory](https://code.claude.com/docs/en/memory)
- [GitHub Copilot response customization](https://docs.github.com/en/copilot/concepts/prompting/response-customization)
- [VS Code MCP configuration](https://code.visualstudio.com/docs/copilot/reference/mcp-configuration)
- [Cursor rules](https://docs.cursor.com/en/context/rules)
- [Kiro steering](https://kiro.dev/docs/steering/)
- [Kiro MCP configuration](https://kiro.dev/docs/mcp/configuration/)
- [Roo Code custom instructions](https://roocodeinc.github.io/Roo-Code/features/custom-instructions)
- [Roo Code MCP configuration](https://roocodeinc.github.io/Roo-Code/features/mcp/using-mcp-in-roo/)
- [Trae Model Context Protocol](https://docs.trae.ai/ide/model-context-protocol)
- [Trae add MCP servers](https://docs.trae.ai/ide/add-mcp-servers)
- [Trae rules](https://docs.trae.ai/ide/rules)
- [Trae skills](https://docs.trae.ai/ide/skills)
- [Qodo Agentic Tools MCP](https://docs.qodo.ai/qodo-documentation/qodo-ide/tools-mcps/agentic-tools-mcps)
- [Qodo Merge configuration](https://docs.qodo.ai/qodo-documentation/qodo-review/configuration/qodo-merge-configuration)
- [Zed MCP](https://zed.dev/docs/ai/mcp)
- [Zed rules](https://zed.dev/docs/ai/rules)
- [Zed tool permissions](https://zed.dev/docs/ai/tool-permissions)
- [Junie MCP settings](https://junie.jetbrains.com/docs/junie-plugin-mcp-settings.html)
- [Junie guidelines and memory](https://junie.jetbrains.com/docs/guidelines-and-memory.html)
- [OpenHands MCP servers](https://docs.openhands.dev/openhands/usage/cli/mcp-servers)
- [OpenHands skills](https://docs.openhands.dev/overview/skills)
- [OpenCode config](https://opencode.ai/docs/config/)
- [OpenCode MCP servers](https://opencode.ai/docs/mcp-servers/)
- [OpenCode agents](https://opencode.ai/docs/agents/)
- [Aider conventions](https://aider.chat/docs/usage/conventions.html)
- [Aider configuration](https://aider.chat/docs/config/aider_conf.html)
- [Windsurf/Cascade memories and rules](https://docs.windsurf.com/windsurf/cascade/memories)
- [Cline rules](https://docs.cline.bot/customization/cline-rules)
- [Continue rules](https://docs.continue.dev/customize/rules)
- [Continue MCP](https://docs.continue.dev/customize/deep-dives/mcp)
- [Gemini CLI GEMINI.md context](https://google-gemini.github.io/gemini-cli/docs/cli/gemini-md.html)

## Related

- [Agent WorkPaper handoff](agent-adoption-kit.md)
- [OpenHands WorkPaper MCP setup](openhands-workpaper-mcp.md)
- [Trae WorkPaper MCP setup](trae-workpaper-mcp.md)
- [OpenCode WorkPaper MCP setup](opencode-workpaper-mcp.md)
- [Aider WorkPaper conventions](aider-workpaper-conventions.md)
- [Agent WorkPaper evaluator matrix](agent-proof-matrix.md)
- [WorkPaper agent handbook](headless-workpaper-agent-handbook.md)
- [Evaluate Bilig as an agent MCP workbook tool](eval-agent-mcp.md)
