---
title: MCP client setup for Bilig WorkPaper
published: true
description: Copy-paste MCP client configuration for running the published @bilig/headless WorkPaper stdio server from Claude, Cursor, VS Code, and Codex.
tags: mcp, claude, cursor, vscode, codex, spreadsheet
canonical_url: https://proompteng.github.io/bilig/mcp-client-setup.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# MCP client setup for Bilig WorkPaper

Use this when you found `io.github.proompteng/bilig-workpaper` in an MCP
directory and want to run it from a local agent client.

The server is the published npm binary from `@bilig/headless`. It starts over
stdio, exposes WorkPaper tools, and returns computed workbook readback after a
write.

## Server command

Every client below starts the same process:

```sh
npm exec --package @bilig/headless -- bilig-workpaper-mcp
```

Quick protocol smoke test:

```sh
printf '%s\n' \
  '{"jsonrpc":"2.0","id":1,"method":"initialize"}' \
  '{"jsonrpc":"2.0","method":"notifications/initialized"}' \
  '{"jsonrpc":"2.0","id":2,"method":"tools/list"}' |
  npm exec --package @bilig/headless -- bilig-workpaper-mcp
```

`tools/list` should include `read_workpaper_summary` and
`set_workpaper_input_cell`.

## Claude Code

Claude Code can add an MCP server from JSON. Add the server to the current
project:

```sh
claude mcp add-json bilig-workpaper '{"type":"stdio","command":"npm","args":["exec","--package","@bilig/headless","--","bilig-workpaper-mcp"],"env":{}}' --scope project
```

Then check it:

```sh
claude mcp get bilig-workpaper
```

Ask Claude:

```text
List the Bilig WorkPaper tools. Then read the sample WorkPaper summary, set the input cell that controls conversion rate to 0.4, and report the before/after expected ARR plus the persistence checks.
```

## Claude Desktop

Add the same stdio server to `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "bilig-workpaper": {
      "type": "stdio",
      "command": "npm",
      "args": ["exec", "--package", "@bilig/headless", "--", "bilig-workpaper-mcp"],
      "env": {}
    }
  }
}
```

Restart Claude Desktop after editing the config. If the client shows the server
but the tools are missing, run the protocol smoke test above in a terminal first
so you know whether the issue is the client config or the npm server command.

## Cursor

For a project-local setup, create `.cursor/mcp.json`:

```json
{
  "mcpServers": {
    "bilig-workpaper": {
      "type": "stdio",
      "command": "npm",
      "args": ["exec", "--package", "@bilig/headless", "--", "bilig-workpaper-mcp"],
      "env": {}
    }
  }
}
```

Use a user-level Cursor MCP config when you want the server available across
projects. Use a project-local config when the workbook tooling should be tied
to one repository.

## VS Code

For GitHub Copilot agent mode in VS Code, add `.vscode/mcp.json`:

```json
{
  "servers": {
    "bilig-workpaper": {
      "type": "stdio",
      "command": "npm",
      "args": ["exec", "--package", "@bilig/headless", "--", "bilig-workpaper-mcp"]
    }
  }
}
```

Open the Command Palette and run `MCP: List Servers` to start, stop, or inspect
the server. VS Code also supports `code --add-mcp` for user-level setup; the
workspace file is easier to review in a repository.

## Codex

For Codex CLI or the Codex IDE extension, add this to `~/.codex/config.toml`:

```toml
[mcp_servers.bilig-workpaper]
command = "npm"
args = ["exec", "--package", "@bilig/headless", "--", "bilig-workpaper-mcp"]
enabled = true
startup_timeout_sec = 30
```

Then check the configured servers:

```sh
codex mcp list
```

Keep this in your user config unless the whole repository needs the same MCP
server. Do not check personal Codex config into the project.

## What the tools prove

The write tool changes one workbook input, recalculates dependent formulas,
saves the WorkPaper document, restores it, and returns checks such as
`formulasPersisted`, `restoredMatchesAfter`, and `expectedArrChanged`.

That is the useful boundary for spreadsheet agents. A tool that only says
`updated` is not enough; the agent needs the edited address, previous value,
new value, before/after computed values, and persistence proof.

## Troubleshooting

| Symptom                    | Check                                                                                     |
| -------------------------- | ----------------------------------------------------------------------------------------- |
| The server never starts    | Run the smoke test in a terminal and confirm `npm` is on your PATH.                       |
| Tools do not appear        | Restart the MCP client after changing config, then reset or refresh cached MCP tools.     |
| `spawn npm ENOENT` appears | Use the absolute path to `npm`, for example the output of `which npm`.                    |
| The client parses nothing  | Make sure the command is `npm` and the package flags are in `args`, not one shell string. |
| A write seems too vague    | Ask for `editedCell`, `before`, `after`, and `checks` in the tool result.                 |

## Client References

- Claude Code MCP configuration:
  <https://code.claude.com/docs/en/mcp>
- Cursor MCP configuration:
  <https://docs.cursor.com/advanced/model-context-protocol>
- VS Code MCP configuration:
  <https://code.visualstudio.com/docs/copilot/reference/mcp-configuration>
- OpenAI Docs MCP setup for Codex, VS Code, and Cursor:
  <https://platform.openai.com/docs/docs-mcp>

For the server-side tool contract, see the
[MCP spreadsheet tool server guide](mcp-workpaper-tool-server.md).

If the setup works for your agent workflow, star the repository so the next
person searching for MCP spreadsheet tools can find it:
<https://github.com/proompteng/bilig/stargazers>.
