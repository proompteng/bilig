---
title: MCP client setup for Bilig WorkPaper
published: true
description: Copy-paste MCP client configuration for running the published @bilig/headless WorkPaper stdio server from Claude, Cursor, VS Code, Cline, and Codex.
tags: mcp, claude, cursor, vscode, cline, codex, spreadsheet
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

For the agent-side write/read/persist loop, use the
[headless WorkPaper agent handbook](headless-workpaper-agent-handbook.md).

## Server command

Every client below starts the same process:

```sh
npm exec --package @bilig/headless -- bilig-workpaper-mcp
npm exec --package @bilig/headless -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable
```

The first command is demo mode. The client configs below use file-backed mode
because that is the useful agent setup: the server owns a real WorkPaper JSON
file, initializes it when missing, writes through tools, recalculates formulas,
and persists edits back to the same path.

Quick protocol smoke test:

```sh
printf '%s\n' \
  '{"jsonrpc":"2.0","id":1,"method":"initialize"}' \
  '{"jsonrpc":"2.0","method":"notifications/initialized"}' \
  '{"jsonrpc":"2.0","id":2,"method":"tools/list"}' |
  npm exec --package @bilig/headless -- bilig-workpaper-mcp
```

`tools/list` should include `read_workpaper_summary` and
`set_workpaper_input_cell` in default demo mode. In file-backed mode,
`tools/list` should include `list_sheets`, `read_range`, `read_cell`,
`set_cell_contents`, `get_cell_display_value`, `export_workpaper_document`, and
`validate_formula`; `--init-demo-workpaper` creates the demo JSON file when it
is missing, and `--writable` persists `set_cell_contents` changes to the same
WorkPaper JSON file.

## Claude Code

Claude Code can add an MCP server from JSON. Add the server to the current
project:

```sh
claude mcp add-json bilig-workpaper '{
  "type": "stdio",
  "command": "npm",
  "args": ["exec", "--package", "@bilig/headless", "--", "bilig-workpaper-mcp", "--workpaper", "./pricing.workpaper.json", "--init-demo-workpaper", "--writable"],
  "env": {}
}' --scope project
```

Then check it:

```sh
claude mcp get bilig-workpaper
```

Ask Claude:

```text
List the Bilig WorkPaper tools.
Then read the sample WorkPaper summary, set the input cell that controls
conversion rate to 0.4, and report the before/after expected ARR plus the
persistence checks.
```

## Claude Desktop

Add the same stdio server to `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "bilig-workpaper": {
      "type": "stdio",
      "command": "npm",
      "args": ["exec", "--package", "@bilig/headless", "--", "bilig-workpaper-mcp", "--workpaper", "./pricing.workpaper.json", "--init-demo-workpaper", "--writable"],
      "env": {}
    }
  }
}
```

Restart Claude Desktop after editing the config. If the client shows the server
but the tools are missing, run the protocol smoke test above in a terminal first
so you know whether the issue is the client config or the npm server command.

### Claude Desktop MCPB bundle

If you prefer a local Claude Desktop bundle, build the MCPB package from this
repository:

```sh
pnpm mcpb:workpaper:build
open build/mcpb/bilig-workpaper.mcpb
```

The bundle installs the same published `@bilig/headless` stdio server, but
ships the package and its production dependencies inside the `.mcpb` file. See
the [Claude Desktop MCPB guide](claude-desktop-mcpb-workpaper.md) for the
manifest shape and verification prompt.

## Cursor

For a project-local setup, create `.cursor/mcp.json`:

```json
{
  "mcpServers": {
    "bilig-workpaper": {
      "type": "stdio",
      "command": "npm",
      "args": ["exec", "--package", "@bilig/headless", "--", "bilig-workpaper-mcp", "--workpaper", "./pricing.workpaper.json", "--init-demo-workpaper", "--writable"],
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
      "args": ["exec", "--package", "@bilig/headless", "--", "bilig-workpaper-mcp", "--workpaper", "./pricing.workpaper.json", "--init-demo-workpaper", "--writable"]
    }
  }
}
```

Open the Command Palette and run `MCP: List Servers` to start, stop, or inspect
the server. VS Code also supports `code --add-mcp` for user-level setup; the
workspace file is easier to review in a repository.

## Cline

Cline can run the published WorkPaper server as a local stdio MCP server. For
the IDE extension, open the MCP Servers icon, choose the Configure tab, click
Configure MCP Servers, and add this entry to `cline_mcp_settings.json` under
`mcpServers`:

```json
{
  "mcpServers": {
    "bilig-workpaper": {
      "command": "npm",
      "args": ["exec", "--package", "@bilig/headless", "--", "bilig-workpaper-mcp", "--workpaper", "./pricing.workpaper.json", "--init-demo-workpaper", "--writable"],
      "env": {},
      "disabled": false
    }
  }
}
```

For Cline CLI, put the same `mcpServers` object in
`~/.cline/data/settings/cline_mcp_settings.json`. If you use `CLINE_DIR` or a
custom config path, edit that config directory's
`data/settings/cline_mcp_settings.json` instead. Then confirm the server is
enabled and ask Cline:

```text
List the Bilig WorkPaper tools.
Read Summary!A1:B5, set Inputs!B3 to 0.4, and return the edited cell,
the before/after expected ARR, and the persistence checks.
```

## Codex

For Codex CLI or the Codex IDE extension, add this to `~/.codex/config.toml`:

```toml
[mcp_servers.bilig-workpaper]
command = "npm"
args = ["exec", "--package", "@bilig/headless", "--", "bilig-workpaper-mcp", "--workpaper", "./pricing.workpaper.json", "--init-demo-workpaper", "--writable"]
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
- Cline MCP configuration:
  <https://docs.cline.bot/mcp/adding-and-configuring-servers>
- OpenAI Docs MCP setup for Codex, VS Code, and Cursor:
  <https://platform.openai.com/docs/docs-mcp>

For the server-side tool contract, see the
[MCP spreadsheet tool server guide](mcp-workpaper-tool-server.md).

If the setup works for your agent workflow, star the repository so the next
person searching for MCP spreadsheet tools can find it:
<https://github.com/proompteng/bilig/stargazers>.
