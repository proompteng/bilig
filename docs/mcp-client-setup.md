---
title: MCP client setup for Bilig WorkPaper
published: true
description: Remote MCP smoke endpoint and local stdio configuration for Bilig WorkPaper in Claude, Cursor, Junie, Trae, Zed, VS Code, Cline, and Codex.
tags: mcp, claude, cursor, junie, trae, zed, vscode, cline, codex, spreadsheet
canonical_url: https://proompteng.github.io/bilig/mcp-client-setup.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# MCP client setup for Bilig WorkPaper

Use this when you found `io.github.proompteng/bilig-workpaper` in an MCP
directory and want to test the hosted endpoint or wire a local agent client to a
project WorkPaper file.

The hosted endpoint is a stateless Streamable HTTP demo for connector smoke
tests. The local server is the published npm binary from `@bilig/workpaper`; it
starts over stdio, owns a real WorkPaper JSON file, writes through tools,
recalculates formulas, and persists edits back to disk.

For the host-side write/read/persist loop, use the
[headless WorkPaper host handbook](headless-workpaper-agent-handbook.md).

## Smithery install

If your agent host uses Smithery, install the hosted Bilig WorkPaper MCP server
directly:

```sh
npx -y smithery mcp add gkonushev/bilig-workpaper
npx -y smithery tool list bilig-workpaper
npx -y smithery tool call bilig-workpaper list_sheets '{}'
```

The Smithery listing is
<https://smithery.ai/servers/gkonushev/bilig-workpaper>. It points at the same
request-local hosted MCP demo as the remote smoke endpoint below. Use local
stdio or the MCPB bundle when the workflow needs a writable project WorkPaper
file.

## Remote smoke in 30 seconds

Clients that support Streamable HTTP MCP can use the hosted stateless demo
endpoint:

```text
https://bilig.proompteng.ai/mcp
```

Protocol smoke:

```sh
curl -fsS https://bilig.proompteng.ai/mcp \
  -H 'content-type: application/json' \
  -H 'accept: application/json, text/event-stream' \
  -H 'mcp-protocol-version: 2025-11-25' \
  --data '{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2025-11-25","capabilities":{},"clientInfo":{"name":"curl","version":"1"}}}' | jq .
```

Directory and connector scanners can read the hosted same-origin server card:

```sh
curl -fsS https://bilig.proompteng.ai/.well-known/mcp/server-card.json | jq '.transport, (.tools | length)'
```

Use the remote endpoint when the client cannot launch `npm` locally or when you
only need tool discovery and write/readback checks. It is request-local: it does
not persist user files and does not issue `MCP-Session-Id`.

For persistent project workflows, use the local stdio config below with
`--workpaper ./pricing.workpaper.json --init-demo-workpaper --writable`.
If the project already has an XLSX file and the agent only needs triage,
readback, or a risk diagnostic, start with `--from-xlsx ./pricing.xlsx`; it
imports into an in-memory WorkPaper server and does not write a sidecar file.
In that direct mode, edits stay in memory.
Add `--workpaper ./.bilig/pricing.workpaper.json --writable` only when the
workflow needs persisted edits.

## Persistent file-backed stdio server

Every client below starts the same process:

```sh
npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp
npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable
npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp --from-xlsx ./pricing.xlsx
npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp --from-xlsx ./pricing.xlsx --workpaper ./.bilig/pricing.workpaper.json --writable
```

The first command is demo mode. The third command imports an XLSX into an in-memory WorkPaper server for readback, throwaway edits, and
`analyze_workbook_risk` without creating a JSON file. The client configs below use file-backed mode because that is the useful agent setup: the server owns a
real WorkPaper JSON file, initializes it when missing, writes through tools,
recalculates formulas, and persists edits back to the same path.
The persistent `--from-xlsx ... --workpaper` command imports an existing
workbook once; pass `--overwrite-workpaper` only when you intentionally want to
replace the generated WorkPaper JSON.
When started with `--from-xlsx`, `tools/list` also includes
`analyze_workbook_risk`. That tool is fixed to the source XLSX passed at
startup, returns workbook risk indicators before a workflow trusts the imported
WorkPaper, and does not certify Excel compatibility.

Quick protocol smoke test:

```sh
printf '%s\n' \
  '{"jsonrpc":"2.0","id":1,"method":"initialize"}' \
  '{"jsonrpc":"2.0","method":"notifications/initialized"}' \
  '{"jsonrpc":"2.0","id":2,"method":"tools/list"}' |
  npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp
```

`tools/list` should include `read_workpaper_summary` and
`set_workpaper_input_cell` in default demo mode. In file-backed mode,
`tools/list` should include `list_sheets`, `read_range`, `read_cell`,
`set_cell_contents`, `set_cell_contents_and_readback`,
`get_cell_display_value`, `export_workpaper_document`, and `validate_formula`.
When the local server starts with `--from-xlsx`, the same `tools/list` response
also includes `analyze_workbook_risk` for the source workbook diagnostic.
`resources/list` should include
`bilig://workpaper/agent-handoff` and `bilig://workpaper/current-document`.
`prompts/list` should include `edit_and_verify_workpaper` and
`debug_workpaper_formula`. `--init-demo-workpaper` creates the demo JSON file
when it is missing, and `--writable` persists `set_cell_contents` or
`set_cell_contents_and_readback` changes to the same WorkPaper JSON file.

## Open WebUI

Open WebUI can use Bilig in two ways:

- native MCP with the hosted Streamable HTTP endpoint,
  `https://bilig.proompteng.ai/mcp`;
- `mcpo` when Open WebUI should consume the local npm stdio server as an
  OpenAPI tool server.

Use the dedicated [Open WebUI WorkPaper MCP setup](open-webui-workpaper-mcp.md)
for the exact Open WebUI settings, Docker URL boundary, `mcpo` command, and
proof prompt.

## Claude Code

Claude Code can add an MCP server from JSON. Add the server to the current
project:

```sh
claude mcp add-json bilig-workpaper '{
  "type": "stdio",
  "command": "npm",
  "args": ["exec", "--package", "@bilig/workpaper@latest", "--", "bilig-workpaper-mcp", "--workpaper", "./pricing.workpaper.json", "--init-demo-workpaper", "--writable"],
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
      "args": [
        "exec",
        "--package",
        "@bilig/workpaper@latest",
        "--",
        "bilig-workpaper-mcp",
        "--workpaper",
        "./pricing.workpaper.json",
        "--init-demo-workpaper",
        "--writable"
      ],
      "env": {}
    }
  }
}
```

Restart Claude Desktop after editing the config. If the client shows the server
but the tools are missing, run the protocol smoke test above in a terminal first
so you know whether the issue is the client config or the npm server command.

### Claude Desktop MCPB bundle

If you prefer a Claude Desktop bundle, download the released MCPB asset:

```text
https://github.com/proompteng/bilig/releases/latest/download/bilig-workpaper.mcpb
```

The checksum is published beside it:

```text
https://github.com/proompteng/bilig/releases/latest/download/bilig-workpaper.mcpb.sha256
```

You can also reproduce the same MCPB package from this repository:

```sh
pnpm mcpb:workpaper:build
open build/mcpb/bilig-workpaper.mcpb
```

The bundle installs the same published `@bilig/workpaper` stdio server, but
ships the package and its production dependencies inside the `.mcpb` file. See
the [Claude Desktop MCPB guide](claude-desktop-mcpb-workpaper.md) for the
manifest shape and verification prompt.

## Cursor

Bilig checkouts include `.cursor/mcp.json` for project-local Cursor MCP usage.
For another repository, copy this shape:

```json
{
  "mcpServers": {
    "biligWorkpaperFile": {
      "type": "stdio",
      "command": "npm",
      "args": [
        "exec",
        "--package",
        "@bilig/workpaper@latest",
        "--",
        "bilig-workpaper-mcp",
        "--workpaper",
        "./pricing.workpaper.json",
        "--init-demo-workpaper",
        "--writable"
      ],
      "env": {}
    }
  }
}
```

Use a user-level Cursor MCP config when you want the server available across
projects. Use a project-local config when the workbook tooling should be tied
to one repository.

After Cursor starts the server, ask for a concrete readback check:

```text
Use the biligWorkpaperFile MCP server. List sheets, read Summary!A1:B5,
set Inputs!B3 to 0.4 with set_cell_contents_and_readback, then report the
edited cell, before and after values, and whether the WorkPaper JSON persisted.
```

The useful Cursor tool set includes `list_sheets`, `read_range`,
`set_cell_contents_and_readback`, `export_workpaper_document`, and
`validate_formula`.

## JetBrains Junie

Bilig checkouts include `.junie/mcp/mcp.json` for project-local Junie MCP
usage. For another repository, copy this shape:

```json
{
  "mcpServers": {
    "biligWorkpaperFile": {
      "type": "stdio",
      "command": "npm",
      "args": [
        "exec",
        "--yes",
        "--package",
        "@bilig/workpaper@latest",
        "--",
        "bilig-workpaper-mcp",
        "--workpaper",
        "./.bilig/pricing.workpaper.json",
        "--init-demo-workpaper",
        "--writable"
      ],
      "env": {}
    }
  }
}
```

Junie reads project guidelines from `.junie/AGENTS.md` when present and root
`AGENTS.md` otherwise. Keep the WorkPaper proof rule in `AGENTS.md` unless you
need Junie-only memory.

Ask Junie for a concrete readback check:

```text
Use the biligWorkpaperFile MCP server from .junie/mcp/mcp.json. List sheets,
read Summary!A1:B5, set Inputs!B3 to 0.4 with set_cell_contents_and_readback,
then report the edited cell, before and after values, and whether the WorkPaper
JSON persisted.
```

## Trae

Bilig checkouts include `.trae/mcp.json` for project-local Trae MCP usage and
`.trae/rules/bilig-workpaper.md` for the matching project rule. Enable Project
MCP in Trae Settings > MCP, then restart or refresh MCP tools.

For another repository, copy this shape:

```json
{
  "mcpServers": {
    "bilig-workpaper": {
      "command": "npm",
      "args": [
        "exec",
        "--yes",
        "--package",
        "@bilig/workpaper@latest",
        "--",
        "bilig-workpaper-mcp",
        "--workpaper",
        "./.bilig/pricing.workpaper.json",
        "--init-demo-workpaper",
        "--writable"
      ],
      "env": {}
    }
  }
}
```

Trae reads the project rule from `.trae/rules/bilig-workpaper.md`. Keep the
rule focused on WorkPaper readback and persistence proof, and keep the hosted
endpoint only for stateless smoke tests.

Ask Trae for a concrete readback check:

```text
Use the bilig-workpaper MCP server from .trae/mcp.json after Project MCP is
enabled. List sheets, read Inputs!B3 and Summary!B3, set Inputs!B3 to 0.4 with
set_cell_contents_and_readback, export the WorkPaper document, and report
editedCell, before, after, afterRestore, persistedDocumentBytes, verified, and
limitations.
```

For a focused walkthrough, use the
[Trae WorkPaper MCP setup](trae-workpaper-mcp.md).

## Zed

Bilig checkouts include `.zed/settings.json` for project-local Zed MCP usage.
Zed calls these MCP servers `context_servers`. For another repository, copy this
shape:

```json
{
  "context_servers": {
    "bilig-workpaper": {
      "command": {
        "path": "npm",
        "args": [
          "exec",
          "--yes",
          "--package",
          "@bilig/workpaper@latest",
          "--",
          "bilig-workpaper-mcp",
          "--workpaper",
          "./.bilig/pricing.workpaper.json",
          "--init-demo-workpaper",
          "--writable"
        ],
        "env": {}
      },
      "settings": {}
    }
  },
  "agent": {
    "default_profile": "write"
  }
}
```

Zed can use root `AGENTS.md` and `.agents/skills/bilig-workpaper/SKILL.md` as
project context. Keep personal tool approvals in your local Zed settings when
needed; useful permission keys include
`mcp:bilig-workpaper:set_cell_contents_and_readback` and
`mcp:bilig-workpaper:export_workpaper_document`.

Ask Zed for a concrete readback check:

```text
Use the bilig-workpaper context server from .zed/settings.json. List sheets,
read Summary!A1:B5, set Inputs!B3 to 0.4 with
set_cell_contents_and_readback, export the WorkPaper document, and report
editedCell, before, after, persistedDocumentBytes, verified, and limitations.
```

## VS Code

Bilig checkouts include `.vscode/mcp.json` for GitHub Copilot agent mode in VS
Code, plus `.mcp.json` for Claude Code and `mcp/bilig-workpaper.mcp.json` as a
reusable file-backed server entry. Copy the VS Code shape below into another
repository when that project should get the same WorkPaper tools:

```json
{
  "servers": {
    "biligWorkpaperDemo": {
      "type": "http",
      "url": "https://bilig.proompteng.ai/mcp"
    },
    "biligWorkpaperFile": {
      "type": "stdio",
      "command": "npm",
      "args": [
        "exec",
        "--package",
        "@bilig/workpaper@latest",
        "--",
        "bilig-workpaper-mcp",
        "--workpaper",
        "${workspaceFolder}/.bilig/pricing.workpaper.json",
        "--init-demo-workpaper",
        "--writable"
      ]
    }
  }
}
```

Open the Command Palette and run `MCP: List Servers` to start, stop, or inspect
the server. Use `biligWorkpaperDemo` for a no-file hosted smoke test and the
file-backed stdio server when the agent must persist a project WorkPaper JSON
file. VS Code also supports `code --add-mcp` for user-level setup; the
workspace file is easier to review in a repository.

For a user-level VS Code install, pass the same reviewed server shape through
the VS Code CLI:

```sh
code --add-mcp '{"name":"biligWorkpaperFile","type":"stdio","command":"npm","args":["exec","--package","@bilig/workpaper@latest","--","bilig-workpaper-mcp","--workpaper","${workspaceFolder}/.bilig/pricing.workpaper.json","--init-demo-workpaper","--writable"]}'
```

Use a second user-level server only for public no-file smoke tests:

```sh
code --add-mcp '{"name":"biligWorkpaperDemo","type":"http","url":"https://bilig.proompteng.ai/mcp"}'
```

## Cline

Cline can run the published WorkPaper server as a local stdio MCP server. For
the IDE extension, open the MCP Servers icon, choose the Configure tab, click
Configure MCP Servers, and add this entry to the MCP settings JSON under
`mcpServers`:

```json
{
  "mcpServers": {
    "bilig-workpaper": {
      "command": "npm",
      "args": [
        "exec",
        "--package",
        "@bilig/workpaper@latest",
        "--",
        "bilig-workpaper-mcp",
        "--workpaper",
        "./pricing.workpaper.json",
        "--init-demo-workpaper",
        "--writable"
      ],
      "env": {},
      "disabled": false,
      "autoApprove": []
    }
  }
}
```

For Cline CLI, manage the same server with `cline config mcp` or inspect it
non-interactively with `cline config mcp --json`. Current Cline CLI installs use
`~/.cline/mcp.json` for MCP server settings unless you have configured a custom
Cline directory. Then confirm the server is enabled and ask Cline:

```text
List the Bilig WorkPaper tools.
Read Summary!A1:B5, set Inputs!B3 to 0.4, and return the edited cell,
the before/after expected ARR, and the persistence checks.
```

## Codex

This is the shortest Codex spreadsheet MCP server path for formula readback
without driving Excel, Sheets, or a browser grid.

For Codex CLI or the Codex IDE extension, add this to `~/.codex/config.toml`:

```toml
[mcp_servers.bilig-workpaper]
command = "npm"
args = ["exec", "--package", "@bilig/workpaper@latest", "--", "bilig-workpaper-mcp", "--workpaper", "./pricing.workpaper.json", "--init-demo-workpaper", "--writable"]
enabled = true
startup_timeout_sec = 30
```

Then check the configured servers:

```sh
codex mcp list
```

Keep this in your user config unless the whole repository needs the same MCP
server. Do not check personal Codex config into the project.

Ask Codex:

```text
Use the Bilig WorkPaper MCP server from Codex. List sheets, read Summary!A1:B5,
set Inputs!B3 to 0.4 with set_cell_contents_and_readback, export the WorkPaper
document, and report editedCell, before, after, afterRestore,
persistedDocumentBytes, verified, and limitations.
```

## What the tools prove

The write tool changes one workbook input, recalculates dependent formulas,
saves the WorkPaper document, restores it, and returns checks such as
`formulasPersisted`, `restoredMatchesAfter`, and `expectedArrChanged`.

That is the useful boundary for spreadsheet agents. A tool that only says
`updated` is not enough; the agent needs the edited address, previous value,
new value, before/after computed values, and persistence readback.

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
- Trae Model Context Protocol:
  <https://docs.trae.ai/ide/model-context-protocol>
- Trae add MCP servers:
  <https://docs.trae.ai/ide/add-mcp-servers>
- Trae rules:
  <https://docs.trae.ai/ide/rules>
- Trae skills:
  <https://docs.trae.ai/ide/skills>
- Qodo Agentic Tools MCP:
  <https://docs.qodo.ai/qodo-documentation/qodo-ide/tools-mcps/agentic-tools-mcps>
- VS Code MCP configuration:
  <https://code.visualstudio.com/docs/copilot/reference/mcp-configuration>
- Cline MCP configuration:
  <https://docs.cline.bot/mcp/configuring-mcp-servers>
- Zed MCP:
  <https://zed.dev/docs/ai/mcp>
- Zed rules:
  <https://zed.dev/docs/ai/rules>
- Zed tool permissions:
  <https://zed.dev/docs/ai/tool-permissions>
- OpenAI Docs MCP setup for Codex, VS Code, and Cursor:
  <https://platform.openai.com/docs/docs-mcp>

For the server-side tool contract, see the
[MCP spreadsheet tool server guide](mcp-workpaper-tool-server.md).

If the setup works for your agent workflow, keep the repository and release feed
nearby:
<https://github.com/proompteng/bilig>.
