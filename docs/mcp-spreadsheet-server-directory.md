---
title: MCP spreadsheet server directory status
published: true
description: Live directory and install status for the Bilig WorkPaper MCP server, including official Registry metadata, Glama indexing, npm stdio install, and PulseMCP follow-up.
tags: mcp, model context protocol, spreadsheet, agents, directory
canonical_url: https://proompteng.github.io/bilig/mcp-spreadsheet-server-directory.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# MCP Spreadsheet Server Directory Status

Bilig WorkPaper is the MCP server for `@bilig/headless`. It gives coding agents
and local MCP clients a formula-backed workbook surface: read a summary, change
an input cell, recalculate formulas, and return the before/after values instead
of a screenshot.

Use this page when you are checking whether a directory listing points to the
real package, or when you want the shortest install command for Claude Desktop,
Cursor, VS Code, Codex, or another stdio MCP client.

## Canonical Package

```sh
npm exec --package @bilig/headless -- bilig-workpaper-mcp
```

Package metadata:

- npm package: <https://www.npmjs.com/package/@bilig/headless>
- GitHub repository: <https://github.com/proompteng/bilig>
- MCP name: `io.github.proompteng/bilig-workpaper`
- Packaged metadata:
  <https://github.com/proompteng/bilig/blob/main/packages/headless/server.json>
- Client setup guide:
  <https://proompteng.github.io/bilig/mcp-client-setup.html>
- Claude Desktop MCPB bundle:
  <https://proompteng.github.io/bilig/claude-desktop-mcpb-workpaper.html>

The server is local-first stdio. It does not need a hosted Bilig account or a
network service to answer `tools/list` and `tools/call`.

## Directory Status

| Directory             | Status                             | Link                                                                                                  |
| --------------------- | ---------------------------------- | ----------------------------------------------------------------------------------------------------- |
| Official MCP Registry | Live                               | <https://registry.modelcontextprotocol.io/v0.1/servers?search=io.github.proompteng%2Fbilig-workpaper> |
| Glama                 | Live                               | <https://glama.ai/mcp/servers/proompteng/bilig>                                                       |
| mcp.so                | Submitted for maintainer review    | <https://github.com/chatmcp/mcpso/issues/2295>                                                        |
| Cline MCP Marketplace | Submitted for maintainer review    | <https://github.com/cline/mcp-marketplace/issues/1557>                                                |
| PulseMCP              | Not indexed yet as of May 13, 2026 | <https://www.pulsemcp.com/servers?search=bilig&q=bilig>                                               |

PulseMCP says server listings are ingested from the official MCP Registry daily
and processed weekly. The Bilig WorkPaper registry entry is already live, so the
right next step is to let that import run, then contact PulseMCP only if the
server still does not appear after their stated review window.

## Verify The Registry Entry

Use the official Registry API when you need a machine-checkable proof:

```sh
curl -fsSL \
  'https://registry.modelcontextprotocol.io/v0/servers?search=io.github.proompteng%2Fbilig-workpaper'
```

A useful result includes:

- `name: io.github.proompteng/bilig-workpaper`
- `identifier: @bilig/headless`
- `transport.type: stdio`
- `repository.url: https://github.com/proompteng/bilig`

The package itself carries the matching `mcpName` field. That is the ownership
signal the registry uses for npm package validation.

## Verify The Tool Surface

From a checkout:

```sh
cd examples/headless-workpaper
npm install
npm run agent:mcp-tools
```

For the packaged stdio server:

```sh
printf '%s\n' \
  '{"jsonrpc":"2.0","id":1,"method":"initialize"}' \
  '{"jsonrpc":"2.0","method":"notifications/initialized"}' \
  '{"jsonrpc":"2.0","id":2,"method":"tools/list"}' |
  npm run --silent agent:mcp-stdio
```

The tool list includes:

- `read_workpaper_summary`
- `set_workpaper_input_cell`

The write path edits one input cell, recalculates dependent formulas, and
returns structured readback. That is the behavior a directory listing should
describe; Bilig is not a generic spreadsheet screenshot tool.

## Short Listing Copy

Bilig WorkPaper is a local stdio MCP server for formula-backed workbook
automation. It lets agents and Node services read workbook summaries, edit input
cells, recalculate formulas, verify readback, and persist WorkPaper JSON through
the published `@bilig/headless` package.

```sh
npm exec --package @bilig/headless -- bilig-workpaper-mcp
```
