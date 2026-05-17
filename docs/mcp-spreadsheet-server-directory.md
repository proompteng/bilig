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

For directory scanners that need a containerized start command, the root
Dockerfile exposes a dedicated MCP target without changing the production app
image:

```sh
docker build --target bilig-workpaper-mcp -t bilig-workpaper-mcp:local .
printf '%s\n' \
  '{"jsonrpc":"2.0","id":1,"method":"initialize"}' \
  '{"jsonrpc":"2.0","method":"notifications/initialized"}' \
  '{"jsonrpc":"2.0","id":2,"method":"tools/list"}' |
  docker run --rm -i bilig-workpaper-mcp:local
```

The target installs `@bilig/headless` from npm, seeds
`/workpaper/pricing.workpaper.json`, starts
`bilig-workpaper-mcp --workpaper /workpaper/pricing.workpaper.json --writable`,
and labels the image with
`io.modelcontextprotocol.server.name=io.github.proompteng/bilig-workpaper`.
That makes `tools/list` expose `list_sheets`, `read_range`, `read_cell`,
`set_cell_contents`, `get_cell_display_value`, `export_workpaper_document`, and
`validate_formula` during registry introspection.

For indexers that cannot execute containers, the docs site also serves a static
MCP server card with the same tool catalog:
<https://proompteng.github.io/bilig/.well-known/mcp/server-card.json>.
The same card is mirrored at
<https://proompteng.github.io/bilig/.well-known/mcp.json> and
<https://proompteng.github.io/bilig/.well-known/mcp-server-card.json> for
crawlers that probe those well-known variants.

## Directory Status

| Directory                       | Status                                       | Link                                                                                                  |
| ------------------------------- | -------------------------------------------- | ----------------------------------------------------------------------------------------------------- |
| Official MCP Registry           | Live, current npm version indexed            | <https://registry.modelcontextprotocol.io/v0.1/servers?search=io.github.proompteng%2Fbilig-workpaper> |
| Static MCP server card          | Live                                         | <https://proompteng.github.io/bilig/.well-known/mcp/server-card.json>                                 |
| Static MCP discovery aliases    | Live                                         | <https://proompteng.github.io/bilig/.well-known/mcp.json>                                             |
| Glama                           | Live, installability and tool indexing pending | <https://glama.ai/mcp/servers/proompteng/bilig>                                                       |
| mcp.so                          | Submitted for maintainer review              | <https://github.com/chatmcp/mcpso/issues/2295>                                                        |
| Cline MCP Marketplace           | Submitted for maintainer review              | <https://github.com/cline/mcp-marketplace/issues/1557>                                                |
| mcpserver.cc                    | Submitted for maintainer review              | <https://mcpserver.cc/en?q=bilig>                                                                     |
| AgentNDX                        | Submitted for review                         | <https://agentndx.ai/browse?q=bilig>                                                                  |
| YuzeHao2023 Awesome MCP Servers | Submitted for maintainer review              | <https://github.com/YuzeHao2023/Awesome-MCP-Servers/pull/244>                                         |
| ToolSDK MCP Registry            | Submitted for maintainer review              | <https://github.com/toolsdk-ai/toolsdk-mcp-registry/pull/309>                                         |
| Ever Works MCP data             | Submitted for maintainer review              | <https://github.com/ever-works/awesome-mcp-servers-data/pull/4>                                       |
| mcpserve.com                    | Submitted for maintainer review              | <https://github.com/jmstfv/mcpserve/pull/19>                                                          |
| MCPFind                         | Submitted for maintainer review              | <https://github.com/MCPFind/mcp-find/pull/37>                                                         |
| mctrinh Awesome MCP Servers     | Submitted for maintainer review              | <https://github.com/mctrinh/awesome-mcp-servers/pull/46>                                              |
| MCPRepository                   | Live                                         | <https://mcprepository.com/proompteng/bilig>                                                          |
| PulseMCP                        | Still not indexed in public search on May 17, 2026 | <https://www.pulsemcp.com/servers?search=bilig&q=bilig>                                               |

PulseMCP says server listings are ingested from the official MCP Registry daily
and processed weekly. The Bilig WorkPaper registry entry is live, and the public
Registry API now includes the current `@bilig/headless@0.18.27` package version.
The current entry is marked `isLatest: true`, so treat PulseMCP absence as a
downstream directory refresh lag rather than an upstream registry lag. Starter
issue [#384](https://github.com/proompteng/bilig/issues/384) captured the first
public verification pass and is now closed; reopen a narrower follow-up only if
PulseMCP remains missing after the official registry refresh window.

Glama lists Bilig WorkPaper publicly in search with TypeScript, Developer
Tools, Workplace & Productivity, and Remote attributes. Its public API still
reports `tools: 0`, `package: null`, and no installability. The score page now
shows the concrete blocker: `No Glama release`. The repository already ships
`glama.json` with maintainer `gregkonush`, so the next directory action is to
claim the Glama listing, configure the Dockerfile admin page to use the existing
`bilig-workpaper-mcp` image target, deploy the build test, and publish a Glama
release so the `@bilig/headless` package and file-backed tool surface become
installable in Glama.

The `mcpserver.cc` submission was accepted for review on May 13, 2026 with
submission UUID `bcdce4e1-3b05-4be2-b611-2a2abb8baf79`. Search still returned no
published Bilig result immediately after submission, so treat that directory as
pending until the listing appears.

The AgentNDX submission was accepted for review on May 13, 2026 through the
public submit endpoint with the GitHub repository, homepage, MCP protocol, and
WorkPaper MCP description. AgentNDX search returned `0` Bilig results before
submission, so treat it as pending until the reviewed listing appears.

The YuzeHao2023 Awesome-MCP-Servers pull request was opened on May 13, 2026
with a Development Tools entry for the Bilig WorkPaper MCP server. Treat it as
pending until the maintainer merges the pull request.

The ToolSDK MCP Registry pull request was opened on May 13, 2026 with a
Developer Tools package entry for `@bilig/headless` and the
`bilig-workpaper-mcp` stdio binary. Biome passed on the pull request; the
integration job failed before package validation because the base workflow used
latest pnpm on Node.js 20 and hit `node:sqlite` before reading the Bilig entry.

The Ever Works awesome-mcp-servers-data pull request was opened on May 13, 2026
with source data for the generated mcpserver.works / Awesome MCP Servers
directory. GitHub reports the pull request as mergeable with no repository
checks configured, so treat it as pending maintainer review until it is merged
and appears in the generated directory.

The mcpserve.com pull request was opened on May 13, 2026 with a
`content/servers/bilig-workpaper.md` listing for Bilig WorkPaper. GitHub reports
the pull request as mergeable; Netlify preview checks were pending immediately
after creation, so treat it as submitted until the maintainer merges it and the
listing appears on mcpserve.com.

The mctrinh Awesome MCP Servers pull request was opened on May 17, 2026 as a
single-line `Bilig WorkPaper` addition in the production-ready servers list. The
PR body links the public repo, npm package, MCP docs/transcript, and official
registry name.

MCPRepository search returns a live Bilig page at
<https://mcprepository.com/proompteng/bilig>. The page title is
`bilig - MCP Server`, and its description mirrors the GitHub repository
positioning: formula WorkPaper runtime for Node.js services and agent tools with
cell edits, recalculation, readback, and JSON persistence.

## Verify The Registry Entry

Use the official Registry API when you need a machine-checkable proof:

```sh
curl -fsSL \
  'https://registry.modelcontextprotocol.io/v0.1/servers?search=io.github.proompteng%2Fbilig-workpaper'
```

A useful result includes:

- `name: io.github.proompteng/bilig-workpaper`
- `identifier: @bilig/headless`
- `transport.type: stdio`
- `repository.url: https://github.com/proompteng/bilig`

Latest checked result on May 17, 2026: npm latest is `@bilig/headless@0.18.27`,
and the official Registry API search now includes Bilig WorkPaper entry version
`0.18.27` with `isLatest: true`. The API also returns historical entries, so
consumers should select the latest-marked entry or prefer the npm package
version when they need a single freshest install coordinate. The last documented
refresh attempt was published by the repository workflow run at
<https://github.com/proompteng/bilig/actions/runs/26002693536>.

The package itself carries the matching `mcpName` field. That is the ownership
signal the registry uses for npm package validation.

## Verify The Tool Surface

From a checkout:

```sh
cd examples/headless-workpaper
npm install
npm run agent:mcp-tools
```

For the default packaged stdio demo server:

```sh
printf '%s\n' \
  '{"jsonrpc":"2.0","id":1,"method":"initialize"}' \
  '{"jsonrpc":"2.0","method":"notifications/initialized"}' \
  '{"jsonrpc":"2.0","id":2,"method":"tools/list"}' |
  npm run --silent agent:mcp-stdio
```

The demo tool list includes:

- `read_workpaper_summary`
- `set_workpaper_input_cell`

The write path edits one input cell, recalculates dependent formulas, and
returns structured readback. That is the behavior a directory listing should
describe; Bilig is not a generic spreadsheet screenshot tool.

For directory scanners and production client configs, prefer file-backed mode:

```sh
printf '%s\n' \
  '{"jsonrpc":"2.0","id":1,"method":"initialize"}' \
  '{"jsonrpc":"2.0","method":"notifications/initialized"}' \
  '{"jsonrpc":"2.0","id":2,"method":"tools/list"}' |
  npm exec --package @bilig/headless -- \
    bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --writable
```

That mode exposes `list_sheets`, `read_range`, `read_cell`,
`set_cell_contents`, `get_cell_display_value`, `export_workpaper_document`, and
`validate_formula`.

## Short Listing Copy

Bilig WorkPaper is a local stdio MCP server for formula-backed workbook
automation. It lets agents and Node services read workbook summaries, edit input
cells, recalculate formulas, verify readback, and persist WorkPaper JSON through
the published `@bilig/headless` package.

```sh
npm exec --package @bilig/headless -- bilig-workpaper-mcp
```
