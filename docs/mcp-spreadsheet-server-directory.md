---
title: MCP spreadsheet server directory status
published: true
description: Live directory and install status for the Bilig WorkPaper MCP server, including official Registry metadata, Smithery install, Glama indexing, npm stdio install, and PulseMCP follow-up.
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
npm exec --package @bilig/headless@0.40.21 -- bilig-workpaper-mcp
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
- Smithery install:
  <https://smithery.ai/servers/gkonushev/bilig-workpaper>

The server is local-first stdio. It does not need a hosted Bilig account or a
network service to answer `tools/list` and `tools/call`.

For directories and clients that prefer hosted MCP, the app runtime exposes a
stateless Streamable HTTP endpoint:

```text
https://bilig.proompteng.ai/mcp
```

That endpoint is for discovery, connector smoke tests, and agent onboarding. It
does not persist user workbooks or issue an MCP session id. The local stdio
server remains the recommended path when an agent needs to own a project
WorkPaper JSON file.

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
`bilig-workpaper-mcp --workpaper /workpaper/pricing.workpaper.json --init-demo-workpaper --writable`,
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

The hosted MCP origin serves remote transport metadata at
<https://bilig.proompteng.ai/.well-known/mcp/server-card.json>. Use that URL
for directories that start from `https://bilig.proompteng.ai/mcp` and expect
same-origin static server-card discovery.

## Directory Status

| Directory                       | Status                                                                        | Link                                                                                                            |
| ------------------------------- | ----------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------- |
| Official MCP Registry           | Live but latest marker lags npm; `0.27.0` is latest-marked while npm latest is `0.40.21` | <https://registry.modelcontextprotocol.io/v0.1/servers?search=io.github.proompteng%2Fbilig-workpaper&limit=100> |
| Hosted Streamable HTTP endpoint | App runtime endpoint for JSON-only stateless MCP smoke tests                  | <https://bilig.proompteng.ai/mcp>                                                                               |
| Hosted MCP server card          | Same-origin server card for Streamable HTTP scanners                          | <https://bilig.proompteng.ai/.well-known/mcp/server-card.json>                                                  |
| Static MCP server card          | Live                                                                          | <https://proompteng.github.io/bilig/.well-known/mcp/server-card.json>                                           |
| Static MCP discovery aliases    | Live                                                                          | <https://proompteng.github.io/bilig/.well-known/mcp.json>                                                       |
| Smithery                        | Live; `smithery mcp add` smoke connected and listed demo workbook sheets      | <https://smithery.ai/servers/gkonushev/bilig-workpaper>                                                         |
| Glama                           | Live with `Try in Browser`; seven tools indexed with A-grade TDQS             | <https://glama.ai/mcp/servers/proompteng/bilig>                                                                 |
| Docker MCP Registry             | Submitted for maintainer review; source commit and readme refreshed on May 19 | <https://github.com/docker/mcp-registry/pull/3606>                                                              |
| Goose MCP catalog               | Closed by maintainer while Goose pauses new MCP server additions              | <https://github.com/aaif-goose/goose/pull/9315>                                                                 |
| mcp.so                          | Submitted for maintainer review; issue body refreshed on May 19               | <https://github.com/chatmcp/mcpso/issues/2295>                                                                  |
| Cline MCP Marketplace           | Submitted for maintainer review; issue body refreshed on May 19               | <https://github.com/cline/mcp-marketplace/issues/1557>                                                          |
| mcpserver.cc                    | Submitted for maintainer review                                               | <https://mcpserver.cc/en?q=bilig>                                                                               |
| AgentNDX                        | Submitted for review                                                          | <https://agentndx.ai/browse?q=bilig>                                                                            |
| YuzeHao2023 Awesome MCP Servers | Submitted for maintainer review                                               | <https://github.com/YuzeHao2023/Awesome-MCP-Servers/pull/244>                                                   |
| ToolSDK MCP Registry            | Submitted for maintainer review                                               | <https://github.com/toolsdk-ai/toolsdk-mcp-registry/pull/309>                                                   |
| Ever Works MCP data             | Submitted for maintainer review                                               | <https://github.com/ever-works/awesome-mcp-servers-data/pull/4>                                                 |
| mcpserve.com                    | Submitted for maintainer review                                               | <https://github.com/jmstfv/mcpserve/pull/19>                                                                    |
| MCPFind                         | Submitted for maintainer review                                               | <https://github.com/MCPFind/mcp-find/pull/37>                                                                   |
| mctrinh Awesome MCP Servers     | Submitted for maintainer review                                               | <https://github.com/mctrinh/awesome-mcp-servers/pull/46>                                                        |
| MCPRepository                   | Live                                                                          | <https://mcprepository.com/proompteng/bilig>                                                                    |
| PulseMCP                        | Live in PulseMCP-backed lookup as `Bilig WorkPaper`                           | <https://www.pulsemcp.com/servers?search=bilig&q=bilig>                                                         |

PulseMCP says server listings are ingested from the official MCP Registry daily
and processed weekly. Live verification on May 19, 2026 found Bilig WorkPaper
through a PulseMCP-backed lookup query for `bilig`, so do not resubmit the same
server there. Starter issue
[#384](https://github.com/proompteng/bilig/issues/384) captured the first public
verification pass and should stay closed unless the live listing regresses.

The Docker MCP Registry pull request was refreshed on May 19, 2026 by updating
the existing PR body, not by opening a duplicate submission. Docker still points
at Bilig source commit `a1ecdd52cda3d54e0254afce129a9012c5027826`; the PR body
now points reviewers at `@bilig/headless@0.40.21` and `libraries-v0.40.21`.

The Goose MCP catalog pull request was closed on May 19, 2026 because Goose is
paused on adding new MCP servers while it works on a more scalable extensions
system. Do not resubmit there until maintainers reopen that path.

The mcp.so and Cline MCP Marketplace submissions were refreshed on May 19, 2026
by editing the existing issue bodies, not by adding more comments. Both now
point reviewers at the file-backed command with `--init-demo-workpaper
--writable`, the seven current tools, the official Registry entry, and the
static MCP server card.

Smithery lists Bilig WorkPaper as `gkonushev/bilig-workpaper` and exposes a
one-command install path:

```sh
npx -y smithery mcp add gkonushev/bilig-workpaper
npx -y smithery tool call bilig-workpaper list_sheets '{}'
```

Live verification on May 19, 2026 returned a connected Smithery server named
`bilig-workpaper-remote-demo`, version `0.25.4`, and `list_sheets` returned the
demo `Inputs` and `Summary` sheets. Smithery's generated remote endpoint is
<https://bilig-workpaper--gkonushev.run.tools>; use the Smithery page as the
stable install URL.

Glama lists Bilig WorkPaper publicly with TypeScript, Developer Tools,
Workplace & Productivity, Remote attributes, `Try in Browser`, and public pages
for the seven file-backed tools: `list_sheets`, `read_range`, `read_cell`,
`set_cell_contents`, `get_cell_display_value`, `export_workpaper_document`, and
`validate_formula`. The public score surface reports A-grade Tool Definition Quality
for the indexed tools. Use npm and the official Registry for the latest
install coordinate because Glama's source crawl, hosted smoke build, and JSON
API can refresh on different cadences.

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

Latest checked result on May 19, 2026: Live but latest marker lags npm; `0.27.0` is latest-marked while npm latest is `0.40.21`. npm latest is `@bilig/headless@0.40.21`,
and the official Registry latest-marked entry is
`io.github.proompteng/bilig-workpaper@0.27.0` with package
`@bilig/headless@0.27.0`, so it does not yet match npm latest `@bilig/headless@0.40.21`. The API also returns historical entries, so
consumers should follow pagination, request a sufficient limit, and select the
latest-marked entry when they need the Registry-owned freshest install
coordinate. The hosted server-card path still advertises remote `https://bilig.proompteng.ai/mcp` for live smoke tests.

The package itself carries the matching `mcpName` field. That is the ownership
signal the registry uses for npm package validation.

## Verify The Tool Surface

From a checkout:

```sh
pnpm --dir examples/headless-workpaper install --ignore-workspace
pnpm --dir examples/headless-workpaper run agent:mcp-tools
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
  npm exec --package @bilig/headless@0.40.21 -- \
    bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable
```

That mode exposes `list_sheets`, `read_range`, `read_cell`,
`set_cell_contents`, `get_cell_display_value`, `export_workpaper_document`, and
`validate_formula`.

For remote Streamable HTTP smoke:

```sh
curl -fsS https://bilig.proompteng.ai/mcp \
  -H 'content-type: application/json' \
  -H 'accept: application/json, text/event-stream' \
  -H 'mcp-protocol-version: 2025-11-25' \
  --data '{"jsonrpc":"2.0","id":1,"method":"tools/list"}' | jq '.result.tools[].name'
```

## Short Listing Copy

Bilig WorkPaper is an MCP server for formula-backed workbook automation. It has
a local file-backed stdio server for project WorkPaper JSON files and a hosted
stateless Streamable HTTP endpoint for connector smoke tests. Agents and Node
services can read workbook summaries, edit input cells, recalculate formulas,
verify readback, and persist WorkPaper JSON through the published
`@bilig/headless` package.

```sh
npm exec --package @bilig/headless@0.40.21 -- bilig-workpaper-mcp
```
