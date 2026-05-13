---
title: Install Bilig WorkPaper in Claude Desktop with MCPB
published: true
description: Build a Claude Desktop MCPB bundle for the published @bilig/headless WorkPaper MCP server and test formula-backed workbook tools locally.
tags: claude, mcpb, mcp, spreadsheet, workbook, agents
canonical_url: https://proompteng.github.io/bilig/claude-desktop-mcpb-workpaper.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Install Bilig WorkPaper in Claude Desktop with MCPB

Use this path when you want a local Claude Desktop bundle instead of editing
`claude_desktop_config.json` by hand. The bundle contains the published
`@bilig/headless` package, runs the WorkPaper MCP stdio server with Node, and
needs no API key.

## Build the bundle

From the repository root:

```sh
pnpm mcpb:workpaper:build
```

The command resolves the latest published `@bilig/headless`, installs its
production dependencies into a local bundle folder, writes a MCPB manifest, and
packs:

```text
build/mcpb/bilig-workpaper.mcpb
```

For a reproducible build, pin the version you intend to ship. This keeps the
guide from baking a stale version into copy-paste setup commands:

```sh
BILIG_HEADLESS_VERSION=$(npm view @bilig/headless version)
pnpm mcpb:workpaper:build -- --package-version "$BILIG_HEADLESS_VERSION"
```

## Install in Claude Desktop

Open the generated file with Claude Desktop:

```sh
open build/mcpb/bilig-workpaper.mcpb
```

Claude should show an install dialog for **Bilig WorkPaper**. After installing,
ask Claude:

```text
List the Bilig WorkPaper tools.
Read the sample WorkPaper summary, set Inputs!B3 to 0.4, and report the
before/after expected ARR plus the persistence checks.
```

The server should expose:

- `read_workpaper_summary`
- `set_workpaper_input_cell`

The write tool changes one input cell, recalculates dependent formulas, saves
the WorkPaper document, restores it, and returns checks such as
`formulasPersisted`, `restoredMatchesAfter`, and `expectedArrChanged`.

## What is inside the bundle

The generated MCPB folder is intentionally small:

```text
build/mcpb/bilig-workpaper/
  manifest.json
  icon.png
  package.json
  README.md
  server/index.js
  node_modules/
```

`server/index.js` imports `runDemoWorkPaperMcpStdioServer` from the packaged
`@bilig/headless` dependency and passes through the bundled package version.
The manifest points Claude Desktop at that launcher with:

```json
{
  "server": {
    "type": "node",
    "entry_point": "server/index.js",
    "mcp_config": {
      "command": "node",
      "args": ["${__dirname}/server/index.js"],
      "env": {}
    }
  }
}
```

If Claude Desktop does not show the tools, run the plain stdio smoke test from
the [MCP client setup guide](mcp-client-setup.md) first. That separates bundle
installation issues from server protocol issues.

## Related links

- [MCP client setup](mcp-client-setup.md)
- [MCP spreadsheet tool server guide](mcp-workpaper-tool-server.md)
- [Official MCP Registry entry](https://registry.modelcontextprotocol.io/v0.1/servers?search=io.github.proompteng%2Fbilig-workpaper)
- [GitHub repository](https://github.com/proompteng/bilig)

If this saves you a custom spreadsheet-tool spike, star the repository so the
next Claude Desktop user can find it:
<https://github.com/proompteng/bilig/stargazers>.
