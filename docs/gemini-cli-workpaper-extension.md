---
title: Gemini CLI WorkPaper extension
published: true
description: Install Bilig as a Gemini CLI extension so Gemini can use WorkPaper MCP tools for formula readback without spreadsheet UI automation.
tags: gemini-cli, mcp, spreadsheet, workpaper, agent-tools
canonical_url: https://proompteng.github.io/bilig/gemini-cli-workpaper-extension.html
image: /assets/github-social-preview.png
---

# Gemini CLI WorkPaper Extension

Gemini CLI extensions can load MCP servers from a repository manifest. Bilig
ships a root `gemini-extension.json` that starts the WorkPaper MCP server with
`@bilig/workpaper@latest`.

The manifest version tracks the published `@bilig/workpaper` release so gallery
crawlers show current package metadata. The MCP command intentionally stays on
`@latest`, because installed extensions should pick up the current WorkPaper
server without editing local config after each Bilig release.

Use this when Gemini needs spreadsheet formulas as a tool contract instead of a
spreadsheet UI session:

- edit an input cell;
- recalculate formulas;
- read the computed value back;
- persist the WorkPaper JSON;
- return proof instead of trusting a write call.

Official Gemini CLI references:

- <https://github.com/google-gemini/gemini-cli/blob/main/docs/extensions/writing-extensions.md>
- <https://github.com/google-gemini/gemini-cli/blob/main/docs/extensions/reference.md>
- <https://github.com/google-gemini/gemini-cli/blob/main/docs/extensions/releasing.md>

## Install

```sh
gemini extensions install https://github.com/proompteng/bilig --ref main
```

Restart Gemini CLI after installing the extension. The manifest starts this MCP
server:

```json
{
  "name": "bilig-workpaper",
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
        "${extensionPath}${/}pricing.workpaper.json",
        "--init-demo-workpaper",
        "--writable"
      ]
    }
  }
}
```

The default workbook path lives inside the installed extension copy. Gemini can
edit it safely for a local smoke test without credentials.

## Ask Gemini

After restart, ask Gemini for a proof-shaped workbook edit:

```text
Use the Bilig WorkPaper tools. List sheets, read Inputs!B3, set Inputs!B3 to =0.4, read the recalculated output, and tell me whether the WorkPaper JSON persisted.
```

Useful answers should include the edited sheet and address, the before and after
cell contents, the dependent output value, and whether the final readback was
verified.

## Discovery

Gemini CLI's extension gallery indexes public GitHub repositories with the
`gemini-cli-extension` topic and a root `gemini-extension.json`. Bilig keeps the
manifest at the repository root so the gallery crawler can validate it without a
separate submission issue.

Bilig also checks the extension manifest in CI. If `@bilig/workpaper` is released
and the manifest version is not updated with it, `pnpm docs:discovery:check`
fails before the stale metadata reaches the public gallery.

## Boundary

This extension exposes Bilig WorkPaper MCP tools to Gemini CLI. It does not
claim desktop Excel macro support, Google Sheets account mutation, external link
refresh, or compatibility with every XLSX feature. Use it for service-owned
formula workbooks where JSON persistence and read-after-write proof matter.

Manifest:
[`gemini-extension.json`](https://github.com/proompteng/bilig/blob/main/gemini-extension.json).

Context:
[`gemini-workpaper-context.md`](https://github.com/proompteng/bilig/blob/main/gemini-workpaper-context.md).
