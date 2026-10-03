---
title: LobeHub WorkPaper MCP setup
published: true
description: Connect Bilig WorkPaper to LobeHub Custom MCP so agents can edit workbook inputs, verify formula readback, and persist JSON proof.
tags: lobehub, mcp, spreadsheet automation, workbook-api, custom-mcp
canonical_url: https://proompteng.github.io/bilig/lobehub-workpaper-mcp.html
image: /assets/github-social-preview.png
---

# LobeHub WorkPaper MCP setup

Use this when a LobeHub agent needs spreadsheet-shaped business logic, but the
formula truth should live in a WorkPaper API instead of Excel UI automation,
browser grid clicks, or unchecked workbook readback.

LobeHub's Custom MCP flow supports:

- **Streamable HTTP** for remote MCP servers available over HTTPS.
- **STDIO** for local desktop MCP servers. LobeHub documents STDIO as desktop
  only, not web.
- **Import JSON config**, which is the fastest way to add a custom MCP server.

Official LobeHub reference:

- <https://lobehub.com/docs/usage/community/custom-mcp>

## Fastest smoke test: hosted Streamable HTTP

In LobeHub, open **Settings -> Skills -> Skill Store -> Custom -> Add custom
skill**, then choose **Import JSON config** and paste:

```json
{
  "mcpServers": {
    "bilig-workpaper": {
      "url": "https://bilig.proompteng.ai/mcp",
      "type": "http"
    }
  }
}
```

Click **Import**, review the generated settings, then click **Test connection**.
The hosted endpoint exposes these tools:

- `list_sheets`
- `read_range`
- `read_cell`
- `set_cell_contents`
- `set_cell_contents_and_readback`
- `get_cell_display_value`
- `export_workpaper_document`
- `validate_formula`

Use this prompt after enabling the custom MCP on an agent:

```text
List the available Bilig WorkPaper tools. Then read the sample sheets, set the
conversion-rate input to 0.4, read the recalculated ARR output, export the
WorkPaper document, and return a compact proof object.
```

The hosted endpoint is stateless and request-local. It proves that LobeHub can
discover and call Bilig WorkPaper tools. It does not persist a private project
file.

## Persistent desktop WorkPaper: STDIO

Use this in the LobeHub desktop app when the WorkPaper JSON file should live on
your machine and survive across turns.

Import this JSON:

```json
{
  "mcpServers": {
    "bilig-workpaper-local": {
      "command": "npx",
      "args": [
        "-y",
        "--package",
        "@bilig/workpaper@latest",
        "bilig-workpaper-mcp",
        "--workpaper",
        "./pricing.workpaper.json",
        "--init-demo-workpaper",
        "--writable"
      ],
      "type": "stdio"
    }
  }
}
```

`--init-demo-workpaper` creates the demo file only when it is missing.
`--writable` lets tool writes persist back to the same WorkPaper JSON file.

Before adding it to LobeHub, you can prove the same local MCP contract from a
terminal:

```sh
npx -y --package @bilig/workpaper@latest bilig-mcp-challenge --json
```

Expected proof fields include:

```json
{
  "transport": "stdio-json-rpc",
  "tools": [
    "list_sheets",
    "read_range",
    "read_cell",
    "set_cell_contents",
    "get_cell_display_value",
    "export_workpaper_document",
    "validate_formula"
  ],
  "editedCell": "Inputs!B3",
  "dependentCell": "Summary!B3",
  "before": 60000,
  "after": 96000,
  "afterRestore": 96000,
  "afterRestart": 96000,
  "persistedDocumentBytes": 1162,
  "verified": true
}
```

`verified` should only be true after the dependent formula output is read back
and the persisted document can be restored.

## Which path to use

| Path                   | Use when                                           | LobeHub surface           |
| ---------------------- | -------------------------------------------------- | ------------------------- |
| Hosted Streamable HTTP | You need a quick remote tool-discovery smoke test. | Web or desktop Custom MCP |
| Local STDIO            | You need a private writable WorkPaper JSON file.   | Desktop Custom MCP        |

Start with hosted HTTP when you only need to verify tool calling. Use local
STDIO when the agent should own durable workbook state on your machine.

## Troubleshooting

- If **Test connection** fails for the hosted endpoint, confirm the URL is
  `https://bilig.proompteng.ai/mcp`, type is `http`, and auth is empty.
- If STDIO fails, run `which npx` and the `bilig-mcp-challenge` command in a
  terminal first. LobeHub needs to find the same executable from its desktop
  process environment.
- If tools are installed but not used, enable the custom MCP on the specific
  LobeHub agent before asking for the proof.
- If you need Excel desktop parity, macros, pivots, charts, or external links,
  keep Excel or a workbook oracle in the loop and treat this as a WorkPaper API
  proof, not a desktop Excel proof.

## Related Bilig docs

- [Agent MCP workbook evaluator](eval-agent-mcp.md)
- [MCP client setup](mcp-client-setup.md)
- [MCP WorkPaper tool server](mcp-workpaper-tool-server.md)
- [Agent framework workbook tools](agent-framework-workbook-tools.md)
- [WorkPaper agent handbook](headless-workpaper-agent-handbook.md)
