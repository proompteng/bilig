---
title: AnythingLLM WorkPaper MCP setup
published: true
description: Connect Bilig WorkPaper to AnythingLLM Agent Skills so agents can edit workbook inputs, verify formula readback, and persist JSON proof.
tags: anythingllm, mcp, spreadsheet automation, workbook-api, agent-skills
canonical_url: https://proompteng.github.io/bilig/anythingllm-workpaper-mcp.html
image: /assets/github-social-preview.png
---

# AnythingLLM WorkPaper MCP setup

Use this when an AnythingLLM agent needs spreadsheet-shaped business logic, but
the formula state should live behind explicit WorkPaper tools instead of Excel
UI automation, browser grid clicks, or unchecked workbook readback.

AnythingLLM loads MCP servers from `plugins/anythingllm_mcp_servers.json` in the
AnythingLLM storage directory. Its MCP integration exposes tools to agents; it
does not expose MCP Resources, Prompts, or Sampling.

Official AnythingLLM references:

- <https://docs.anythingllm.com/mcp-compatibility/overview>
- <https://docs.anythingllm.com/mcp-compatibility/desktop>
- <https://docs.anythingllm.com/mcp-compatibility/docker>

## Fastest smoke test: hosted Streamable HTTP

Use this when you only need to prove that AnythingLLM can discover and call the
Bilig WorkPaper tools.

Edit `plugins/anythingllm_mcp_servers.json`:

```json
{
  "mcpServers": {
    "bilig-workpaper": {
      "type": "streamable",
      "url": "https://bilig.proompteng.ai/mcp"
    }
  }
}
```

Then open **Agent Skills** and refresh MCP servers, or invoke the agent so
AnythingLLM starts the configured server. The hosted endpoint is stateless and
request-local. It proves tool discovery and formula readback, but it does not
persist a private project file.

## Persistent Desktop WorkPaper: stdio

Use this in AnythingLLM Desktop when the WorkPaper JSON file should live on the
host machine and survive across turns.

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
      ]
    }
  }
}
```

`--init-demo-workpaper` creates the demo file only when it is missing.
`--writable` persists tool writes back to the same WorkPaper JSON file.

AnythingLLM Desktop runs MCP commands on the host machine. Make sure `npx` works
from a normal terminal before refreshing Agent Skills.

## Persistent Docker WorkPaper: stdio

Use this when AnythingLLM runs in Docker and the WorkPaper file should persist
inside AnythingLLM storage. AnythingLLM documents that Docker MCP servers can
use paths under `/app/server/storage/...`, which map back to the host
`STORAGE_LOCATION`.

```json
{
  "mcpServers": {
    "bilig-workpaper-docker": {
      "command": "npx",
      "args": [
        "-y",
        "--package",
        "@bilig/workpaper@latest",
        "bilig-workpaper-mcp",
        "--workpaper",
        "/app/server/storage/workpapers/pricing.workpaper.json",
        "--init-demo-workpaper",
        "--writable"
      ]
    }
  }
}
```

Create the `workpapers` directory under the same host `STORAGE_LOCATION` that
AnythingLLM mounts for `/app/server/storage`.

If startup cost matters, keep the MCP server enabled but opt out of automatic
startup until the agent needs workbook tools:

```json
{
  "mcpServers": {
    "bilig-workpaper-docker": {
      "command": "npx",
      "args": [
        "-y",
        "--package",
        "@bilig/workpaper@latest",
        "bilig-workpaper-mcp",
        "--workpaper",
        "/app/server/storage/workpapers/pricing.workpaper.json",
        "--init-demo-workpaper",
        "--writable"
      ],
      "anythingllm": {
        "autoStart": false
      }
    }
  }
}
```

## Proof prompt

After refreshing Agent Skills, ask in an agent-enabled thread:

```text
@agent Use the Bilig WorkPaper MCP tools. List the tools, read the sample sheets,
set Inputs!B3 to 0.4, read Summary!B3, export the WorkPaper document, and return
editedCell, before, after, afterRestore, persistedDocumentBytes, verified, and
limitations.
```

The Bilig server exposes these tools:

- `list_sheets`
- `read_range`
- `read_cell`
- `set_cell_contents`
- `set_cell_contents_and_readback`
- `get_cell_display_value`
- `export_workpaper_document`
- `validate_formula`

Before wiring it into AnythingLLM, you can prove the same local MCP contract from
a terminal:

```sh
npx -y --package @bilig/workpaper@latest bilig-mcp-challenge --json
```

Expected proof fields include:

```json
{
  "transport": "stdio-json-rpc",
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

## Boundaries

- AnythingLLM MCP exposes tools only. Do not rely on MCP Resources, Prompts, or
  Sampling for this integration.
- Hosted Streamable HTTP is stateless. Use Desktop or Docker stdio for private
  writable WorkPaper files.
- Desktop paths are host paths. Docker paths should live under
  `/app/server/storage/...` when the file must persist through the mounted
  storage directory.
- Keep Excel or another workbook oracle in the loop for macros, pivots, charts,
  external links, and layout fidelity.

## Related Bilig docs

- [Agent MCP workbook evaluator](eval-agent-mcp.md)
- [MCP client setup](mcp-client-setup.md)
- [MCP WorkPaper tool server](mcp-workpaper-tool-server.md)
- [Agent framework workbook tools](agent-framework-workbook-tools.md)
- [WorkPaper agent handbook](headless-workpaper-agent-handbook.md)
