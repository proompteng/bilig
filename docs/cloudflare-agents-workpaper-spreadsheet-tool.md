---
title: Cloudflare Agents WorkPaper spreadsheet tool
published: true
description: Use @bilig/headless inside Cloudflare Agents with narrow workbook tools, formula readback, and saved WorkPaper state.
tags: cloudflare agents, agentTool, spreadsheet, workpaper, typescript
canonical_url: https://proompteng.github.io/bilig/cloudflare-agents-workpaper-spreadsheet-tool.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Cloudflare Agents WorkPaper Spreadsheet Tool

Cloudflare Agents can keep state per customer, workspace, or planning session.
That fits workbook-backed workflows: store the WorkPaper document with the
agent, expose a small read tool, and expose one validated write tool.

Use `@bilig/headless` for the spreadsheet part: read a computed range, write one
input cell, verify the dependent formulas, serialize the document, and restore
it.

## Run the checked adapter

```sh
git clone https://github.com/proompteng/bilig.git
cd bilig/examples/headless-workpaper
npm install
npm run agent:framework-adapters
```

The Cloudflare Agents lane exposes AI SDK-style tools and a verified write:

```json
{
  "toolNames": ["readWorkPaperSummary", "setWorkPaperInputCell"],
  "writeResult": {
    "editedCell": "Inputs!B3",
    "checks": {
      "formulasPersisted": true,
      "restoredMatchesAfter": true,
      "expectedArrChanged": true
    }
  }
}
```

## Cloudflare Agents shape

Cloudflare's Agents docs describe `AIChatAgent`, server-side tools, and the
`agentTool` helper for retained sub-agent calls. This WorkPaper example keeps
the integration simpler: expose ordinary AI SDK-style tools from the agent
runtime and keep the mutation behind one small function.

```ts
const tools = {
  setWorkPaperInputCell: {
    description: 'Set one WorkPaper input cell and return formula readback.',
    inputSchema: setInputCellInputSchema,
    execute: setWorkPaperInputCell,
  },
}
```

If the WorkPaper document is stored in the Agent instance, save only after the
tool returns a valid readback. That makes reconnects and later tool calls start
from a verified workbook state.

## What to copy

- Use Agent state for the current WorkPaper document when each user or team has
  an isolated workbook.
- Keep tool arguments narrow: sheet, address, value.
- Return before/after computed values and restored readback equality.
- Use the same WorkPaper functions locally before deploying the Agent.

Official Cloudflare references:
<https://developers.cloudflare.com/agents/api-reference/agents-api/> and
<https://developers.cloudflare.com/agents/api-reference/agent-tools/>.

Runnable source:
[`examples/headless-workpaper/agent-framework-adapters.ts`](../examples/headless-workpaper/agent-framework-adapters.ts).
