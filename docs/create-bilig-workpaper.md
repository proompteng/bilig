---
title: Create a Bilig WorkPaper starter
published: true
description: Create a runnable @bilig/workpaper WorkPaper starter from a blank directory with npm create.
tags: typescript, node, spreadsheet, formulas, opensource
canonical_url: https://proompteng.github.io/bilig/create-bilig-workpaper.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Create a Bilig WorkPaper starter

Use this path when you want a runnable project instead of a pasted snippet.

The starter package is `@bilig/create-workpaper`, exposed by
`npm create @bilig/workpaper@latest`. Use it when you want the quote approval
WorkPaper API shape without cloning the full monorepo.

The starter creates a quote approval API with `@bilig/workpaper`. It writes
quote inputs through an API-style handler, recalculates workbook formulas,
persists the WorkPaper as JSON, restores it, and verifies that the restored
formula output still matches the live result.

## Run It

Generated-project path:

```sh
npm create @bilig/workpaper@latest pricing-workpaper
cd pricing-workpaper
npm install
npm run smoke
```

Agent-ready MCP project:

```sh
npm create @bilig/workpaper@latest pricing-agent -- --agent
cd pricing-agent
npm install
npm run agent:verify
npm run mcp:server
```

Full-repo flagship example:

```sh
git clone --depth 1 https://github.com/proompteng/bilig.git
cd bilig
pnpm --dir examples/serverless-workpaper-api install --ignore-workspace
pnpm --dir examples/serverless-workpaper-api run smoke
```

Expected output includes:

```json
{
  "verified": true,
  "nextStep": "If this proof matches your service or agent workflow, star or bookmark Bilig: https://github.com/proompteng/bilig/stargazers"
}
```

The generated project also includes a local API server:

```sh
npm run dev
curl http://localhost:8788/api/quote/approval
curl -X POST http://localhost:8788/api/quote/approval \
  -H 'content-type: application/json' \
  -d '{"units":40,"listPrice":1200,"discount":0.05,"unitCost":760,"minimumMargin":0.3}'
```

## What It Proves

- a real service shape, not only a formula evaluator snippet;
- input writes into named workbook cells;
- formula readback after mutation;
- JSON persistence and restore;
- a narrow API surface that an agent tool or backend route can own.
- an optional agent starter with `AGENTS.md`, `CLAUDE.md`, Cursor and VS Code
  MCP configs, a reusable MCP config file, and an `agent:verify` command that
  proves both direct TypeScript and file-backed MCP paths.

## Source

- package source:
  [`packages/create-workpaper`](https://github.com/proompteng/bilig/tree/main/packages/create-workpaper)
- generated API source:
  [`packages/create-workpaper/template/src/index.ts`](https://github.com/proompteng/bilig/blob/main/packages/create-workpaper/template/src/index.ts)
- generated agent overlay:
  [`packages/create-workpaper/agent-overlay`](https://github.com/proompteng/bilig/tree/main/packages/create-workpaper/agent-overlay)
- full flagship example:
  [quote approval WorkPaper API](quote-approval-workpaper-api.md)

If this starter matches a service or agent workflow you maintain, star or
bookmark the repo so the package is easier to find later:
<https://github.com/proompteng/bilig/stargazers>.
