---
title: Create a Bilig WorkPaper starter
published: true
description: Create a runnable @bilig/workpaper WorkPaper starter for Node services and MCP/tool integrations.
tags: typescript, node, formulas, workpaper, mcp, opensource
canonical_url: https://proompteng.github.io/bilig/create-bilig-workpaper.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Create a Bilig WorkPaper starter

Use this path when you want a runnable project instead of a pasted snippet.

> [!WARNING]
> Do not use the commands below while `@bilig/create-workpaper@latest` resolves
> to `0.164.11`. That release's generated smoke reports
> `formulasPersisted: false`. The restored-formula fix is in source and must
> ship in a newer release, then pass a fresh consumer smoke.

The starter package is `@bilig/create-workpaper`, exposed by
`npm create @bilig/workpaper@latest`. Use it when you want the quote approval
WorkPaper API shape without cloning the full monorepo.

The fixed starter creates a quote approval API with `@bilig/workpaper`. It uses
the A1 facade, writes quote inputs through one atomic `editManyAndReadback`
proof, recalculates workbook formulas, persists the WorkPaper as JSON, restores
it, and verifies that the restored formula output still matches the live
result. Generated projects pin `@bilig/workpaper` to the generator package
version and use exact dev-tool versions instead of `latest`, so the smoke proof
is reproducible.

## Run It After The Fixed Release

Generated-project path:

```sh
npm create @bilig/workpaper@latest pricing-workpaper
cd pricing-workpaper
npm install
npm run smoke
```

MCP-enabled project with host integration files:

```sh
npm create @bilig/workpaper@latest pricing-agent -- --agent
cd pricing-agent
npm install
npm run agent:verify
npm run mcp:server
```

Existing Node project:

```sh
npm create @bilig/workpaper@latest . -- --add-agent
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --scenario revenue-plan --json
npm exec --yes --package @bilig/workpaper@latest -- bilig-workpaper-mcp --workpaper ./.bilig/pricing.workpaper.json --init-demo-workpaper --writable
```

`--add-agent` only adds Bilig MCP and host-integration instructions. It keeps
your existing app template and `README.md`, writes `BILIG_WORKPAPER.md`, keeps
WorkPaper state under `./.bilig/pricing.workpaper.json`, does not edit
`package.json`, and skips existing files unless you pass `--force`. If an
existing host policy blocks part of the overlay, the CLI writes
`BILIG_WORKPAPER_INSTALL.md` with the skipped paths and a compact handoff
snippet for your current policy file.
For `npm create @bilig/workpaper@latest . -- --add-agent`, the overlay labels
`BILIG_WORKPAPER.md` from your existing `package.json` name instead of `.`.

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
  "verified": true
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
- atomic input writes into named workbook cells through A1 addresses;
- formula readback after mutation;
- JSON persistence and restore;
- a narrow API surface that an MCP tool, host integration, or backend route can
  own.
- an optional integration starter with `AGENTS.md`, `CONVENTIONS.md`,
  `.aider.conf.yml`, `CLAUDE.md`, `GEMINI.md`, host rule files, project-root
  MCP configs, and an `agent:verify` command that runs the service smoke plus
  the package-owned basic and revenue-plan MCP evaluator proofs. The
  revenue-plan evaluator checks MCP tool discovery, mutation, recalculated
  `SUM`, `SUMIF`, `XLOOKUP`, `FILTER`, a named expression, persistence, and
  restart readback. Use `npm run mcp:challenge` only when you need the
  lower-level JSON-RPC transcript.

Representative host config files include `.kiro/settings/mcp.json`,
`.trae/mcp.json`, `.zed/settings.json`, `.junie/mcp/mcp.json`, and
`mcp/bilig-workpaper.mcp.json`.

## Source

- package source:
  [`packages/create-workpaper`](https://github.com/proompteng/bilig/tree/main/packages/create-workpaper)
- generated API source:
  [`packages/create-workpaper/template/src/index.ts`](https://github.com/proompteng/bilig/blob/main/packages/create-workpaper/template/src/index.ts)
- generated host integration overlay:
  [`packages/create-workpaper/agent-overlay`](https://github.com/proompteng/bilig/tree/main/packages/create-workpaper/agent-overlay)
- full flagship example:
  [quote approval WorkPaper API](quote-approval-workpaper-api.md)

If this starter almost matches a service or integration workflow you maintain,
open one concrete implementation gap so the package becomes easier to evaluate:
<https://github.com/proompteng/bilig/discussions/new?category=general>.

If the proof already matches your workflow, watch releases for API and formula
compatibility updates: <https://github.com/proompteng/bilig/subscription>.
