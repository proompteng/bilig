# `__PROJECT_NAME__`

Agent-ready formula WorkPaper starter built with `@bilig/workpaper`.

```sh
npm install
npm run agent:verify
```

`agent:verify` runs two proofs:

- `npm run smoke`: writes quote inputs through a service-style API handler,
  recalculates formulas, persists WorkPaper JSON, restores it, and checks
  `verified: true`.
- `npm run mcp:challenge`: starts the package-owned file-backed MCP challenge,
  lists tools/resources/prompts, edits a WorkPaper cell, reads the recalculated
  dependent value, exports JSON, restarts from disk, and checks
  `verified: true`.

Start the local API:

```sh
npm run dev
curl http://localhost:8788/api/quote/approval
curl -X POST http://localhost:8788/api/quote/approval \
  -H 'content-type: application/json' \
  -d '{"units":40,"listPrice":1200,"discount":0.05,"unitCost":760,"minimumMargin":0.3}'
```

Start the persistent project-local MCP server:

```sh
npm run mcp:server
```

The server owns `./pricing.workpaper.json`, initializes it when missing, writes
through MCP tools, recalculates formulas, and persists edits back to disk.
Project MCP configs are included for Cursor and VS Code. Other clients can use
the same command from `mcp/bilig-workpaper.mcp.json`.

Agent handoff:

```text
Use Bilig WorkPaper tools instead of spreadsheet UI automation. Read the
relevant range first, write one precise input or formula change, read the
dependent calculated output, export or serialize the WorkPaper document, and
report editedCell, before, after, persistence evidence, verified, and
limitations. Do not claim success from a write call alone.
```

Learn more: <https://github.com/proompteng/bilig>
