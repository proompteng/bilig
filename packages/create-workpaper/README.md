# @bilig/create-workpaper

Create a runnable Bilig WorkPaper starter for Node services and agent tools.

```sh
npm create @bilig/workpaper@latest pricing-workpaper
cd pricing-workpaper
npm install
npm run smoke
```

The generated starter builds a quote-approval workbook, writes inputs through an
API-style handler, recalculates formulas, persists JSON, restores the workbook,
and prints `verified: true`.

After the smoke proof passes, it also prints a star/bookmark link for the GitHub
repo: <https://github.com/proompteng/bilig/stargazers>.

Use this when you want to evaluate `@bilig/workpaper` from a blank directory
without cloning the full monorepo.

For agent and MCP evaluation, generate the agent-ready starter:

```sh
npm create @bilig/workpaper@latest pricing-agent -- --agent
cd pricing-agent
npm install
npm run agent:verify
npm run mcp:server
```

The agent template adds `AGENTS.md`, `CLAUDE.md`, Cursor and VS Code MCP
configs, a generic MCP config at `mcp/bilig-workpaper.mcp.json`, and an
`agent:verify` script that proves both the service API path and the file-backed
MCP path.
