# @bilig/create-workpaper

Create a runnable Bilig WorkPaper starter for Node services and tool integrations.

> [!WARNING]
> Do not use the command below while `@bilig/create-workpaper@latest` resolves
> to `0.164.11`. That release's generated smoke reports
> `formulasPersisted: false`. The restored-formula fix is in source and must
> ship in a newer release, then pass a fresh consumer smoke.

After that fixed release is published:

```sh
npm create @bilig/workpaper@latest pricing-workpaper
cd pricing-workpaper
npm install
npm run smoke
```

The fixed generated starter builds a quote-approval workbook with the A1
facade, writes quote inputs through one atomic `editManyAndReadback` proof,
recalculates formulas, persists JSON, restores the workbook, and prints
`verified: true`. Generated projects pin `@bilig/workpaper` to the generator
package version and use exact dev-tool versions instead of `latest`, so the
smoke proof is reproducible.

After the smoke proof passes, keep the JSON output limited to proof fields.
If the workflow is relevant, watch releases for API and formula compatibility
updates: <https://github.com/proompteng/bilig/subscription>.
If it almost works, open one concrete blocker or integration note:
<https://github.com/proompteng/bilig/discussions/new?category=general>.

Use this when you want to evaluate `@bilig/workpaper` from a blank directory
without cloning the full monorepo.

For an MCP-enabled starter with host integration files, use the `--agent`
template:

```sh
npm create @bilig/workpaper@latest pricing-agent -- --agent
cd pricing-agent
npm install
npm run agent:verify
npm run mcp:server
```

The integration template keeps the service smoke test and adds project-local
MCP config, host rule files, and an `agent:verify` script. That verification
script runs the service smoke plus the package-owned basic and revenue-plan MCP
evaluator proofs. The revenue-plan evaluator checks MCP tool discovery,
mutation, recalculated `SUM`, `SUMIF`, `XLOOKUP`, `FILTER`, a named expression,
persistence, and restart readback. The raw MCP challenge remains available as
`npm run mcp:challenge` when you need the lower-level JSON-RPC transcript.

Included host files cover `AGENTS.md`, `CLAUDE.md`, `GEMINI.md`, Claude Code,
GitHub Copilot / VS Code, Cursor, Kiro, Roo Code, Trae, Qodo, Zed, Junie,
Aider, Cline, Continue, Cascade/Devin, Windsurf, OpenHands, and OpenCode. The
shared MCP command is also written to `mcp/bilig-workpaper.mcp.json`.
Representative host config files include `.kiro/settings/mcp.json`,
`.trae/mcp.json`, `.zed/settings.json`, and `.junie/mcp/mcp.json`.

To add the same MCP proof loop and host files to an existing Node repo without
replacing its app, run:

```sh
npm create @bilig/workpaper@latest . -- --add-agent
npm exec --yes --package @bilig/workpaper@latest -- bilig-evaluate --door agent-mcp --scenario revenue-plan --json
npm exec --yes --package @bilig/workpaper@latest -- bilig-workpaper-mcp --workpaper ./.bilig/pricing.workpaper.json --init-demo-workpaper --writable
```

That writes Bilig-specific host files, keeps an existing `README.md`, does not
edit `package.json`, keeps WorkPaper state under
`./.bilig/pricing.workpaper.json`, and skips existing files unless you pass
`--force`. When existing host instruction files block part of the install, the
CLI writes `BILIG_WORKPAPER_INSTALL.md` with the skipped paths and a short
handoff snippet you can paste into your current policy. The overlay uses
your existing `package.json` name for `BILIG_WORKPAPER.md` instead of rendering
the directory argument.

For an existing MCP client or host integration that does not need a generated
project yet, use the handoff checklist first:
<https://proompteng.github.io/bilig/agent-adoption-kit.html>.
