---
title: MCP spreadsheet tool server for WorkPaper agents
published: true
description: Expose @bilig/headless workbook reads, verified edits, formula contracts, persistence checks, resources, and prompts through MCP.
tags: mcp, model context protocol, spreadsheet, tool calling, node
canonical_url: https://proompteng.github.io/bilig/mcp-workpaper-tool-server.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# MCP Spreadsheet Tool Server For WorkPaper Agents

This page is for agent builders who want workbook formulas behind a Model
Context Protocol surface. The useful boundary is small: list the tools, read
the workbook context resources, invoke a reusable workflow prompt, call one
tool, return exact cell readback, and include enough structured output for the
agent to verify the edit.

`@bilig/headless` owns the workbook behavior. MCP should stay as the transport
and discovery layer around ordinary Node functions.

If you need the short agent decision path before the protocol details, start
with the [headless WorkPaper agent handbook](headless-workpaper-agent-handbook.md).

## Runnable MCP-Style Example

Run the dependency-free example from a clean checkout:

```sh
git clone https://github.com/proompteng/bilig.git
cd bilig
pnpm --dir examples/headless-workpaper install --ignore-workspace
pnpm --dir examples/headless-workpaper run agent:mcp-tools
```

For a local stdio transport, pipe newline-delimited JSON-RPC requests into the
stdio entrypoint:

```sh
printf '%s\n' \
  '{"jsonrpc":"2.0","id":1,"method":"initialize"}' \
  '{"jsonrpc":"2.0","method":"notifications/initialized"}' \
  '{"jsonrpc":"2.0","id":2,"method":"tools/list"}' |
  npm run --silent agent:mcp-stdio
```

## Copy-Paste JSON-RPC Transcript

Use the maintained transcript smoke when reviewing the server from an MCP
client, directory submission, or HN-style launch thread:

```sh
pnpm --dir examples/headless-workpaper install --ignore-workspace
pnpm --dir examples/headless-workpaper run agent:mcp-transcript
```

The script starts the stdio server, sends `initialize`, `tools/list`, and
`tools/call`, parses the JSON-RPC responses, asserts the formula readback, and
prints a compact transcript summary. The important response is the `tools/call`
result. A passing run returns structured content like this:

```json
{
  "jsonrpc": "2.0",
  "id": 3,
  "result": {
    "structuredContent": {
      "editedCell": "Inputs!B3",
      "before": {
        "expectedCustomers": 5,
        "expectedArr": 60000,
        "expansionArr": 66000,
        "targetGap": -34000
      },
      "after": {
        "expectedCustomers": 8,
        "expectedArr": 96000,
        "expansionArr": 105600,
        "targetGap": 5600
      },
      "restored": {
        "expectedCustomers": 8,
        "expectedArr": 96000,
        "expansionArr": 105600,
        "targetGap": 5600
      },
      "formulaContracts": {
        "expectedCustomers": "=Inputs!B2*Inputs!B3",
        "expectedArr": "=B2*Inputs!B4",
        "expansionArr": "=B3*Inputs!B5",
        "targetGap": "=B4-100000"
      },
      "checks": {
        "previousValue": 0.25,
        "newValue": 0.4,
        "formulasPersisted": true,
        "restoredMatchesAfter": true,
        "expectedArrChanged": true,
        "serializedBytes": 1163
      }
    },
    "isError": false
  }
}
```

That single response proves the tool changed one input cell, recalculated
dependent formulas, preserved the formulas through WorkPaper JSON
serialization, restored the document, and returned machine-checkable readback.

If you want the raw newline-delimited JSON-RPC request stream instead of the
maintained transcript wrapper, use:

```sh
printf '%s\n' \
  '{"jsonrpc":"2.0","id":1,"method":"initialize"}' \
  '{"jsonrpc":"2.0","method":"notifications/initialized"}' \
  '{"jsonrpc":"2.0","id":2,"method":"tools/list"}' \
  '{"jsonrpc":"2.0","id":3,"method":"tools/call","params":{"name":"set_workpaper_input_cell","arguments":{"sheetName":"Inputs","address":"B3","value":0.4}}}' |
  NODE_NO_WARNINGS=1 npm run --silent agent:mcp-stdio
```

The npm package exposes the demo server as `bilig-workpaper-mcp` by default:

```sh
npm exec --package @bilig/headless@0.40.19 -- bilig-workpaper-mcp
```

## Remote Stateless Endpoint

The hosted app runtime also exposes a JSON-only Streamable HTTP MCP endpoint for
clients that cannot launch a local stdio process:

```text
https://bilig.proompteng.ai/mcp
```

There is also a compatibility alias:

```text
https://bilig.proompteng.ai/mcp/workpaper
```

The endpoint is stateless and request-local. It loads the packaged demo
WorkPaper for each JSON-RPC request, exposes the same file-backed tool catalog,
resources, and prompts, and returns write/readback proof without writing user
files or issuing an MCP session id. Use it for Claude custom connector smoke
tests, directory probes, and agent onboarding. Use local file-backed stdio when
an agent needs to persist a real project WorkPaper JSON file.

Protocol smoke:

```sh
curl -fsS https://bilig.proompteng.ai/mcp \
  -H 'content-type: application/json' \
  -H 'accept: application/json, text/event-stream' \
  -H 'mcp-protocol-version: 2025-11-25' \
  --data '{"jsonrpc":"2.0","id":1,"method":"tools/list"}' | jq '.result.tools[].name'
```

For server-to-server clients, omit `Origin`. Browser-based clients must send an
allowed `Origin`; Claude origins are allowed by default.

For a real agent workflow, point the same binary at a persisted WorkPaper JSON
document:

```sh
npm exec --package @bilig/headless@0.40.19 -- bilig-mcp-challenge
npm exec --package @bilig/headless@0.40.19 -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --init-demo-workpaper --writable
```

`bilig-mcp-challenge` is the one-command evaluator path. It initializes the
file-backed MCP server, lists tools/resources/prompts, edits `Inputs!B3`, reads
recalculated `Summary!B3`, exports WorkPaper JSON, restarts from disk, and
prints `verified: true`.

File-backed mode loads `./pricing.workpaper.json`, exposes `list_sheets`,
`read_range`, `read_cell`, `set_cell_contents`, `get_cell_display_value`,
`export_workpaper_document`, and `validate_formula`, then writes the updated
WorkPaper JSON back to the same file after `set_cell_contents` when `--writable`
is present. It also exposes `resources/list`, `resources/read`,
`prompts/list`, and `prompts/get` so clients can discover the live workbook
manifest, agent handoff instructions, current document JSON, and reusable edit
or formula-debug prompts. Omit `--writable` for read-only inspection.

The high-signal runtime resources are:

- `bilig://workpaper/manifest`
- `bilig://workpaper/agent-handoff`
- `bilig://workpaper/sheets`
- `bilig://workpaper/current-document`

The reusable prompts are:

- `edit_and_verify_workpaper`
- `debug_workpaper_formula`

Every file-backed tool includes an MCP `outputSchema`, parameter descriptions,
and safety annotations (`readOnlyHint`, `destructiveHint`, `idempotentHint`,
and `openWorldHint`). That is deliberate: directory scanners and coding agents
should be able to pick the workbook read, write, display, export, or formula
validation tool without treating the description as a vague demo.

Use the maintained file-backed transcript when a directory reviewer or agent
builder needs proof that the packaged binary mutates a real WorkPaper JSON file:

```sh
pnpm --dir examples/headless-workpaper install --ignore-workspace
pnpm --dir examples/headless-workpaper run agent:mcp-file-transcript
```

A passing run starts `npm exec --package @bilig/headless@latest --
bilig-workpaper-mcp --workpaper pricing.workpaper.json --init-demo-workpaper --writable`, lists the
file-backed tool surface, writes `Inputs!B3`, persists the JSON file, reads
`Summary!B3`, and asserts that the recalculated value is `96000`.

## Docker Target For Directory Introspection

MCP directories such as Glama need to start the server and run `tools/list`
without cloning the monorepo or building the web app. The root Dockerfile keeps
the production web image as `--target bilig-runtime` and adds a separate MCP
target for directory scanners:

```sh
docker build --target bilig-workpaper-mcp -t bilig-workpaper-mcp:local .
printf '%s\n' \
  '{"jsonrpc":"2.0","id":1,"method":"initialize"}' \
  '{"jsonrpc":"2.0","method":"notifications/initialized"}' \
  '{"jsonrpc":"2.0","id":2,"method":"tools/list"}' |
  docker run --rm -i bilig-workpaper-mcp:local
```

The target installs `@bilig/headless` from npm, seeds
`/workpaper/pricing.workpaper.json`, and starts
`bilig-workpaper-mcp --workpaper /workpaper/pricing.workpaper.json --init-demo-workpaper --writable`
over stdio. That makes directory introspection see the general WorkPaper tools:
`list_sheets`, `read_range`, `read_cell`, `set_cell_contents`,
`get_cell_display_value`, `export_workpaper_document`, and `validate_formula`.
It also carries the OCI label
`io.modelcontextprotocol.server.name=io.github.proompteng/bilig-workpaper`, so
registry and directory tooling can match the container target to the official
MCP Registry name.

For crawlers that cannot run Docker or stdio, the docs site also publishes a
static MCP server card at
`https://proompteng.github.io/bilig/.well-known/mcp/server-card.json`. The card
lists the same `list_sheets`, `read_range`, `read_cell`, `set_cell_contents`,
`get_cell_display_value`, `export_workpaper_document`, and `validate_formula`
tools, plus the WorkPaper resources and prompts, without requiring account auth
or a live server connection.

The hosted endpoint origin serves the same crawler-friendly card at
`https://bilig.proompteng.ai/.well-known/mcp/server-card.json`, with
`streamable-http` transport metadata for `https://bilig.proompteng.ai/mcp`.
That gives Smithery-style scanners a same-origin metadata path when they start
from the remote MCP URL rather than the documentation site.

The package carries `mcpName: io.github.proompteng/bilig-workpaper` and a
matching `server.json`. It is published in the official MCP Registry as
`io.github.proompteng/bilig-workpaper`:
<https://registry.modelcontextprotocol.io/v0.1/servers?search=io.github.proompteng%2Fbilig-workpaper>.

If you already know which client you want to use, start with the
[MCP client setup guide](mcp-client-setup.md) for Claude, Cursor, VS Code, and
Codex config snippets.

If you are checking a directory listing or preparing one, use the
[MCP spreadsheet server directory status page](mcp-spreadsheet-server-directory.md)
for the canonical npm command, official Registry proof, Glama listing, and
pending directory-review status.

Before submitting the server to an MCP registry, verify this repo-specific
readiness checklist:

- `packages/headless/server.json` exists and describes the packaged stdio
  server.
- `packages/headless/package.json` exposes `bilig-workpaper-mcp` in `bin`.
- `packages/headless/package.json` includes
  `mcpName: io.github.proompteng/bilig-workpaper`.
- `pnpm publish:runtime:check` passes against the runtime packages.
- `pnpm workpaper:smoke:external` passes against packed local runtime packages.

Passing the checklist means the repository metadata and smoke checks are ready
for registry submission; it does not mean the package has already been
published.

## Vercel AI SDK MCP Client Recipe

If your agent loop already uses the Vercel AI SDK, keep the MCP client thin and
let the WorkPaper server own the spreadsheet reads and writes:

```ts
import { createMCPClient } from '@ai-sdk/mcp'
import { Experimental_StdioMCPTransport } from '@ai-sdk/mcp/mcp-stdio'
import { generateText } from 'ai'

const client = await createMCPClient({
  transport: new Experimental_StdioMCPTransport({
    command: 'npm',
    args: ['exec', '--package', '@bilig/headless@0.32.6', '--', 'bilig-workpaper-mcp'],
  }),
})

try {
  const tools = await client.tools()
  const { text } = await generateText({
    model: 'your-model',
    tools,
    prompt: [
      'Read the WorkPaper summary with read_workpaper_summary for Summary!A1:B5.',
      'Then set Inputs!B3 to 0.4 with set_workpaper_input_cell.',
      'Return editedCell plus the before and after expectedArr values.',
    ].join('\n'),
  })

  console.log(text)
} finally {
  await client.close()
}
```

The server command is `bilig-workpaper-mcp`; the `npm exec --package
@bilig/headless -- bilig-workpaper-mcp` wrapper only resolves the published npm
package for a clean checkout. The stdio transport receives `npm` as the command
and the rest as `args`, so shell parsing does not sit between the AI SDK client
and the MCP server. The two tool calls prove the useful workflow: read a
formula-backed summary, set one input cell, and return computed before/after
readback.

Verify the docs links and discovery metadata after editing this page:

```sh
pnpm docs:discovery:check
```

The script implements the JSON-RPC methods needed for the file-backed WorkPaper
agent surface:

- `tools/list` returns `read_workpaper_summary` and
  `set_workpaper_input_cell` with JSON Schema inputs and MCP tool annotations.
- `tools/call` invokes the requested WorkPaper tool and returns text content
  plus structured formula readback.
- `resources/list` and `resources/read` expose the live WorkPaper manifest,
  sheet summary, current document JSON, and compact agent handoff.
- `prompts/list` and `prompts/get` expose the edit-and-verify and formula-debug
  workflows as reusable client prompts.

The packaged binary has two tool sets:

- default demo mode: `read_workpaper_summary` and `set_workpaper_input_cell`
- file-backed mode: `list_sheets`, `read_range`, `read_cell`,
  `set_cell_contents`, `get_cell_display_value`, `export_workpaper_document`,
  and `validate_formula`

The annotations are explicit for directory reviewers and cautious MCP clients:
`read_workpaper_summary` is read-only, idempotent, and closed-world.
`set_workpaper_input_cell` mutates the local WorkPaper state, is idempotent for
the same cell/value arguments, and is closed-world rather than a network or
filesystem tool.
In file-backed mode, `set_cell_contents` is annotated as destructive only when
the server starts with `--writable`.

### MCP Stdio Troubleshooting

| Symptom                        | What to check                                                                                   |
| ------------------------------ | ----------------------------------------------------------------------------------------------- |
| `Parse error` response         | Make sure each stdin line is valid JSON before it reaches the server.                           |
| No response appears            | End each JSON-RPC message with a newline; the server waits for newline-delimited input.         |
| Notification has no output     | `notifications/initialized` is intentionally one-way and does not produce a JSON-RPC response.  |
| `Invalid params` or tool error | Check that `tools/call` includes a supported `name` and the required `arguments` for that tool. |

The example deliberately avoids an MCP SDK dependency so the workbook contract
is visible. Put the same handlers behind stdio, HTTP, or your MCP SDK adapter
when you wire it into a production agent host.

## What A Passing Run Proves

The write tool edits `Inputs!B3`, recalculates dependent formulas, serializes
the WorkPaper document, restores it, and checks that formulas and computed
values survived the round trip:

```json
{
  "editedCell": "Inputs!B3",
  "before": {
    "expectedCustomers": 5,
    "expectedArr": 60000,
    "expansionArr": 66000,
    "targetGap": -34000
  },
  "after": {
    "expectedCustomers": 8,
    "expectedArr": 96000,
    "expansionArr": 105600,
    "targetGap": 5600
  },
  "checks": {
    "previousValue": 0.25,
    "newValue": 0.4,
    "formulasPersisted": true,
    "restoredMatchesAfter": true,
    "expectedArrChanged": true
  }
}
```

That is the part spreadsheet agents need. A tool that only says "updated" is
not enough. Return the edited address, previous value, new value, before/after
computed values, formula contracts, and persistence proof.

## Tool Boundary

Expose only the minimum useful surface first:

1. `read_workpaper_summary` reads a bounded range and returns computed values
   plus serialized cell contents.
2. `set_workpaper_input_cell` validates the sheet and A1 address before a
   write, then returns formula readback and persistence checks.
3. Everything outside that boundary stays in your MCP host: auth, transport,
   rate limits, logging, and user approval policy.

The official MCP specification describes tool discovery through `tools/list`,
tool invocation through `tools/call`, input schemas, and tool annotations:
<https://modelcontextprotocol.io/specification/2025-11-25/server/tools>.
It also defines server resources through `resources/list` and
`resources/read`, and reusable prompt templates through `prompts/list` and
`prompts/get`:
<https://modelcontextprotocol.io/specification/2025-11-25/server/resources>
and
<https://modelcontextprotocol.io/specification/2025-11-25/server/prompts>.

## Files To Inspect

- MCP-style adapter script:
  [`examples/headless-workpaper/mcp-tool-server.ts`](https://github.com/proompteng/bilig/blob/main/examples/headless-workpaper/mcp-tool-server.ts)
- stdio adapter script:
  [`examples/headless-workpaper/mcp-stdio-server.ts`](https://github.com/proompteng/bilig/blob/main/examples/headless-workpaper/mcp-stdio-server.ts)
- official MCP Registry entry:
  [`io.github.proompteng/bilig-workpaper`](https://registry.modelcontextprotocol.io/v0.1/servers?search=io.github.proompteng%2Fbilig-workpaper)
- example README:
  [`examples/headless-workpaper/README.md#mcp-tool-server-shape`](https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper#mcp-tool-server-shape)
- SDK-neutral tool-calling recipe:
  [`docs/agent-workpaper-tool-calling-recipe.md`](agent-workpaper-tool-calling-recipe.md)
- Vercel AI SDK and LangChain wrappers:
  [`docs/vercel-ai-sdk-langchain-spreadsheet-tool.md`](vercel-ai-sdk-langchain-spreadsheet-tool.md)

## Feedback Thread

Use the
[MCP spreadsheet tool server discussion](https://github.com/proompteng/bilig/discussions/230)
for adapter feedback. The open questions are deliberately concrete: stdio,
HTTP/SSE, or SDK adapter next; which spreadsheet workflow should be proven
next; and which structured fields every write tool should return.

## When This Is A Good Fit

Use this pattern when an agent needs to edit a forecast, pricing workbook,
quote approval rule, budget check, or service-side spreadsheet model and prove
the formulas reacted. Keep the MCP layer thin, keep the workbook logic
testable, and make every write return structured verification.

Start with the adapter command above. If it saves you a spreadsheet-tooling
spike, star the repository so the next person searching for MCP spreadsheet
tools can find it:
<https://github.com/proompteng/bilig/stargazers>.

If it almost matches but a gap blocks adoption, use the adoption blocker form:
<https://github.com/proompteng/bilig/discussions/new?category=general>.
