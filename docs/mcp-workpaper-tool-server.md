---
title: MCP spreadsheet tool server for WorkPaper agents
published: true
description: Expose @bilig/headless workbook reads, verified edits, formula contracts, and persistence checks through MCP-style tools/list and tools/call handlers.
tags: mcp, model context protocol, spreadsheet, tool calling, node
canonical_url: https://proompteng.github.io/bilig/mcp-workpaper-tool-server.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# MCP Spreadsheet Tool Server For WorkPaper Agents

This page is for agent builders who want workbook formulas behind a Model
Context Protocol tool surface. The useful boundary is small: list the tools,
call one tool, return exact cell readback, and include enough structured output
for the agent to verify the edit.

`@bilig/headless` owns the workbook behavior. MCP should stay as the transport
and discovery layer around ordinary Node functions.

## Runnable MCP-Style Example

Run the dependency-free example from a clean checkout:

```sh
git clone https://github.com/proompteng/bilig.git
cd bilig/examples/headless-workpaper
npm install
npm run agent:mcp-tools
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
cd examples/headless-workpaper
npm install
NODE_NO_WARNINGS=1 npm run --silent agent:mcp-transcript
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
npm exec --package @bilig/headless -- bilig-workpaper-mcp
```

For a real agent workflow, point the same binary at a persisted WorkPaper JSON
document:

```sh
npm exec --package @bilig/headless -- bilig-workpaper-mcp --workpaper ./pricing.workpaper.json --writable
```

File-backed mode loads `./pricing.workpaper.json`, exposes `list_sheets`,
`read_range`, `read_cell`, `set_cell_contents`, `get_cell_display_value`,
`export_workpaper_document`, and `validate_formula`, then writes the updated
WorkPaper JSON back to the same file after `set_cell_contents` when `--writable`
is present. Omit `--writable` for read-only inspection.

Use the maintained file-backed transcript when a directory reviewer or agent
builder needs proof that the packaged binary mutates a real WorkPaper JSON file:

```sh
cd examples/headless-workpaper
npm install
NODE_NO_WARNINGS=1 npm run --silent agent:mcp-file-transcript
```

A passing run starts `npm exec --package @bilig/headless@latest --
bilig-workpaper-mcp --workpaper pricing.workpaper.json --writable`, lists the
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
`bilig-workpaper-mcp --workpaper /workpaper/pricing.workpaper.json --writable`
over stdio. That makes directory introspection see the general WorkPaper tools:
`list_sheets`, `read_range`, `read_cell`, `set_cell_contents`,
`get_cell_display_value`, `export_workpaper_document`, and `validate_formula`.
It also carries the OCI label
`io.modelcontextprotocol.server.name=io.github.proompteng/bilig-workpaper`, so
registry and directory tooling can match the container target to the official
MCP Registry name.

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
    args: ['exec', '--package', '@bilig/headless', '--', 'bilig-workpaper-mcp'],
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

The script implements two JSON-RPC methods shaped around the MCP tool model:

- `tools/list` returns `read_workpaper_summary` and
  `set_workpaper_input_cell` with JSON Schema inputs and MCP tool annotations.
- `tools/call` invokes the requested WorkPaper tool and returns text content
  plus structured formula readback.

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
<https://modelcontextprotocol.io/specification/2025-06-18/server/tools>.

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
