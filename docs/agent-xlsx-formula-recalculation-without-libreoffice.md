---
title: Agent XLSX formula recalculation without LibreOffice
published: true
description: A focused Node.js path for agents that need to edit XLSX inputs, recalculate formulas, verify readback, and export the edited workbook without shelling out to LibreOffice.
tags: typescript, node, xlsx, agents, formulas, recalculation
canonical_url: https://proompteng.github.io/bilig/agent-xlsx-formula-recalculation-without-libreoffice.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Agent XLSX formula recalculation without LibreOffice

If an agent edits an `.xlsx` file and then acts on a formula result, it needs a
fresh value before the next tool call. Returning the old cached value is worse
than an error because the agent thinks the workbook agreed with it.

Many spreadsheet-agent recipes solve this by running Excel, LibreOffice,
Microsoft Graph, or a Python recalculation helper after every file write. That
is a reasonable choice when exact Excel behavior matters. It is also a heavy
boundary for a Node agent tool that only needs a supported formula workbook,
verified readback, and an exported `.xlsx` at the edge.

Bilig's narrower path is:

1. import the `.xlsx` into a WorkPaper;
2. write the agent's input cells;
3. recalculate in the Node process;
4. read the output cells;
5. export the edited `.xlsx`;
6. reimport it in a smoke test to prove the boundary still works.

## Run the proof

This is the smallest useful check. It starts from a blank directory, downloads
one TypeScript file, creates an XLSX quote workbook, edits inputs, reads the
calculated approval result, exports the edited XLSX, and reimports it.

```sh
mkdir bilig-agent-xlsx-proof
cd bilig-agent-xlsx-proof
npm init -y >/dev/null
npm pkg set type=module
npm install @bilig/headless tsx
curl -fsSLO https://proompteng.github.io/bilig/xlsx-recalculation-proof.ts
npx tsx xlsx-recalculation-proof.ts
```

The run is useful only if it ends with:

```json
{
  "checks": {
    "decisionChanged": true,
    "recalculatedMargin": true,
    "exportedReimportMatchesAfter": true,
    "formulasSurvivedXlsxRoundTrip": true,
    "verified": true
  }
}
```

## Tool contract

For an agent, keep the tool surface boring:

```ts
type WorkbookEditRequest = {
  file: string
  writes: Array<{ sheet: string; cell: string; value: string | number | boolean }>
  reads: Array<{ sheet: string; cell: string }>
}

type WorkbookEditResult = {
  values: Array<{ sheet: string; cell: string; value: unknown }>
  exportedFile: string
  verified: true
}
```

The tool should refuse to return `verified: true` unless all of these happened:

- the target sheets and cells existed;
- every requested write was applied;
- formula output cells were read after the writes;
- the edited workbook was exported;
- the exported workbook could be imported again;
- the reimported values matched the values returned to the agent.

That contract is more important than the model prompt. The agent needs a
tool-shaped invariant it cannot hand-wave past.

## When not to use this

Keep Excel, LibreOffice, or Microsoft Graph in the loop when the workbook
depends on macros, pivots, charts, external links, unsupported functions, or
exact Excel UI behavior.

Use Bilig when the formulas are in the supported runtime surface and the job is
a backend or agent workflow: pricing checks, payout approvals, import
validation, budget gates, quote models, or fixture-driven workbook tests.

## Where this fits

This page exists for the same class of problem documented by spreadsheet-agent
tooling that shells out to a recalculation step after writing formulas. If your
agent already has LibreOffice available and the latency is acceptable, keep it.
If you want a TypeScript runtime that can be tested inside the agent tool loop,
run the proof above and inspect the emitted XLSX files.

Related:

- [curlable XLSX recalculation proof](xlsx-recalculation-proof.md)
- [XLSX formula recalculation in Node.js](xlsx-formula-recalculation-node.md)
- [stale XLSX formula cache in Node.js](stale-xlsx-formula-cache-node.md)
- [agent spreadsheet tool-call loop](agent-spreadsheet-tool-call-loop.md)
- [MCP spreadsheet tool server](mcp-workpaper-tool-server.md)
- [compatibility limits](where-bilig-is-not-excel-compatible-yet.md)

If this is the exact agent spreadsheet loop you are trying to avoid rebuilding,
star or bookmark Bilig so the next developer can find it:
<https://github.com/proompteng/bilig/stargazers>.
