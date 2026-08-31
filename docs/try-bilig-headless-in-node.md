---
title: Try Bilig WorkPaper in Node
published: true
description: Try @bilig/workpaper from an empty Node.js directory, edit one workbook input, read the calculated value, and verify JSON restore.
tags: typescript, node, spreadsheet, formulas, opensource
canonical_url: https://proompteng.github.io/bilig/try-bilig-headless-in-node.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Try Bilig WorkPaper in Node

> [!WARNING]
> Do not use the `npm create` path below while
> `@bilig/create-workpaper@latest` resolves to `0.164.11`. That release's
> generated smoke reports `formulasPersisted: false`. Use the direct evaluator
> path until the restored-formula fix ships in a newer release and passes a
> fresh consumer smoke.

This page is for people who want to try the package before reading the whole
repo. After the fixed generator release, it starts from an empty directory,
installs the published npm package, builds a tiny WorkPaper, edits an input
cell, reads the recalculated formula result, serializes the document, restores
it, and reads the result again.

No browser UI, account, server, or clone is required.

## Quickstart After The Fixed Release

```sh
npm create @bilig/workpaper@latest pricing-workpaper
cd pricing-workpaper
npm install
npm run smoke
```

Expected output includes this proof shape:

```json
{
  "before": {
    "summary": {
      "decision": "review"
    },
    "inputCells": {
      "units": "Inputs!B2",
      "listPrice": "Inputs!B3"
    }
  },
  "edit": {
    "before": {
      "decision": "review"
    },
    "after": {
      "decision": "approved"
    },
    "restored": {
      "decision": "approved"
    },
    "checks": {
      "decisionChanged": true,
      "formulasPersisted": true,
      "restoredMatchesAfter": true,
      "serializedBytes": 1242
    }
  },
  "verified": true
}
```

The exact byte count can change between package versions. The important part is
that `verified` is `true`, `decisionChanged` is `true`, and
`restoredMatchesAfter` is `true`.

The generated starter uses the same maintained TypeScript proof shape as the
public mirror at <https://proompteng.github.io/bilig/npm-eval.ts> and
[`examples/headless-workpaper/npm-eval.ts`](https://github.com/proompteng/bilig/blob/main/examples/headless-workpaper/npm-eval.ts).

## Try it in Docker (optional)

> **Note:** pnpm is the primary recommended path. This section is for
> evaluators who prefer not to change their local Node version.

After completing the **Quickstart** step above you will have a generated
`pricing-workpaper/` project. Mount that directory into an official Node 24
container and run the same smoke script:

```sh
docker run --rm \
  -v "$(pwd)":/eval \
  -w /eval \
  node:24-slim \
  bash -c "npm install --silent && npm run smoke"
```

Expected output uses the same proof shape as above. The local run must set
`verified`, `decisionChanged`, `formulasPersisted`, and `restoredMatchesAfter`
to `true`:

```json
{
  "edit": {
    "after": {
      "decision": "approved"
    },
    "restored": {
      "decision": "approved"
    },
    "checks": {
      "decisionChanged": true,
      "formulasPersisted": true,
      "restoredMatchesAfter": true,
      "serializedBytes": 1242
    }
  },
  "verified": true
}
```

No repo clone is needed. The container installs dependencies from npm and exits
cleanly after printing the result.

## What this proves

- multi-sheet workbook creation from plain arrays
- formula evaluation without a browser grid
- input edits through the workbook API
- computed value readback after the edit
- JSON document export, parse, restore, and readback

This is the core shape behind the larger examples for service routes, MCP tools,
agent writeback, and workbook automation.

## What this does not prove

`bilig` is not a finished Excel clone. It is useful when a TypeScript service or
agent needs a formula-backed workbook object it can mutate and persist. For full
Excel compatibility or XLSX layout fidelity, check the comparison and
compatibility pages before adopting it.

## Next paths

- [GitHub repository](https://github.com/proompteng/bilig)
- [@bilig/workpaper npm package](https://www.npmjs.com/package/@bilig/workpaper)
- [Five Node.js workbook automation examples](workbook-automation-examples-node.md)
- [Node.js spreadsheet formula engine guide](node-spreadsheet-formula-engine.md)
- [WorkPaper service recipe](node-service-workpaper-recipe.md)
- [MCP spreadsheet tool server](mcp-workpaper-tool-server.md)
- [Where bilig is not Excel-compatible yet](where-bilig-is-not-excel-compatible-yet.md)
- [Where bilig is not Excel-compatible yet](where-bilig-is-not-excel-compatible-yet.md)

If it almost matches but a gap blocks adoption, open an implementation gap discussion:
<https://github.com/proompteng/bilig/discussions/new?category=general>.
