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

This page is for people who want to try the package before reading the whole
repo. It starts from an empty directory, installs the published npm package,
builds a tiny WorkPaper, edits an input cell, reads the recalculated formula
result, serializes the document, restores it, and reads the result again.

No browser UI, account, server, or clone is required.

## Quickstart

```sh
mkdir bilig-headless-eval
cd bilig-headless-eval
npm init -y
npm pkg set type=module
npm install @bilig/workpaper
npm install -D tsx typescript @types/node
curl -fsSLo quickstart.ts https://proompteng.github.io/bilig/npm-eval.ts
npx tsx quickstart.ts
```

Expected output:

```json
{
  "before": 24000,
  "after": 38400,
  "afterRestore": 38400,
  "sheets": ["Inputs", "Summary"],
  "bytes": 999,
  "verified": true,
  "nextStep": "If this proof matches your service or agent workflow, star or bookmark Bilig: https://github.com/proompteng/bilig/stargazers"
}
```

The exact byte count can change between package versions. The important part is
that `verified` is `true` and `afterRestore` matches `after`.

The downloaded file is the maintained TypeScript example at
[`examples/headless-workpaper/npm-eval.ts`](https://github.com/proompteng/bilig/blob/main/examples/headless-workpaper/npm-eval.ts).

## Try it in Docker (optional)

> **Note:** pnpm is the primary recommended path. This section is for
> evaluators who prefer not to change their local Node version.

After completing the **Quickstart** step above you will have a `quickstart.ts` file
inside `bilig-headless-eval/`. Mount that directory into an official Node 24
container and run the same script:

```sh
docker run --rm \
  -v "$(pwd)":/eval \
  -w /eval \
  node:24-slim \
  bash -c "npm install --silent && npx tsx quickstart.ts"
```

Expected output (same as above; `verified` must be `true`):

```json
{
  "before": 24000,
  "after": 38400,
  "afterRestore": 38400,
  "sheets": ["Inputs", "Summary"],
  "bytes": 999,
  "verified": true,
  "nextStep": "If this proof matches your service or agent workflow, star or bookmark Bilig: https://github.com/proompteng/bilig/stargazers"
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
- [@bilig/headless npm package](https://www.npmjs.com/package/@bilig/headless)
- [Five Node.js workbook automation examples](workbook-automation-examples-node.md)
- [Node.js spreadsheet formula engine guide](node-spreadsheet-formula-engine.md)
- [WorkPaper service recipe](node-service-workpaper-recipe.md)
- [MCP spreadsheet tool server](mcp-workpaper-tool-server.md)
- [What the WorkPaper benchmark proves](what-workpaper-benchmark-proves.md)
- [Where bilig is not Excel-compatible yet](where-bilig-is-not-excel-compatible-yet.md)

If the quickstart matches a backend or agent workflow you are building, star the
repo so the package is easier to find later:
<https://github.com/proompteng/bilig/stargazers>.

If it almost matches but a gap blocks adoption, use the adoption blocker form:
<https://github.com/proompteng/bilig/discussions/new?category=general>.
