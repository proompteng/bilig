---
title: Try bilig headless spreadsheet engine in Node
published: true
description: Run a small @bilig/headless WorkPaper smoke test from an empty Node.js directory and verify formula readback after JSON persistence.
tags: typescript, node, spreadsheet, formulas, opensource
canonical_url: https://proompteng.github.io/bilig/try-bilig-headless-in-node.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Try bilig headless spreadsheet engine in Node

This page is for people who want to try the package before reading the whole
repo. It starts from an empty directory, installs the published npm package,
builds a tiny WorkPaper, edits an input cell, reads the recalculated formula
result, serializes the document, restores it, and reads the result again.

No browser UI, account, server, or clone is required.

## Run it

```sh
mkdir bilig-headless-eval
cd bilig-headless-eval
npm init -y
npm pkg set type=module
npm install @bilig/headless
npm install -D tsx typescript @types/node
cat > eval.ts <<'EOF'
import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from '@bilig/headless'

type NumericCell = {
  value: number
}

function numberValue(cell: unknown): number {
  if (typeof cell === 'object' && cell !== null && typeof (cell as NumericCell).value === 'number') {
    return (cell as NumericCell).value
  }

  throw new Error(`expected numeric cell value, got ${JSON.stringify(cell)}`)
}

const workbook = WorkPaper.buildFromSheets({
  Inputs: [
    ['Metric', 'Value'],
    ['Customers', 20],
    ['Average revenue', 1200],
  ],
  Summary: [
    ['Metric', 'Value'],
    ['Revenue', '=Inputs!B2*Inputs!B3'],
  ],
})

const inputs = workbook.getSheetId('Inputs')
const summary = workbook.getSheetId('Summary')
if (inputs === undefined || summary === undefined) {
  throw new Error('missing sheet')
}

const before = numberValue(workbook.getCellValue({ sheet: summary, row: 1, col: 1 }))
workbook.setCellContents({ sheet: inputs, row: 1, col: 1 }, 32)

const after = numberValue(workbook.getCellValue({ sheet: summary, row: 1, col: 1 }))
const saved = serializeWorkPaperDocument(exportWorkPaperDocument(workbook, { includeConfig: true }))
const restored = createWorkPaperFromDocument(parseWorkPaperDocument(saved))
const restoredSummary = restored.getSheetId('Summary')
if (restoredSummary === undefined) {
  throw new Error('missing restored Summary sheet')
}

const afterRestore = numberValue(restored.getCellValue({ sheet: restoredSummary, row: 1, col: 1 }))
console.log(
  JSON.stringify(
    {
      before,
      after,
      afterRestore,
      bytes: saved.length,
      verified: before === 24000 && after === 38400 && afterRestore === 38400,
    },
    null,
    2,
  ),
)
EOF
npx tsx eval.ts
```

Expected output:

```json
{
  "before": 24000,
  "after": 38400,
  "afterRestore": 38400,
  "bytes": 1000,
  "verified": true
}
```

The exact byte count can change between package versions. The important part is
that `verified` is `true` and `afterRestore` matches `after`.

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
- [npm package](https://www.npmjs.com/package/@bilig/headless)
- [Five Node.js workbook automation examples](workbook-automation-examples-node.md)
- [Node.js spreadsheet formula engine guide](node-spreadsheet-formula-engine.md)
- [WorkPaper service recipe](node-service-workpaper-recipe.md)
- [MCP spreadsheet tool server](mcp-workpaper-tool-server.md)
- [What the WorkPaper benchmark proves](what-workpaper-benchmark-proves.md)
- [Where bilig is not Excel-compatible yet](where-bilig-is-not-excel-compatible-yet.md)

If the smoke test matches a backend or agent workflow you are building, star the
repo so the package is easier to find later:
<https://github.com/proompteng/bilig/stargazers>.
