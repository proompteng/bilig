# Persisting Formula-Backed WorkPaper Documents In Node

`@bilig/headless` can run spreadsheet logic without opening a browser grid, but
the useful boundary for services and agents is not just formula evaluation. A
workflow also needs a way to save workbook state, restore it later, and prove
that formulas still recalculate after restore.

This is the persistence path in `bilig`: export a WorkPaper document, serialize
it as JSON, store it wherever the host application stores durable state, then
parse and restore it before the next operation.

## The Shape

The public persistence helpers are exported from `@bilig/headless`:

- `exportWorkPaperDocument(workbook, { includeConfig: true })`
- `serializeWorkPaperDocument(document)`
- `parseWorkPaperDocument(json)`
- `createWorkPaperFromDocument(document)`

The exported document contains sheets, cell contents, named expressions, and
optionally the persistable WorkPaper config. It is not a screenshot, CSV dump,
or browser session snapshot. Formulas remain formulas, and the restored
workbook can continue to evaluate and mutate through the WorkPaper API.

## Minimal Node Flow

Install the package:

```sh
pnpm add @bilig/headless
```

Build a workbook, write it to disk, restore it, and apply a new edit:

```ts
import { readFileSync, writeFileSync } from 'node:fs'

import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from '@bilig/headless'

const workbook = WorkPaper.buildFromSheets({
  Plan: [
    ['Month', 'Bookings', 'Churn', 'Net MRR'],
    ['January', 12000, 800, '=B2-C2'],
    ['February', 15000, 900, '=B3-C3'],
    ['March', 18000, 1200, '=B4-C4'],
  ],
  Summary: [
    ['Metric', 'Value'],
    ['Quarter net MRR', '=SUM(Plan!D2:D4)'],
    ['Annualized run rate', '=B2*12'],
  ],
})

const saved = serializeWorkPaperDocument(exportWorkPaperDocument(workbook, { includeConfig: true }))
writeFileSync('workpaper.json', saved)

const restored = createWorkPaperFromDocument(parseWorkPaperDocument(readFileSync('workpaper.json', 'utf8')))

const plan = restored.getSheetId('Plan')
if (plan === undefined) {
  throw new Error('Plan sheet was not restored')
}

restored.setCellContents({ sheet: plan, row: 3, col: 1 }, 21000)
```

The important point is that the edit after restore still goes through the same
workbook model. Formula cells can recalculate from the restored document instead
of relying on a stale display value.

## Runnable Example

The focused example is checked in at
[`examples/headless-workpaper/persistence-roundtrip.mjs`](../examples/headless-workpaper/persistence-roundtrip.mjs).

Run it from the example package:

```sh
cd examples/headless-workpaper
npm install
npm run persistence
```

It verifies:

- formulas before save
- JSON serialization to a temporary file
- parse and restore through public helpers
- formula recalculation after an edit to the restored workbook
- sheet and named-expression preservation

The expected summary is:

```json
{
  "beforeSave": {
    "quarterNetMrr": 42100,
    "annualizedRunRate": 505200,
    "expansionAdjustedArr": 545616
  },
  "afterRestoreAndEdit": {
    "quarterNetMrr": 45100,
    "annualizedRunRate": 541200,
    "expansionAdjustedArr": 584496
  },
  "persistedSheets": ["Plan", "Summary"],
  "persistedNamedExpressions": ["ExpansionRatePercent"]
}
```

## Notes For Services And Agents

Keep persistence at the workbook-document boundary:

- Save the serialized WorkPaper document in your normal durable store.
- Treat `parseWorkPaperDocument()` as the validation step for loaded JSON.
- Register application-specific custom functions in code before restoring
  documents that depend on them.
- Use explicit readback after restore for values that matter to the workflow.
- Keep screenshots as human review artifacts, not as the saved state.

For a broader package overview, start with
[`packages/headless/README.md`](../packages/headless/README.md).
