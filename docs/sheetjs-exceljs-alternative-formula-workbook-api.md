---
title: SheetJS and ExcelJS alternative for formula-backed workbook APIs
published: true
description: Decide when SheetJS, ExcelJS, xlsx-populate, xlsx-formula-recalc, exceljs-formula-recalc, or @bilig/workpaper is the right fit for stale formula results and verified Node.js workbook execution.
tags: typescript, node, spreadsheet, formulas, xlsx, opensource
canonical_url: https://proompteng.github.io/bilig/sheetjs-exceljs-alternative-formula-workbook-api.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# SheetJS and ExcelJS alternative for formula-backed workbook APIs

Status: public comparison guide for developers evaluating high-traffic
spreadsheet-file libraries and stale formula result fixes.

Research date: 2026-05-20.

If you are searching for a SheetJS alternative or ExcelJS alternative, do not
start with a package name. Start with the job:

- read and write workbook files
- generate an XLSX report for Excel to open later
- keep formula text and cached or supplied formula results in the file
- refresh stale formula results after a Node service edits workbook inputs
- run a workbook inside a Node.js service, edit inputs, and verify the new
  result

SheetJS, `xlsx-populate`, and ExcelJS are strong tools for workbook-file
workflows. `bilig` is not trying to replace that whole layer. The useful Bilig
slice is narrower:

- `@bilig/xlsx-formula-recalc` refreshes formulas from raw XLSX bytes produced by
  SheetJS, `xlsx-populate`, template builders, or file uploads.
- `@bilig/exceljs-formula-recalc` keeps ExcelJS as the workbook authoring layer and
  adds recalculated readback for ExcelJS workflows.
- `@bilig/workpaper` gives a Node service or coding agent a WorkPaper object it
  can build, mutate, evaluate, persist, restore, and verify without opening
  Excel or a browser grid.

## Short Version

Use SheetJS when you need broad spreadsheet-file parsing and export.

Use ExcelJS when you need to create or edit XLSX workbooks with workbook-file
features such as sheets, rows, styles, and formula records.

Use `@bilig/xlsx-formula-recalc` when a Node service already has XLSX bytes and must
read fresh formula values after changing inputs.

Use `@bilig/exceljs-formula-recalc` when the workbook is already moving through
ExcelJS and the missing piece is in-process formula readback.

Use `@bilig/workpaper` when the service must own the formula-backed workbook as
runtime state, persist it as JSON, restore it, and verify recalculated values in
the same process.

That is the boundary. If a backend only needs a file, stay with a file library.
If it needs a recalculated answer now, add the narrow recalculation bridge. If
it needs durable formula state, move to WorkPaper.

For the broader library choice, start with the
[headless spreadsheet engine use-case chooser](headless-spreadsheet-engine-comparison.md#use-case-chooser).

## The Boundary That Matters

SheetJS Community Edition stores cell formulas in the `f` field and cell values
in the `v` field. Its formula docs explain that, when actual results are needed
in JavaScript, SheetJS Pro has a formula calculator component.

ExcelJS can store formulas and supplied results, but its public package docs say
it cannot process a formula to generate a result.

Those are reasonable design choices for file-centric libraries. They become a
problem only when your app needs to change an input, recalculate dependent
cells, and reject a workflow when computed readback does not match.

That is the place to evaluate `@bilig/workpaper`.

For the more common file-boundary problem, evaluate the narrow packages first:

```sh
npx --package @bilig/xlsx-formula-recalc xlsx-recalc --demo --json
npx --package @bilig/exceljs-formula-recalc exceljs-recalc --demo --json
```

Both demos print `verified: true` when the service changes input cells and
reads recalculated formula output without opening Excel, LibreOffice, or a
browser.

## Traffic Reality

This page intentionally targets the big spreadsheet-library paths, not low
traffic integration directories. On the research date, the live npm download
API showed the real audience is already concentrated around these packages:

| Package | Last-week npm downloads on 2026-05-20 | Practical implication |
| --- | ---: | --- |
| `xlsx` / SheetJS | 10,608,303 | Optimize for SheetJS-style XLSX buffers and stale formula cache searches. |
| `exceljs` | 8,133,216 | Keep ExcelJS in the workflow; add recalculated readback at the missing boundary. |
| `@formulajs/formulajs` | 344,141 | Formula-function users may need workbook semantics, dependency tracking, and verification. |
| `hyperformula` | 305,054 | Compare honestly against mature formula-engine use cases. |
| `xlsx-populate` | 201,621 | Generated-workbook users often need fresh formula results before sending the file. |
| `xlsx-calc` | 150,686 | Migration pages should focus on unsupported formulas, workbook size, and verification. |

The growth surface is not another generic "spreadsheet engine" post. It is the
exact failure mode those users search for: "I edited an XLSX in Node and the
formula result is stale."

## Comparison Table

| Need | Start with | Reason |
| --- | --- | --- |
| Parse many spreadsheet file formats into JavaScript data | SheetJS | It is built around file-format import/export and a common spreadsheet object model. |
| Generate XLSX reports with workbook structure and styling | ExcelJS | It focuses on reading, manipulating, and writing XLSX workbook files. |
| Store formulas in a workbook file and let Excel calculate later | SheetJS or ExcelJS | Both can represent formula text and cached or supplied values in workbook data. |
| Recalculate a SheetJS / `xlsx` pipeline after changing inputs | `@bilig/xlsx-formula-recalc` | It accepts the XLSX bytes already produced by SheetJS and returns fresh formula readback plus exported bytes. |
| Recalculate raw XLSX bytes after changing inputs | `@bilig/xlsx-formula-recalc` | It accepts the XLSX bytes already produced by SheetJS, `xlsx-populate`, or template tools and returns fresh readback plus exported bytes. |
| Recalculate an existing ExcelJS workbook after changing inputs | `@bilig/exceljs-formula-recalc` | It preserves the ExcelJS authoring boundary and patches requested formula cells with fresh results. |
| Recalculate formulas inside a Node service after changing inputs | `@bilig/workpaper` or `@bilig/workpaper` | It exposes a WorkPaper runtime with formula readback, JSON persistence, and restore verification after edits. |
| Give a coding agent a spreadsheet tool it can mutate and verify | `@bilig/workpaper` | The maintained examples prove writeback, dependent formulas, persistence, and restore. |

## Use The Narrow Bridge First

If the `.xlsx` file already exists, start here:

```sh
npm install @bilig/xlsx-formula-recalc
npx --package @bilig/xlsx-formula-recalc xlsx-recalc pricing.xlsx \
  --set Inputs!B2=48 \
  --set Inputs!B3=1500 \
  --read Summary!B7 \
  --out pricing.recalculated.xlsx \
  --json
```

If the workbook already lives in ExcelJS, keep ExcelJS:

```sh
npm install exceljs @bilig/exceljs-formula-recalc
npx --package @bilig/exceljs-formula-recalc exceljs-recalc --demo --json
```

If you want one checkout-level proof across the common incumbents, run the
bridge smoke test. It edits the same workbook through SheetJS/`xlsx`,
`xlsx-populate`, and ExcelJS, then verifies Bilig refreshes the stale `48000`
result to `72000` in all three paths:

```sh
git clone https://github.com/proompteng/bilig.git
cd bilig
npm --prefix examples/recalc-bridge-workflows install
npm --prefix examples/recalc-bridge-workflows run smoke
```

That is the conversion path for most file-library users. Reach for
`@bilig/workpaper` only when the workbook itself becomes service-owned runtime
state rather than a file artifact.

## TypeScript WorkPaper Evaluation Path

Install the full WorkPaper runtime in a scratch project:

```sh
mkdir bilig-headless-eval
cd bilig-headless-eval
npm init -y
npm pkg set type=module
npm install @bilig/workpaper
npm install -D tsx typescript @types/node
```

Create `workbook-runtime-check.ts`:

```ts
import {
  WorkPaper,
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
} from '@bilig/workpaper'

type NumericCell = {
  value: number
}

function readNumber(cell: unknown, label: string): number {
  if (typeof cell === 'object' && cell !== null && typeof (cell as NumericCell).value === 'number') {
    return (cell as NumericCell).value
  }

  throw new Error(`Expected ${label} to be numeric, got ${JSON.stringify(cell)}`)
}

const workbook = WorkPaper.buildFromSheets({
  Inputs: [
    ['Metric', 'Value'],
    ['Customers', 32],
    ['ARPA', 1200],
    ['Discount', 0.04],
  ],
  Summary: [
    ['Metric', 'Value'],
    ['Net revenue', '=Inputs!B2*Inputs!B3*(1-Inputs!B4)'],
  ],
})

const inputs = workbook.getSheetId('Inputs')
const summary = workbook.getSheetId('Summary')
if (inputs === undefined || summary === undefined) {
  throw new Error('Expected Inputs and Summary sheets')
}

const revenue = { sheet: summary, row: 1, col: 1 }
const before = readNumber(workbook.getCellValue(revenue), 'before revenue')

workbook.setCellContents({ sheet: inputs, row: 1, col: 1 }, 40)
const after = readNumber(workbook.getCellValue(revenue), 'after revenue')

const saved = serializeWorkPaperDocument(exportWorkPaperDocument(workbook, { includeConfig: true }))
const restored = createWorkPaperFromDocument(parseWorkPaperDocument(saved))
const restoredSummary = restored.getSheetId('Summary')
if (restoredSummary === undefined) {
  throw new Error('Expected restored Summary sheet')
}

const afterRestore = readNumber(restored.getCellValue({ sheet: restoredSummary, row: 1, col: 1 }), 'restored revenue')

console.log({
  before,
  after,
  afterRestore,
  verified: before === 36864 && after === 46080 && afterRestore === after,
})
```

Run it:

```sh
npx tsx workbook-runtime-check.ts
```

Expected output:

```json
{ "before": 36864, "after": 46080, "afterRestore": 46080, "verified": true }
```

That check is intentionally small. It proves the part that file libraries and
the narrow XLSX bridges do not try to own: a Node process changed an input,
read a dependent formula value, serialized the workbook document, restored it,
and read the same calculated value again.

The maintained repository example adds more workflows:

```sh
git clone https://github.com/proompteng/bilig.git
cd bilig
pnpm --dir examples/workpaper-workpaper install --ignore-workspace
pnpm --dir examples/workpaper-workpaper run agent:tool-call
pnpm --dir examples/workpaper-workpaper run agent:verify
```

The agent tool-call loop changes input cells, reads dependent formula outputs,
persists the workbook, restores it, and fails if the restored formulas or
values do not match.

## When To Combine The Tools

Use file libraries at the boundary and Bilig for the runtime model:

1. Use SheetJS, `xlsx-populate`, or ExcelJS where the product is an `.xlsx`
   file.
2. Use `@bilig/xlsx-formula-recalc` or `@bilig/exceljs-formula-recalc` when the file workflow
   needs fresh formula readback inside Node.
3. Use `@bilig/workpaper` where the product is trusted computed workbook state.
4. Keep compatibility tests around the boundary so import/export and formula
   runtime behavior are not confused.

This is the honest architecture for many services. File libraries are still
useful. Bilig earns its keep when the service needs an auditable workbook-state
transition, not just a generated spreadsheet file.

## When Not To Choose Bilig

Do not choose Bilig first if the main requirement is broad XLSX styling,
images, charts, pivot tables, or complete Excel compatibility.

Do not choose it if a cached formula result is enough and Excel can calculate
later.

Do not choose it if the workload needs a mature commercial spreadsheet-file
support channel today.

## Related Proof

- [`docs/xlsx-formula-recalculation-node.md`](xlsx-formula-recalculation-node.md)
- [`docs/exceljs-formula-recalculation-node.md`](exceljs-formula-recalculation-node.md)
- [`docs/stale-xlsx-formula-cache-node.md`](stale-xlsx-formula-cache-node.md)
- [`examples/recalc-bridge-workflows`](https://github.com/proompteng/bilig/tree/main/examples/recalc-bridge-workflows)
- [`docs/workpaper-spreadsheet-engine-comparison.md`](headless-spreadsheet-engine-comparison.md)
- [`docs/agent-spreadsheet-tool-call-loop.md`](agent-spreadsheet-tool-call-loop.md)
- [`docs/persisting-formula-backed-workpaper-documents-in-node.md`](persisting-formula-backed-workpaper-documents-in-node.md)
- [`docs/where-bilig-is-not-excel-compatible-yet.md`](where-bilig-is-not-excel-compatible-yet.md)
- [`examples/workpaper-workpaper`](https://github.com/proompteng/bilig/tree/main/examples/workpaper-workpaper)

## Sources

- SheetJS Cell Objects:
  <https://docs.sheetjs.com/docs/csf/cell/>
- SheetJS Formulae:
  <https://docs.sheetjs.com/docs/csf/features/formulae>
- SheetJS Parse Options:
  <https://docs.sheetjs.com/docs/api/parse-options>
- ExcelJS package docs:
  <https://www.npmjs.com/package/exceljs>
- npm downloads API:
  <https://api.npmjs.org/downloads/point/last-week/xlsx>
  <https://api.npmjs.org/downloads/point/last-week/exceljs>
  <https://api.npmjs.org/downloads/point/last-week/xlsx-populate>
  <https://api.npmjs.org/downloads/point/last-week/xlsx-calc>
  <https://api.npmjs.org/downloads/point/last-week/hyperformula>
  <https://api.npmjs.org/downloads/point/last-week/%40formulajs%2Fformulajs>
