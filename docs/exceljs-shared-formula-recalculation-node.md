---
title: ExcelJS shared formulas and Node.js recalculation
published: true
description: What to do when an ExcelJS or XLSX workflow has shared formulas and a Node.js service needs recalculated workbook values.
tags: typescript, node, exceljs, xlsx, shared formulas, spreadsheet
canonical_url: https://proompteng.github.io/bilig/exceljs-shared-formula-recalculation-node.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# ExcelJS shared formulas and Node.js recalculation

This page is for one narrow XLSX problem: a workbook stores repeated formulas as
Excel shared formulas, then a Node.js service needs to edit inputs and read
fresh calculated values.

That is two separate jobs:

1. Decode the XLSX shared-formula representation into ordinary per-cell formulas.
2. Recalculate the workbook state after service-side edits.

ExcelJS is useful for workbook files, but it is not a calculation engine. If
the service must return a computed answer before Excel or LibreOffice opens the
file, put a formula runtime in the path.

## What shared formulas change

In XLSX, Excel can store one master formula and let nearby cells reference it by
shared index. That keeps the file smaller, but it means a downstream formula
runtime often needs an expansion step before recalculation.

For example, a workbook may store one formula for `B2:B3`, while `B3` only says
"use shared formula 0." A runtime that expects every formula cell to contain
formula text has to translate the master formula from `B2` to `B3`.

## What Bilig does

`@bilig/headless/xlsx` imports XLSX files through the Bilig Excel import layer.
That layer reads worksheet formula XML, tracks shared-formula bases, and expands
follower cells with translated references before the snapshot reaches
`WorkPaper`.

The implementation is in:

- [`packages/excel-import/src/xlsx-formulas.ts`](https://github.com/proompteng/bilig/blob/main/packages/excel-import/src/xlsx-formulas.ts)
- [`packages/excel-import/src/__tests__/xlsx-formula-cache-roundtrip.test.ts`](https://github.com/proompteng/bilig/blob/main/packages/excel-import/src/__tests__/xlsx-formula-cache-roundtrip.test.ts)

The shared-formula regression test covers an `INDIRECT` formula where the second
row must become `INDIRECT("'"&A3&"'!A2")`, not a broken reference into `A3`.

## When this is a fit

Try `@bilig/headless` when:

- the service owns the workbook state;
- the workflow needs write, recalculate, readback, and persistence in Node;
- JSON WorkPaper state is an acceptable internal representation;
- XLSX import/export is an edge format, not the only source of truth.

Keep Excel, LibreOffice, or a dedicated XLSX pipeline in the loop when exact
file fidelity is the product requirement. Shared formulas are only one XLSX
detail; charts, pivots, macros, styles, and unusual formula families can matter
just as much.

## Minimal path

Use the runnable XLSX proof if you need to test the full loop:

```sh
git clone https://github.com/proompteng/bilig.git
cd bilig/examples/xlsx-recalculation-node
pnpm install
pnpm run smoke
```

The example imports an XLSX workbook, edits inputs, reads recalculated formulas,
exports XLSX, reimports it, and verifies that the calculated readback still
matches.

For the smaller package-only recalculation check, use the
[ExcelJS formula recalculation guide](exceljs-formula-recalculation-node.md).

## Boundary

This is not a claim that Bilig is a drop-in replacement for ExcelJS,
HyperFormula, SheetJS, Excel, or LibreOffice.

Use ExcelJS when the job is producing or modifying an `.xlsx` file. Use
HyperFormula when broad mature formula-engine coverage is the main requirement.
Use Bilig when the service needs a WorkPaper-style state object it can mutate,
recalculate, serialize, restore, and test.

If this is the exact class of bug you are trying to avoid, star or bookmark the
repository so the next backend developer can find it:
<https://github.com/proompteng/bilig/stargazers>.
