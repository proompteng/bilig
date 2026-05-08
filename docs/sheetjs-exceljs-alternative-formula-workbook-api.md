# SheetJS And ExcelJS Alternative For Formula-Backed Workbook APIs

Status: public comparison guide for developers evaluating spreadsheet
automation libraries.

Research date: 2026-05-08.

If you are searching for a SheetJS alternative or ExcelJS alternative, first
split the job into two questions:

1. Do you need to read and write workbook files?
2. Do you need a service-side workbook object that recalculates formulas after
   edits and proves the computed state?

SheetJS and ExcelJS are strong tools for workbook-file workflows. `bilig` is
not trying to replace that whole layer. The useful Bilig slice is narrower:
`@bilig/headless` gives a Node service or coding agent a WorkPaper object it can
build, mutate, evaluate, persist, restore, and verify without opening Excel or a
browser grid.

## Short Version

Use SheetJS when you need broad spreadsheet-file parsing and export across
formats.

Use ExcelJS when you need to create or edit XLSX workbooks with workbook-file
features such as sheets, rows, styles, and formula records.

Use `@bilig/headless` when the service must mutate a formula-backed workbook
and read the recalculated values back in the same process.

## The Boundary That Matters

SheetJS Community Edition stores cell formulas in the `f` field and cell values
in the `v` field. Its formula docs say actual formula results in JavaScript are
handled by a SheetJS Pro formula calculator component.

ExcelJS can store formulas and supplied results, but its public package docs say
it cannot process a formula to generate a result.

Those are reasonable design choices for file-centric libraries. They become a
problem only when your app needs to change an input, recalculate dependent
cells, and reject a workflow when computed readback does not match.

That is the place to evaluate `@bilig/headless`.

## Comparison Table

| Need | Start with | Reason |
| --- | --- | --- |
| Parse many spreadsheet file formats into JavaScript data | SheetJS | It is built around file-format import/export and a common spreadsheet object model. |
| Generate XLSX reports with workbook structure and styling | ExcelJS | It focuses on reading, manipulating, and writing XLSX workbook files. |
| Store formulas in a workbook file and let Excel calculate later | SheetJS or ExcelJS | Both can represent formula text and cached or supplied values in workbook data. |
| Recalculate formulas inside a Node service after changing inputs | `@bilig/headless` | It exposes a WorkPaper runtime with formula readback after edits. |
| Give a coding agent a spreadsheet tool it can mutate and verify | `@bilig/headless` | The maintained examples prove writeback, dependent formulas, persistence, and restore. |

## Example Evaluation Path

Install the package in a scratch project:

```sh
mkdir bilig-headless-eval
cd bilig-headless-eval
npm init -y
npm pkg set type=module
npm install @bilig/headless
```

Or run the maintained example:

```sh
git clone https://github.com/proompteng/bilig.git
cd bilig/examples/headless-workpaper
npm install
npm run agent:tool-call
npm run agent:verify
```

The agent tool-call loop changes input cells, reads dependent formula outputs,
persists the workbook, restores it, and fails if the restored formulas or values
do not match.

## When To Combine The Tools

Use file libraries at the boundary and Bilig for the runtime model:

1. Use SheetJS or ExcelJS where the product is an `.xlsx` file.
2. Use `@bilig/headless` where the product is trusted computed workbook state.
3. Keep compatibility tests around the boundary so import/export and formula
   runtime behavior are not confused.

This is the honest architecture for many services. File libraries are still
useful. Bilig earns its keep when the service needs an auditable workbook state
transition, not just a generated spreadsheet file.

## When Not To Choose Bilig

Do not choose Bilig first if the main requirement is broad XLSX styling,
images, charts, pivot tables, or complete Excel compatibility.

Do not choose it if a cached formula result is enough and Excel can calculate
later.

Do not choose it if the workload needs a mature commercial spreadsheet-file
support channel today.

## Related Proof

- [`docs/headless-spreadsheet-engine-comparison.md`](headless-spreadsheet-engine-comparison.md)
- [`docs/agent-spreadsheet-tool-call-loop.md`](agent-spreadsheet-tool-call-loop.md)
- [`docs/persisting-formula-backed-workpaper-documents-in-node.md`](persisting-formula-backed-workpaper-documents-in-node.md)
- [`docs/where-bilig-is-not-excel-compatible-yet.md`](where-bilig-is-not-excel-compatible-yet.md)
- [`examples/headless-workpaper`](https://github.com/proompteng/bilig/tree/main/examples/headless-workpaper)

## Sources

- SheetJS Cell Objects:
  <https://docs.sheetjs.com/docs/csf/cell/>
- SheetJS Formulae:
  <https://docs.sheetjs.com/docs/csf/features/formulae>
- SheetJS Parse Options:
  <https://docs.sheetjs.com/docs/api/parse-options>
- ExcelJS package docs:
  <https://www.npmjs.com/package/exceljs>
