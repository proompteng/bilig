---
title: JavaScript spreadsheet library for Node services
published: true
description: Decide whether a JavaScript spreadsheet library should be a browser grid, an XLSX file toolkit, a formula engine, or a headless WorkPaper runtime for Node.js services.
tags: javascript, typescript, node, spreadsheet, formulas, opensource
canonical_url: https://proompteng.github.io/bilig/javascript-spreadsheet-library-headless-node.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# JavaScript spreadsheet library for Node services

Most searches for a JavaScript spreadsheet library mix four different jobs:

- a browser grid where people edit cells
- an XLSX reader or writer
- a formula engine
- a backend workbook object that a service or agent can mutate and verify

Those are not the same product. A polished grid can still be the wrong backend
primitive. A great file library can still leave a service unable to recalculate
formulas before Excel opens the file. A formula-function package can be useful
without giving you a workbook document, dependency graph, persistence, or
post-write readback.

`@bilig/workpaper` fits the fourth job. It is a TypeScript WorkPaper runtime for
Node.js services, workbook automation, and coding-agent tools. It is not a
visual spreadsheet grid, and it is not a finished Excel clone.

Research date: 2026-05-13.

## Short version

If people need to edit cells in the browser, start with a grid or spreadsheet
component such as Handsontable, JSpreadsheet, AG Grid, or another UI-first
library.

If the workbook file is the product, start with SheetJS, ExcelJS, or another
XLSX-focused toolkit.

If formula calculation is the product, start with HyperFormula when mature
spreadsheet-engine coverage matters most, or Formula.js when you only need
Excel-like functions as direct JavaScript calls.

If a Node.js service or coding agent needs to create a workbook, change inputs,
read calculated cells, persist state, restore it, and prove the values after the
edit, try `@bilig/workpaper`.

## Choose by job

| Job                                                                                  | Better first tool                            | Why                                                                                                              |
| ------------------------------------------------------------------------------------ | -------------------------------------------- | ---------------------------------------------------------------------------------------------------------------- |
| Let people edit a spreadsheet-like table in React, Vue, Angular, or plain JavaScript | A browser data grid or spreadsheet component | The main problem is UI behavior: rendering, keyboard navigation, copy/paste, selection, validation, and editing. |
| Import, export, or transform spreadsheet files                                       | SheetJS or ExcelJS-style tooling             | The main problem is file interchange, workbook structure, styles, tables, or generated reports.                  |
| Evaluate many spreadsheet formulas inside JavaScript                                 | HyperFormula                                 | The main problem is calculation-engine maturity and formula coverage.                                            |
| Call individual Excel-style functions from code                                      | Formula.js                                   | The main problem is function calls, not a workbook document.                                                     |
| Put formula-backed workbook state behind an API, queue worker, or agent tool         | `@bilig/workpaper`                           | The main problem is a mutable workbook object with formula readback, persistence, restore, and verifiable edits. |

The decision gets easier when you name the user. If the user is a person at a
browser grid, use a grid. If the user is Excel, use file tooling. If the user is
another backend process or a coding agent, use a headless workbook runtime.

## Where bilig fits

`@bilig/workpaper` gives a Node process a WorkPaper object. A WorkPaper has
sheets, cells, formulas, computed values, structural operations, JSON
persistence, and restore paths. That makes it useful for service code such as:

- pricing, quote, commission, invoice, and capacity checks
- finance or operations models that still need spreadsheet formulas
- serverless routes that accept inputs and return calculated cells
- agents that must prove which cells changed and what recalculated afterward
- tests that need formula-backed fixtures instead of hand-coded math clones

The important habit is readback. Do not stop at "the service wrote a formula."
Read the calculated value from the WorkPaper, persist the document, restore it,
and read it again.

## Quick evaluation path

Start from an empty directory:

```sh
mkdir bilig-javascript-spreadsheet-eval
cd bilig-javascript-spreadsheet-eval
npm init -y
npm pkg set type=module
npm install @bilig/workpaper
```

Then run the npm-only smoke test:

- [try `@bilig/workpaper` in Node.js](try-bilig-headless-in-node.md)

The maintained repo example is the next step:

```sh
git clone https://github.com/proompteng/bilig.git
cd bilig
pnpm --dir examples/headless-workpaper install --ignore-workspace
pnpm --dir examples/headless-workpaper run start
pnpm --dir examples/headless-workpaper run agent:verify
```

`agent:verify` changes workbook inputs, checks dependent formula readback,
serializes the document, restores it, and verifies that formulas and
values survived the round trip.

## When not to use bilig

Do not choose `@bilig/workpaper` just because the phrase "spreadsheet library"
appears in a search result.

Use a browser grid when the product is human editing. Use XLSX tooling when the
product is a file. Use HyperFormula when broad formula compatibility is the
deciding constraint today. Use Formula.js when isolated function calls are
enough.

Use `@bilig/workpaper` when the product needs a workbook-shaped backend object
that can be changed, recalculated, saved, restored, and inspected by code.

## Related bilig pages

- [Node.js spreadsheet formula engine for services](node-spreadsheet-formula-engine.md)
- [Server-side spreadsheet automation in Node.js](server-side-spreadsheet-automation-node.md)
- [Evaluate Excel formulas in Node.js with TypeScript](evaluate-excel-formulas-in-node-typescript.md)
- [Headless spreadsheet engine comparison for Node services and agents](headless-spreadsheet-engine-comparison.md)
- [SheetJS and ExcelJS alternative for formula-backed workbook APIs](sheetjs-exceljs-alternative-formula-workbook-api.md)
- [Where bilig is not Excel-compatible yet](where-bilig-is-not-excel-compatible-yet.md)

## Sources

- Handsontable JavaScript data grid docs:
  <https://handsontable.com/docs/javascript-data-grid/>
- JSpreadsheet Formula Pro docs:
  <https://jspreadsheet.com/products/formulas>
- AG Grid JavaScript data grid quick start:
  <https://www.ag-grid.com/javascript-data-grid/getting-started/>
- SheetJS Community Edition docs:
  <https://docs.sheetjs.com/>
- ExcelJS package README:
  <https://www.npmjs.com/package/exceljs>
- HyperFormula official docs:
  <https://hyperformula.handsontable.com/docs/>
- Formula.js function docs:
  <https://formulajs.info/functions/>
