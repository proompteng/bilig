# @bilig/headless

WorkPaper workbook facade for `bilig`.

`@bilig/headless` exposes `WorkPaper`, a HyperFormula-style workbook API on top of
`@bilig/core`'s `SpreadsheetEngine` for server-side and headless workflows.

## Install

```sh
pnpm add @bilig/headless
```

The package is also verified in CI through clean external Node and Vite consumer projects
built from packed runtime tarballs.

## Usage

```ts
import { WorkPaper } from "@bilig/headless";

const workbook = WorkPaper.buildFromSheets({
  Sheet1: [[1, "=A1*2"]],
});

const sheetId = workbook.getSheetId("Sheet1")!;
const value = workbook.getCellValue({ sheet: sheetId, row: 0, col: 1 });
```

## Supported Workflows

- Build empty workbooks or initialize from arrays or named sheets.
- Read cell, range, and sheet values, formulas, and serialized contents.
- Mutate cells, rows, columns, sheets, and named expressions with change tracking.
- Use `batch()`, `undo()`, `redo()`, `suspendEvaluation()`, and `resumeEvaluation()`.
- Register custom functions and language translations before workbook construction.
- Use copy, cut, paste, fill-range translation, and formula normalization helpers.
- Persist and restore WorkPaper documents with:
  - `exportWorkPaperDocument()`
  - `createWorkPaperFromDocument()`
  - `serializeWorkPaperDocument()`
  - `parseWorkPaperDocument()`
- Subscribe with HyperFormula-style positional listeners through `on()`, `once()`, and `off()`.
- Subscribe with richer payload objects through `onDetailed()`, `onceDetailed()`, and `offDetailed()`.
- Use stable compatibility adapters through `graph`, `rangeMapping`, `arrayMapping`,
  `sheetMapping`, `addressMapping`, `dependencyGraph`, `evaluator`,
  `columnSearch`, and `lazilyTransformingAstService`.

## Persistence

`@bilig/headless` follows the HyperFormula-style persistence model:

- sheets are serialized as ordered sheet-content arrays
- named expressions are serialized separately
- only the JSON-safe subset of `WorkPaperConfig` is persisted automatically
- custom function plugins and callback hooks should be registered in code before restore

## Compatibility Notes

- `WorkPaper` is the canonical top-level interface.
- The facade follows HyperFormula's public workbook workflow closely, but it is not a
  byte-for-byte drop-in replacement.
- Public lookup helpers such as `getSheetId()`, `getSheetName()`,
  `simpleCellAddressFromString()`, `simpleCellRangeFromString()`, and named-expression
  reads return `undefined` on misses in the HyperFormula style.
- `@bilig/headless` keeps `bilig`'s richer change arrays and also exposes additive
  detailed-event payloads instead of cloning HyperFormula's exact exported change types.
