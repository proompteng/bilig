# @bilig/excel-import

CSV/XLSX-to-`WorkbookSnapshot` import helpers and supported-subset XLSX export helpers for bilig.

## Package Status

This package is part of the `bilig` monorepo runtime package set, but the
`@bilig/excel-import` npm name is not provisioned yet. Use it from a repository
checkout for now. The external npm install path will be documented here after
the package is published.

From the repository:

```sh
pnpm install
pnpm --filter @bilig/excel-import build
pnpm exec vitest run packages/excel-import/src/__tests__/excel-import.test.ts
```

## XLSX To WorkPaper

```ts
import { readFileSync } from "node:fs";
import { WorkPaper } from "@bilig/headless";
import { importXlsx } from "@bilig/excel-import";

const imported = importXlsx(
  new Uint8Array(readFileSync("model.xlsx")),
  "model.xlsx",
);

const workbook = WorkPaper.buildFromSnapshot(imported.snapshot, {
  evaluationTimeoutMs: 30_000,
  useColumnIndex: true,
});
```

Use `WorkPaper.buildFromSnapshot()` for imported XLSX files. It preserves the
workbook metadata that Excel formulas need, including defined names, table
metadata, and structured-reference translations. `WorkPaper.buildFromSheets()`
is intentionally metadata-free.
