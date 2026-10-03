---
name: @bilig/xlsx-formula-recalc
version: 0.1.0
description: Recalculate XLSX formula outputs in Node.js after cell edits without opening Excel, LibreOffice, or browser automation.
tags:
  - xlsx
  - excel
  - formula-recalculation
  - node
  - spreadsheet-automation
  - typescript
---

# XLSX Formula Recalculation

Use `@bilig/xlsx-formula-recalc` when an agent or Node.js service has changed
XLSX inputs and must read recalculated formula outputs without opening Excel,
LibreOffice, or a browser. The unscoped `@bilig/xlsx-formula-recalc` package remains a
compatibility/search alias; new agent instructions should use the scoped package
first.

## First Check

```sh
npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- bilig-evaluate --door workbook-compatibility --json
npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- bilig-evaluate --door xlsx-cache --json
```

Use `workbook-compatibility` when you need a local preflight report for a real
workbook before a Node service or agent trusts it. The report lists unsupported
functions, external links, macro payloads, pivots, volatile functions, stale
caches, and concrete risk reasons. It does not certify Excel compatibility and
must not grow a compatibility score.

The cache evaluator should print `schemaVersion: "bilig-evaluator.v1"`,
`door: "xlsx-cache"`, `verified: true`, and evidence for a stale cached
formula corrected to a `Summary!B2` value of `72000`.

For SheetJS / `xlsx` stale-formula issues, use the SheetJS-named binary from the
same package:

```sh
npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- sheetjs-recalc --demo --json
```

If you have a real workbook but do not yet know which formula cells to verify,
inspect it without writing an output file:

```sh
npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- xlsx-recalc workbook.xlsx --inspect --json
```

For CI or issue triage, use the cache-diagnosis alias:

```sh
npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- xlsx-cache-doctor workbook.xlsx --json
```

For workbook risk review, use the compatibility report:

```sh
npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- workbook-compatibility-report workbook.xlsx --json
```

The inspection output includes `formulaCellCount`, `staleCachedFormulaCount`,
`suggestedReads`, and a `nextStep.command` for the recalculation proof. Default
inspection checks every formula. If a workflow sets `--inspect-limit`, require
`uninspectedFormulaCellCount: 0` before treating the report as complete
coverage.

## Real Workbook

```sh
npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- xlsx-recalc workbook.xlsx \
  --set Inputs!B2=48 \
  --read Summary!B7 \
  --out workbook.recalculated.xlsx \
  --json
```

For linked workbooks, pass companion files:

```sh
npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- xlsx-recalc workbook.xlsx \
  --external-workbook rates.xlsx \
  --read Summary!B7 \
  --json
```

Use an exact target binding when the workbook link target differs from the local
filename:

```sh
npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- xlsx-recalc workbook.xlsx \
  --external-workbook-target ./fixtures/rates-current.xlsx file:///tmp/rates.xlsx \
  --read Summary!B7 \
  --json
```

## TypeScript

```ts
import { recalculateXlsxFileToFile } from '@bilig/xlsx-formula-recalc'

const result = await recalculateXlsxFileToFile('workbook.xlsx', {
  outputPath: 'workbook.recalculated.xlsx',
  edits: [{ target: 'Inputs!B2', value: 48 }],
  reads: ['Summary!B7'],
  engine: 'streaming-native',
})
```

For the older bytes-in/bytes-out WorkPaper path, import from
`@bilig/workpaper/xlsx` and install `@bilig/workpaper`
explicitly.

Prefer `@bilig/exceljs-formula-recalc` when the caller already owns an ExcelJS
`Workbook` object and wants read results patched back into that object.
