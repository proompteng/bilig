# @bilig/xlsx-formula-recalc

Diagnose stale cached XLSX formula values in Node and CI, then recalculate the
cells your service actually reads without opening Excel, LibreOffice, or a
browser.

This package is a native file-backed recalculation path for the high-friction
Node XLSX workflow:

1. import an XLSX workbook,
2. edit input cells,
3. recalculate formulas,
4. read proof values,
5. export an updated XLSX.

It fits the common `xlsx-populate`, SheetJS, and template-generation case
where the file writer can create or edit the XLSX, but the Node service also
needs fresh formula readback before returning.

## If You Arrived From SheetJS or xlsx-populate

`xlsx`, SheetJS-style workbook objects, and `xlsx-populate` are good at file
I/O. They can read workbook bytes, write cells, preserve formulas, and export
an `.xlsx` artifact.

They do not make stale cached formula values fresh inside your Node process.
That is the failure behind issues and searches like:

- `xlsx-populate formula calculated value`
- `SheetJS formula result not updating`
- `xlsx formula recalculation Node.js`
- `get computed value from xlsx formula cell`

Use this package at the file boundary:

1. let your existing library produce XLSX bytes;
2. write them to the XLSX file path your service owns;
3. call `recalculateXlsxFileToFile(...)`;
4. read the proof cells from `result.reads`;
5. return the recalculated output file.

That keeps your current file-writer choice intact and adds only the missing
calculation/readback step.

For a cross-library proof, run
[`examples/recalc-bridge-workflows`](../../examples/recalc-bridge-workflows).
It edits the same workbook through SheetJS/`xlsx`, `xlsx-populate`, and
ExcelJS, then verifies that Bilig refreshes the stale formula result.

For the SheetJS-specific boundary, read
[SheetJS formula result not updating in Node.js](../../docs/sheetjs-formula-result-not-updating-node.md).
The SheetJS-named `@bilig/sheetjs-formula-recalc` package is also published for teams
that search and install through the SheetJS / `xlsx` pipeline. It uses the same
underlying recalculation implementation as this package.

```sh
npx --package @bilig/sheetjs-formula-recalc sheetjs-recalc --demo --json
```

If you are not sure whether the fix belongs at the XLSX file, SheetJS, ExcelJS,
template, CI, WorkPaper, or agent boundary, start with the
[stale formula readback chooser](https://proompteng.github.io/bilig/stale-formula-readback-chooser.html).

## Install

```sh
npm install @bilig/xlsx-formula-recalc
```

## CLI

Run the evaluator first when you want one JSON shape for XLSX stale-cache
checks, WorkPaper services, and agent/MCP tooling:

```sh
npx --package @bilig/xlsx-formula-recalc bilig-evaluate --door workbook-compatibility --json
npx --package @bilig/xlsx-formula-recalc bilig-evaluate --door xlsx-cache --json
```

Use `workbook-compatibility` when the question is whether a specific workbook
has known Bilig integration risks before a Node service or agent trusts it. The
report includes unsupported functions, external links, VBA payloads, pivots,
volatile functions, stale caches, and concrete `risk.reasons`. It does not
certify Excel compatibility and it does not print a compatibility score.

Expected evaluator shape:

```json
{
  "schemaVersion": "bilig-evaluator.v1",
  "door": "workbook-compatibility",
  "verified": true,
  "evidence": {
    "riskLevel": "high",
    "unsupportedFunctions": [{ "name": "CUBEVALUE", "count": 1 }],
    "volatileFunctions": [{ "name": "NOW", "count": 1 }],
    "formulaCellCount": 3,
    "staleCachedFormulaCount": 2
  }
}
```

Run the workbook report directly when you have a file:

```sh
npx --package @bilig/xlsx-formula-recalc workbook-compatibility-report pricing.xlsx --json
```

Run the cache-doctor check directly when you need the full formula-cache report:

```sh
npx --package @bilig/xlsx-formula-recalc xlsx-cache-doctor --demo --json
```

Expected shape:

```json
{
  "schemaVersion": "xlsx-cache-doctor.v1",
  "formulaCellCount": 1,
  "inspectedFormulaCellCount": 1,
  "uninspectedFormulaCellCount": 0,
  "staleCachedFormulaCount": 1,
  "cacheStatusSummary": {
    "inspected": 1,
    "stale": 1,
    "fresh": 0,
    "missingCache": 0,
    "unsupportedRecalculation": 0
  },
  "suggestedReads": ["Summary!B2"],
  "commandSucceeded": true,
  "inspectionCompleted": true
}
```

Use `cacheStatusSummary` and per-formula `cacheStatus` in full reports to
separate confirmed stale caches from missing cached values or formulas without a
comparable recalculated value.

When you know which cells matter, run the recalculation check:

```sh
npx --package @bilig/xlsx-formula-recalc xlsx-recalc --demo --json
```

That command creates a tiny workbook, changes `Inputs!B2` and `Inputs!B3`,
recalculates `Summary!B2`, writes `bilig-formula-recalc-demo.xlsx`, and prints
a proof object with explicit recalculation fields and the recalculated value:

```json
{
  "reads": {
    "Summary!B2": {
      "value": 72000
    }
  },
  "warnings": [],
  "commandSucceeded": true,
  "recalculationCompleted": true,
  "excelParity": "not_proven",
  "expectedReadback": {
    "Summary!B2": 72000
  },
  "expectedValueMatched": true
}
```

Keep the JSON first. It contains proof fields only; it does not include star,
release-watch, or discussion links. Use README or docs links after the
recalculated value and warnings match the workflow.

If you have a real workbook but do not yet know which formula cells matter,
diagnose it without writing an output file:

```sh
npx --package @bilig/xlsx-formula-recalc xlsx-cache-doctor pricing.xlsx --json
```

Inspection imports the workbook, lists formula cells, recomputes every formula
by default, reports stale cached values, and returns suggested `--read` targets
for the recalculation command. If you intentionally pass `--inspect-limit 50`,
the JSON includes the skipped count as `uninspectedFormulaCellCount`.

```json
{
  "schemaVersion": "xlsx-cache-doctor.v1",
  "formulaCellCount": 12,
  "inspectedFormulaCellCount": 12,
  "uninspectedFormulaCellCount": 0,
  "inspectionLimit": "all",
  "staleCachedFormulaCount": 3,
  "cacheStatusSummary": {
    "inspected": 12,
    "stale": 3,
    "fresh": 9,
    "missingCache": 0,
    "unsupportedRecalculation": 0
  },
  "suggestedReads": ["Summary!B7"],
  "formulas": [
    {
      "target": "Summary!B7",
      "formula": "=Inputs!B2*Inputs!B3",
      "cachedValue": 60000,
      "literalRecalculatedValue": 72000,
      "cacheStatus": "stale",
      "staleCachedValue": true
    }
  ],
  "commandSucceeded": true,
  "inspectionCompleted": true,
  "recalculationCompleted": true,
  "excelParity": "not_proven"
}
```

`xlsx-cache-doctor` is a readable alias for
`xlsx-recalc pricing.xlsx --inspect --json`. It is meant for issue triage, CI,
and pull-request checks where you only want to know whether committed XLSX files
carry stale cached formula values. The GitHub Action wrapper lives at
[`actions/xlsx-cache-doctor`](../../actions/xlsx-cache-doctor), and the runnable
fixture/workflow example lives at
[`examples/xlsx-cache-doctor-ci`](../../examples/xlsx-cache-doctor-ci).

To create the GitHub Actions workflow from npm:

```sh
mkdir -p .github/workflows
npx --package @bilig/xlsx-formula-recalc xlsx-cache-doctor --print-github-action "**/*.xlsx" \
  > .github/workflows/xlsx-cache-doctor.yml
```

The generated workflow is read-only and report-only by default, sets up Node.js
22, and pins the Bilig npm runtime with `package-version`. Add
`--fail-on-stale true` when stale formula caches should block pull requests.
Use `--inspect-limit`, `--json-output`, and `--markdown-output` when your first
run should sample formulas or write the JSON and Markdown reports somewhere
specific. For production, pin both the Action ref and `package-version`.

For an existing workbook:

```sh
npx --package @bilig/xlsx-formula-recalc xlsx-recalc pricing.xlsx \
  --set Inputs!B2=48 \
  --set Inputs!B3=1500 \
  --read Summary!B7 \
  --out pricing.recalculated.xlsx \
  --json
```

The CLI writes a recalculated workbook and prints readback values. Cell targets must be sheet-qualified A1 references such as `Inputs!B2` or `'Pricing Model'!F12`.

For workbooks with external links, pass companion workbook files so cached link
values can be refreshed before recalculation:

```sh
npx --package @bilig/xlsx-formula-recalc xlsx-recalc model.xlsx \
  --external-workbook rates.xlsx \
  --read Model!C1 \
  --out model.recalculated.xlsx \
  --json
```

When the link target in the workbook is an exact path or URI that does not match
the local companion filename, bind the companion explicitly:

```sh
npx --package @bilig/xlsx-formula-recalc xlsx-recalc model.xlsx \
  --external-workbook-target ./fixtures/rates-current.xlsx file:///tmp/rates.xlsx \
  --read Model!C1 \
  --json
```

Ambiguous companion matches fail closed: the command preserves existing
external-link cache values, emits a warning, and includes hydration diagnostics
in JSON output.

For a maintained external-workbook proof with companion hydration diagnostics,
run
[external workbook recalculation proof in Node.js](https://proompteng.github.io/bilig/external-workbook-recalc-proof.html).

## API

Use `inspectXlsxCacheFile` when a service or test runner needs the
cache-doctor report without shelling out to the CLI. It stays on the same
file-backed streaming-native path as `xlsx-cache-doctor`:

```ts
import { inspectXlsxCacheFile } from '@bilig/xlsx-formula-recalc'

const report = await inspectXlsxCacheFile('pricing.xlsx', {
  maxRssBytes: 350 * 1024 * 1024,
})

if (report.staleCachedFormulaCount > 0) {
  throw new Error(
    `stale XLSX formula cache: ${report.formulas
      .filter((formula) => formula.cacheStatus === 'stale')
      .map((formula) => formula.target)
      .join(', ')}`,
  )
}
```

The API returns the same `schemaVersion`, `cacheStatusSummary`, per-formula
`cacheStatus`, and `suggestedReads` fields as the JSON CLI report.

```ts
import { recalculateXlsxFileToFile } from '@bilig/xlsx-formula-recalc'

const result = await recalculateXlsxFileToFile('pricing.xlsx', {
  outputPath: 'pricing.recalculated.xlsx',
  edits: [
    { target: 'Inputs!B2', value: 48 },
    { target: 'Inputs!B3', value: 1500 },
  ],
  reads: ['Summary!B7'],
  engine: 'streaming-native',
})

console.log(result.reads['Summary!B7'])
```

External companion workbooks stay on the native file-to-file path:

```ts
import { readFile } from 'node:fs/promises'

import { recalculateXlsxFileToFile } from '@bilig/xlsx-formula-recalc'

const result = await recalculateXlsxFileToFile('model.xlsx', {
  outputPath: 'model.recalculated.xlsx',
  externalWorkbooks: [
    {
      bytes: await readFile('rates.xlsx'),
      fileName: 'rates.xlsx',
      target: 'file:///tmp/rates.xlsx',
    },
  ],
  reads: ['Model!C1'],
  engine: 'streaming-native',
})

console.log(result.diagnostics?.externalWorkbookHydration)
```

If another library already produced workbook bytes instead of a file path, use
the explicit legacy compatibility import:

```ts
import { recalculateXlsx } from '@bilig/workpaper/xlsx'

const output = await workbook.outputAsync('nodebuffer') // for example, from xlsx-populate

const result = recalculateXlsx(output, {
  reads: ['Summary!B7'],
})
```

For the full workbook API, use `@bilig/workpaper`. For the old bytes-in,
bytes-out compatibility API, import from
`@bilig/workpaper/xlsx` and install `@bilig/workpaper`
explicitly.

## Common Boundaries

| Existing tool                           | Keep using it for                                      | Add this package when                                      |
| --------------------------------------- | ------------------------------------------------------ | ---------------------------------------------------------- |
| `xlsx-populate`                         | template editing and workbook generation               | formula cells need fresh cached values in Node             |
| SheetJS / `xlsx`                        | broad XLSX parsing, writing, and file interchange      | edited inputs must update dependent formulas now           |
| ExcelJS                                 | styled reports, sheets, tables, and ExcelJS workbooks  | use `@bilig/exceljs-formula-recalc` for the ExcelJS object |
| Excel, LibreOffice, Microsoft Graph     | exact spreadsheet application behavior                 | you cannot depend on an external app or API call           |
| `@bilig/workpaper` or `bilig-workpaper` | service-owned formula workbook state with JSON storage | the workbook does not have to stay XLSX-first              |

## Scope

Use this when a Node service needs deterministic formula readback after it changes XLSX inputs. It is not a full Excel clone: unsupported Excel functions, external workbook links, macros, and volatile functions may need review. Import warnings are returned in `result.warnings`.
