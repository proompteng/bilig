# @bilig/xlsx-formula-recalc Agent Notes

Use this package when a Node.js task edits an `.xlsx` workbook and needs fresh
formula results before returning the file or reading output cells.

The canonical public package is `@bilig/xlsx-formula-recalc`. The unscoped
`@bilig/xlsx-formula-recalc` package remains a compatibility and search alias, but new
agent instructions should use the scoped package first.

Start with the no-project evaluator:

```sh
npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- bilig-evaluate --door workbook-compatibility --json
npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- bilig-evaluate --door xlsx-cache --json
```

Use `workbook-compatibility` when a caller asks whether a specific `.xlsx` is
safe to use with Bilig before an agent or Node service trusts it. The report is
an inspector, not a grader: it lists unsupported functions, external links,
macro payloads, pivots, volatile functions, stale caches, and risk reasons. Do
not add or claim a compatibility score.

If the task came from a SheetJS / `xlsx` stale-formula issue, the same package
also exposes a SheetJS-named command:

```sh
npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- sheetjs-recalc --demo --json
```

If the caller has a real workbook but does not know which formulas to read yet,
inspect it first without writing an output file:

```sh
npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- xlsx-recalc workbook.xlsx --inspect --json
```

Use the memorable cache-diagnosis alias for issue triage or CI:

```sh
npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- xlsx-cache-doctor workbook.xlsx --json
```

Use the workbook compatibility report for preflight risk review:

```sh
npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- workbook-compatibility-report workbook.xlsx --json
```

Use the returned `suggestedReads`, stale cached formula counts, and warnings to
decide the narrow proof command. By default the doctor recomputes every formula;
if a caller passes `--inspect-limit`, require `uninspectedFormulaCellCount: 0`
before treating the report as complete coverage.

For a real workbook, use sheet-qualified A1 targets:

```sh
npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- xlsx-recalc quote.xlsx \
  --set Inputs!B2=48 \
  --read Summary!B7 \
  --out quote.recalculated.xlsx \
  --json
```

For external workbook links, pass companion files with `--external-workbook`
or exact link bindings with `--external-workbook-target <path> <target>`:

```sh
npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- xlsx-recalc quote.xlsx \
  --external-workbook-target ./fixtures/rates-current.xlsx file:///tmp/rates.xlsx \
  --read Summary!B7 \
  --json
```

Use the production API when code has input and output file paths. The root
package path is file-backed and uses the native streaming engine first:

```ts
import { recalculateXlsxFileToFile } from '@bilig/xlsx-formula-recalc'

const result = await recalculateXlsxFileToFile('quote.xlsx', {
  outputPath: 'quote.recalculated.xlsx',
  edits: [{ target: 'Inputs!B2', value: 48 }],
  reads: ['Summary!B7'],
  engine: 'streaming-native',
})
```

The older bytes-in/bytes-out WorkPaper API is explicit legacy compatibility:
import it from `@bilig/workpaper/xlsx` and install `@bilig/workpaper` when that fallback is intentional.

Do not claim this is a full Excel clone. Review `result.warnings` and reduce
unsupported functions, external links, macros, and volatile formula cases into
fixtures before promising production behavior.
