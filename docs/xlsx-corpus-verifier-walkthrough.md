# XLSX Corpus Verifier Walkthrough

Status: public verifier walkthrough for `@bilig/headless`

The WorkPaper XLSX corpus verifier answers a narrow question:

> when a workbook file contains cached formula results, does `bilig` calculate
> the comparable formula cells to the same values?

It is not a blanket Excel-compatibility claim. It is a way to turn real workbook
files into concrete matched, skipped, or mismatched evidence.

## Run The Checked-In Corpus

From the repository root:

```sh
pnpm workpaper:xlsx-corpus:check -- packages/headless/fixtures/xlsx-corpus
```

The current checked-in reduction corpus returns:

```json
{
  "summary": {
    "totalFiles": 1,
    "filesProcessed": 1,
    "failedErrors": 0,
    "failedTimeouts": 0,
    "formulaCells": 14,
    "comparableFormulaCells": 14,
    "matchingFormulaCells": 14,
    "mismatchedFormulaCells": 0,
    "ok": 1,
    "skippedFormulaCells": 0,
    "matchRate": 1
  },
  "mismatches": [],
  "skippedByReason": {
    "missing-cached-result": 0,
    "unsupported-cached-result-type": 0,
    "volatile-or-environment-dependent-formula": 0
  }
}
```

That means the fixture has `14` formula cells, all `14` had comparable cached
results, and all `14` matched.

## How To Read The Report

`formulaCells` is the total formula-cell count found in the workbook files.

`comparableFormulaCells` is the subset with a cached result that can be compared
deterministically.

`matchingFormulaCells` is the number of comparable formula cells where `bilig`
matched the cached workbook result.

`mismatchedFormulaCells` is the number of comparable formula cells where `bilig`
calculated a different result.

`skippedFormulaCells` is the number of formula cells excluded from direct
comparison. Skips are not successes. They are explicit exclusions.

`matchRate` is `matchingFormulaCells / comparableFormulaCells`.

## Skips Are Evidence Boundaries

The verifier separates skipped formulas by reason:

- `missing-cached-result`: the workbook did not contain a cached value for that
  formula cell
- `unsupported-cached-result-type`: the cached result type cannot be compared
  by the verifier
- `volatile-or-environment-dependent-formula`: the formula depends on runtime
  state, clock time, random values, file names, or another environment-specific
  value

Common examples of environment-dependent formulas are `NOW()` and
`CELL("filename")`.

Skipped formulas are useful because they prevent false confidence. A workbook
with many skipped formulas needs a narrower claim than a workbook with a high
comparable-cell count and a high match rate.

## Mismatches Are Reproduction Seeds

When `mismatches` is non-empty, treat the report as a debugging input:

- keep the workbook file or reduce it to a smaller fixture
- record the sheet, address, formula, cached result, and calculated result
- add a focused regression test or canonical fixture
- link the mismatch in a GitHub issue

The goal is not to hide mismatches. The goal is to turn them into small,
reproducible compatibility work.

## Use It On External Workbooks

Point the verifier at a file or directory:

```sh
pnpm workpaper:xlsx-corpus:check -- /path/to/workbooks
```

The verifier accepts `.xlsx`, `.xlsm`, and `.xls` paths. For macro-enabled
workbooks, `bilig` can preserve macro payload metadata through import/export
paths, but it does not execute native macro code.

## Related Public Notes

- [`docs/where-bilig-is-not-excel-compatible-yet.md`](where-bilig-is-not-excel-compatible-yet.md)
- [`packages/headless/README.md`](../packages/headless/README.md)
