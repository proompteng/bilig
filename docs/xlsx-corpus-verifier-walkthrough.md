# XLSX Corpus Verifier Walkthrough

Status: public verifier walkthrough for `@bilig/headless`

`bilig` should earn trust on real workbooks, not on a vague
Excel-compatible badge. The WorkPaper XLSX corpus verifier is the quickest way
to check a directory of workbook files against the cached formula results that
Excel or another spreadsheet app already wrote into those files.

It answers one narrow question:

> when a workbook file contains cached formula results, does `bilig` calculate
> the comparable formula cells to the same values?

It is not a blanket Excel-compatibility claim. It is a way to turn real workbook
files into concrete matched, skipped, or mismatched evidence.

Use it when you are deciding whether a workbook model can move into a Node.js
service, checking an upgrade before release, or preparing a small reproduction
for a formula issue.

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

## Run It Against Your Files

Put the files you care about in a local directory and point the verifier at it:

```sh
pnpm workpaper:xlsx-corpus:check -- /path/to/workbooks
```

The verifier accepts `.xlsx`, `.xlsm`, and `.xls` paths. For macro-enabled
workbooks, `bilig` can preserve macro payload metadata through import/export
paths, but it does not execute native macro code.

Start with workbooks that have been opened and saved by Excel, Google Sheets, or
LibreOffice. The cached results in the file are the comparison target. If a file
has formulas but no cached results, the verifier will say that instead of
pretending it proved compatibility.

## Put It In CI

For a service that depends on workbook logic, keep a small private corpus in the
repo or in a CI-only fixture bundle and run the verifier before release:

```sh
pnpm workpaper:xlsx-corpus:check -- ./fixtures/workbooks
```

A useful corpus is not huge. Five to ten representative workbooks are usually
better than hundreds of stale exports nobody understands. Keep the files small,
name them after the business case they cover, and include at least one workbook
for every formula family you rely on.

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
reproducible compatibility work. A good report includes the command you ran, the
summary block, and one reduced workbook if the original file is private.

## Share A Useful Result

If you are comparing `bilig` for a project, paste the short result instead of a
sales claim:

```text
Checked 7 finance workbooks with 412 comparable formula cells.
408 matched cached workbook results.
4 mismatches are reduced in issue links.
23 formulas were skipped because the files had no cached results.
```

That tells another engineer what was tested, what was not tested, and where to
look next.

## Turn A Miss Into A Contribution

Formula compatibility work is one of the easiest places to make a first
contribution because a mismatch already gives you the expected value, the
calculated value, and the formula text.

Open an issue with the smallest workbook or fixture you can share:

- [new issue](https://github.com/proompteng/bilig/issues/new/choose)
- [first-timers-only queue](https://github.com/proompteng/bilig/issues?q=is%3Aissue%20state%3Aopen%20label%3Afirst-timers-only)
- [GitHub Discussions](https://github.com/proompteng/bilig/discussions)

## Related Public Notes

- [`docs/where-bilig-is-not-excel-compatible-yet.md`](where-bilig-is-not-excel-compatible-yet.md)
- [`packages/headless/README.md`](https://github.com/proompteng/bilig/blob/main/packages/headless/README.md)
