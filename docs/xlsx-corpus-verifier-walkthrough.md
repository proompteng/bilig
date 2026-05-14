# XLSX Corpus Verifier Walkthrough

Status: public verifier walkthrough for `@bilig/headless`

`bilig` should earn trust on real workbooks, not on a vague
Excel-compatible badge. The original WorkPaper XLSX corpus verifier is useful,
but it compares against cached formula results embedded in workbook files.
Those cached values can be stale.

Treat that checker as a cache diagnostic. It answers one narrow question:

> when a workbook file contains cached formula results, does `bilig` calculate
> the comparable formula cells to the same values?

It is not an accuracy verdict. A Bilig correctness bug needs a fresh
recalculation oracle: open the workbook in Microsoft Excel, force recalculation,
save a recalculated copy, and compare Bilig against that recalculated copy's
formula results.

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

## Run The Excel Oracle Harness

Use the durable harness when you need an accuracy report instead of a cache
diagnostic:

```sh
OUT=.cache/excel-oracle-evaluation
pnpm workpaper:xlsx-oracle -- prepare-oracle /path/to/workbooks "$OUT"
pnpm workpaper:xlsx-oracle -- evaluate-cache /path/to/workbooks "$OUT"
pnpm workpaper:xlsx-oracle -- evaluate-oracle /path/to/workbooks "$OUT/recalculated" "$OUT"
pnpm workpaper:xlsx-oracle -- summarize "$OUT"
```

The commands write derived files under the output folder:

- `cache-diagnostic.json`: Bilig compared with embedded XLSX cached values,
  explicitly non-authoritative
- `excel-oracle-report.json`: Bilig compared with fresh Excel-recalculated
  formula results
- `summary.md`: human-readable counts and sanitized true-mismatch samples
- `github-issues/`: optional sanitized drafts, written only for true
  Excel-oracle mismatches

If Microsoft Excel automation is unavailable, `prepare-oracle` records that and
`evaluate-oracle` marks cells as `missing_excel_oracle`. Cache-only mismatches
stay diagnostic; they are not promoted to correctness bugs.

## Run It Against Your Files

Put the files you care about in a local directory and point the verifier at it:

```sh
pnpm workpaper:xlsx-corpus:check -- /path/to/workbooks
```

The verifier accepts `.xlsx`, `.xlsm`, and `.xls` paths. For macro-enabled
workbooks, `bilig` can preserve macro payload metadata through import/export
paths, but it does not execute native macro code.

Start with workbooks that have been opened and saved by Excel, Google Sheets, or
LibreOffice if you only need cache diagnostics. For accuracy, use the Excel
oracle harness and compare against the recalculated output folder. If a file has
formulas but no fresh Excel oracle, the harness says `missing_excel_oracle`
instead of pretending it proved compatibility.

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

`matchRate` in the cache checker is `matchingFormulaCells /
comparableFormulaCells`. In the Excel oracle harness, the primary metric is
`Bilig vs fresh Excel match rate`; embedded-cache freshness is reported
separately.

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

When cache-only `mismatches` is non-empty, treat the report as a debugging
input, not an accuracy verdict:

- keep the workbook file or reduce it to a smaller fixture
- refresh the workbook through the Excel oracle harness
- record the sheet, address, formula, fresh Excel result, and Bilig result
- add a focused regression test or canonical fixture
- link the mismatch in a GitHub issue

Only open a correctness issue when a fresh Excel expected value, Bilig actual
value, formula text, and repro notes are present. The goal is not to hide
mismatches. The goal is to turn them into small, reproducible compatibility
work.

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
