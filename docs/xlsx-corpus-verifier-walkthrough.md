---
title: Do not trust stale XLSX cached formula values
published: true
description: A TypeScript harness for separating XLSX import success, cache diagnostics, timeouts, and real formula accuracy against fresh Microsoft Excel recalculation.
tags: typescript, node, excel, xlsx, spreadsheet, testing
canonical_url: https://proompteng.github.io/bilig/xlsx-corpus-verifier-walkthrough.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Do not trust stale XLSX cached formula values

An `.xlsx` file can contain formula text and a cached result from the last app
that saved it. That cached result is convenient for previewing a workbook
without recalculating it. It is not proof that the formula still returns that
value.

For example, a stale cached XLSX value can differ from a fresh Excel
recalculation even when the underlying formula is correct.

That distinction matters when you are evaluating `@bilig/workpaper` for a Node.js
service, an agent tool, or a workbook automation job. A stale cache can make a
correct engine look wrong. It can also make a wrong engine look correct.

Use two reports:

- cache diagnostic: compare `bilig` with embedded XLSX cached values
- Excel oracle: compare `bilig` with a workbook copy freshly recalculated and
  saved by Microsoft Excel

Only the second report is an accuracy verdict.

## The rule

Do not call something a Bilig accuracy bug unless the expected value came from a
fresh recalculation oracle.

For this harness, the preferred oracle is Microsoft Excel:

1. Open the workbook.
2. Force a full recalculation.
3. Save a recalculated copy into a local output folder.
4. Compare Bilig output with that recalculated copy's formula results.

OpenPyXL, SheetJS, and similar libraries are useful for extracting formulas and
cached values. They are not used here as formula-calculation oracles.

If Excel is unavailable, the harness marks cells as `missing_excel_oracle`.
Cache-only mismatches stay diagnostic.

## Why stale caches happen

XLSX cached formula values can drift for ordinary reasons:

- a workbook was saved before dependent cells changed
- formulas were edited by a tool that did not recalculate
- the workbook uses manual calculation mode
- an external-link value was saved into the file
- an app wrote formula text but left an old cached result behind
- Excel rewrote a function into an unsupported UDF wrapper during recalculation

The last case is subtle. If the recalculated workbook no longer contains the
same formula meaning, the harness does not use that cell as an oracle. It is
reported as `missing_excel_oracle` instead of being promoted into a fake
correctness bug.

## Run the Excel oracle harness

The Excel oracle harness belongs in a workspace package, not as a root-level
one-off script. Use the packaged harness once it is installed in this repo.

The commands preserve your original files. All derived workbooks and reports go
under the output folder.

The output is deliberately split:

- `cache-diagnostic.json`: Bilig compared with embedded XLSX cached values;
  useful, but non-authoritative
- `excel-oracle-report.json`: Bilig compared with fresh Excel-recalculated
  formula results
- `summary.md`: human-readable counts and sanitized true-mismatch samples
- `github-issues/`: optional sanitized issue drafts, written only for true
  Excel-oracle mismatches

## What the summary tells you

The summary separates import, execution, cache freshness, and accuracy:

- total workbooks evaluated
- import/parser failures
- timeout failures
- total formula cells
- comparable formula cells
- Bilig vs fresh Excel match rate
- embedded-cache freshness rate
- stale-cache false positives
- real Bilig mismatches
- top formula/function families for true mismatches
- sanitized formula samples with expected Excel value and actual Bilig value

`Bilig vs fresh Excel match rate` is the primary accuracy metric.
`Embedded-cache freshness rate` is a cache-quality metric.

## What counts as a real mismatch

A correctness issue needs all of this:

- original formula text
- fresh Excel expected value
- Bilig actual value
- workbook/cell repro notes that can be shared safely

The harness only writes `github-issues/` drafts for those true Excel-oracle
mismatches. File paths, customer names, organization identifiers, and private
metadata are redacted from GitHub-ready output.

## Cache diagnostic still has value

The older corpus checker is still useful when you understand its boundary:

```sh
pnpm workpaper:xlsx-corpus:check -- /path/to/workbooks
```

It answers this narrower question:

> For formula cells that have embedded cached values, does Bilig currently
> calculate the same values?

That is helpful for triage and regression reduction. It is not enough to claim
Excel accuracy.

## Deterministic Fixture Budget

The checked reduction corpus is intentionally small and tracked in the repo.
When a new real workbook exposes a regression, reduce it to a deterministic
fixture under `packages/workpaper/fixtures/xlsx-corpus` and refresh the fixture
manifest:

```sh
pnpm workpaper:xlsx-corpus:fixtures:generate
pnpm workpaper:xlsx-corpus:fixtures:check
```

This keeps CI focused on product behavior instead of external corpus status.

## Checked-in fixture result

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
  "mismatches": []
}
```

Read that as cache-diagnostic evidence. It says Bilig matches the fixture's
embedded cached values. It does not replace the Excel oracle harness above.

## Put it in CI

For a service that depends on workbook logic, keep a small private corpus in the
repo or in a CI-only fixture bundle:

```sh
pnpm workpaper:xlsx-corpus:check -- ./fixtures/workbooks
```

Five to ten representative workbooks are usually better than hundreds of files
nobody understands. Include at least one workbook for every formula family your
service depends on.

## Turn a miss into a contribution

Formula compatibility work is one of the easiest first contributions because a
good mismatch already gives you the formula, the expected value, the actual
value, and a small repro path.

Useful links:

- [oracle design feedback thread](https://github.com/proompteng/bilig/discussions/382)
- [new issue](https://github.com/proompteng/bilig/issues/new/choose)
- [first-timers-only queue](https://github.com/proompteng/bilig/issues?q=is%3Aissue%20state%3Aopen%20label%3Afirst-timers-only)
- [GitHub Discussions](https://github.com/proompteng/bilig/discussions)
- [`@bilig/workpaper` README](https://github.com/proompteng/bilig/blob/main/packages/workpaper/README.md)
- [compatibility limits](where-bilig-is-not-excel-compatible-yet.md)
