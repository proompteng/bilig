---
title: XLSX Cache Doctor proof transcript
published: true
description: Terminal transcript for proving stale cached XLSX formula values with xlsx-cache-doctor before a Node service, CI job, or agent trusts the workbook.
tags: xlsx, formulas, stale-cache, ci, proof
canonical_url: https://proompteng.github.io/bilig/xlsx-cache-doctor-proof-transcript.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# XLSX Cache Doctor Proof Transcript

Use this when a reviewer asks what the stale-cache detector actually returns on
the first run. This is a terminal transcript, not a screenshot, not spreadsheet
UI automation, and not a hosted workbook upload.

The transcript is diagnostic by default. A stale cached value is reported with
an exact cell address, cached value, recalculated value, and follow-up read
target. It does not fail CI unless the GitHub Action or wrapper is configured to
fail on stale values.

## Demo Command

```sh
npm exec --package @bilig/xlsx-formula-recalc@latest -- xlsx-cache-doctor --demo --json
```

Observed from `@bilig/xlsx-formula-recalc` `0.157.0`:

```json
{
  "schemaVersion": "xlsx-cache-doctor.v1",
  "mode": "demo",
  "input": "generated demo workbook",
  "edits": 2,
  "externalWorkbooks": 0,
  "sheetNames": ["Inputs", "Summary"],
  "formulaCellCount": 1,
  "inspectedFormulaCellCount": 1,
  "uninspectedFormulaCellCount": 0,
  "inspectionLimit": "all",
  "staleCachedFormulaCount": 1,
  "cacheStatusSummary": {
    "inspected": 1,
    "stale": 1,
    "fresh": 0,
    "missingCache": 0,
    "unsupportedRecalculation": 0
  },
  "suggestedReads": ["Summary!B2"],
  "formulas": [
    {
      "target": "Summary!B2",
      "formula": "=Inputs!B2*Inputs!B3",
      "cachedValue": 60000,
      "literalRecalculatedValue": 72000,
      "cacheStatus": "stale",
      "staleCachedValue": true
    }
  ],
  "warnings": [],
  "commandSucceeded": true,
  "inspectionCompleted": true,
  "recalculationCompleted": true,
  "excelParity": "not_proven"
}
```

## Fixture Command

The committed CI example has a real `.xlsx` fixture with 60 formula cells and
one intentionally stale cached value:

```sh
npm exec --package @bilig/xlsx-formula-recalc@latest -- \
  xlsx-cache-doctor examples/xlsx-cache-doctor-ci/fixtures/stale-pricing.xlsx --json
```

The important readback fields are:

```json
{
  "schemaVersion": "xlsx-cache-doctor.v1",
  "mode": "file",
  "input": "examples/xlsx-cache-doctor-ci/fixtures/stale-pricing.xlsx",
  "formulaCellCount": 60,
  "inspectedFormulaCellCount": 60,
  "uninspectedFormulaCellCount": 0,
  "staleCachedFormulaCount": 1,
  "suggestedReadsCount": 60,
  "stale": [
    {
      "target": "Sheet1!B61",
      "formula": "=A61*10",
      "cachedValue": 999,
      "literalRecalculatedValue": 600,
      "cacheStatus": "stale"
    }
  ],
  "commandSucceeded": true,
  "inspectionCompleted": true,
  "recalculationCompleted": true,
  "excelParity": "not_proven"
}
```

That means the workbook stored `Sheet1!B61 = 999`, but recalculation in Node
returns `600`. The next narrow proof is to read the cell your service depends
on:

```sh
npm exec --package @bilig/xlsx-formula-recalc@latest -- \
  xlsx-recalc examples/xlsx-cache-doctor-ci/fixtures/stale-pricing.xlsx \
  --read Sheet1!B61 \
  --out examples/xlsx-cache-doctor-ci/fixtures/stale-pricing.recalculated.xlsx \
  --json
```

## CI Wrapper

For pull requests, use the GitHub Action in report-only mode first:

```yaml
- uses: proompteng/bilig@v1
  with:
    workbooks: 'fixtures/**/*.xlsx'
    changed-files-only: 'true'
    fail-on-stale: 'false'
```

The action writes JSON and Markdown reports and exposes `stale-count`,
`uninspected-count`, and `suggested-reads`. Set `fail-on-stale: "true"` only
after the team agrees that stale workbook caches should block merges.

## What This Proves

- XLSX formula cells can be enumerated in Node.
- Cached values in the file are compared with recalculated values.
- Stale cells are reported with exact addresses.
- The report separates stale, fresh, missing-cache, and unsupported
  recalculation states.
- `suggestedReads` gives the next `xlsx-recalc --read` targets.

## What This Does Not Prove

This is not a complete Excel compatibility claim. It does not prove macros,
pivots, charts, unsupported formulas, external-link freshness, locale-specific
Excel behavior, or desktop Excel UI behavior.

## Related

- [Evaluate stale XLSX formula caches](eval-xlsx-cache-doctor.md)
- [XLSX Cache Doctor GitHub Action](xlsx-cache-doctor-github-action.md)
- [Stale formula readback chooser](stale-formula-readback-chooser.md)
- [XLSX formula recalculation in Node.js](xlsx-formula-recalculation-node.md)
- [Where Bilig is not Excel-compatible yet](where-bilig-is-not-excel-compatible-yet.md)
