# XLSX Cache Doctor Launch Guide

This is the public launch surface for the root GitHub Action:

```yaml
- uses: proompteng/bilig@v1
  with:
    workbooks: '**/*.xlsx'
    changed-files-only: true
```

Use this guide when updating GitHub Marketplace copy, release notes, demo repos,
or one-off launch posts. Keep it narrow. Bilig can do WorkPaper, MCP, and
formula-runtime work, but the current adoption wedge is simple:

**Catch stale XLSX formula caches in CI before production reads the wrong
number.**

## Marketplace Listing

- **Name:** XLSX Cache Doctor
- **Short description:** Diagnose stale cached XLSX formula values in CI without
  Excel, LibreOffice, or browser automation.
- **Category:** Code quality
- **Secondary category:** Testing
- **Pricing:** Free
- **License:** MIT
- **Repository:** https://github.com/proompteng/bilig
- **Marketplace:** https://github.com/marketplace/actions/xlsx-cache-doctor
- **Demo PR:** https://github.com/proompteng/xlsx-cache-doctor-demo/pull/1
- **Docs:** https://proompteng.github.io/bilig/xlsx-cache-doctor-github-action.html
- **Package:** https://www.npmjs.com/package/@bilig/xlsx-formula-recalc

## Search Terms

Use these sparingly in public copy. Do not stuff them into prose.

```text
xlsx, excel, formula, formula cache, stale formula, cached formula value,
github actions, ci, exceljs, sheetjs, xlsx-populate, nodejs, typescript
```

## What It Does

XLSX files can store both a formula and a cached result. If a Node job edits an
input cell with SheetJS, ExcelJS, `xlsx-populate`, or custom ZIP/XML code, the
formula cell can still contain the old cached value. A backend, API, queue job,
or test can then read the wrong number.

XLSX Cache Doctor scans committed workbooks, recalculates formula cells in Node,
and reports where the cached value disagrees with the recalculated value.

It starts report-only by default. Teams opt into blocking pull requests with
`fail-on-stale: true`.

## One-Minute Proof

No clone:

```sh
npm exec --package @bilig/xlsx-formula-recalc@latest -- xlsx-cache-doctor --demo --json
```

Expected proof shape:

```json
{
  "formulaCellCount": 1,
  "inspectedFormulaCellCount": 1,
  "staleCachedFormulaCount": 1,
  "suggestedReads": ["Summary!B2"],
  "formulas": [
    {
      "target": "Summary!B2",
      "cachedValue": 60000,
      "literalRecalculatedValue": 72000,
      "staleCachedValue": true
    }
  ],
  "inspectionCompleted": true,
  "recalculationCompleted": true,
  "excelParity": "not_proven"
}
```

Real workbook:

```sh
npm exec --package @bilig/xlsx-formula-recalc@latest -- xlsx-cache-doctor pricing.xlsx --json
```

Generate the workflow:

```sh
mkdir -p .github/workflows
npm exec --package @bilig/xlsx-formula-recalc@latest -- \
  xlsx-cache-doctor --print-github-action "**/*.xlsx" \
  > .github/workflows/xlsx-cache-doctor.yml
```

## Good Fits

- repositories that store `.xlsx` fixtures, reports, templates, pricing models,
  or spreadsheet-backed test data;
- services that edit XLSX inputs in Node and then read formula outputs;
- CI checks where a stale cached value should be visible before merge;
- teams using SheetJS, ExcelJS, or `xlsx-populate` for file I/O but not formula
  calculation.

## Not A Fit

- desktop Excel automation;
- Office macros, charts, pivots, or full Excel compatibility claims;
- private workbook uploads to a hosted service;
- workflows where the workbook is only a manual editing artifact and no backend
  reads formula outputs.

## Trust Boundaries

Say these plainly. They improve trust.

- The Action is read-only by default.
- It does not need secrets.
- It uploads JSON and Markdown artifacts for review.
- It does not claim full Excel parity.
- Unsupported formulas are reported as warnings or skipped counts, not hidden.
- `fail-on-stale` is opt-in.

## Primary Copy

### README Or Hero

```text
Catch stale XLSX formula values before production reads them.

If a Node job edits an input cell, the formula cell can still carry the old
cached value. XLSX Cache Doctor runs in CI, recalculates the workbook in Node,
and reports the cells where the saved cache disagrees with the fresh result.
```

### HN Or Lobsters Title

```text
Show HN: XLSX Cache Doctor catches stale Excel formula values in CI
```

### HN Or Lobsters Body

```text
I built a small GitHub Action for a bug that is easy to miss in Node XLSX
pipelines.

An .xlsx formula cell can store a cached result. If a script edits an input cell
with SheetJS, ExcelJS, xlsx-populate, or raw XML, the formula may still carry the
old cached value. If the backend reads that value, it can return the wrong
number without any failing test.

XLSX Cache Doctor scans workbooks in a pull request, recalculates formula cells
in Node, and reports stale cached values with JSON and Markdown artifacts. It is
report-only by default; teams can turn on fail-on-stale after they trust it.

No-clone demo:

npm exec --package @bilig/xlsx-formula-recalc@latest -- xlsx-cache-doctor --demo --json

Action:

- uses: proompteng/bilig@v1
  with:
    workbooks: "**/*.xlsx"
    changed-files-only: true

It is not a full Excel clone. Unsupported formulas are surfaced as warnings or
skipped cells. The goal is to make stale cached formula values visible before a
service trusts them.
```

### Maintainer-To-Maintainer Note

Use only after checking that there is no existing Bilig issue, PR, or discussion
on that project.

```text
Hi, I maintain Bilig / XLSX Cache Doctor. I noticed this project keeps .xlsx
fixtures or generated workbooks in the repo.

One failure mode we see in Node XLSX pipelines is a formula cell keeping an old
cached result after a script edits an input cell. I opened a small PR/example
that runs a read-only check on changed .xlsx files and uploads a JSON/Markdown
report. It starts report-only and does not require secrets.

If this is not a fit, no worries. I mainly wanted to make the stale-cache check
easy to review rather than pitch a broad spreadsheet library.
```

## Do Not Do

- Do not pitch Bilig as a generic spreadsheet engine in the first sentence.
- Do not lead with MCP or agents for this launch.
- Do not ask for stars before the proof command.
- Do not post the same note across repositories.
- Do not open another PR where a Bilig PR, issue, discussion, registry entry, or
  manual review is already active.
- Do not claim Excel parity.
- Do not hide skipped formulas or unsupported functions.

## Release Checklist

Before public launch work:

- `v1` points at the current green commit on Forgejo and GitHub.
- The Marketplace page title and description say "XLSX Cache Doctor".
- The README first command is `xlsx-cache-doctor --demo --json`.
- The demo PR uses `proompteng/bilig@v1`.
- The demo JSON shows at least one stale cached formula value.
- `@bilig/xlsx-formula-recalc@latest` matches the README proof.
- `pnpm docs:discovery:check`, `pnpm lint`, and `git diff --check` pass.
