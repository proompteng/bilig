---
title: Where Bilig is not Excel-compatible yet
published: true
description: Review Bilig WorkPaper Excel compatibility gaps before importing real workbooks or relying on formula parity.
tags: excel, compatibility, spreadsheet, formulas, workpaper
canonical_url: https://proompteng.github.io/bilig/where-bilig-is-not-excel-compatible-yet.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Where bilig Is Not Excel-Compatible Yet

Status: public compatibility boundary for `@bilig/workpaper`

`bilig` is not a complete Excel clone. The current adoption wedge is narrower:
`@bilig/workpaper` gives Node services and tool hosts a workbook API with formulas,
structural edits, persistence, validation, and auditable compatibility fixtures.

This page names the main compatibility boundaries so people can evaluate the
project without reading generated contract JSON first.

## Current Evidence Snapshot

The repository keeps compatibility claims tied to checked-in fixtures and
reports:

- formula inventory breadth is `100%` for the current office-listed and tracked
  formula inventory in
  [`packages/formula/src/__tests__/fixtures/formula-compatibility-snapshot.json`](../packages/formula/src/__tests__/fixtures/formula-compatibility-snapshot.json)
- formula semantics coverage has `431` canonical fixtures and `12` workbook
  semantics fixtures, with no missing committed fixture ids in
  [`packages/benchmarks/baselines/calculation-semantics-contract.json`](../packages/benchmarks/baselines/calculation-semantics-contract.json)
- import/export fidelity passes required CSV/XLSX cases, reports no unsupported
  import/export features, and explicitly declines native macro execution in
  [`packages/benchmarks/baselines/import-export-fidelity-contract.json`](../packages/benchmarks/baselines/import-export-fidelity-contract.json)
- the headless product claim is API behavior: WorkPaper edit/readback,
  recalculation, persistence, restore, and deterministic import/export checks.

Those artifacts are useful evidence. They are not a blanket promise that every
Excel workbook, every formula argument shape, every UI interaction, or every
third-party file behaves exactly like desktop Excel.

If you need to preflight a specific workbook before integrating it, run the
[Workbook Compatibility Report](workbook-compatibility-report.md):

```sh
npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- workbook-compatibility-report workbook.xlsx --json
```

That report lists unsupported functions, external links, macro payloads,
pivots, volatile formulas, stored formula results, and risk reasons. It is
not an Excel compatibility certification and does not include a compatibility
score.

## The Biggest Non-Goals

### Native macro execution

`bilig` does not execute VBA or spreadsheet macro code.

The XLSM path detects macro-enabled workbooks, preserves safe workbook cells,
preserves the original VBA payload and code names for round trips, and records a
non-execution warning. Native macro execution remains a deliberately declined
runtime feature: `xlsx.macros.execution`.

That boundary is security posture, not a missing convenience feature.

### Full Excel application parity

`@bilig/workpaper` is a workbook engine package, not a replacement for the full
Excel desktop application.

It does not claim complete parity for:

- ribbon behavior, dialog behavior, add-ins, and desktop automation surfaces
- arbitrary interactive chart editing
- arbitrary interactive PivotTable refresh behavior
- Excel's full UI collaboration surface
- every file produced by every Excel-compatible application

The current XLSX contract covers round trips for values, formulas, formats,
defined names, comments, styles, conditional formats, dimensions, merges,
freeze panes, filters, sorts, sheet protection, protected ranges, data
validations, tables, charts, pivots, multi-sheet workbooks, and macro payload
preservation. It does not turn charts and pivots into a promise of full desktop
Excel interactivity.

### Universal formula-behavior parity

The formula registry and fixture suite are broad, and the current tracked
Office formula inventory is production-routed. The formula-behavior claim is
still evidence-scoped.

The current formula semantics artifact covers the committed canonical fixtures
and workbook semantics fixtures. It should not be read as "every Excel formula
argument combination and locale/date edge case is already covered." New edge
cases should become fixtures, and unsupported deterministic formulas in an XLSX
corpus should show up as mismatches rather than being silently accepted.

### Cached XLSX result parity for arbitrary corpora

Cached-result parity is a corpus property, not a universal package guarantee.

Use:

```sh
pnpm workpaper:xlsx-corpus:check -- /path/to/xlsx-corpus
```

The verifier reads `.xlsx`, `.xlsm`, and `.xls` files and compares formula
cells against cached workbook results where that comparison is meaningful.
Missing cached results and volatile or environment-dependent formulas such as
`NOW()` and `CELL()` are counted as skipped, not as proof of parity.

For a concrete report walkthrough, see
[`docs/xlsx-corpus-verifier-walkthrough.md`](xlsx-corpus-verifier-walkthrough.md).

### Browser Claim Boundaries

The local browser grid and WorkPaper headless engine are different surfaces.

Browser performance gates in this repository are deterministic product budgets
for the local workbook shell, including large-workbook load and headed
Playwright scroll behavior. They are not cross-product benchmark claims against
Google Sheets, Microsoft Excel Web, or every real user workflow.

Do not use the headless WorkPaper benchmark to claim the browser grid is faster
than every spreadsheet UI. Keep those claims separated.

## When bilig Is A Good Fit Today

`@bilig/workpaper` is a good fit when you need:

- a Node workbook engine for formula-backed business workflows
- agent-controlled workbook edits with explicit readback
- structural edits without driving a browser UI
- JSON persistence and restore for workbook state
- compatibility reports and deterministic fixtures you can inspect and rerun
- import/export paths that surface compatibility warnings instead of hiding
  them

Start with:

- [`docs/why-agents-need-workbook-apis.md`](why-agents-need-workbook-apis.md)
- [`docs/building-a-revenue-model-with-headless-workpaper.md`](building-a-revenue-model-with-headless-workpaper.md)
- [`examples/headless-workpaper`](../examples/headless-workpaper)

## How To Improve Compatibility

The right contribution is usually not a vague "support Excel better" issue.
Use one of these shapes:

- add a minimal workbook fixture that exposes a real mismatch
- add a canonical formula fixture for a missing semantic edge
- add an XLSX round-trip case with a specific expected metadata surface
- extend the corpus verifier report when a skipped or mismatched case needs a
  clearer explanation
- add a focused public example that shows a supported workflow end to end

Small, reproducible compatibility reports are much more useful than screenshots
or broad parity claims.
