# Where bilig Is Not Excel-Compatible Yet

Status: public compatibility boundary for `@bilig/headless`

`bilig` is not a complete Excel clone. The current adoption wedge is narrower:
`@bilig/headless` gives Node services and agents a workbook API with formulas,
structural edits, persistence, validation, and auditable benchmark artifacts.

This page names the main compatibility boundaries so people can evaluate the
project without reading a pile of benchmark JSON first.

## Current Evidence Snapshot

The repository keeps compatibility and performance claims tied to checked-in
artifacts:

- formula inventory breadth is `100%` for the current office-listed and tracked
  formula inventory in
  [`packages/benchmarks/baselines/bilig-dominance-scorecard.json`](../packages/benchmarks/baselines/bilig-dominance-scorecard.json)
- formula semantics coverage has `300` canonical fixtures and `10` workbook
  semantics fixtures, with no missing committed fixture ids in
  [`packages/benchmarks/baselines/calculation-semantics-scorecard.json`](../packages/benchmarks/baselines/calculation-semantics-scorecard.json)
- import/export fidelity passes required CSV/XLSX cases, reports no unsupported
  import/export features, and explicitly declines native macro execution in
  [`packages/benchmarks/baselines/import-export-fidelity-scorecard.json`](../packages/benchmarks/baselines/import-export-fidelity-scorecard.json)
- the headless benchmark claim is `48/57` mean wins against the current
  HyperFormula-style comparable workload scorecard, with the p95 caveat kept
  visible in
  [`docs/what-workpaper-benchmark-proves.md`](what-workpaper-benchmark-proves.md)

Those artifacts are useful evidence. They are not a blanket promise that every
Excel workbook, every formula argument shape, every UI interaction, or every
third-party file behaves exactly like desktop Excel.

## The Biggest Non-Goals

### Native macro execution

`bilig` does not execute VBA or spreadsheet macro code.

The XLSM path detects macro-enabled workbooks, preserves safe workbook cells,
preserves the original VBA payload and code names for round trips, and records a
non-execution warning. Native macro execution remains a deliberately declined
runtime feature: `xlsx.macros.execution`.

That boundary is security posture, not a missing convenience feature.

### Full Excel application parity

`@bilig/headless` is a workbook engine package, not a replacement for the full
Excel desktop application.

It does not claim complete parity for:

- ribbon behavior, dialog behavior, add-ins, and desktop automation surfaces
- arbitrary interactive chart editing
- arbitrary interactive PivotTable refresh behavior
- Excel's full UI collaboration surface
- every file produced by every Excel-compatible application

The current XLSX scorecard proves round trips for values, formulas, formats,
defined names, comments, styles, conditional formats, dimensions, merges,
freeze panes, filters, sorts, sheet protection, protected ranges, data
validations, tables, charts, pivots, multi-sheet workbooks, and macro payload
preservation. It does not turn charts and pivots into a promise of full desktop
Excel interactivity.

### Blanket formula parity

The formula registry and fixture suite are broad, but the claim is still
evidence-scoped.

The current formula semantics artifact proves the committed canonical fixtures
and workbook semantics fixtures. It should not be read as "every Excel formula
argument combination and locale/date edge case is already proven." New edge
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

### UI dominance claims

The local browser grid and WorkPaper headless engine are different surfaces.

The live browser scorecard currently covers public unauthenticated browser load
and viewport scroll timing for Google Sheets and Microsoft Excel Web. Its own
limitations say it does not cover authenticated edit latency, equivalent
tenants, every browser-cache condition, or every real user workflow.

Do not use the headless WorkPaper benchmark to claim the browser grid is faster
than every spreadsheet UI. Keep those claims separated.

## When bilig Is A Good Fit Today

`@bilig/headless` is a good fit when you need:

- a Node workbook engine for formula-backed business workflows
- agent-controlled workbook edits with explicit readback
- structural edits without driving a browser UI
- JSON persistence and restore for workbook state
- benchmark artifacts you can inspect and rerun
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
