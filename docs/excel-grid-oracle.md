# Excel Grid Oracle

This document records the evidence sources for `apps/web` grid parity.

## Product oracle

`bilig` grid parity targets Excel for the web on desktop Chrome first.

The source of truth is:

1. live Excel for the web behavior on public probe workbooks,
2. official Microsoft support documentation,
3. downloaded workbook metadata and XML only when it clarifies saved workbook state.

## Probe workbook

Primary public probe workbook:

- OneDrive workbook share used during the reverse-engineering tranche on `2026-03-18`

Observed workbook facts from the downloaded `.xlsx`:

- workbook application: `Microsoft Excel Online`
- app version: `16.0300`
- sheet count: `1`
- visible data is intentionally sparse
- saved active selection in workbook XML: `I12`
- workbook calc feature flags advertise modern Excel behavior including `LET_WF`, `LAMBDA_WF`, and `ARRAYTEXT_WF`

This confirms the workbook is useful as a product probe for Excel Online shell behavior, not as a rich business workbook to reproduce cell-for-cell.

## Official behavior references

- selection semantics: Microsoft support article for selecting ranges of cells
- formula entry semantics: Microsoft support article for creating a formula by using a function
- keyboard model: Microsoft support article for keyboard shortcuts in Excel
- row and column sizing: Microsoft support article for changing column width or row height
- error rendering: Microsoft support article for correcting a `#VALUE!` error

## Locked parity decisions

- visible target is the sheet surface and editing model, not the full ribbon
- default product chrome is only name box, `fx`, formula bar, headers, grid, sheet tabs, and status bar
- the product shell must not show a separate resolved-value box in the visible formula row
- active cell and range styling must be restrained and must not use host-level browser focus outlines
- grid gutter clicks must not select cells
- string click-away commits and visible Excel-style error rendering are required parity behaviors

## Exit gate

The oracle is considered established when:

- every Excel-visible grid behavior in `docs/excel-grid-parity-matrix.md` links to either an official doc or a live probe observation
- every shipped parity behavior has a browser acceptance test in `apps/web`
- the product shell no longer depends on legacy demo behavior or chrome for validation
