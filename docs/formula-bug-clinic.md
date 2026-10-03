---
title: Bilig formula bug clinic
published: true
description: Send a reduced XLSX fixture, ExcelJS shared-formula, import, formula, or WorkPaper blocker so it can become a Bilig fixture, test, docs page, or example.
tags: typescript, node, exceljs, xlsx, formulas, open source
canonical_url: https://proompteng.github.io/bilig/formula-bug-clinic.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Bilig formula bug clinic

If a workbook formula bug is blocking your Node service, send the smallest
public case that proves it. The goal is not to collect private spreadsheets.
The goal is to turn real failures into public fixtures that future evaluators
can run.

Start with the WorkPaper clinic report when the blocker is an import, formula,
persistence, or agent readback gap:

```sh
npm exec --package @bilig/workpaper@latest -- bilig-formula-clinic ./reduced.xlsx \
  --cells "Summary!B7,Inputs!B2"
```

That command runs locally, does not upload the workbook, imports the file into
WorkPaper, samples formulas, reads the requested cells, and prints a Markdown
report you can paste into the fixture form or discussion.

Good cases:

- an ExcelJS, SheetJS, or `xlsx-populate` pipeline writes inputs but cannot prove
  formula readback;
- an ExcelJS workflow writes inputs but needs recalculated output evidence;
- an XLSX uses shared formulas and the imported formula text is wrong;
- a workbook works in Excel but fails in a local Node formula runtime;
- a WorkPaper JSON restore changes a calculated value;
- an agent or MCP tool writes a cell but cannot prove the recalculated output;
- a service route needs one missing formula family, import detail, or example.

Open the fixture form when the reduced public fixture is ready:
<https://github.com/proompteng/bilig/issues/new?template=workbook_fixture.yml>.

Discuss the shape first if you are still reducing the case:
<https://github.com/proompteng/bilig/discussions/414>.

## Generate a local report

Use the narrowest command that matches the blocker:

| Blocker                                            | First local command                                                                                                 | What to paste                                                                                                         |
| -------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------- |
| WorkPaper import, formula, or persistence mismatch | `npm exec --package @bilig/workpaper@latest -- bilig-formula-clinic ./reduced.xlsx --cells "Summary!B7,Inputs!B2"`  | The Markdown clinic report with requested cells, formula samples, warnings, and actual readback.                      |
| Saved workbook compatibility question              | `npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- workbook-compatibility-report ./reduced.xlsx --json` | The compatibility report with unsupported functions, external links, volatile formulas, and inspected formula counts. |

If the workbook is already reduced, run the clinic reporter locally and paste
the Markdown output into the fixture form. It reads the file on your machine and
does not upload workbook contents.

```sh
npm exec --package @bilig/workpaper@latest -- bilig-formula-clinic ./reduced.xlsx \
  --cells "Summary!B7,Inputs!B2"
```

That is the lowest-friction path for package users. It imports the workbook,
samples formulas, reads the requested cells through WorkPaper, and prints a
Markdown report.

If you want to pin or edit the reporter script directly:

```sh
mkdir bilig-formula-clinic
cd bilig-formula-clinic
npm init -y
npm pkg set type=module
npm install @bilig/workpaper
npm install --save-dev tsx typescript @types/node
curl -fsSLo formula-clinic-report.ts \
  https://proompteng.github.io/bilig/formula-clinic-report.ts
npx tsx formula-clinic-report.ts ./reduced.xlsx \
  --cells "Summary!B7,Inputs!B2"
```

Use `--cells` for the output cells that prove the bug. The report includes
import warnings, formula samples, requested readback, and a fixture checklist.

## What to send

Send one reduced public fixture, not the whole production workbook.

Include:

- package version or commit tested;
- sheet names and exact cells or ranges;
- formulas involved;
- input values before and after the edit;
- expected output from Excel, LibreOffice, Graph, an existing service, or a
  manual check;
- actual Bilig output, import error, unsupported function, or missing API;
- the shortest command or script that maintainers can run.

Do not attach confidential workbooks, customer data, financial models, or files
that cannot be redistributed in a public test corpus. Replace names and numbers
with neutral values while keeping the same formula shape.

Good discussion summary:

```text
Reduced XLSX fixture attached or linked. The clinic report reads Summary!B7
as "review" after the input edit, while Excel returns "approved" for the same
fixture. The service should return "approved".
```

Bad discussion summary:

```text
My spreadsheet is wrong. Can Bilig support it?
```

## Why this helps

A reduced workbook fixture is better than a broad bug report because it gives
maintainers something concrete to merge:

- a regression test;
- an XLSX import/export corpus case;
- a formula compatibility note;
- a WorkPaper JSON persistence fixture;
- a service-route example;
- an MCP or agent-tool transcript.

When a case lands, the issue can point to the commit, release, and docs page
that fixed it. That is the evidence a skeptical backend developer can inspect
before adopting the package.

## Fast local check

For saved workbook boundaries, first inspect compatibility risks:

```sh
npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- \
  workbook-compatibility-report ./reduced.xlsx --json
```

If the report shows unsupported formulas, external links, macros, pivots, or
volatile formulas, include that output with the reduced fixture.

For a pure WorkPaper case, reduce it to a script:

```sh
mkdir bilig-fixture-check
cd bilig-fixture-check
npm init -y
npm pkg set type=module
npm install @bilig/workpaper
npm install --save-dev tsx typescript @types/node
```

If the script is short enough to paste into an issue, it is probably a good
fixture.

## Useful references

- [Submit a workbook fixture](submit-workbook-fixture.md)
- [ExcelJS shared formulas and Node.js recalculation](exceljs-shared-formula-recalculation-node.md)
- [XLSX formula recalculation in Node.js](xlsx-formula-recalculation-node.md)
- [Where Bilig is not Excel-compatible yet](where-bilig-is-not-excel-compatible-yet.md)

If this helped you reduce a workbook bug but a gap still blocks adoption,
open one concrete blocker or fixture note:
<https://github.com/proompteng/bilig/discussions/new?category=general>.
