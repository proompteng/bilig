---
title: Show HN: Bilig runs small formula workbooks in Node
published: true
description: A maintainer note for Bilig with the npm check, benchmark numbers, limits, and the feedback that would make the project more useful.
tags: show-hn, typescript, node, spreadsheet, agents
canonical_url: https://proompteng.github.io/bilig/show-hn-formula-workbooks-node-services.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Show HN: Bilig runs small formula workbooks in Node

I built Bilig for one boring case I kept running into: a pricing rule, payout
check, or import validator is easier to review as cells and formulas, but the
production code still has to run in Node.

I do not want a service clicking around Excel or Google Sheets. I want it to
load a workbook-shaped object, write a few input cells, recalculate, read the
answer back, and save the state as JSON.

## Try the npm package

This starts from an empty directory and uses the published npm package. The
current checked package version is `@bilig/headless@0.16.24`.

```sh
mkdir bilig-headless-eval
cd bilig-headless-eval
npm init -y
npm pkg set type=module
npm install @bilig/headless
npm install -D tsx typescript @types/node
curl -fsSLo quickstart.ts https://proompteng.github.io/bilig/npm-eval.ts
npx tsx quickstart.ts
```

Expected shape:

```json
{
  "before": 24000,
  "after": 38400,
  "afterRestore": 38400,
  "sheets": ["Inputs", "Summary"],
  "verified": true
}
```

The important line is `"verified": true`. The script changed an input cell,
read the recalculated formula value, saved WorkPaper JSON, restored it, and got
the same calculated output again.

## Why this is not just a formula parser

Evaluating `=A1+B1` is not the hard part. The useful loop is:

- put typed inputs into stable cells
- recalculate dependent formulas after edits
- read computed values back from the same workbook state
- save formulas and values as JSON
- restore the workbook in CI and prove the answer did not change

Bilig exposes a `WorkPaper` object because the workbook state is part of the
contract.

## Evidence

The checked benchmark artifact currently records `76/100` mean-latency wins
against HyperFormula-style comparable workloads, and `75/100` workloads winning
both mean and p95.

The caveat is visible on purpose:
`lookup-approximate-sorted-large` is slower at p95 by `2.626x`.
Browser grid rendering is not part of this benchmark.

Read the benchmark note:
[what the WorkPaper benchmark proves](what-workpaper-benchmark-proves.md).

## What it is not

Bilig is not a finished Excel clone. It does not claim full Excel formula
parity, chart fidelity, macro execution, collaborative editing, or
faster p95 on every workload.

Use HyperFormula first when you primarily need a mature broad formula engine.
Use SheetJS or ExcelJS first when the main job is file reading, writing, or
styling. Use Google Sheets API first when a shared hosted spreadsheet and human
collaboration are the product requirement.

Use `@bilig/headless` when your Node code owns the workbook state and
needs formula readback, persistence, and restore checks.

## If you are evaluating it

The most useful feedback is concrete:

- the workflow you tried
- the formula or workbook shape that blocked you
- whether the npm check worked on your machine
- the smallest example that would make you consider it for a real service

Open feedback here:
<https://github.com/proompteng/bilig/discussions/new?category=general>.

If this matches a service workflow you want to revisit later, star or bookmark
the repository:
<https://github.com/proompteng/bilig/stargazers>.

## Shareable post

Suggested HN title:

```text
Show HN: Bilig runs small formula workbooks in Node
```

Suggested short body:

```text
I maintain Bilig.

The use case is narrow: a pricing rule, payout check, or import validator is
easiest to review as cells and formulas, but the production path has to run in
Node.

The npm check starts from an empty project, edits one input cell, reads the
recalculated value, saves WorkPaper JSON, restores it, and checks the same value
again.

It is not an Excel clone. It will not run macros or preserve every weird XLSX
artifact. The current benchmark artifact says 76/100 mean wins against
HyperFormula-style comparable workloads, with the p95 misses called out.

I am looking for blunt feedback from people who have shipped spreadsheet-backed
services: missing formulas, XLSX cases, API shape, or a benchmark that would
make you reject this quickly.
```
