---
title: Show HN: Bilig runs small formula workbooks in Node
published: true
description: A plain maintainer note for Bilig with the npm check, benchmark artifact, limits, and concrete feedback ask.
tags: show-hn, typescript, node, spreadsheet, agents
canonical_url: https://proompteng.github.io/bilig/show-hn-formula-workbooks-node-services.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Show HN: Bilig runs small formula workbooks in Node

Bilig is a Node library for a boring thing I keep needing: put a small business
model in cells and formulas, then run it from code without asking somebody to
open Excel or Google Sheets.

The first targets are quote checks, payout rules, import validators, budget
gates, and little revenue models. The finance/operator side wants formulas
because they can read them. The service side needs a normal API path: write
inputs, recalculate, read the result, save the state, and test the same thing
later.

That is all I am trying to make solid.

## Try the npm package

This starts from an empty directory and uses the published package. The current
checked version is `@bilig/headless@0.16.25`.

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

The output should look like this:

```json
{
  "before": 24000,
  "after": 38400,
  "afterRestore": 38400,
  "sheets": ["Inputs", "Summary"],
  "verified": true
}
```

The main bit is `"verified": true`. The script changes an input, reads a
formula result, serializes the workbook JSON, restores it, and gets the same
answer again. That restore check is the part I care about for service code.

## What this is

Bilig is a workbook-state API for Node. The point is not evaluating `=A1+B1`
in isolation. It is the whole loop:

- write typed inputs into known cells
- recalculate dependent formulas
- read calculated values back from the same state
- save formulas and values as JSON
- restore the workbook later and check the answer again

The API is built around a `WorkPaper` object because the workbook state should
be the artifact under test. Screenshots and cached XLSX values are too easy to
fool yourself with.

## Current numbers

The checked benchmark artifact currently records `78/100` mean-latency wins on
HyperFormula-style comparable workloads, with `74/100` workloads winning on
both mean and p95.

The caveat is real: `single-formula-edit-recalc` is slower at p95 by `2.608x`.
Browser grid rendering is not measured here.

Read the benchmark note:
[what the WorkPaper benchmark proves](what-workpaper-benchmark-proves.md).

## What this is not

Bilig is not Excel in Node. It does not run macros, preserve every workbook
artifact, cover every Excel formula, do collaborative editing, or win every p95
case.

If you mainly need a mature broad formula engine, start with HyperFormula. If
the problem is XLSX reading, writing, or styling, start with SheetJS or ExcelJS.
If the product is a shared hosted spreadsheet, use Google Sheets.

Use `@bilig/headless` only when your Node code can own the workbook state and
you need formula readback, persistence, and restore checks.

## What would help

I am looking for rejection reasons, not compliments:

- a formula family that blocks a real workbook
- a workbook shape that breaks the model
- a runtime or deployment target where the package is painful
- an API shape that makes this awkward in a real service
- a benchmark you would need before trusting it

Open feedback here:
<https://github.com/proompteng/bilig/discussions/new?category=general>.

If this is a problem you might come back to, star or bookmark the repo:
<https://github.com/proompteng/bilig/stargazers>.

## Shareable post

Suggested HN title:

```text
Show HN: Bilig runs small formula workbooks in Node
```

Suggested short body:

```text
I maintain Bilig. It is a Node library for running small formula-backed
workbooks without opening Excel or Google Sheets.

The first targets are quote checks, payout rules, import validators, budget
gates, and small revenue models. The formulas are useful because non-engineers
can review them. The backend still needs a real API path: write inputs,
recalculate, read the result, save the state, and test it again later.

Bilig exposes that as a WorkPaper object. The quick npm check starts from an
empty directory, changes an input cell, reads the recalculated value, serializes
the workbook JSON, restores it, and checks the same answer again.

It is not Excel in Node. It does not run macros, preserve every XLSX artifact,
or claim full Excel compatibility. The current benchmark artifact says 78/100
mean-latency wins on HyperFormula-style comparable workloads, and the p95 miss is
called out on the page.

I am looking for rejection reasons from people who have shipped spreadsheet-ish
backend workflows: missing formulas, XLSX cases, bad API shape, runtime pain, or
the benchmark that would make you trust or reject it faster.
```
