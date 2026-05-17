---
title: 'Show HN: Bilig runs small formula workbooks in Node'
published: true
description: A plain maintainer note for Bilig with the npm check, benchmark artifact, limits, and concrete feedback ask.
tags: show-hn, typescript, node, spreadsheet, agents
canonical_url: https://proompteng.github.io/bilig/show-hn-formula-workbooks-node-services.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Show HN: Bilig runs small formula workbooks in Node

I built Bilig for a specific mess I kept running into: a spreadsheet has the
business logic, but the product needs the answer inside a Node service.

The usual options are awkward. Drive Excel somehow. Push the model into Google
Sheets and call an API. Reimplement the spreadsheet in application code. Or
read an `.xlsx` file and accidentally trust stale cached formula values.

Bilig is the smaller thing I wanted: keep the model as sheets and formulas,
write inputs from code, read the calculated output, and save the workbook state
as JSON so tests can replay it.

That is the whole pitch.

## Try the npm package

This starts from an empty directory and uses the published package. The version
checked by this page is `@bilig/headless@0.18.5`.

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

The line that matters is `"verified": true`. The script changes an input, reads
a formula result, serializes the workbook JSON, restores it, and gets the same
answer again. That last check is there because backend spreadsheet bugs often
show up only after the state crosses a boundary.

## What this is

Bilig is a workbook-state API for Node. It is not mainly about evaluating
`=A1+B1` in isolation. It is about this loop:

- write typed inputs into known cells
- recalculate dependent formulas
- read calculated values back from the same state
- save formulas and values as JSON
- restore the workbook later and check the answer again

The API is built around a `WorkPaper` object because the workbook state is the
artifact I want under test. Screenshots are not enough, and cached XLSX formula
values are a common footgun.

## Current numbers

The checked benchmark artifact currently says Bilig wins `77/100` comparable
workloads on mean latency against the HyperFormula-style baseline. It wins
`76/100` on both mean and p95.

The miss is not hidden: `structural-append-formula-rows` is slower at p95 by
`1.785x`. Browser grid rendering is not part of this benchmark.

Read the benchmark note:
[what the WorkPaper benchmark proves](what-workpaper-benchmark-proves.md).

## What this is not

Bilig is not Excel in Node. It does not run macros, preserve every workbook
artifact, cover every Excel formula, do collaborative editing, or win every p95
case.

If you mainly need a mature broad formula engine, start with HyperFormula. If
the problem is XLSX reading, writing, or styling, start with SheetJS or ExcelJS.
If the product is a shared hosted spreadsheet, use Google Sheets.

Use `@bilig/headless` when your Node code can own the workbook state and you
care about formula readback, persistence, and restore checks.

## What would help

I am looking for rejection reasons:

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
I maintain Bilig. I built it for the annoying case where a spreadsheet owns a
small piece of business logic, but the product needs the answer in a Node
service.

The package gives you a WorkPaper object: write input cells, recalculate, read
output cells, serialize the workbook JSON, restore it, and test that the answer
is still the same after the boundary.

The quick npm check starts from an empty directory and does exactly that.

It is not Excel in Node. No macros, no full XLSX preservation claim, no full
Excel compatibility claim. If you need a mature broad formula engine,
HyperFormula is probably the first thing to test. If you need file manipulation,
start with SheetJS or ExcelJS.

The current benchmark artifact says 77/100 mean-latency wins on comparable
workloads, with the p95 miss called out on the page.

I am looking for rejection reasons from people who have shipped this kind of
thing: missing formulas, XLSX cases, bad API shape, runtime pain, or the
benchmark that would make you trust or reject it faster.
```
