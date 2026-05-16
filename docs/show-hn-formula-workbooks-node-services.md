---
title: Show HN proof for formula workbooks in Node services
published: true
description: A compact launch proof for Bilig: a runnable @bilig/headless npm check, honest benchmark evidence, known limits, and a feedback ask for Hacker News-style evaluators.
tags: show-hn, typescript, node, spreadsheet, agents
canonical_url: https://proompteng.github.io/bilig/show-hn-formula-workbooks-node-services.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Show HN proof for formula workbooks in Node services

Bilig is a TypeScript WorkPaper runtime for backend and agent workflows where a
calculation is easiest to review as cells and formulas, but it needs to run from
Node instead of from Excel, Google Sheets, or a browser automation script.

Use it when code owns the workflow: pricing rules, quote approval, payout
checks, budget guardrails, import validation, and agent tools that need
read-after-write proof.

## Run the proof

This starts from an empty directory and uses the published npm package. The
current checked package proof is `@bilig/headless@0.16.21`.

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

The important line is `"verified": true`: code edited an input cell, the formula
recalculated, the workbook serialized as WorkPaper JSON, and a restored
workbook returned the same calculated value.

## What is different from a formula parser

The useful boundary is not just evaluating `=A1+B1`. A service or agent usually
needs the whole loop:

- map typed inputs to stable workbook cells
- recalculate dependent formulas after edits
- read computed values back from the workbook runtime
- persist formulas and values as JSON
- restore the workbook and prove the same output in CI

That is the reason Bilig exposes a `WorkPaper` object instead of only a scalar
formula evaluator.

## Evidence

The checked benchmark artifact currently records `73/94` mean-latency wins
against HyperFormula-style comparable workloads, and `69/94` workloads winning
both mean and p95.

The caveat is intentionally visible: `structural-insert-columns-large` is slower at
p95 by `2.311x`. Browser grid rendering is outside this benchmark.

Read the benchmark note:
[what the WorkPaper benchmark proves](what-workpaper-benchmark-proves.md).

## What it is not

Bilig is not a finished Excel clone. It does not claim full Excel formula
parity, chart fidelity, macro execution, collaborative spreadsheet editing, or
faster p95 on every workload.

Use HyperFormula first when you primarily need a mature broad formula engine.
Use SheetJS or ExcelJS first when the main job is file reading, writing, or
styling. Use Google Sheets API first when a shared hosted spreadsheet and human
collaboration are the product requirement.

Use `@bilig/headless` when a Node service or agent tool owns the workbook state
and needs formula readback, persistence, and restore proof.

## If you are evaluating it

The most useful feedback is concrete:

- the workflow you tried
- the formula or workbook shape that blocked you
- whether the npm proof worked on your machine
- the smallest example that would make adoption easier

Open feedback here:
<https://github.com/proompteng/bilig/discussions/new?category=general>.

If the proof matches a service or agent workflow you want to revisit, star or
bookmark the repository:
<https://github.com/proompteng/bilig/stargazers>.

## Shareable post

Suggested HN title:

```text
Show HN: Formula workbooks for Node services and agent tools
```

Suggested short body:

```text
I built Bilig, an open-source TypeScript WorkPaper runtime for cases where
business logic is easiest to review as spreadsheet cells and formulas, but it
needs to run in a Node service or agent tool.

The proof starts from an empty npm project, edits an input cell, reads the
recalculated formula value, serializes WorkPaper JSON, restores it, and verifies
the same output.

It is not a finished Excel clone. The benchmark and caveats are public: 73/94
mean-latency wins against HyperFormula-style comparable workloads, 69/94
mean+p95 wins, and structural-insert-columns-large is the visible p95 holdout.

I am looking for concrete adoption blockers: which backend or agent workbook
workflow would this need to support before you would try it?
```
