---
title: Five Node.js workbook automation examples
published: true
description: Runnable WorkPaper examples for invoice totals, budget variance alerts, subscription MRR, quote approval, and fulfillment capacity planning.
tags: typescript, node, spreadsheet, opensource
canonical_url: https://proompteng.github.io/bilig/workbook-automation-examples-node.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Five Node.js workbook automation examples

These examples are for the common case where spreadsheet logic belongs inside a
service, queue worker, API route, or agent tool instead of a browser grid. Each
script builds a small WorkPaper, writes formulas, reads computed values, and
prints `verified: true` after the expected checks pass.

The examples use the published `@bilig/headless` package from npm.

## Clone and run

```sh
git clone https://github.com/proompteng/bilig.git
cd bilig/examples/headless-workpaper
npm install
npm run invoice-totals
npm run budget-variance
npm run subscription-mrr
npm run quote-approval
npm run fulfillment-capacity
```

## What each example proves

| Workflow | Command | What to inspect |
| --- | --- | --- |
| Invoice totals | `npm run invoice-totals` | line totals, subtotal, tax, grand total, and serialized formulas |
| Budget variance alerts | `npm run budget-variance` | budget vs actual rows, variance percent, and review flags |
| Subscription MRR forecast | `npm run subscription-mrr` | churn, expansion, new customers, ending MRR, and forecast formulas |
| Quote approval threshold | `npm run quote-approval` | discount amount, quote total, max line discount, and approval status |
| Fulfillment capacity plan | `npm run fulfillment-capacity` | forecast orders, required hours, available labor, capacity gap, and short days |

## Output shape

The full JSON output is intentionally small enough to paste into an issue,
agent trace, or service log. Current example outputs include these checks.

`npm run invoice-totals`:

```json
{
  "invoiceNumber": "INV-2026-001",
  "lineItems": 4,
  "subtotal": 1890,
  "tax": 151.2,
  "total": 2041.2,
  "verified": true
}
```

`npm run budget-variance`:

```json
{
  "flaggedDepartment": "Marketing",
  "varianceAmount": 7500,
  "variancePercent": 0.15,
  "summary": {
    "totalBudget": 185000,
    "totalActual": 196600,
    "reviewCount": 1
  },
  "verified": true
}
```

`npm run subscription-mrr`:

```json
{
  "months": 4,
  "startingMrr": 5880,
  "endingMrr": 9604.03,
  "mrrDelta": 3724.03,
  "verified": true
}
```

`npm run quote-approval`:

```json
{
  "quoteId": "Q-2026-041",
  "listTotal": 6980,
  "discountAmount": 993,
  "quoteTotal": 5987,
  "approvalRequired": "Review",
  "verified": true
}
```

`npm run fulfillment-capacity`:

```json
{
  "days": 4,
  "forecastOrders": 2020,
  "requiredHours": 61.0318,
  "availableHours": 60,
  "capacityGap": -1.0318,
  "status": "Short",
  "verified": true
}
```

## Why this matters

Most backend spreadsheet automation examples stop at printing a table. These
scripts are more useful because they keep the workbook behavior visible:

- formulas stay in the workbook instead of being rewritten as application code
- computed values are read back through the engine
- the output includes the business result and the formula row that produced it
- each workflow can become a small API response, job result, or agent tool log

If one of these workflows is close to a real system you are building, use the
[show-and-tell discussion](https://github.com/proompteng/bilig/discussions/213)
to ask for the next example. If the package saves you a workbook automation
spike, star the repo so it is easier to find later:
<https://github.com/proompteng/bilig/stargazers>.
