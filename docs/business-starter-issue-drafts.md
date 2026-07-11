# Business Starter Issue Drafts

These are copy-paste starter issue drafts for small public examples that make
`@bilig/workpaper` easier to evaluate for real business automation. Keep each
task small enough for one focused pull request.

## 1. Commission Payout Calculator Example

Title:

```text
docs(examples): add commission payout WorkPaper smoke
```

Outcome:

Add a runnable example that takes a few sales rows, applies a commission formula,
and prints a deterministic payout summary.

Likely files:

```text
- examples/headless-workpaper/
- examples/headless-workpaper/README.md
```

Suggested approach:

Follow the existing example shape in `examples/headless-workpaper`. Keep the
input data tiny and sanitized, use public `@bilig/workpaper` exports, and print a
JSON object with a `verified: true` field.

Acceptance proof:

```text
npm run commission-payout

Expected output includes:
{ "verified": true }
```

## 2. CSV Report Validator Example

Title:

```text
docs(examples): add CSV report validation WorkPaper smoke
```

Outcome:

Add a small example that converts CSV-shaped rows into a WorkPaper, checks
required totals or thresholds with formulas, and prints validation results.

Likely files:

```text
- examples/headless-workpaper/
- examples/headless-workpaper/README.md
```

Suggested approach:

Reuse the existing CSV-shaped example style. Keep the fixture inline or local to
the example, avoid external services, and make the pass/fail result
deterministic.

Acceptance proof:

```text
npm run csv-report-validator

Expected output includes:
{ "valid": true, "verified": true }
```

## 3. Quote Approval Workflow Example

Title:

```text
docs(examples): add quote approval threshold WorkPaper smoke
```

Outcome:

Add a compact quote approval example where inputs such as quantity, discount,
and margin produce an approve/review result.

Likely files:

```text
- examples/headless-workpaper/
- examples/headless-workpaper/README.md
```

Suggested approach:

Model the workflow as workbook cells instead of hard-coded branching. The
example should show how a service or agent can edit inputs, recalculate, and
read the approval result.

Acceptance proof:

```text
npm run quote-approval-threshold

Expected output includes:
{ "decision": "review", "verified": true }
```
