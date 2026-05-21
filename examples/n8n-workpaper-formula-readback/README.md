# Bilig WorkPaper Formula Readback For n8n

This example is an importable n8n workflow for the spreadsheet formula problem
n8n users keep hitting: write an input value, recalculate formulas, and verify
the computed output without opening Excel, LibreOffice, Google Sheets, or a
browser spreadsheet UI.

The workflow uses only built-in n8n nodes:

- Manual Trigger
- Code
- HTTP Request
- Code

It calls a Bilig app route. The importable workflow defaults to a local or
self-hosted app so the artifact does not depend on a hosted route that may not
be deployed yet:

```text
POST http://localhost:4321/api/workpaper/n8n/forecast
```

Change `baseUrl` in the `Choose forecast input` node when you deploy the route
behind another Bilig app URL.

The route edits one input cell in a demo forecast WorkPaper, recalculates the
summary formulas, exports and restores the WorkPaper JSON, and returns proof
that the formula output changed and survived restore.

## Import

1. Open n8n.
2. Choose Import from File.
3. Select `bilig-workpaper-formula-readback.n8n.json`.
4. Run the workflow manually.

n8n documents workflow import/export as JSON:
<https://docs.n8n.io/workflows/export-import/>.

## Expected Proof

The final node returns a compact object like:

```json
{
  "verdict": "verified",
  "editedCell": "Inputs!B3",
  "beforeExpectedArr": 60000,
  "afterExpectedArr": 96000,
  "targetGap": 5600,
  "checks": {
    "formulasPersisted": true,
    "restoredMatchesAfter": true,
    "computedOutputChanged": true
  }
}
```

Change the input in the `Choose forecast input` node if you want to test a
different editable cell:

- `B2`: qualified opportunities
- `B3`: win rate
- `B4`: average ARR
- `B5`: expansion multiplier

This is intentionally not a custom n8n node yet. It is the smallest reproducible
workflow that proves the formula-workbook value path in n8n before asking users
to install anything.
