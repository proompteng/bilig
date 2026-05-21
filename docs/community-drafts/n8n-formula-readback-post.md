# n8n Forum Draft: Formula Readback Without Spreadsheet UI Automation

Title:

```text
Write an Excel-style formula in n8n, recalculate it, and verify the computed value without Excel
```

Post:

```markdown
I hit a common n8n spreadsheet boundary: generating an XLSX/workbook-shaped workflow is easy, but proving the formula result changed is usually where things get awkward.

If you write `=B2*B3*B4` into a file, you still need something to calculate it. The usual options are Excel, LibreOffice, Google Sheets, or browser automation. That is a lot of moving parts for a workflow step.

I made a small importable n8n workflow that does the direct version:

1. sends one forecast input edit to a public Bilig WorkPaper endpoint
2. recalculates the dependent formulas
3. returns before/after computed values
4. exports and restores the WorkPaper JSON
5. fails the workflow unless the restored output still matches

Workflow JSON:

https://github.com/proompteng/bilig/tree/main/examples/n8n-workpaper-formula-readback

The hosted proof endpoint it calls:

```text
POST https://bilig.proompteng.ai/api/workpaper/n8n/forecast
```

Example request:

```json
{
  "sheetName": "Inputs",
  "address": "B3",
  "value": 0.4
}
```

Example proof fields:

```json
{
  "verified": true,
  "editedCell": "Inputs!B3",
  "checks": {
    "formulasPersisted": true,
    "restoredMatchesAfter": true,
    "computedOutputChanged": true
  }
}
```

This is not trying to replace Excel as a UI. The use case is backend or agent workflows where the spreadsheet is the business-logic model, and the automation needs a direct readback check before it keeps going.

I’m also working on a scoped n8n community node for the same endpoint, but the workflow above uses only built-in n8n nodes so it is easier to inspect.
```

Notes before posting:

- Post under n8n Community `Tips & Tricks` or equivalent workflow-building category.
- Do not post as a launch announcement.
- Reply only in exact-match threads where the problem is formula recalculation,
  XLSX cached formula values, or avoiding spreadsheet UI automation.
