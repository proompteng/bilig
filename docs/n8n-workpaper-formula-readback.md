# n8n WorkPaper Formula Readback

Use this when an n8n workflow needs spreadsheet formulas but the important
operation is not editing a visible Excel grid. The workflow writes one input,
recalculates dependent formulas, reads the computed outputs, and checks that the
WorkPaper JSON restores to the same result.

## Importable Workflow

The example workflow lives in:

```text
examples/n8n-workpaper-formula-readback/bilig-workpaper-formula-readback.n8n.json
```

It uses only built-in n8n nodes:

- Manual Trigger
- Code
- HTTP Request
- Code

n8n imports workflows as JSON, so the file can be imported directly from the
editor. See the n8n workflow import/export docs:
<https://docs.n8n.io/workflows/export-import/>.

## Proof Route

The workflow calls a Bilig app route. It defaults to the hosted demo endpoint so
someone can import it and run the proof before deploying Bilig:

```text
POST https://bilig.proompteng.ai/api/workpaper/n8n/forecast
```

Change `baseUrl` in the `Choose forecast input` node if you want to use your
own Bilig app.

Request:

```json
{
  "sheetName": "Inputs",
  "address": "B3",
  "value": 0.4
}
```

Response shape:

```json
{
  "verified": true,
  "editedCell": "Inputs!B3",
  "before": {
    "expectedArr": 60000
  },
  "after": {
    "expectedArr": 96000,
    "targetGap": 5600
  },
  "checks": {
    "formulasPersisted": true,
    "restoredMatchesAfter": true,
    "computedOutputChanged": true
  }
}
```

Editable inputs in the demo forecast WorkPaper:

| Cell | Meaning |
| --- | --- |
| `B2` | Qualified opportunities |
| `B3` | Win rate |
| `B4` | Average ARR |
| `B5` | Expansion multiplier |

## Why This Fits n8n

n8n should orchestrate the workflow. Bilig owns the formula workbook step:

1. receive one spreadsheet-shaped input edit;
2. recalculate formulas in Node;
3. return the computed readback;
4. export and restore WorkPaper JSON as proof.

That keeps the n8n surface small and reproducible before a custom community node
exists.
