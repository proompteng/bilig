# @bilig/n8n-nodes-workpaper

This is a Bilig WorkPaper community node for n8n.

It gives an n8n workflow one small spreadsheet-shaped operation:

1. write a forecast input cell;
2. recalculate dependent formulas in Bilig;
3. return before/after values plus proof that formula output changed and the
   exported WorkPaper JSON restores to the same result.

Use it when an n8n automation needs formula readback without driving Excel,
LibreOffice, Google Sheets, or a browser spreadsheet UI.

## Installation

This package is not published yet. The intended package name is scoped:

```sh
npm install @bilig/n8n-nodes-workpaper
```

Do not publish or install an unscoped `n8n-nodes-workpaper` package for Bilig.
The scoped package is the canonical name.

## Operations

### Forecast: Verify Formula Readback

Posts to:

```text
POST https://bilig.proompteng.ai/api/workpaper/n8n/forecast
```

Default parameters:

```json
{
  "sheetName": "Inputs",
  "address": "B3",
  "value": 0.4
}
```

The response includes:

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

## Credentials

No credentials are required for the public hosted demo endpoint.

## Compatibility

Built with the official `@n8n/node-cli` scaffold. The node is a thin HTTP
integration with no runtime dependencies, matching n8n verification guidance for
community nodes.

## Usage

1. Add the Bilig WorkPaper node to a workflow.
2. Choose `Forecast` as the resource.
3. Choose `Verify Formula Readback` as the operation.
4. Keep the default hosted base URL or point `Bilig Base URL` at your own Bilig
   app.
5. Pick an editable input cell:
   - `B2`: qualified opportunities
   - `B3`: win rate
   - `B4`: average ARR
   - `B5`: expansion multiplier
6. Use the returned `verified` and `checks` fields as a gate before the workflow
   continues.

For a no-install workflow, see:

```text
examples/n8n-workpaper-formula-readback/bilig-workpaper-formula-readback.n8n.json
```

## Resources

- [Bilig GitHub repository](https://github.com/proompteng/bilig)
- [Bilig n8n workflow example](https://github.com/proompteng/bilig/tree/main/examples/n8n-workpaper-formula-readback)
- [n8n community nodes documentation](https://docs.n8n.io/integrations/community-nodes/)

## Version history

- `0.1.0`: initial forecast formula-readback action.
