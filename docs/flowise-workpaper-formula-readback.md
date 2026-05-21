# Flowise WorkPaper Formula Readback

Bilig can be used from Flowise as a custom tool. The agent sends one forecast
input edit, Bilig recalculates dependent formulas, and the tool returns a JSON
string with computed readback proof.

The Flowise custom-tool artifact lives at:

```text
examples/flowise-workpaper-formula-readback/bilig-workpaper-formula-readback.flowise-tool.json
```

It uses the Flowise custom tool JSON shape:

- `name`
- `description`
- `schema` as a stringified array
- `func` as JavaScript

Flowise documents custom tools here:
<https://docs.flowiseai.com/integrations/langchain/tools/custom-tool>.

## Tool Inputs

```json
{
  "baseUrl": "http://localhost:4321",
  "sheetName": "Inputs",
  "address": "B3",
  "valueJson": "0.4"
}
```

The tool calls:

```text
POST http://localhost:4321/api/workpaper/n8n/forecast
```

Use a hosted Bilig app URL in `baseUrl` after this route is deployed outside
local development.

and returns:

```json
{
  "verified": true,
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

Attach the imported tool to a Flowise Tool Agent when the agent needs
spreadsheet-style formula state but should not control Excel or a spreadsheet
UI.
