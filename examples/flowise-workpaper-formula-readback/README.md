# Bilig WorkPaper Formula Readback For Flowise

This example is a standalone Flowise custom-tool template. Import it as a
Flowise tool, attach it to a Tool Agent, and the agent can call Bilig WorkPaper
for spreadsheet formula readback without driving Excel, Google Sheets, or a
browser grid.

The tool uses Flowise's built-in Custom Tool pattern and `node-fetch`, matching
the official marketplace tool examples.

## Import

1. Open Flowise.
2. Open Tools.
3. Import or load `bilig-workpaper-formula-readback.flowise-tool.json`.
4. Add the tool to a Tool Agent.
5. Ask the agent to update the forecast win rate to `0.4` and report the
   formula readback proof.

Flowise documents custom tools here:
<https://docs.flowiseai.com/integrations/langchain/tools/custom-tool>.

## Tool Inputs

| Input | Example | Meaning |
| --- | --- | --- |
| `baseUrl` | `http://localhost:4321` | Local or self-hosted Bilig app base URL |
| `sheetName` | `Inputs` | Public demo input sheet |
| `address` | `B3` | One of `B2`, `B3`, `B4`, or `B5` |
| `valueJson` | `0.4` | JSON value to write |

## Output

The tool returns a JSON string with the edited cell, before/after outputs, and
checks:

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
