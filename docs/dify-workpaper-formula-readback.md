# Dify WorkPaper Formula Readback

Bilig can be exposed to Dify as a small tool plugin: one tool writes a workbook
input cell, recalculates dependent formulas, and returns JSON proof that the
computed output changed and the WorkPaper document restores to the same value.

The plugin source artifact lives at:

```text
examples/dify-workpaper-formula-readback
```

It follows Dify's tool-plugin shape: `manifest.yaml`, `provider/*.yaml`, one
tool YAML file, and one Python implementation file.

## Tool

`forecast_formula_readback` calls:

```text
POST http://localhost:4321/api/workpaper/n8n/forecast
```

Set the provider `base_url` to your hosted Bilig app when this route is deployed
outside local development.

Example input:

```json
{
  "address": "B3",
  "value": 0.4
}
```

Example output:

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

## Why This Exists

Dify should orchestrate the agent workflow. Bilig should own spreadsheet formula
state: write the input, recalculate, read the computed output, and return proof.

That avoids a spreadsheet UI dependency and gives the agent a compact, auditable
tool result.

## Package

Dify documents plugin manifests and packaging through its CLI:

- Manifest: <https://docs.dify.ai/en/develop-plugin/features-and-specs/plugin-types/plugin-info-by-manifest>
- CLI: <https://docs.dify.ai/en/develop-plugin/getting-started/cli>
- Local package file: <https://docs.dify.ai/en/develop-plugin/publishing/marketplace-listing/release-by-file>

From the example directory:

```sh
uv lock
dify plugin package .
```
