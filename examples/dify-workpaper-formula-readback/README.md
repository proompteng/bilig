# Bilig WorkPaper Formula Readback For Dify

This is a standalone Dify tool-plugin source artifact. It gives a Dify agent or
workflow one focused spreadsheet tool: edit a forecast input cell through Bilig
WorkPaper, recalculate formulas, and return verified readback.

It intentionally lives outside the Bilig pnpm workspace. Packaging this plugin
does not require changing Bilig's root `package.json` or `pnpm-lock.yaml`.

## Tool

`forecast_formula_readback` calls:

```text
POST http://localhost:4321/api/workpaper/n8n/forecast
```

Set the provider `base_url` to your hosted Bilig app when this route is deployed
outside local development.

with:

```json
{
  "sheetName": "Inputs",
  "address": "B3",
  "value": 0.4
}
```

The tool returns JSON proof:

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

## Package

Dify packages plugins with the Dify CLI:
<https://docs.dify.ai/en/develop-plugin/getting-started/cli>.

From this directory:

```sh
uv lock
dify plugin package .
```

Dify Marketplace submissions require the plugin source plus the generated
`.difypkg` file in a directory under `langgenius/dify-plugins`, and the Dify
plugin repository documents that review flow:
<https://github.com/langgenius/dify-plugins>.

## Local Checks

```sh
python3 -m py_compile main.py provider/bilig.py tools/forecast_formula_readback.py
ruby -e 'require "yaml"; %w[manifest.yaml provider/bilig.yaml tools/forecast_formula_readback.yaml].each { |p| YAML.load_file(p) }'
```
