# Overview

Bilig WorkPaper runs spreadsheet-style formulas in backend services and agent
tools. This Pipedream app lets workflows edit a WorkPaper input cell, recalc
dependent formulas, and verify readback without driving Excel, LibreOffice,
Google Sheets, or a browser UI.

# Example Use Cases

1. Gate a quote, forecast, payout, or approval workflow on a formula workbook.
2. Recalculate formula-backed business logic from webhook or CRM data.
3. Let an automation verify computed output before sending a downstream update.

# Getting Started

Use the `Verify Formula Readback` action with the default hosted Bilig endpoint
for a no-credential proof run.

Editable cells in the demo forecast:

- `B2`: qualified opportunities
- `B3`: win rate
- `B4`: average ARR
- `B5`: expansion multiplier

The action returns the edited cell, before/after computed values, and checks for
formula persistence, restored-document equality, and changed computed output.

# Troubleshooting

If `verified` is false, inspect the returned `checks` object. A usable run
should report:

- `formulasPersisted: true`
- `restoredMatchesAfter: true`
- `computedOutputChanged: true`

If you are using a self-hosted Bilig app, make sure `baseUrl` points at the app
origin and does not include the `/api/workpaper/n8n/forecast` path.
