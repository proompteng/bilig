# Overview

Write one forecast input cell in Bilig WorkPaper, recalculate dependent
formulas, and verify that the computed output changed and survives a
WorkPaper JSON export/restore cycle.

# Example Use Cases

1. Check a quote or forecast calculation before sending an approval request.
2. Turn CRM or webhook input into a formula-backed computed result.
3. Replace fragile spreadsheet UI automation with a direct workbook API call.

# Getting Started

Use the hosted demo endpoint for a no-credential proof run:

```text
https://bilig.proompteng.ai
```

The default input writes `0.4` to `Inputs!B3`, which changes the forecast win
rate and returns recalculated forecast output.

# Troubleshooting

The action fails if Bilig does not return all required proof checks:

- formulas persisted in the exported WorkPaper JSON
- restored WorkPaper output matched the post-edit output
- at least one computed output changed
