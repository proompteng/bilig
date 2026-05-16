# XLSX Formula Recalculation In Node

This example shows the concrete server-side XLSX workflow:

1. build a pricing workbook,
2. export it as `pricing-model-source.xlsx`,
3. import the XLSX through `@bilig/headless/xlsx`,
4. edit input cells through the WorkPaper API,
5. read recalculated formula values,
6. export `pricing-model-edited.xlsx`,
7. reimport the edited XLSX and verify formulas still calculate to the same values.

Run it from a clean clone:

```sh
cd examples/xlsx-recalculation-node
npm install
npm start
```

Expected output includes:

```json
{
  "before": {
    "decision": "review"
  },
  "after": {
    "decision": "approved"
  },
  "checks": {
    "decisionChanged": true,
    "exportedReimportMatchesAfter": true,
    "formulasSurvivedXlsxRoundTrip": true,
    "verified": true
  }
}
```

Generated files are written to `examples/xlsx-recalculation-node/dist/` and are
ignored by git.
