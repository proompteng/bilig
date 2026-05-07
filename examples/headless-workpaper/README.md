# Headless WorkPaper Example

This example shows `@bilig/headless` running as a Node library with no browser
UI. It builds a small revenue workbook, evaluates formulas, uses a named
expression in the revenue plan, applies an agent-style edit, persists the
workbook, restores it, and prints the verified summary.

Run it outside the monorepo with the published package:

```sh
npm install
npm start
```

Expected output:

```json
{
  "initial": {
    "totalRevenue": 27300,
    "westCustomers": 30,
    "targetRevenue": 30576
  },
  "afterAgentEdit": {
    "totalRevenue": 36900,
    "westCustomers": 38,
    "enterpriseArpa": 1200,
    "targetRevenue": 41328,
    "qualifiedCustomerCounts": [20, 30, 18]
  },
  "persistedSheets": ["Deals", "Summary"],
  "persistedNamedExpressions": ["GrowthRatePercent"],
  "restoredGrowthRatePercent": 12
}
```

The repository smoke test runs this same example against packed local runtime
packages through `pnpm workpaper:smoke:external`.

## Persistence Round Trip

Run the focused persistence example when you want to see a WorkPaper document
written to disk, restored, edited, and exported again:

```sh
npm run persistence
```

Expected output:

```json
{
  "beforeSave": {
    "quarterNetMrr": 42100,
    "annualizedRunRate": 505200,
    "expansionAdjustedArr": 545616
  },
  "afterRestoreAndEdit": {
    "quarterNetMrr": 45100,
    "annualizedRunRate": 541200,
    "expansionAdjustedArr": 584496
  },
  "persistedSheets": ["Plan", "Summary"],
  "persistedNamedExpressions": ["ExpansionRatePercent"],
  "saveFileBytes": 1209
}
```
