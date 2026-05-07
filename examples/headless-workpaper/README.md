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

## Revenue Scenarios

Run the scenario model when you want to see a multi-sheet revenue workbook,
formula-backed projections, an agent-style planning edit, and persistence
readback:

```sh
npm run scenarios
```

Expected output:

```json
{
  "beforeEdit": {
    "totalNetMrr": 119267.2,
    "annualRunRate": 1431206.4,
    "enterpriseNetMrr": 57456,
    "expansionTarget": 1688823.55,
    "scenarios": {
      "conservativeNetMrr": 107340.48,
      "expansionNetMrr": 137157.28,
      "stretchNetMrr": 161010.72
    }
  },
  "afterEdit": {
    "totalNetMrr": 136791.2,
    "annualRunRate": 1641494.4,
    "enterpriseNetMrr": 66074.4,
    "expansionTarget": 1936963.39,
    "scenarios": {
      "conservativeNetMrr": 123112.08,
      "expansionNetMrr": 157309.88,
      "stretchNetMrr": 184668.12
    }
  },
  "persistedSheets": ["Pipeline", "Summary", "Scenarios"],
  "serializedBytes": 1594
}
```

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
