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

## Agent Tool Call Loop

Run the tool-call loop example when you want a small SDK-neutral artifact for
wrapping WorkPaper operations as agent tools. It reads a summary range, applies
a planned input edit through a `setInputCell` tool, verifies formula readback,
persists the workbook, restores it, and checks that computed outputs survive
the round trip:

```sh
npm run agent:tool-call
```

Expected output:

```json
{
  "toolCall": {
    "toolName": "setInputCell",
    "arguments": {
      "sheetName": "Inputs",
      "address": "B3",
      "value": 0.4,
      "reason": "Use the latest qualified pipeline conversion estimate."
    }
  },
  "toolResult": {
    "editedCell": "Inputs!B3",
    "before": {
      "expectedCustomers": 5,
      "expectedArr": 60000,
      "expansionArr": 66000,
      "targetGap": -34000
    },
    "after": {
      "expectedCustomers": 8,
      "expectedArr": 96000,
      "expansionArr": 105600,
      "targetGap": 5600
    },
    "verified": {
      "previousValue": 0.25,
      "newValue": 0.4,
      "formulasPersisted": true,
      "restoredMatchesAfter": true,
      "expectedArrImproved": true,
      "targetGapClosed": true
    }
  }
}
```

The actual output also includes the initial range read, formula contracts, the
restored summary, and serialized byte count.

## Agent Writeback Verification

Run the agent verification demo when you want a small artifact for the claim
that spreadsheet agents need workbook APIs, not screenshots. It applies an
agent-style assumption edit, records the exact input cells changed, verifies the
dependent formulas and readback values, persists the workbook, restores it, and
checks that formulas and outputs survived the round trip:

```sh
npm run agent:verify
```

Expected output:

```json
{
  "edits": [
    { "cell": "Assumptions!B2", "before": 500, "after": 650 },
    { "cell": "Assumptions!B3", "before": 0.08, "after": 0.1 },
    { "cell": "Assumptions!B5", "before": 1.1, "after": 1.2 }
  ],
  "before": {
    "customers": 40,
    "grossMrr": 9600,
    "expansionMrr": 10560,
    "annualizedArr": 126720,
    "arrTargetDelta": -23280
  },
  "after": {
    "customers": 65,
    "grossMrr": 15600,
    "expansionMrr": 18720,
    "annualizedArr": 224640,
    "arrTargetDelta": 74640
  },
  "restored": {
    "customers": 65,
    "grossMrr": 15600,
    "expansionMrr": 18720,
    "annualizedArr": 224640,
    "arrTargetDelta": 74640
  },
  "formulaContracts": {
    "customers": "=Assumptions!B2*Assumptions!B3",
    "grossMrr": "=B2*Assumptions!B4",
    "expansionMrr": "=B3*Assumptions!B5",
    "annualizedArr": "=B4*12",
    "arrTargetDelta": "=Plan!B5-150000"
  },
  "verified": {
    "formulasUnchanged": true,
    "formulasPersisted": true,
    "restoredMatchesAfter": true,
    "serializedBytes": 1237
  }
}
```

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

## Named Expression Update

Run the named expression example when you want to see a service or agent change
a workbook-scoped named expression, recalculate dependent formulas, persist the
workbook, restore it, and verify the restored value still matches the edited
state:

```sh
npm run named-expression
```

Expected output:

```json
{
  "verified": true,
  "namedExpression": "GrowthRatePercent",
  "before": {
    "baseRevenue": 36000,
    "growthAdjustedRevenue": 39600
  },
  "after": {
    "baseRevenue": 36000,
    "growthAdjustedRevenue": 45000
  },
  "restored": {
    "baseRevenue": 36000,
    "growthAdjustedRevenue": 45000
  },
  "namedExpressionValues": {
    "before": 10,
    "after": 25,
    "restored": 25
  },
  "persistedNamedExpressions": [
    "GrowthRatePercent"
  ],
  "restoredMatchesAfter": true
}
```

## CSV Shaped Input

Run the CSV shaped input example when you want to see how to load a simple array or CSV-shaped data into a WorkPaper, add a formula-backed summary cell, and read back the result:

```sh
node csv-shaped-input.mjs
```

Expected output:

```json
{
  "success": true,
  "totalQ1": 480
}
```

## JSON Records Input

Run the JSON records input example when a Node service or agent already has an
array of API records and needs to turn it into a formula-backed WorkPaper
without writing an import subsystem:

```sh
npm run json-records
```

Expected output:

```json
{
  "sourceRecords": 3,
  "computed": {
    "committedMrr": 39600,
    "weightedPipelineMrr": 43400,
    "westSeats": 27,
    "largestOpportunityMrr": 21600
  },
  "serializedFirstDataRow": [
    "Acme Manufacturing",
    "West",
    "Committed",
    12,
    1800,
    1,
    "=D2*E2",
    "=G2*F2"
  ],
  "verified": true
}
```
