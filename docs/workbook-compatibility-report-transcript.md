---
title: Workbook Compatibility Report transcript
published: true
description: Terminal transcript for the local workbook compatibility report and evaluator door.
tags: xlsx, formulas, compatibility, proof, transcript
canonical_url: https://proompteng.github.io/bilig/workbook-compatibility-report-transcript.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Workbook Compatibility Report Transcript

This is the local terminal proof for the workbook compatibility report. It is
not a screenshot, not a hosted upload, and not spreadsheet UI automation.

## Report command

```sh
npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- \
  workbook-compatibility-report --demo --json
```

Observed shape from `@bilig/xlsx-formula-recalc` `0.157.0`:

```json
{
  "schemaVersion": "bilig-workbook-compatibility-report.v1",
  "verified": true,
  "input": {
    "fileName": "bilig-workbook-compatibility-demo.xlsx",
    "externalWorkbookCount": 0,
    "inspectLimit": "all"
  },
  "workbook": {
    "sheetCount": 2,
    "sheetNames": ["Inputs", "Summary"],
    "nonEmptyCellCount": 14,
    "formulaCellCount": 3,
    "definedNameCount": 0,
    "tableCount": 0,
    "pivotTableCount": 0,
    "chartCount": 0,
    "macroModuleCount": 0
  },
  "findings": {
    "unsupportedFunctions": [{ "name": "CUBEVALUE", "count": 1 }],
    "externalLinks": { "count": 0, "unresolvedCount": 0, "refreshedCount": 0 },
    "macroModules": { "count": 0, "byteLength": 0 },
    "volatileFunctions": [{ "name": "NOW", "count": 1 }],
    "pivotTables": { "count": 0, "unsupportedCount": 0, "cacheOnlyCount": 0 },
    "staleCachedFormulas": { "count": 2 },
    "missingCachedFormulaValues": { "count": 1 },
    "unsupportedRecalculations": { "count": 0 },
    "warnings": ["Volatile formulas were preserved during XLSX import; cached formula values may depend on workbook calculation time."]
  },
  "cacheInspection": {
    "inspectedFormulaCellCount": 3,
    "uninspectedFormulaCellCount": 0,
    "inspectionLimit": "all",
    "suggestedReads": ["Summary!B2", "Summary!B3", "Summary!B4"]
  },
  "commandSucceeded": true,
  "inspectionCompleted": true,
  "recalculationCompleted": true,
  "excelParity": "not_proven",
  "risk": {
    "level": "high",
    "reasons": ["unsupported functions: CUBEVALUE (1)"]
  }
}
```

## Evaluator command

```sh
npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- \
  bilig-evaluate --door workbook-compatibility --json
```

The evaluator wraps the same report in `bilig-evaluator.v1`:

```json
{
  "schemaVersion": "bilig-evaluator.v1",
  "door": "workbook-compatibility",
  "doorName": "Workbook compatibility risk report",
  "verified": true,
  "evidence": {
    "riskLevel": "high",
    "unsupportedFunctions": [{ "name": "CUBEVALUE", "count": 1 }],
    "volatileFunctions": [{ "name": "NOW", "count": 1 }],
    "formulaCellCount": 3,
    "staleCachedFormulaCount": 2,
    "checks": {
      "commandSucceeded": true,
      "inspectionCompleted": true,
      "recalculationCompleted": true,
      "riskReasonsExplainFindings": true,
      "noCompatibilityScore": true,
      "unsupportedFunctionsReported": true
    }
  }
}
```

The important evaluator check is `noCompatibilityScore: true`. This proof path
must remain a risk inspector. It must not grow a `compatibilityScore`,
`excelCompatibilityPercent`, or similar field that implies a defensible Excel
parity percentage.

## Related

- [Workbook Compatibility Report](workbook-compatibility-report.md)
