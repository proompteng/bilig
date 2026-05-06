import type { BuildScorecardInput } from '../gen-bilig-dominance-scorecard.ts'

export function sloMeasurement(
  id: string,
  category: BuildScorecardInput['largeWorkbookSloScorecard']['measurements'][number]['category'],
  materializedCells: number,
  actualP95: number,
  budgetP95: number,
): BuildScorecardInput['largeWorkbookSloScorecard']['measurements'][number] {
  return {
    id,
    category,
    label: id,
    materializedCells,
    corpusCaseId: null,
    metric: 'elapsedMs.p95',
    actualP95,
    budgetP95,
    gateBudgetP95: budgetP95,
    sampleCount: 3,
    passed: true,
    gatePassed: true,
  }
}

export function headedBrowserContract(
  id: string,
  category: BuildScorecardInput['largeWorkbookSloScorecard']['headedBrowserFrameP95Contracts'][number]['category'],
  materializedCells: number,
  corpusCaseId: string,
  metric: BuildScorecardInput['largeWorkbookSloScorecard']['headedBrowserFrameP95Contracts'][number]['metric'],
  budgetP95: number,
): BuildScorecardInput['largeWorkbookSloScorecard']['headedBrowserFrameP95Contracts'][number] {
  return {
    id,
    category,
    label: id,
    materializedCells,
    corpusCaseId,
    metric,
    budgetP95,
    minSampleCount: metric === 'frameMs.p95' ? 120 : 1,
    playwrightTestFile: 'e2e/tests/web-shell-scroll-performance.pw.ts',
    playwrightArtifactFile: `${id}.json`,
    command: 'pnpm test:browser:full',
    passed: true,
    findings: [],
  }
}

export function family(familyName: string, ratio: number): BuildScorecardInput['competitiveArtifact']['families'][number] {
  return {
    family: familyName,
    scorecardEligible: true,
    comparableCount: 1,
    workpaperWins: 1,
    hyperformulaWins: 0,
    worstWorkpaperToHyperFormulaMeanRatio: ratio,
    worstMeanRatioWorkload: `${familyName}-workload`,
    worstWorkpaperToHyperFormulaP95Ratio: ratio,
    worstP95RatioWorkload: `${familyName}-workload`,
  }
}

export function googleSheetsCalculationCase(
  id: string,
  formula: string,
  formulaCell: string,
  coveredFeature: string,
  value: number,
): BuildScorecardInput['googleSheetsLiveCalculationScorecard']['cases'][number] {
  return {
    id,
    formula,
    formulaCell,
    coveredFeature,
    biligValue: value,
    googleSheetsRawValue: String(value),
    googleSheetsValue: value,
    passed: true,
  }
}

export function googleSheetsLargeWorkbookSpreadsheet(
  caseId: string,
  sampleIndex: number,
): BuildScorecardInput['googleSheetsLiveLargeWorkbookScorecard']['googleSheets']['spreadsheets'][number] {
  const corpusLabel = caseId.includes('250k') ? '250k' : '100k'
  return {
    caseId,
    sampleIndex,
    spreadsheetId: `google-sheets-${corpusLabel}-sample-${String(sampleIndex)}`,
    spreadsheetUrl: `https://docs.google.com/spreadsheets/d/google-sheets-${corpusLabel}-sample-${String(sampleIndex)}`,
    title: `Google Sheets ${corpusLabel} sample ${String(sampleIndex)}`,
  }
}

export function googleSheetsLargeWorkbookCase(
  id: string,
  workload: BuildScorecardInput['googleSheetsLiveLargeWorkbookScorecard']['cases'][number]['workload'],
  corpusCaseId: BuildScorecardInput['googleSheetsLiveLargeWorkbookScorecard']['cases'][number]['corpusCaseId'],
  materializedCells: number,
): BuildScorecardInput['googleSheetsLiveLargeWorkbookScorecard']['cases'][number] {
  const biligElapsedMs = numericSummary(40)
  const googleSheetsElapsedMs = numericSummary(4_000)
  return {
    id,
    workload,
    corpusCaseId,
    materializedCells,
    sampleCount: 3,
    biligElapsedMs,
    googleSheetsElapsedMs,
    biligToGoogleSheetsMeanRatio: biligElapsedMs.mean / googleSheetsElapsedMs.mean,
    biligToGoogleSheetsP95Ratio: biligElapsedMs.p95 / googleSheetsElapsedMs.p95,
    tenXMeanAndP95: true,
    verification: {
      bilig: {
        sheetCount: 1,
        height: materializedCells / 4,
        width: 4,
        usedRangeCells: materializedCells,
        terminalAddress: corpusCaseId === 'dense-mixed-250k' ? 'C62500' : 'C25000',
        terminalValue: 500,
      },
      googleSheets: {
        sheetCount: 1,
        height: materializedCells / 4,
        width: 4,
        usedRangeCells: materializedCells,
        terminalAddress: corpusCaseId === 'dense-mixed-250k' ? 'C62500' : 'C25000',
        terminalValue: 500,
      },
      equivalent: true,
    },
    passed: true,
  }
}

export function googleSheetsRecalculationSpreadsheet(
  caseId: string,
  sampleIndex: number,
): BuildScorecardInput['googleSheetsLiveRecalculationScorecard']['googleSheets']['spreadsheets'][number] {
  const workloadLabel = caseId.replace('google-sheets-live-recalculation-', '')
  return {
    caseId,
    sampleIndex,
    spreadsheetId: `google-sheets-recalc-${workloadLabel}-sample-${String(sampleIndex)}`,
    spreadsheetUrl: `https://docs.google.com/spreadsheets/d/google-sheets-recalc-${workloadLabel}-sample-${String(sampleIndex)}`,
    title: `Google Sheets recalculation ${workloadLabel} sample ${String(sampleIndex)}`,
  }
}

export function googleSheetsRecalculationCase(
  id: string,
  workload: BuildScorecardInput['googleSheetsLiveRecalculationScorecard']['cases'][number]['workload'],
): BuildScorecardInput['googleSheetsLiveRecalculationScorecard']['cases'][number] {
  const workpaperElapsedMs = numericSummary(1)
  const googleSheetsElapsedMs = numericSummary(30)
  return {
    id,
    workload,
    fixture: {
      rowCount: 1_000,
      formulaCount: 1_000,
      materializedCells: 2_000,
    },
    sampleCount: 3,
    workpaperElapsedMs,
    googleSheetsElapsedMs,
    workpaperToGoogleSheetsMeanRatio: workpaperElapsedMs.mean / googleSheetsElapsedMs.mean,
    workpaperToGoogleSheetsP95Ratio: workpaperElapsedMs.p95 / googleSheetsElapsedMs.p95,
    tenXMeanAndP95: true,
    verification: {
      workpaper: {
        value: 500,
      },
      googleSheets: {
        value: 500,
      },
      equivalent: true,
    },
    passed: true,
  }
}

export function googleSheetsStructuralSpreadsheet(
  caseId: string,
  sampleIndex: number,
): BuildScorecardInput['googleSheetsLiveStructuralScorecard']['googleSheets']['spreadsheets'][number] {
  const operationLabel = caseId.replace('google-sheets-live-structural-', '')
  return {
    caseId,
    sampleIndex,
    spreadsheetId: `google-sheets-structural-${operationLabel}-sample-${String(sampleIndex)}`,
    spreadsheetUrl: `https://docs.google.com/spreadsheets/d/google-sheets-structural-${operationLabel}-sample-${String(sampleIndex)}`,
    title: `Google Sheets structural ${operationLabel} sample ${String(sampleIndex)}`,
  }
}

export function googleSheetsStructuralCase(
  id: string,
  operation: BuildScorecardInput['googleSheetsLiveStructuralScorecard']['cases'][number]['operation'],
  axis: BuildScorecardInput['googleSheetsLiveStructuralScorecard']['cases'][number]['axis'],
): BuildScorecardInput['googleSheetsLiveStructuralScorecard']['cases'][number] {
  const workpaperElapsedMs = numericSummary(1)
  const googleSheetsElapsedMs = numericSummary(30)
  const verification = googleSheetsStructuralVerification(operation)
  return {
    id,
    operation,
    axis,
    rowCount: 500,
    sampleCount: 3,
    workpaperElapsedMs,
    googleSheetsElapsedMs,
    workpaperToGoogleSheetsMeanRatio: workpaperElapsedMs.mean / googleSheetsElapsedMs.mean,
    workpaperToGoogleSheetsP95Ratio: workpaperElapsedMs.p95 / googleSheetsElapsedMs.p95,
    tenXMeanAndP95: true,
    verification: {
      workpaper: verification,
      googleSheets: verification,
      equivalent: true,
    },
    passed: true,
  }
}

function googleSheetsStructuralVerification(
  operation: BuildScorecardInput['googleSheetsLiveStructuralScorecard']['cases'][number]['operation'],
): Record<string, number | string> {
  switch (operation) {
    case 'insert-rows':
      return { targetCell: 'A501', value: 500 }
    case 'delete-rows':
      return { targetCell: 'A499', value: 500 }
    case 'move-rows':
      return { targetCell: 'A1', value: 251 }
    case 'insert-columns':
    case 'delete-columns':
      return { targetCell: 'A500', value: 500 }
    case 'move-columns':
      return { targetCell: 'A500', value: 1_000 }
  }
}

export function recalculationCase(
  id: string,
  workload: BuildScorecardInput['microsoftExcelLiveRecalculationScorecard']['cases'][number]['workload'],
  tenXMeanAndP95: boolean,
): BuildScorecardInput['microsoftExcelLiveRecalculationScorecard']['cases'][number] {
  const workpaperElapsedMs = numericSummary(tenXMeanAndP95 ? 1 : 3)
  const microsoftExcelElapsedMs = numericSummary(20)
  return {
    id,
    workload,
    fixture: {
      rowCount: 1_000,
      formulaCount: 1_000,
      materializedCells: 2_000,
    },
    sampleCount: 5,
    workpaperElapsedMs,
    microsoftExcelElapsedMs,
    workpaperToMicrosoftExcelMeanRatio: workpaperElapsedMs.mean / microsoftExcelElapsedMs.mean,
    workpaperToMicrosoftExcelP95Ratio: workpaperElapsedMs.p95 / microsoftExcelElapsedMs.p95,
    tenXMeanAndP95,
    verification: {
      workpaper: {
        value: 500,
      },
      microsoftExcel: {
        value: 500,
      },
      equivalent: true,
    },
    passed: true,
  }
}

export function largeWorkbookCase(
  id: string,
  workload: BuildScorecardInput['microsoftExcelLiveLargeWorkbookScorecard']['cases'][number]['workload'],
  corpusCaseId: BuildScorecardInput['microsoftExcelLiveLargeWorkbookScorecard']['cases'][number]['corpusCaseId'],
  materializedCells: number,
  tenXMeanAndP95: boolean,
): BuildScorecardInput['microsoftExcelLiveLargeWorkbookScorecard']['cases'][number] {
  const biligElapsedMs = numericSummary(tenXMeanAndP95 ? 100 : 220)
  const microsoftExcelElapsedMs = numericSummary(2_000)
  return {
    id,
    workload,
    corpusCaseId,
    materializedCells,
    sampleCount: 3,
    biligElapsedMs,
    microsoftExcelElapsedMs,
    biligToMicrosoftExcelMeanRatio: biligElapsedMs.mean / microsoftExcelElapsedMs.mean,
    biligToMicrosoftExcelP95Ratio: biligElapsedMs.p95 / microsoftExcelElapsedMs.p95,
    tenXMeanAndP95,
    verification: {
      bilig: {
        sheetCount: 1,
        terminalValue: 500,
      },
      microsoftExcel: {
        sheetCount: 1,
        terminalValue: 500,
      },
      equivalent: true,
    },
    passed: true,
  }
}

export function structuralCase(
  id: string,
  operation: BuildScorecardInput['microsoftExcelLiveStructuralScorecard']['cases'][number]['operation'],
  axis: BuildScorecardInput['microsoftExcelLiveStructuralScorecard']['cases'][number]['axis'],
): BuildScorecardInput['microsoftExcelLiveStructuralScorecard']['cases'][number] {
  return {
    id,
    operation,
    axis,
    rowCount: 500,
    sampleCount: 5,
    workpaperElapsedMs: numericSummary(1),
    microsoftExcelElapsedMs: numericSummary(20),
    workpaperToMicrosoftExcelMeanRatio: 0.05,
    workpaperToMicrosoftExcelP95Ratio: 0.05,
    tenXMeanAndP95: true,
    verification: {
      workpaper: {
        height: 500,
        width: axis === 'columns' ? 4 : 2,
        value: 500,
      },
      microsoftExcel: {
        height: 500,
        width: axis === 'columns' ? 4 : 2,
        value: 500,
      },
      equivalent: true,
    },
    passed: true,
  }
}

export function numericSummary(
  value: number,
): BuildScorecardInput['microsoftExcelLiveStructuralScorecard']['cases'][number]['workpaperElapsedMs'] {
  return {
    samples: [value, value, value, value, value],
    min: value,
    median: value,
    p95: value,
    max: value,
    mean: value,
    standardDeviation: 0,
    relativeStandardDeviation: 0,
    standardError: 0,
    confidence95: {
      low: value,
      high: value,
    },
  }
}
