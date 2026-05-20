import { describe, expect, it } from 'vitest'

import { buildImportExportFidelityScorecard, validateImportExportFidelityScorecard } from '../gen-import-export-fidelity-scorecard.ts'

describe('import/export fidelity scorecard', () => {
  it('generates a checked artifact from real CSV and XLSX import/export round trips', async () => {
    const scorecard = await buildImportExportFidelityScorecard('2026-05-06T08:00:00.000Z')

    expect(scorecard).toMatchObject({
      schemaVersion: 1,
      suite: 'import-export-fidelity',
      generatedAt: '2026-05-06T08:00:00.000Z',
      summary: {
        allRequiredCasesPassed: true,
        csvRoundTripPassed: true,
        xlsxImportPassed: true,
        xlsxSnapshotRoundTripPassed: true,
        externalGoogleSheetsEvidence: 'official-docs-comparison-artifact',
        externalMicrosoftExcelEvidence: 'official-docs-comparison-artifact',
      },
    })
    expect(scorecard.source.externalImportExportComparisonArtifact).toBe(
      'packages/benchmarks/baselines/import-export-external-sheets-excel-comparison.json',
    )
    expect(scorecard.cases.map((entry) => entry.id)).toEqual([
      'csv-import-preview',
      'csv-engine-roundtrip',
      'xlsx-import-preview',
      'xlsx-snapshot-roundtrip-values-formulas-formats',
      'xlsx-snapshot-roundtrip-dimensions-merges',
      'xlsx-snapshot-roundtrip-freeze-panes',
      'xlsx-snapshot-roundtrip-filters',
      'xlsx-snapshot-roundtrip-sorts',
      'xlsx-snapshot-roundtrip-sheet-protection',
      'xlsx-snapshot-roundtrip-protected-ranges',
      'xlsx-snapshot-roundtrip-data-validations',
      'xlsx-snapshot-roundtrip-tables',
      'xlsx-snapshot-roundtrip-charts',
      'xlsx-snapshot-roundtrip-pivots',
      'xlsx-formula-context-audit',
      'xlsx-pivot-cache-semantics',
      'xlsx-external-data-provenance',
      'xlsx-macro-payload-preserved-without-execution',
      'xlsx-runtime-feature-policy-warning',
      'external-sheets-excel-import-export-comparison',
    ])
    expect(scorecard.cases.every((entry) => entry.required && entry.passed)).toBe(true)
    expect(scorecard.cases.find((entry) => entry.id === 'xlsx-macro-payload-preserved-without-execution')).toMatchObject({
      format: 'xlsx',
      direction: 'import-export-import',
      coveredFeatures: [
        'xlsx.macros.detectedNoExecution',
        'xlsx.macros.payloadRoundtrip',
        'xlsx.macros.codeNameRoundtrip',
        'xlsx.runtimeFeaturePolicyWarnings',
      ],
      missingFeatures: [],
    })
    expect(scorecard.cases.find((entry) => entry.id === 'external-sheets-excel-import-export-comparison')).toMatchObject({
      format: 'external-docs',
      direction: 'comparison',
      coveredFeatures: [
        'external.googleSheetsImportExportDocs',
        'external.microsoftExcelImportExportDocs',
        'external.sheetsExcelImportExportComparison',
      ],
    })
    expect(scorecard.summary.coveredFeatures).toEqual([
      'csv.import',
      'csv.preview',
      'csv.export',
      'csv.roundtrip',
      'xlsx.import',
      'xlsx.preview',
      'xlsx.export',
      'xlsx.roundtrip',
      'xlsx.values',
      'xlsx.formulas',
      'xlsx.formulaAudit.context',
      'xlsx.formulaAudit.cacheStatus',
      'xlsx.numberFormats',
      'xlsx.workbookProperties',
      'xlsx.calculationSettings',
      'xlsx.calculationSettings.calcChainDiagnostics',
      'xlsx.definedNames',
      'xlsx.comments',
      'xlsx.styles',
      'xlsx.conditionalFormats.roundtrip',
      'xlsx.rowColumnDimensions',
      'xlsx.merges',
      'xlsx.freezePanes.roundtrip',
      'xlsx.filters.roundtrip',
      'xlsx.sorts.roundtrip',
      'xlsx.sheetProtection.roundtrip',
      'xlsx.protectedRanges.roundtrip',
      'xlsx.dataValidations.roundtrip',
      'xlsx.tables.roundtrip',
      'xlsx.charts.roundtrip',
      'xlsx.pivots.roundtrip',
      'xlsx.pivots.cacheSemantics',
      'xlsx.pivots.externalCacheOnlySemantics',
      'xlsx.multiSheet',
      'xlsx.macros.detectedNoExecution',
      'xlsx.macros.payloadRoundtrip',
      'xlsx.macros.codeNameRoundtrip',
      'xlsx.externalData.provenance',
      'xlsx.runtimeFeaturePolicyWarnings',
      'external.googleSheetsImportExportDocs',
      'external.microsoftExcelImportExportDocs',
      'external.sheetsExcelImportExportComparison',
    ])
    expect(scorecard.summary.unsupportedFeatures).toEqual([])
    expect(scorecard.summary.declinedRuntimeFeatures).toEqual(['xlsx.macros.execution'])
    expect(scorecard.semanticLedger).toContainEqual({
      feature: 'xlsx.macros.execution',
      disposition: 'declined-runtime',
      reason: 'Bilig preserves macro payload metadata but intentionally never executes workbook macros.',
    })
    expect(scorecard.semanticLedger).toContainEqual({
      feature: 'external.sheetsExcelImportExportComparison',
      disposition: 'external',
      reason: 'Tracked by the official Google Sheets and Microsoft Excel import/export comparison artifact.',
    })
    expect(scorecard.semanticLedger).toContainEqual({
      feature: 'xlsx.values',
      disposition: 'preserved',
      reason: 'Preserved by required import/export fidelity case evidence.',
    })
  })

  it('keeps unsupported and declined import/export semantics explicit', async () => {
    const scorecard = await buildImportExportFidelityScorecard('test-generated')

    expect(scorecard.summary.unsupportedFeatures).toEqual([])
    expect(scorecard.summary.declinedRuntimeFeatures).toEqual(['xlsx.macros.execution'])
    expect(new Set(scorecard.semanticLedger.map((entry) => entry.disposition))).toEqual(
      new Set(['preserved', 'external', 'declined-runtime']),
    )
  })

  it('rejects stale artifacts missing required fidelity cases', async () => {
    const scorecard = await buildImportExportFidelityScorecard('2026-05-06T08:00:00.000Z')
    const staleScorecard = {
      ...scorecard,
      cases: scorecard.cases.filter((entry) => entry.id !== 'xlsx-snapshot-roundtrip-dimensions-merges'),
    }

    expect(() => validateImportExportFidelityScorecard(staleScorecard)).toThrow(
      'Import/export fidelity scorecard is missing required case: xlsx-snapshot-roundtrip-dimensions-merges',
    )
  })

  it('rejects artifacts whose summary feature coverage drifts from case evidence', async () => {
    const scorecard = await buildImportExportFidelityScorecard('2026-05-06T08:00:00.000Z')
    const missingFeatureScorecard = {
      ...scorecard,
      summary: {
        ...scorecard.summary,
        coveredFeatures: scorecard.summary.coveredFeatures.filter((feature) => feature !== 'xlsx.styles'),
      },
    }
    const extraFeatureScorecard = {
      ...scorecard,
      summary: {
        ...scorecard.summary,
        coveredFeatures: [...scorecard.summary.coveredFeatures, 'xlsx.unbackedClaim'],
      },
    }

    expect(() => validateImportExportFidelityScorecard(missingFeatureScorecard)).toThrow(
      'Import/export fidelity scorecard summary is missing covered feature: xlsx.styles',
    )
    expect(() => validateImportExportFidelityScorecard(extraFeatureScorecard)).toThrow(
      'Import/export fidelity scorecard summary reports uncovered feature: xlsx.unbackedClaim',
    )
  })
})
