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
        externalGoogleSheetsEvidence: 'not-captured',
        externalMicrosoftExcelEvidence: 'not-captured',
      },
    })
    expect(scorecard.cases.map((entry) => entry.id)).toEqual([
      'csv-import-preview',
      'csv-engine-roundtrip',
      'xlsx-import-preview',
      'xlsx-snapshot-roundtrip-values-formulas-formats',
      'xlsx-snapshot-roundtrip-dimensions-merges',
      'xlsx-snapshot-roundtrip-freeze-panes',
      'xlsx-snapshot-roundtrip-filters',
      'xlsx-snapshot-roundtrip-sheet-protection',
      'xlsx-snapshot-roundtrip-protected-ranges',
      'xlsx-snapshot-roundtrip-data-validations',
      'xlsx-snapshot-roundtrip-tables',
      'xlsx-snapshot-roundtrip-charts',
      'xlsx-snapshot-roundtrip-pivots',
      'xlsx-unsupported-features-warning',
    ])
    expect(scorecard.cases.every((entry) => entry.required && entry.passed)).toBe(true)
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
      'xlsx.numberFormats',
      'xlsx.definedNames',
      'xlsx.comments',
      'xlsx.styles',
      'xlsx.rowColumnDimensions',
      'xlsx.merges',
      'xlsx.freezePanes.roundtrip',
      'xlsx.filters.roundtrip',
      'xlsx.sheetProtection.roundtrip',
      'xlsx.protectedRanges.roundtrip',
      'xlsx.dataValidations.roundtrip',
      'xlsx.tables.roundtrip',
      'xlsx.charts.roundtrip',
      'xlsx.pivots.roundtrip',
      'xlsx.multiSheet',
      'xlsx.unsupportedFeatureWarnings',
    ])
    expect(scorecard.summary.unsupportedFeatures).toEqual(['xlsx.macros.execution'])
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
})
