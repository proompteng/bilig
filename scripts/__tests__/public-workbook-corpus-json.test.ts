import { describe, expect, it } from 'vitest'

import { parsePublicWorkbookCorpusCase } from '../public-workbook-corpus-json.ts'

describe('public workbook corpus JSON parsing', () => {
  it('preserves whitespace-only worksheet names in workbook metadata', () => {
    const parsed = parsePublicWorkbookCorpusCase({
      id: 'workbook-whitespace-sheet',
      sourceId: 'source-whitespace-sheet',
      sourceUrl: 'https://example.com/workbook.xlsx',
      fileName: 'workbook.xlsx',
      sha256: '0'.repeat(64),
      byteSize: 128,
      license: {
        spdxId: 'OGL',
        title: 'Open Government Licence Ontario',
        evidenceUrl: 'https://example.com/license',
      },
      status: 'passed',
      passed: true,
      featureCounts: {
        sheetCount: 1,
        cellCount: 1,
        formulaCellCount: 0,
        valueCellCount: 1,
        definedNameCount: 0,
        tableCount: 0,
        chartCount: 0,
        pivotCount: 0,
        mergeCount: 0,
        styleRangeCount: 0,
        conditionalFormatCount: 0,
        dataValidationCount: 0,
        macroPayloadCount: 0,
        warningCount: 0,
      },
      workbookMetadata: {
        workbookName: 'workbook',
        sheetNames: [' '],
        dimensions: [
          {
            sheetName: ' ',
            rowCount: 1,
            columnCount: 1,
            nonEmptyCellCount: 1,
            usedRange: {
              startRow: 0,
              startColumn: 0,
              endRow: 0,
              endColumn: 0,
            },
          },
        ],
      },
      validation: {
        importPassed: true,
        formulaOraclePassed: true,
        formulaOracleComparisons: 0,
        formulaOracleMismatches: [],
        roundTripPassed: true,
        structuralSmokePassed: null,
      },
      unsupportedFeatureClassifications: [],
      evidence: [],
    })

    expect(parsed.workbookMetadata.sheetNames).toEqual([' '])
    expect(parsed.workbookMetadata.dimensions[0]?.sheetName).toBe(' ')
  })
})
