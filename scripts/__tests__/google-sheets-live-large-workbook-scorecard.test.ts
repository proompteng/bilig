import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

import {
  parseGoogleSheetsLiveLargeWorkbookScorecard,
  validateGoogleSheetsLiveLargeWorkbookScorecard,
} from '../gen-google-sheets-live-large-workbook-scorecard.ts'
import { readJsonObject } from '../json-scorecard-helpers.ts'

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '../..')
const scorecardPath = resolve(repoRoot, 'packages/benchmarks/baselines/google-sheets-live-large-workbook-scorecard.json')

describe('Google Sheets live large-workbook scorecard', () => {
  it('validates the committed live Google Sheets large-workbook timing artifact without calling Google APIs', () => {
    const scorecard = parseGoogleSheetsLiveLargeWorkbookScorecard(readJsonObject(scorecardPath))

    expect(scorecard).toMatchObject({
      schemaVersion: 1,
      suite: 'google-sheets-live-large-workbook-performance',
      source: {
        artifactGenerator: 'scripts/gen-google-sheets-live-large-workbook-scorecard.ts',
        implementationPackage: 'packages/core',
        xlsxExportPackage: 'packages/excel-import',
        corpusPackage: 'packages/benchmarks',
        evidenceKind: 'live-google-sheets-native-conversion-via-google-drive-connector',
        captureTransport: 'google-drive-connector',
      },
      benchmark: {
        sampleCount: 3,
        valueRenderOption: 'UNFORMATTED_VALUE',
        measuredGoogleSheetsOperation: 'native-xlsx-import-and-read-terminal-cell',
        measuredBiligOperation: 'import-snapshot',
        samplingOrder: 'engine-isolated-bilig-then-google-sheets',
      },
      summary: {
        allRequiredCasesPassed: true,
        requiredCaseCount: 2,
        tenXMeanAndP95CaseCount: 2,
        biligWins: 2,
        microsoftExcelEvidence: 'not-covered-by-this-artifact',
      },
    })
    expect(scorecard.summary.coveredCorpusCaseIds).toEqual(['dense-mixed-100k', 'dense-mixed-250k'])
    expect(scorecard.summary.coveredMaterializedCells).toEqual([100_000, 250_000])
    expect(scorecard.googleSheets.spreadsheets).toHaveLength(6)
    expect(scorecard.cases.map((entry) => entry.id)).toEqual([
      'google-sheets-live-large-workbook-import-read-dense-mixed-100k',
      'google-sheets-live-large-workbook-import-read-dense-mixed-250k',
    ])
    expect(scorecard.cases.every((entry) => entry.passed && entry.verification.equivalent && entry.tenXMeanAndP95)).toBe(true)
    validateGoogleSheetsLiveLargeWorkbookScorecard(scorecard)
  })

  it('rejects stale artifacts missing per-sample spreadsheet evidence', () => {
    const scorecard = parseGoogleSheetsLiveLargeWorkbookScorecard(readJsonObject(scorecardPath))

    expect(() =>
      validateGoogleSheetsLiveLargeWorkbookScorecard({
        ...scorecard,
        googleSheets: {
          spreadsheets: scorecard.googleSheets.spreadsheets.filter(
            (entry) => !(entry.caseId === 'google-sheets-live-large-workbook-import-read-dense-mixed-250k' && entry.sampleIndex === 2),
          ),
        },
      }),
    ).toThrow('Google Sheets live large-workbook scorecard spreadsheet evidence is stale')
  })

  it('rejects forged 10x flags that do not match captured timings', () => {
    const scorecard = parseGoogleSheetsLiveLargeWorkbookScorecard(readJsonObject(scorecardPath))

    expect(() =>
      validateGoogleSheetsLiveLargeWorkbookScorecard({
        ...scorecard,
        cases: [
          {
            ...scorecard.cases[0],
            biligToGoogleSheetsMeanRatio: 0.2,
            tenXMeanAndP95: true,
          },
          ...scorecard.cases.slice(1),
        ],
      }),
    ).toThrow('Google Sheets live large-workbook 10x flag is stale')
  })
})
