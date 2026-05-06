import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

import {
  parseGoogleSheetsLiveStructuralScorecard,
  validateGoogleSheetsLiveStructuralScorecard,
} from '../gen-google-sheets-live-structural-scorecard.ts'
import { readJsonObject } from '../json-scorecard-helpers.ts'

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '../..')
const scorecardPath = resolve(repoRoot, 'packages/benchmarks/baselines/google-sheets-live-structural-scorecard.json')

describe('Google Sheets live structural scorecard', () => {
  it('validates the committed live Google Sheets structural timing artifact without calling Google APIs', () => {
    const scorecard = parseGoogleSheetsLiveStructuralScorecard(readJsonObject(scorecardPath))

    expect(scorecard).toMatchObject({
      schemaVersion: 1,
      suite: 'google-sheets-live-structural-performance',
      source: {
        artifactGenerator: 'scripts/gen-google-sheets-live-structural-scorecard.ts',
        implementationPackage: 'packages/headless',
        evidenceKind: 'live-google-sheets-native-conversion-via-google-drive-connector',
        captureTransport: 'google-drive-connector',
      },
      benchmark: {
        rowCount: 500,
        sampleCount: 3,
        valueRenderOption: 'UNFORMATTED_VALUE',
        measuredGoogleSheetsOperation: 'structural-edit-and-read-verification-values',
        measuredWorkpaperOperation: 'structural-edit',
        samplingOrder: 'engine-isolated-workpaper-then-google-sheets',
      },
      summary: {
        allRequiredCasesPassed: true,
        requiredCaseCount: 6,
        tenXMeanAndP95CaseCount: 6,
        workpaperWins: 6,
        microsoftExcelEvidence: 'not-covered-by-this-artifact',
      },
    })
    expect(scorecard.summary.coveredOperations).toEqual([
      'insert-rows',
      'delete-rows',
      'move-rows',
      'insert-columns',
      'delete-columns',
      'move-columns',
    ])
    expect(scorecard.googleSheets.spreadsheets).toHaveLength(18)
    expect(scorecard.cases.map((entry) => entry.id)).toEqual([
      'google-sheets-live-structural-insert-rows',
      'google-sheets-live-structural-delete-rows',
      'google-sheets-live-structural-move-rows',
      'google-sheets-live-structural-insert-columns',
      'google-sheets-live-structural-delete-columns',
      'google-sheets-live-structural-move-columns',
    ])
    expect(scorecard.cases.every((entry) => entry.passed && entry.verification.equivalent && entry.tenXMeanAndP95)).toBe(true)
    validateGoogleSheetsLiveStructuralScorecard(scorecard)
  })

  it('rejects stale artifacts missing per-sample spreadsheet evidence', () => {
    const scorecard = parseGoogleSheetsLiveStructuralScorecard(readJsonObject(scorecardPath))

    expect(() =>
      validateGoogleSheetsLiveStructuralScorecard({
        ...scorecard,
        googleSheets: {
          spreadsheets: scorecard.googleSheets.spreadsheets.filter(
            (entry) => !(entry.caseId === 'google-sheets-live-structural-move-columns' && entry.sampleIndex === 2),
          ),
        },
      }),
    ).toThrow('Google Sheets live structural scorecard spreadsheet evidence is stale')
  })

  it('rejects forged 10x flags that do not match captured timings', () => {
    const scorecard = parseGoogleSheetsLiveStructuralScorecard(readJsonObject(scorecardPath))

    expect(() =>
      validateGoogleSheetsLiveStructuralScorecard({
        ...scorecard,
        cases: [
          {
            ...scorecard.cases[0],
            workpaperToGoogleSheetsMeanRatio: 0.2,
            tenXMeanAndP95: true,
          },
          ...scorecard.cases.slice(1),
        ],
      }),
    ).toThrow('Google Sheets live structural 10x flag is stale')
  })
})
