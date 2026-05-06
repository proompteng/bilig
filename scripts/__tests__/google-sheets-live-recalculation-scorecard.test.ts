import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

import {
  parseGoogleSheetsLiveRecalculationScorecard,
  validateGoogleSheetsLiveRecalculationScorecard,
} from '../gen-google-sheets-live-recalculation-scorecard.ts'
import { readJsonObject } from '../json-scorecard-helpers.ts'

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '../..')
const scorecardPath = resolve(repoRoot, 'packages/benchmarks/baselines/google-sheets-live-recalculation-scorecard.json')

describe('Google Sheets live recalculation scorecard', () => {
  it('validates the committed live Google Sheets recalculation timing artifact without calling Google APIs', () => {
    const scorecard = parseGoogleSheetsLiveRecalculationScorecard(readJsonObject(scorecardPath))

    expect(scorecard).toMatchObject({
      schemaVersion: 1,
      suite: 'google-sheets-live-recalculation-performance',
      source: {
        artifactGenerator: 'scripts/gen-google-sheets-live-recalculation-scorecard.ts',
        implementationPackage: 'packages/headless',
        evidenceKind: 'live-google-sheets-native-conversion-via-google-drive-connector',
        captureTransport: 'google-drive-connector',
      },
      benchmark: {
        sampleCount: 3,
        warmupCount: 0,
        valueRenderOption: 'UNFORMATTED_VALUE',
        measuredGoogleSheetsOperation: 'edit-and-read-recalculated-values',
        measuredWorkpaperOperation: 'mutate-and-recalculate',
        samplingOrder: 'engine-isolated-workpaper-then-google-sheets',
      },
      summary: {
        allRequiredCasesPassed: true,
        requiredCaseCount: 4,
        tenXMeanAndP95CaseCount: 4,
        workpaperWins: 4,
        microsoftExcelEvidence: 'not-covered-by-this-artifact',
      },
    })
    expect(scorecard.summary.coveredWorkloads).toEqual([
      'dirty-fanout-edit',
      'suspended-batch-single-column-edit',
      'conditional-aggregation-criteria-edit',
      'full-rebuild-recalculate',
    ])
    expect(scorecard.googleSheets.spreadsheets).toHaveLength(12)
    expect(scorecard.cases.map((entry) => entry.id)).toEqual([
      'google-sheets-live-recalculation-dirty-fanout-edit',
      'google-sheets-live-recalculation-suspended-batch-single-column-edit',
      'google-sheets-live-recalculation-conditional-aggregation-criteria-edit',
      'google-sheets-live-recalculation-full-rebuild-recalculate',
    ])
    expect(scorecard.cases.every((entry) => entry.passed && entry.verification.equivalent && entry.tenXMeanAndP95)).toBe(true)
    validateGoogleSheetsLiveRecalculationScorecard(scorecard)
  })

  it('rejects stale artifacts missing per-sample spreadsheet evidence', () => {
    const scorecard = parseGoogleSheetsLiveRecalculationScorecard(readJsonObject(scorecardPath))

    expect(() =>
      validateGoogleSheetsLiveRecalculationScorecard({
        ...scorecard,
        googleSheets: {
          spreadsheets: scorecard.googleSheets.spreadsheets.filter(
            (entry) => !(entry.caseId === 'google-sheets-live-recalculation-full-rebuild-recalculate' && entry.sampleIndex === 2),
          ),
        },
      }),
    ).toThrow('Google Sheets live recalculation scorecard spreadsheet evidence is stale')
  })

  it('rejects forged 10x flags that do not match captured timings', () => {
    const scorecard = parseGoogleSheetsLiveRecalculationScorecard(readJsonObject(scorecardPath))

    expect(() =>
      validateGoogleSheetsLiveRecalculationScorecard({
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
    ).toThrow('Google Sheets live recalculation 10x flag is stale')
  })
})
