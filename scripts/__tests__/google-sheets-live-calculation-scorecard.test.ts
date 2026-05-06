import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

import {
  parseGoogleSheetsLiveCalculationScorecard,
  validateGoogleSheetsLiveCalculationScorecard,
} from '../gen-google-sheets-live-calculation-scorecard.ts'
import { readJsonObject } from '../json-scorecard-helpers.ts'

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '../..')
const scorecardPath = resolve(repoRoot, 'packages/benchmarks/baselines/google-sheets-live-calculation-scorecard.json')

describe('Google Sheets live calculation scorecard', () => {
  it('validates the committed live Google Sheets calculation artifact without calling Google APIs', () => {
    const scorecard = parseGoogleSheetsLiveCalculationScorecard(readJsonObject(scorecardPath))

    expect(scorecard).toMatchObject({
      schemaVersion: 1,
      suite: 'google-sheets-live-calculation-correctness',
      source: {
        artifactGenerator: 'scripts/gen-google-sheets-live-calculation-scorecard.ts',
        evidenceKind: 'live-google-sheets-native-conversion-via-google-drive-connector',
        captureTransport: 'google-drive-connector',
      },
      googleSheets: {
        worksheetName: 'Cases',
        valueRenderOption: 'UNFORMATTED_VALUE',
      },
      summary: {
        allRequiredCasesPassed: true,
        requiredCaseCount: 36,
        matchingCaseCount: 36,
        microsoftExcelEvidence: 'not-covered-by-this-artifact',
      },
    })
    expect(scorecard.summary.coveredFeatures).toEqual([
      'excelLive.arithmeticPrecedence',
      'excelLive.aggregateSumRange',
      'excelLive.conditionalIfComparison',
      'excelLive.numericRounding',
      'excelLive.aggregateCountRange',
      'excelLive.textConcat',
      'excelLive.textJoinIgnoreEmptyRange',
      'excelLive.lookupXlookupExact',
      'excelLive.lookupMatchExact',
      'excelLive.lookupIndexRange',
      'excelLive.booleanAnd',
      'excelLive.booleanOr',
      'excelLive.dateSerial',
      'excelLive.dateYearFromSerial',
      'excelLive.conditionalSumifRange',
      'excelLive.mathAbsSqrt',
      'excelLive.aggregateAverageRange',
      'excelLive.aggregateMinRange',
      'excelLive.aggregateMaxRange',
      'excelLive.aggregateCountaRange',
      'excelLive.aggregateCountblankRange',
      'excelLive.mathProductRange',
      'excelLive.mathSumsqRange',
      'excelLive.mathMod',
      'excelLive.mathPower',
      'excelLive.mathGcdRange',
      'excelLive.textLeft',
      'excelLive.textRight',
      'excelLive.textMid',
      'excelLive.textLen',
      'excelLive.textTrimUpper',
      'excelLive.textSearchCaseInsensitive',
      'excelLive.lookupVlookupExact',
      'excelLive.lookupChoose',
      'excelLive.statisticalCountif',
      'excelLive.statisticalAverageif',
    ])
    expect(scorecard.cases.every((entry) => entry.passed)).toBe(true)
    validateGoogleSheetsLiveCalculationScorecard(scorecard)
  })

  it('rejects stale artifacts missing required calculation cases', () => {
    const scorecard = parseGoogleSheetsLiveCalculationScorecard(readJsonObject(scorecardPath))

    expect(() =>
      validateGoogleSheetsLiveCalculationScorecard({
        ...scorecard,
        summary: {
          ...scorecard.summary,
          requiredCaseCount: 35,
        },
        cases: scorecard.cases.filter((entry) => entry.id !== 'math-abs-sqrt'),
      }),
    ).toThrow('Google Sheets live calculation scorecard required case count is stale')
  })

  it('rejects forged pass flags that do not match captured Google Sheets values', () => {
    const scorecard = parseGoogleSheetsLiveCalculationScorecard(readJsonObject(scorecardPath))

    expect(() =>
      validateGoogleSheetsLiveCalculationScorecard({
        ...scorecard,
        cases: [
          {
            ...scorecard.cases[0],
            googleSheetsRawValue: '999',
            googleSheetsValue: 999,
            passed: true,
          },
          ...scorecard.cases.slice(1),
        ],
      }),
    ).toThrow('Google Sheets live calculation pass flag is stale')
  })
})
