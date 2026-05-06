import { describe, expect, it } from 'vitest'

import {
  buildMicrosoftExcelLiveCalculationScorecard,
  evaluateBiligCases,
  validateMicrosoftExcelLiveCalculationScorecard,
  type CalculationScalarValue,
} from '../gen-microsoft-excel-live-calculation-scorecard.ts'

describe('Microsoft Excel live calculation scorecard', () => {
  it('builds a required-case scorecard from Bilig and live-Excel-shaped values', () => {
    const biligValues = evaluateBiligCases()
    const scorecard = buildMicrosoftExcelLiveCalculationScorecard('2026-05-06T09:00:00.000Z', {
      excelVersion: '16.test',
      rawValuesByCaseId: new Map([...biligValues].map(([caseId, value]) => [caseId, toExcelRawValue(value)])),
    })

    expect(scorecard).toMatchObject({
      schemaVersion: 1,
      suite: 'microsoft-excel-live-calculation-correctness',
      generatedAt: '2026-05-06T09:00:00.000Z',
      source: {
        artifactGenerator: 'scripts/gen-microsoft-excel-live-calculation-scorecard.ts',
        evidenceKind: 'live-local-microsoft-excel-automation',
        appleScriptTransport: 'osascript',
      },
      microsoftExcel: {
        appPath: '/Applications/Microsoft Excel.app',
        version: '16.test',
      },
      summary: {
        allRequiredCasesPassed: true,
        requiredCaseCount: 36,
        matchingCaseCount: 36,
        googleSheetsEvidence: 'not-covered-by-this-artifact',
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
    expect(scorecard.cases.map((entry) => entry.id)).toEqual([
      'arithmetic-precedence',
      'aggregate-sum-range',
      'conditional-if-comparison',
      'numeric-rounding',
      'aggregate-count-range',
      'text-concat',
      'textjoin-ignore-empty-range',
      'lookup-xlookup-exact',
      'lookup-match-exact',
      'lookup-index-range',
      'boolean-and',
      'boolean-or',
      'date-serial',
      'date-year-from-serial',
      'conditional-sumif-range',
      'math-abs-sqrt',
      'aggregate-average-range',
      'aggregate-min-range',
      'aggregate-max-range',
      'aggregate-counta-range',
      'aggregate-countblank-range',
      'math-product-range',
      'math-sumsq-range',
      'math-mod',
      'math-power',
      'math-gcd-range',
      'text-left',
      'text-right',
      'text-mid',
      'text-len',
      'text-trim-upper',
      'text-search-case-insensitive',
      'lookup-vlookup-exact',
      'lookup-choose',
      'statistical-countif',
      'statistical-averageif',
    ])
    expect(scorecard.cases.every((entry) => entry.passed)).toBe(true)
    validateMicrosoftExcelLiveCalculationScorecard(scorecard)
  })

  it('rejects stale artifacts missing required cases', () => {
    const biligValues = evaluateBiligCases()
    const scorecard = buildMicrosoftExcelLiveCalculationScorecard('2026-05-06T09:00:00.000Z', {
      excelVersion: '16.test',
      rawValuesByCaseId: new Map([...biligValues].map(([caseId, value]) => [caseId, toExcelRawValue(value)])),
    })

    expect(() =>
      validateMicrosoftExcelLiveCalculationScorecard({
        ...scorecard,
        cases: scorecard.cases.filter((entry) => entry.id !== 'lookup-xlookup-exact'),
      }),
    ).toThrow('Microsoft Excel live calculation scorecard required cases are stale')
  })

  it('rejects forged pass flags that do not match captured Excel values', () => {
    const biligValues = evaluateBiligCases()
    const scorecard = buildMicrosoftExcelLiveCalculationScorecard('2026-05-06T09:00:00.000Z', {
      excelVersion: '16.test',
      rawValuesByCaseId: new Map([...biligValues].map(([caseId, value]) => [caseId, toExcelRawValue(value)])),
    })

    expect(() =>
      validateMicrosoftExcelLiveCalculationScorecard({
        ...scorecard,
        cases: [
          {
            ...scorecard.cases[0],
            microsoftExcelRawValue: '999',
            microsoftExcelValue: 999,
            passed: true,
          },
          ...scorecard.cases.slice(1),
        ],
      }),
    ).toThrow('Microsoft Excel live calculation pass flag is stale')
  })
})

function toExcelRawValue(value: CalculationScalarValue): string {
  if (value === null) {
    return ''
  }
  if (typeof value === 'boolean') {
    return value ? 'true' : 'false'
  }
  if (typeof value === 'number' || typeof value === 'string') {
    return String(value)
  }
  return value.error
}
