import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

import {
  parseMicrosoftExcelLiveStructuralScorecard,
  validateMicrosoftExcelLiveStructuralScorecard,
} from '../gen-microsoft-excel-live-structural-scorecard.ts'
import { readJsonObject } from '../json-scorecard-helpers.ts'

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '../..')
const scorecardPath = resolve(repoRoot, 'packages/benchmarks/baselines/microsoft-excel-live-structural-scorecard.json')

describe('Microsoft Excel live structural scorecard', () => {
  it('validates the committed live Excel structural timing artifact without launching Excel', () => {
    const scorecard = parseMicrosoftExcelLiveStructuralScorecard(readJsonObject(scorecardPath))

    expect(scorecard).toMatchObject({
      schemaVersion: 1,
      suite: 'microsoft-excel-live-structural-performance',
      source: {
        artifactGenerator: 'scripts/gen-microsoft-excel-live-structural-scorecard.ts',
        evidenceKind: 'live-local-microsoft-excel-automation',
        appleScriptTransport: 'osascript',
      },
      benchmark: {
        rowCount: 500,
        sampleCount: 5,
        screenUpdating: false,
      },
      summary: {
        allRequiredCasesPassed: true,
        requiredCaseCount: 6,
        tenXMeanAndP95CaseCount: 6,
        workpaperWins: 6,
        googleSheetsEvidence: 'not-covered-by-this-artifact',
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
    expect(scorecard.cases.map((entry) => entry.id)).toEqual([
      'excel-live-structural-insert-rows',
      'excel-live-structural-delete-rows',
      'excel-live-structural-move-rows',
      'excel-live-structural-insert-columns',
      'excel-live-structural-delete-columns',
      'excel-live-structural-move-columns',
    ])
    expect(scorecard.cases.every((entry) => entry.passed && entry.tenXMeanAndP95 && entry.verification.equivalent)).toBe(true)
    validateMicrosoftExcelLiveStructuralScorecard(scorecard)
  })

  it('rejects stale artifacts missing a required structural operation', () => {
    const scorecard = parseMicrosoftExcelLiveStructuralScorecard(readJsonObject(scorecardPath))

    expect(() =>
      validateMicrosoftExcelLiveStructuralScorecard({
        ...scorecard,
        summary: {
          ...scorecard.summary,
          requiredCaseCount: 5,
          coveredOperations: scorecard.summary.coveredOperations.filter((entry) => entry !== 'move-columns'),
        },
        cases: scorecard.cases.filter((entry) => entry.id !== 'excel-live-structural-move-columns'),
      }),
    ).toThrow('Microsoft Excel live structural scorecard required cases are stale')
  })
})
