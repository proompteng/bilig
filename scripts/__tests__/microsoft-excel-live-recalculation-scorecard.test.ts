import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

import {
  parseMicrosoftExcelLiveRecalculationScorecard,
  validateMicrosoftExcelLiveRecalculationScorecard,
} from '../gen-microsoft-excel-live-recalculation-scorecard.ts'
import { readJsonObject } from '../json-scorecard-helpers.ts'

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '../..')
const scorecardPath = resolve(repoRoot, 'packages/benchmarks/baselines/microsoft-excel-live-recalculation-scorecard.json')

describe('Microsoft Excel live recalculation scorecard', () => {
  it('validates the committed live Excel recalculation timing artifact without launching Excel', () => {
    const scorecard = parseMicrosoftExcelLiveRecalculationScorecard(readJsonObject(scorecardPath))

    expect(scorecard).toMatchObject({
      schemaVersion: 1,
      suite: 'microsoft-excel-live-recalculation-performance',
      source: {
        artifactGenerator: 'scripts/gen-microsoft-excel-live-recalculation-scorecard.ts',
        evidenceKind: 'live-local-microsoft-excel-automation',
        appleScriptTransport: 'osascript',
      },
      benchmark: {
        sampleCount: 5,
        screenUpdating: false,
        calculationMode: 'manual-during-measurement',
      },
      summary: {
        allRequiredCasesPassed: true,
        requiredCaseCount: 4,
        workpaperWins: 4,
        googleSheetsEvidence: 'not-covered-by-this-artifact',
      },
    })
    expect(scorecard.summary.tenXMeanAndP95CaseCount).toBe(2)
    expect(scorecard.summary.coveredWorkloads).toEqual([
      'dirty-fanout-edit',
      'suspended-batch-single-column-edit',
      'conditional-aggregation-criteria-edit',
      'full-rebuild-recalculate',
    ])
    expect(scorecard.cases.map((entry) => entry.id)).toEqual([
      'excel-live-recalculation-dirty-fanout-edit',
      'excel-live-recalculation-suspended-batch-single-column-edit',
      'excel-live-recalculation-conditional-aggregation-criteria-edit',
      'excel-live-recalculation-full-rebuild-recalculate',
    ])
    expect(scorecard.cases.every((entry) => entry.passed && entry.verification.equivalent)).toBe(true)
    validateMicrosoftExcelLiveRecalculationScorecard(scorecard)
  })

  it('rejects stale artifacts missing a required recalculation workload', () => {
    const scorecard = parseMicrosoftExcelLiveRecalculationScorecard(readJsonObject(scorecardPath))

    expect(() =>
      validateMicrosoftExcelLiveRecalculationScorecard({
        ...scorecard,
        summary: {
          ...scorecard.summary,
          requiredCaseCount: 3,
          coveredWorkloads: scorecard.summary.coveredWorkloads.filter((entry) => entry !== 'full-rebuild-recalculate'),
        },
        cases: scorecard.cases.filter((entry) => entry.id !== 'excel-live-recalculation-full-rebuild-recalculate'),
      }),
    ).toThrow('Microsoft Excel live recalculation scorecard required cases are stale')
  })
})
