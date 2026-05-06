import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

import {
  parseMicrosoftExcelLiveLargeWorkbookScorecard,
  validateMicrosoftExcelLiveLargeWorkbookScorecard,
} from '../gen-microsoft-excel-live-large-workbook-scorecard.ts'
import { readJsonObject } from '../json-scorecard-helpers.ts'

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '../..')
const scorecardPath = resolve(repoRoot, 'packages/benchmarks/baselines/microsoft-excel-live-large-workbook-scorecard.json')

describe('Microsoft Excel live large-workbook scorecard', () => {
  it('validates the committed live Excel large-workbook timing artifact without launching Excel', () => {
    const scorecard = parseMicrosoftExcelLiveLargeWorkbookScorecard(readJsonObject(scorecardPath))

    expect(scorecard).toMatchObject({
      schemaVersion: 1,
      suite: 'microsoft-excel-live-large-workbook-performance',
      source: {
        artifactGenerator: 'scripts/gen-microsoft-excel-live-large-workbook-scorecard.ts',
        implementationPackage: 'packages/core',
        xlsxExportPackage: 'packages/excel-import',
        corpusPackage: 'packages/benchmarks',
        evidenceKind: 'live-local-microsoft-excel-automation',
        appleScriptTransport: 'osascript',
      },
      benchmark: {
        sampleCount: 3,
        screenUpdating: false,
        calculationMode: 'manual-during-open-and-calculate',
        measuredExcelOperation: 'open-workbook-and-calculate-full-rebuild',
        measuredBiligOperation: 'import-snapshot',
        samplingOrder: 'engine-isolated-bilig-then-excel',
      },
      summary: {
        allRequiredCasesPassed: true,
        requiredCaseCount: 2,
        googleSheetsEvidence: 'not-covered-by-this-artifact',
      },
    })
    expect(scorecard.summary.coveredCorpusCaseIds).toEqual(['dense-mixed-100k', 'dense-mixed-250k'])
    expect(scorecard.summary.coveredMaterializedCells).toEqual([100_000, 250_000])
    expect(scorecard.cases.map((entry) => entry.id)).toEqual([
      'excel-live-large-workbook-open-calculate-dense-mixed-100k',
      'excel-live-large-workbook-open-calculate-dense-mixed-250k',
    ])
    expect(scorecard.cases.every((entry) => entry.passed && entry.verification.equivalent)).toBe(true)
    validateMicrosoftExcelLiveLargeWorkbookScorecard(scorecard)
  })

  it('rejects stale artifacts missing a required large-workbook corpus', () => {
    const scorecard = parseMicrosoftExcelLiveLargeWorkbookScorecard(readJsonObject(scorecardPath))

    expect(() =>
      validateMicrosoftExcelLiveLargeWorkbookScorecard({
        ...scorecard,
        summary: {
          ...scorecard.summary,
          requiredCaseCount: 1,
          coveredCorpusCaseIds: scorecard.summary.coveredCorpusCaseIds.filter((entry) => entry !== 'dense-mixed-250k'),
          coveredMaterializedCells: scorecard.summary.coveredMaterializedCells.filter((entry) => entry !== 250_000),
        },
        cases: scorecard.cases.filter((entry) => entry.id !== 'excel-live-large-workbook-open-calculate-dense-mixed-250k'),
      }),
    ).toThrow('Microsoft Excel live large-workbook scorecard required cases are stale')
  })
})
