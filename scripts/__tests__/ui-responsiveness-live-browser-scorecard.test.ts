import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

import {
  parseUiResponsivenessLiveBrowserScorecard,
  validateUiResponsivenessLiveBrowserScorecard,
  type UiResponsivenessLiveBrowserScorecard,
} from '../gen-ui-responsiveness-live-browser-scorecard.ts'
import { readJsonObject } from '../json-scorecard-helpers.ts'

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '../..')

describe('UI responsiveness live browser scorecard', () => {
  it('validates the checked-in direct incumbent browser timing artifact', () => {
    const scorecard = parseUiResponsivenessLiveBrowserScorecard(
      readJsonObject(resolve(repoRoot, 'packages/benchmarks/baselines/ui-responsiveness-live-browser-scorecard.json')),
    )

    expect(scorecard.summary).toMatchObject({
      directBrowserTimingCaptured: true,
      allRequiredCasesPassed: true,
      requiredVendorCount: 2,
      capturedVendors: ['google-sheets', 'microsoft-excel-web'],
    })
    expect(scorecard.cases.map((entry) => entry.id)).toEqual(['google-sheets-public-grid-scroll', 'microsoft-excel-web-public-xlsx-scroll'])
    expect(scorecard.cases.every((entry) => entry.sampleCount >= 3 && entry.limitations.length > 0)).toBe(true)
    validateUiResponsivenessLiveBrowserScorecard(scorecard)
  })

  it('rejects missing incumbent vendors', () => {
    const scorecard = parseUiResponsivenessLiveBrowserScorecard(
      readJsonObject(resolve(repoRoot, 'packages/benchmarks/baselines/ui-responsiveness-live-browser-scorecard.json')),
    )
    const staleScorecard: UiResponsivenessLiveBrowserScorecard = {
      ...scorecard,
      summary: {
        ...scorecard.summary,
        capturedVendors: ['google-sheets'],
      },
    }

    expect(() => validateUiResponsivenessLiveBrowserScorecard(staleScorecard)).toThrow(
      'UI responsiveness live browser scorecard is missing vendor: microsoft-excel-web',
    )
  })
})
