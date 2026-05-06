import { describe, expect, it } from 'vitest'

import { buildReliabilityScorecard, validateReliabilityScorecard } from '../gen-reliability-scorecard.ts'

describe('reliability scorecard', () => {
  it('generates a checked artifact from executable pending-op durability controls', async () => {
    const scorecard = await buildReliabilityScorecard('2026-05-06T11:00:00.000Z')

    expect(scorecard).toMatchObject({
      schemaVersion: 1,
      suite: 'reliability-posture',
      generatedAt: '2026-05-06T11:00:00.000Z',
      summary: {
        allRequiredControlsPassed: true,
        pendingReloadPassed: true,
        authoritativeAckPassed: true,
        authoritativeRebasePassed: true,
        failedRetryPassed: true,
        headedBrowserReloadPassed: true,
        externalGoogleSheetsEvidence: 'not-captured',
        externalMicrosoftExcelEvidence: 'not-captured',
      },
    })
    expect(scorecard.controls.map((control) => control.id)).toEqual([
      'pending-mutations-survive-reload',
      'submitted-mutations-absorb-authoritative-ack',
      'authoritative-rebase-preserves-unsent-mutations',
      'failed-mutations-survive-reload-and-retry',
      'headed-browser-reload-persistence-flow',
    ])
    expect(scorecard.controls.every((control) => control.required && control.passed)).toBe(true)
    expect(scorecard.summary.coveredControls).toEqual([
      'pending.localReloadSurvival',
      'pending.submittedReloadSurvival',
      'pending.authoritativeAckAbsorption',
      'pending.authoritativeRebasePreservesLocal',
      'pending.failedRetrySurvival',
      'localStore.journalActiveView',
      'headedBrowser.reloadPersistence',
    ])
    expect(scorecard.summary.uncoveredControls).toEqual([
      'headedBrowser.crashSoak',
      'offlineNetworkPartitionSoak',
      'externalSheetsExcelReliabilityComparison',
    ])
  })

  it('rejects stale artifacts missing required reliability controls', async () => {
    const scorecard = await buildReliabilityScorecard('2026-05-06T11:00:00.000Z')
    const staleScorecard = {
      ...scorecard,
      controls: scorecard.controls.filter((control) => control.id !== 'authoritative-rebase-preserves-unsent-mutations'),
    }

    expect(() => validateReliabilityScorecard(staleScorecard)).toThrow(
      'Reliability scorecard is missing required control: authoritative-rebase-preserves-unsent-mutations',
    )
  })
})
