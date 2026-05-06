import { describe, expect, it } from 'vitest'

import { buildCollaborationScorecard, validateCollaborationScorecard } from '../gen-collaboration-scorecard.ts'

describe('collaboration scorecard', () => {
  it('executes real collaboration controls for sync, presence, conflict, and viewport behavior', async () => {
    const scorecard = await buildCollaborationScorecard('2026-05-06T13:00:00.000Z')

    expect(scorecard).toMatchObject({
      schemaVersion: 1,
      suite: 'collaboration-posture',
      generatedAt: '2026-05-06T13:00:00.000Z',
      summary: {
        allRequiredControlsPassed: true,
        syncRebaseAckPassed: true,
        presenceSelectionPassed: true,
        conflictViewportPassed: true,
        headedBrowserViewportPassed: true,
        longRunningConflictRatePassed: true,
        externalGoogleSheetsEvidence: 'official-docs-comparison-artifact',
        externalMicrosoftExcelEvidence: 'official-docs-comparison-artifact',
      },
    })
    expect(scorecard.source.externalCollaborationComparisonArtifact).toBe(
      'packages/benchmarks/baselines/collaboration-external-sheets-excel-comparison.json',
    )
    expect(scorecard.controls.map((control) => control.id)).toEqual([
      'worker-sync-rebase-ack-roundtrip',
      'presence-session-selection-filtering',
      'editor-conflict-and-viewport-protection',
      'headed-browser-multi-user-viewport-soak',
      'long-running-collaboration-conflict-rate',
      'external-sheets-excel-collaboration-comparison',
    ])
    expect(scorecard.controls.every((control) => control.required && control.passed)).toBe(true)
    expect(scorecard.controls.find((control) => control.id === 'external-sheets-excel-collaboration-comparison')).toMatchObject({
      category: 'external-comparison',
      coveredControls: [
        'external.googleSheetsCollaborationDocs',
        'external.microsoftExcelCollaborationDocs',
        'external.sheetsExcelCollaborationComparison',
      ],
    })
    expect(scorecard.summary.coveredControls).toEqual([
      'sync.pendingRebase',
      'sync.authoritativeAck',
      'sync.noAcceptedOpLoss',
      'presence.sessionLifecycle',
      'presence.selectionSchema',
      'presence.collaboratorFiltering',
      'conflict.authoritativeDriftDetection',
      'viewport.optimisticAxisProtection',
      'viewport.authoritativeCatchupClearsPending',
      'headedBrowser.multiUserViewportSoak',
      'conflict.longRunningZeroUnexpectedConflicts',
      'sync.longRunningAcceptedOpConvergence',
      'external.googleSheetsCollaborationDocs',
      'external.microsoftExcelCollaborationDocs',
      'external.sheetsExcelCollaborationComparison',
    ])
    expect(scorecard.summary.uncoveredControls).toEqual([])

    validateCollaborationScorecard(scorecard)
  })

  it('fails validation when a required collaboration control is missing', async () => {
    const scorecard = await buildCollaborationScorecard('2026-05-06T13:00:00.000Z')

    expect(() =>
      validateCollaborationScorecard({
        ...scorecard,
        controls: scorecard.controls.filter((control) => control.id !== 'presence-session-selection-filtering'),
      }),
    ).toThrow('Collaboration scorecard is missing required control: presence-session-selection-filtering')
  })
})
