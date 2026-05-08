import { describe, expect, it } from 'vitest'

import { buildBiligDominanceStatus, formatBiligDominanceStatusPathForMessage } from '../bilig-dominance-status.ts'
import type { PublicWorkbookCorpusStatus } from '../public-workbook-corpus-status.ts'
import { buildFixtureInput } from './bilig-dominance-scorecard.fixture.ts'

describe('bilig dominance status', () => {
  it('exposes actionable same-corpus UI proof setup commands', () => {
    const status = buildBiligDominanceStatus({
      input: buildFixtureInput(),
      financialCorpusStatus: completeFinancialCorpusStatus(),
      publicWorkbookCorpusStatus: completePublicWorkbookCorpusStatus(),
      stopMarkerActive: false,
      stopMarkerPath: '/repo/.agent-coordination/stop.md',
    })

    expect(status.uiSameCorpus).toMatchObject({
      captured: false,
      evidenceKind: 'not-captured',
      requiredWorkloads: ['visible-scroll-response'],
      missingRequiredWorkloads: ['visible-scroll-response'],
      missingInputs: ['googleSheetsUrlForUploadedSameCorpusWorkbook'],
      fixture: {
        corpusCaseId: 'wide-mixed-250k',
        materializedCells: 250_000,
        localXlsxPath: 'packages/benchmarks/baselines/ui-same-corpus/wide-mixed-250k.xlsx',
      },
      nextFixtureCheckCommand: 'pnpm ui:same-corpus:fixture:check',
      nextScorecardGenerateCommand: 'pnpm ui:browser-live:generate -- --capture .cache/ui-responsiveness/same-corpus-capture.json',
      nextDominanceCheckCommand: 'pnpm dominance:generate && pnpm dominance:check && pnpm dominance:audit:check',
    })
    expect(status.uiSameCorpus.fixture.microsoftExcelWebUrl).toContain('view.officeapps.live.com/op/view.aspx')
    expect(status.uiSameCorpus.nextPreflightCommand).toContain('--google-sheets-url')
    expect(status.uiSameCorpus.nextPreflightCommand).toContain('<google-sheets-url>')
    expect(status.uiSameCorpus.nextCaptureCommand).toContain('.cache/ui-responsiveness/same-corpus-capture.json')
    expect(status.uiSameCorpus.nextGoogleSheetsUploadInstruction).toContain('share it to anyone with the link')
  })

  it('surfaces financial workbook corpus blockers in dominance status', () => {
    const status = buildBiligDominanceStatus({
      input: buildFixtureInput(),
      financialCorpusStatus: {
        targetWorkbookCount: 5_000,
        sourceCount: 9_824,
        cachedArtifactCount: 0,
        recordedManifestArtifactCount: 0,
        recordedNonPassingCaseCount: 0,
      },
      publicWorkbookCorpusStatus: completePublicWorkbookCorpusStatus(),
      stopMarkerActive: true,
      stopMarkerPath: '.agent-coordination/stop.md',
    })

    expect(status.publicWorkbookCorpus).toMatchObject({
      financialWorkbookTargetCount: 5_000,
      financialSourceCount: 9_824,
      financialCachedArtifactCount: 0,
      recordedFinancialManifestArtifactCount: 0,
    })
    expect(status.importExportBlockers).toEqual(
      expect.arrayContaining([
        'financial/accounting corpus cached artifacts below target: 0/5000',
        'financial/accounting corpus recorded verification cases below target: 0/5000',
      ]),
    )
    expect(status.goalStatus).toBe('active-not-achieved')
  })

  it('formats repo-local status paths without exposing the checkout root', () => {
    expect(formatBiligDominanceStatusPathForMessage('/repo/.cache/public-workbook-corpus/manifest.json', '/repo')).toBe(
      '.cache/public-workbook-corpus/manifest.json',
    )
    expect(formatBiligDominanceStatusPathForMessage('/tmp/public-workbook-corpus/manifest.json', '/repo')).toBe(
      '/tmp/public-workbook-corpus/manifest.json',
    )
  })
})

function completePublicWorkbookCorpusStatus(): PublicWorkbookCorpusStatus {
  return {
    targetWorkbookCount: 10_000,
    sourceCount: 10_000,
    cachedArtifactCount: 10_000,
    scorecardCaseCount: 10_000,
    checkpointCaseCount: 10_000,
    recordedManifestArtifactCount: 10_000,
    missingManifestArtifactCount: 0,
    staleRecordedVerificationCount: 0,
    recordedPassedCaseCount: 10_000,
    recordedUnsupportedCaseCount: 0,
    recordedFailedCaseCount: 0,
    recordedErrorCaseCount: 0,
    recordedCoversManifest: true,
    recordedAllCasesPassed: true,
    missingManifestArtifactSample: [],
    staleRecordedVerificationSample: [],
    nextMissingVerificationCommand: null,
    nextMissingVerificationPlanCommand: null,
    nextStaleVerificationCommand: null,
    nextStaleVerificationPlanCommand: null,
    scorecardCoversManifest: true,
    targetComplete: true,
    gaps: [],
  }
}

function completeFinancialCorpusStatus() {
  return {
    targetWorkbookCount: 5_000,
    sourceCount: 5_000,
    cachedArtifactCount: 5_000,
    recordedManifestArtifactCount: 5_000,
    recordedNonPassingCaseCount: 0,
  }
}
