import { describe, expect, it } from 'vitest'

import { buildBiligDominanceStatus, formatBiligDominanceStatusPathForMessage } from '../bilig-dominance-status.ts'
import { buildSameCorpusProof, type SameCorpusCapture } from '../gen-ui-responsiveness-live-browser-scorecard.ts'
import type { PublicWorkbookCorpusFinancialPlan } from '../public-workbook-corpus-financial-plan.ts'
import type { PublicWorkbookCorpusFeatureWitnessPlan } from '../public-workbook-corpus-feature-witness-plan.ts'
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
      googleSheetsUrl: null,
      googleSheetsUrlEnvVar: 'BILIG_UI_SAME_CORPUS_GOOGLE_SHEETS_URL',
      missingInputs: ['googleSheetsUrlForUploadedSameCorpusWorkbook'],
      fixture: {
        corpusCaseId: 'wide-mixed-250k',
        materializedCells: 250_000,
        localXlsxPath: 'packages/benchmarks/baselines/ui-same-corpus/wide-mixed-250k.xlsx',
      },
      nextFixtureCheckCommand: 'pnpm ui:same-corpus:fixture:check',
      nextPublicAccessCheckCommand: expect.stringContaining('pnpm ui:same-corpus:public-check'),
      nextScorecardGenerateCommand: 'pnpm ui:browser-live:generate -- --capture .cache/ui-responsiveness/same-corpus-capture.json',
      nextDominanceCheckCommand: 'pnpm dominance:generate && pnpm dominance:check && pnpm dominance:audit:check',
    })
    expect(status.uiSameCorpus.fixture.microsoftExcelWebUrl).toContain('view.officeapps.live.com/op/view.aspx')
    expect(status.uiSameCorpus.nextPreflightCommand).toContain('--google-sheets-url')
    expect(status.uiSameCorpus.nextPreflightCommand).toContain('<google-sheets-url>')
    expect(status.uiSameCorpus.nextCaptureCommand).toContain('.cache/ui-responsiveness/same-corpus-capture.json')
    expect(status.uiSameCorpus.nextGoogleSheetsUploadInstruction).toContain('share it to anyone with the link')
  })

  it('fills same-corpus UI proof commands when the Google Sheets URL is known', () => {
    const googleSheetsUrl = 'https://docs.google.com/spreadsheets/d/sameCorpusSheet/edit'
    const status = buildBiligDominanceStatus({
      input: buildFixtureInput(),
      financialCorpusStatus: completeFinancialCorpusStatus(),
      publicWorkbookCorpusStatus: completePublicWorkbookCorpusStatus(),
      stopMarkerActive: false,
      stopMarkerPath: '/repo/.agent-coordination/stop.md',
      uiSameCorpusGoogleSheetsUrl: googleSheetsUrl,
    })

    expect(status.uiSameCorpus.googleSheetsUrl).toBe(googleSheetsUrl)
    expect(status.uiSameCorpus.missingInputs).toEqual([])
    expect(status.uiSameCorpus.nextPreflightCommand).toContain(googleSheetsUrl)
    expect(status.uiSameCorpus.nextPreflightCommand).not.toContain('<google-sheets-url>')
    expect(status.uiSameCorpus.nextPublicAccessCheckCommand).toContain(googleSheetsUrl)
    expect(status.uiSameCorpus.nextPublicAccessCheckCommand).not.toContain('<google-sheets-url>')
    expect(status.uiSameCorpus.nextCaptureCommand).toContain(googleSheetsUrl)
    expect(status.uiSameCorpus.nextCaptureCommand).not.toContain('<google-sheets-url>')
  })

  it('keeps asking for the Google Sheets URL when captured same-corpus proof is not 10x', () => {
    const fixtureInput = buildFixtureInput()
    const status = buildBiligDominanceStatus({
      input: {
        ...fixtureInput,
        uiResponsivenessLiveBrowserScorecard: {
          ...fixtureInput.uiResponsivenessLiveBrowserScorecard,
          sameCorpusProof: buildSameCorpusProof(failingSameCorpusCapture()),
        },
      },
      financialCorpusStatus: completeFinancialCorpusStatus(),
      publicWorkbookCorpusStatus: completePublicWorkbookCorpusStatus(),
      stopMarkerActive: false,
      stopMarkerPath: '/repo/.agent-coordination/stop.md',
    })

    expect(status.uiSameCorpus).toMatchObject({
      captured: true,
      evidenceKind: 'same-corpus-browser-capture',
      requiredCaseCount: 1,
      tenXMeanAndP95CaseCount: 0,
      tenXRequirementSatisfied: false,
      missingRequiredWorkloads: [],
      missingInputs: ['googleSheetsUrlForUploadedSameCorpusWorkbook'],
    })
    expect(status.uiSameCorpus.nextPreflightCommand).toContain('<google-sheets-url>')
    expect(status.uiSameCorpus.nextCaptureCommand).toContain('<google-sheets-url>')
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

  it('exposes the guarded financial corpus plan in dominance status', () => {
    const status = buildBiligDominanceStatus({
      input: buildFixtureInput(),
      financialCorpusPlan: financialPlanFixture(),
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

    expect(status.publicWorkbookCorpus.financialPlan).toMatchObject({
      sourceCount: 9_824,
      targetArtifactCount: 5_000,
      cachedArtifactCount: 0,
      remainingArtifactSlots: 5_000,
      candidateSourceCount: 8_949,
      candidateSourceDeficitCount: 0,
      recommendedFetchLimit: 20,
      needsAdditionalDiscovery: false,
      targetReachableFromKnownCandidates: true,
      nextPlanCommand: 'pnpm public-workbook-corpus:discover-financial:plan',
      nextCheckCommand: 'pnpm public-workbook-corpus:discover-financial:check',
      nextFetchPlanCommand: expect.stringContaining('public-workbook-corpus:fetch-financial:plan'),
      nextFetchCommand: expect.stringContaining('--limit 20'),
      nextVerifyCommand: expect.stringContaining('public-workbook-corpus:verify-financial'),
    })
  })

  it('exposes feature witness coverage and targeted discovery commands in dominance status', () => {
    const status = buildBiligDominanceStatus({
      input: buildFixtureInput(),
      featureWitnessPlan: featureWitnessPlanFixture(),
      financialCorpusStatus: completeFinancialCorpusStatus(),
      publicWorkbookCorpusStatus: completePublicWorkbookCorpusStatus(),
      stopMarkerActive: true,
      stopMarkerPath: '.agent-coordination/stop.md',
    })

    expect(status.publicWorkbookCorpus.featureWitnessPlan).toMatchObject({
      recordedCaseCount: 9_880,
      missingWitnessCount: 1,
      nextPlanCommand: 'pnpm public-workbook-corpus:feature-witness:plan',
      nextCheckCommand: 'pnpm public-workbook-corpus:feature-witness:check',
      coverage: [
        {
          id: 'pivots',
          label: 'pivots',
          totalCount: 0,
          witnessCaseCount: 0,
          needsWitness: true,
          discoveryQuery: 'pivot table xlsx',
          nextDiscoverCommand: expect.stringContaining("--query 'pivot table xlsx'"),
        },
      ],
    })
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

function financialPlanFixture(): PublicWorkbookCorpusFinancialPlan {
  return {
    schemaVersion: 1,
    mode: 'plan',
    corpus: 'financial-accounting-workpapers',
    generatedAt: '2026-05-08T10:00:00.000Z',
    manifestExists: true,
    targetWorkbookCount: 5_000,
    manifestPath: '.cache/public-workbook-corpus-financial/manifest.json',
    cacheDir: '.cache/public-workbook-corpus-financial',
    scorecardPath: '.cache/public-workbook-corpus-financial/scorecard.json',
    verifyCheckpointPath: '.cache/public-workbook-corpus-financial/verification-checkpoint.json',
    stopMarker: {
      active: true,
      path: '.agent-coordination/stop.md',
      overrideFlag: '--allow-active-stop-marker',
      overrideEnvVar: 'BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE',
    },
    sourceCount: 9_824,
    targetArtifactCount: 5_000,
    cachedArtifactCount: 0,
    remainingArtifactSlots: 5_000,
    candidateSourceCount: 8_949,
    candidateSourceDeficitCount: 0,
    minimumAdditionalSourceCount: 0,
    recommendedDiscoveryLimit: 9_824,
    recommendedFetchTrancheSize: 20,
    recommendedFetchLimit: 20,
    needsAdditionalDiscovery: false,
    targetReachableFromKnownCandidates: true,
    commands: {
      discoverPlan: null,
      discover: null,
      fetchPlan:
        'pnpm public-workbook-corpus:fetch-financial:plan -- --manifest .cache/public-workbook-corpus-financial/manifest.json --cache-dir .cache/public-workbook-corpus-financial --limit 5000',
      fetch:
        'BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE=1 pnpm public-workbook-corpus:fetch-financial -- --manifest .cache/public-workbook-corpus-financial/manifest.json --cache-dir .cache/public-workbook-corpus-financial --limit 20 --allow-active-stop-marker',
      fetchAll:
        'BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE=1 pnpm public-workbook-corpus:fetch-financial -- --manifest .cache/public-workbook-corpus-financial/manifest.json --cache-dir .cache/public-workbook-corpus-financial --limit 5000 --allow-active-stop-marker',
      resumePlan: 'pnpm public-workbook-corpus:resume-financial:plan',
      resumeCheck: 'pnpm public-workbook-corpus:resume-financial:check',
      verify:
        'BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE=1 pnpm public-workbook-corpus:verify-financial -- --manifest .cache/public-workbook-corpus-financial/manifest.json --cache-dir .cache/public-workbook-corpus-financial --scorecard .cache/public-workbook-corpus-financial/scorecard.json --verify-checkpoint .cache/public-workbook-corpus-financial/verification-checkpoint.json --allow-active-stop-marker',
      check: 'pnpm public-workbook-corpus:check-financial',
    },
    sampledCandidateSources: [],
  }
}

function featureWitnessPlanFixture(): PublicWorkbookCorpusFeatureWitnessPlan {
  return {
    schemaVersion: 1,
    mode: 'feature-witness-plan',
    generatedAt: '2026-05-08T10:00:00.000Z',
    stopMarker: {
      active: true,
      path: '.agent-coordination/stop.md',
      overrideFlag: '--allow-active-stop-marker',
      overrideEnvVar: 'BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE',
    },
    recordedCaseCount: 9_880,
    coverage: [
      {
        id: 'pivots',
        label: 'pivots',
        totalCount: 0,
        witnessCaseCount: 0,
        needsWitness: true,
        discoveryQuery: 'pivot table xlsx',
        commands: {
          discover:
            "BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE=1 pnpm public-workbook-corpus:discover -- --manifest .cache/public-workbook-corpus/manifest.json --cache-dir .cache/public-workbook-corpus --query 'pivot table xlsx' --limit 10000 --allow-active-stop-marker",
        },
      },
    ],
    missingWitnessCount: 1,
  }
}

function failingSameCorpusCapture(): SameCorpusCapture {
  return {
    schemaVersion: 1,
    suite: 'ui-responsiveness-same-corpus-capture',
    sampleCount: 3,
    limitations: [],
    cases: [
      {
        id: 'same-corpus-wide-mixed-250k-visible-scroll-response',
        corpusCaseId: 'wide-mixed-250k',
        materializedCells: 250_000,
        workload: 'visible-scroll-response',
        bilig: {
          product: 'bilig',
          source: 'e2e/tests/web-shell-scroll-performance.pw.ts',
          operationResponseMsSamples: [200, 200, 200],
          postOperationFrameMsSamples: [12, 12, 12],
          corpusVerification: sameCorpusVerification('bilig-benchmark-state'),
          limitations: [],
        },
        googleSheets: {
          product: 'google-sheets',
          source: 'https://docs.google.com/spreadsheets/d/sameCorpusSheet/edit',
          operationResponseMsSamples: [100, 100, 100],
          postOperationFrameMsSamples: [16, 16, 16],
          corpusVerification: sameCorpusVerification('google-sheets-xlsx-export'),
          limitations: [],
        },
        microsoftExcelWeb: {
          product: 'microsoft-excel-web',
          source: 'https://view.officeapps.live.com/op/view.aspx?src=sameCorpusWorkbook',
          operationResponseMsSamples: [100, 100, 100],
          postOperationFrameMsSamples: [16, 16, 16],
          corpusVerification: sameCorpusVerification('microsoft-excel-web-source-xlsx'),
          limitations: [],
        },
      },
    ],
  }
}

function sameCorpusVerification(method: SameCorpusCapture['cases'][number]['bilig']['corpusVerification']['method']) {
  return {
    verified: true,
    method,
    sheetName: 'WideGrid',
    materializedCells: 250_000,
    checkedCells: [
      { address: 'A1', expected: 'metric-1', actual: 'metric-1' },
      { address: 'B1', expected: 'metric-2', actual: 'metric-2' },
      { address: 'F2', expected: 'note-1-5', actual: 'note-1-5' },
    ],
  }
}
