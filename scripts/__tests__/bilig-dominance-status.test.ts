import { describe, expect, it } from 'vitest'

import { buildBiligDominanceStatus, formatBiligDominanceStatusPathForMessage } from '../bilig-dominance-status.ts'
import { buildSameCorpusProof, type SameCorpusCapture } from '../gen-ui-responsiveness-live-browser-scorecard.ts'
import type { PublicWorkbookCorpusFetchPlan } from '../public-workbook-corpus-fetch.ts'
import type { PublicWorkbookCorpusFinancialPlan } from '../public-workbook-corpus-financial-plan.ts'
import type { PublicWorkbookCorpusFeatureWitnessPlan } from '../public-workbook-corpus-feature-witness-plan.ts'
import type { PublicWorkbookCorpusStatus } from '../public-workbook-corpus-status.ts'
import type { SameCorpusPublicAccessCheck } from '../ui-responsiveness-same-corpus-public-access-check.ts'
import { requiredUiResponsivenessSameCorpusWorkloads } from '../ui-responsiveness-same-corpus-workloads.ts'
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
      requiredCaseCount: requiredUiResponsivenessSameCorpusWorkloads.length,
      tenXMeanAndP95CaseCount: 0,
      requiredWorkloads: [...requiredUiResponsivenessSameCorpusWorkloads],
      missingRequiredWorkloads: [...requiredUiResponsivenessSameCorpusWorkloads],
      googleSheetsUrl: null,
      googleSheetsUrlSource: 'missing',
      googleSheetsUrlEnvVar: 'BILIG_UI_SAME_CORPUS_GOOGLE_SHEETS_URL',
      microsoftExcelWebEditableUrl: null,
      microsoftExcelWebEditableUrlEnvVar: 'BILIG_UI_SAME_CORPUS_MICROSOFT_EXCEL_WEB_URL',
      publicAccessCheckPath: '.cache/ui-responsiveness/same-corpus-public-access-check.json',
      missingInputs: ['googleSheetsUrlForUploadedSameCorpusWorkbook'],
      fixture: {
        corpusCaseId: 'wide-mixed-250k',
        materializedCells: 250_000,
        localXlsxPath: 'packages/benchmarks/baselines/ui-same-corpus/wide-mixed-250k.xlsx',
      },
      nextFixtureCheckCommand: 'pnpm ui:same-corpus:fixture:check',
      nextPublicAccessCheckCommand: expect.stringContaining('pnpm ui:same-corpus:public-check'),
      nextGoogleSheetsStorageStateCommand: expect.stringContaining('--auth-product google-sheets'),
      nextMicrosoftExcelWebStorageStateCommand: expect.stringContaining('--auth-product microsoft-excel-web'),
      browserCaptureGuard: {
        active: false,
        activeMarkerPaths: [],
        overrideEnvVar: 'BILIG_ALLOW_LOCAL_CI_RESOURCE_GUARD',
        overridePrefix: null,
        nextPreflightRequiresOverride: false,
        nextCaptureRequiresOverride: false,
      },
      nextScorecardGenerateCommand: 'pnpm ui:browser-live:generate -- --capture .cache/ui-responsiveness/same-corpus-capture.json',
      nextDominanceCheckCommand: 'pnpm dominance:generate && pnpm dominance:check && pnpm dominance:audit:check',
      blockedCommands: [],
    })
    expect(status.uiSameCorpus.fixture.microsoftExcelWebUrl).toContain('view.officeapps.live.com/op/view.aspx')
    expect(status.uiSameCorpus.nextPublicAccessCheckCommand).toContain(
      '--output .cache/ui-responsiveness/same-corpus-public-access-check.json',
    )
    expect(status.uiSameCorpus.nextPreflightCommand).toContain('--google-sheets-url')
    expect(status.uiSameCorpus.nextPreflightCommand).toContain('<google-sheets-url>')
    expect(status.uiSameCorpus.nextPreflightCommand).not.toContain('<microsoft-excel-web-editable-url>')
    expect(status.uiSameCorpus.nextAuthenticatedPreflightCommand).toContain('--google-sheets-storage-state')
    expect(status.uiSameCorpus.nextAuthenticatedPreflightCommand).toContain('.cache/ui-responsiveness/google-sheets-storage-state.json')
    expect(status.uiSameCorpus.nextAuthenticatedPreflightCommand).not.toContain('--microsoft-excel-web-storage-state')
    expect(status.uiSameCorpus.nextCaptureCommand).toContain('.cache/ui-responsiveness/same-corpus-capture.json')
    expect(status.uiSameCorpus.nextCaptureCommand).not.toContain('<microsoft-excel-web-editable-url>')
    expect(status.uiSameCorpus.nextCaptureCommand).not.toContain('view.officeapps.live.com/op/view.aspx')
    expect(status.uiSameCorpus.nextAuthenticatedCaptureCommand).toContain('.cache/ui-responsiveness/same-corpus-capture.json')
    expect(status.uiSameCorpus.nextAuthenticatedCaptureCommand).toContain('--google-sheets-storage-state')
    expect(status.uiSameCorpus.nextAuthenticatedCaptureCommand).not.toContain('--microsoft-excel-web-storage-state')
    expect(status.uiSameCorpus.nextGoogleSheetsStorageStateCommand).toContain('.cache/ui-responsiveness/google-sheets-storage-state.json')
    expect(status.uiSameCorpus.nextGoogleSheetsStorageStateCommand).toContain('<google-sheets-url>')
    expect(status.uiSameCorpus.nextMicrosoftExcelWebStorageStateCommand).toContain(
      '.cache/ui-responsiveness/microsoft-excel-web-storage-state.json',
    )
    expect(status.uiSameCorpus.nextMicrosoftExcelWebStorageStateCommand).toContain('<microsoft-excel-web-editable-url>')
    expect(status.uiSameCorpus.nextGoogleSheetsUploadInstruction).toContain('share it to anyone with the link')
    expect(status.uiSameCorpus.nextMicrosoftExcelWebUploadInstruction).toBeNull()
  })

  it('fills same-corpus UI proof commands when incumbent editable URLs are known', () => {
    const googleSheetsUrl = 'https://docs.google.com/spreadsheets/d/sameCorpusSheet/edit'
    const microsoftExcelWebUrl = 'https://m365.cloud.microsoft/launch/excel?sameCorpusWorkbook'
    const status = buildBiligDominanceStatus({
      input: buildFixtureInput(),
      financialCorpusStatus: completeFinancialCorpusStatus(),
      publicWorkbookCorpusStatus: completePublicWorkbookCorpusStatus(),
      stopMarkerActive: false,
      stopMarkerPath: '/repo/.agent-coordination/stop.md',
      uiSameCorpusGoogleSheetsUrl: googleSheetsUrl,
      uiSameCorpusMicrosoftExcelWebUrl: microsoftExcelWebUrl,
    })

    expect(status.uiSameCorpus.googleSheetsUrl).toBe(googleSheetsUrl)
    expect(status.uiSameCorpus.googleSheetsUrlSource).toBe('argument-or-environment')
    expect(status.uiSameCorpus.microsoftExcelWebEditableUrl).toBe(microsoftExcelWebUrl)
    expect(status.uiSameCorpus.missingInputs).toEqual([])
    expect(status.uiSameCorpus.nextPreflightCommand).toContain(googleSheetsUrl)
    expect(status.uiSameCorpus.nextPreflightCommand).not.toContain(microsoftExcelWebUrl)
    expect(status.uiSameCorpus.nextPreflightCommand).not.toContain('<google-sheets-url>')
    expect(status.uiSameCorpus.nextPreflightCommand).not.toContain('<microsoft-excel-web-editable-url>')
    expect(status.uiSameCorpus.nextAuthenticatedPreflightCommand).toContain(googleSheetsUrl)
    expect(status.uiSameCorpus.nextAuthenticatedPreflightCommand).not.toContain(microsoftExcelWebUrl)
    expect(status.uiSameCorpus.nextPublicAccessCheckCommand).toContain(googleSheetsUrl)
    expect(status.uiSameCorpus.nextPublicAccessCheckCommand).not.toContain('<google-sheets-url>')
    expect(status.uiSameCorpus.nextGoogleSheetsStorageStateCommand).toContain(googleSheetsUrl)
    expect(status.uiSameCorpus.nextMicrosoftExcelWebStorageStateCommand).toContain(microsoftExcelWebUrl)
    expect(status.uiSameCorpus.nextCaptureCommand).toContain(googleSheetsUrl)
    expect(status.uiSameCorpus.nextCaptureCommand).not.toContain(microsoftExcelWebUrl)
    expect(status.uiSameCorpus.nextAuthenticatedCaptureCommand).toContain(googleSheetsUrl)
    expect(status.uiSameCorpus.nextAuthenticatedCaptureCommand).not.toContain(microsoftExcelWebUrl)
    expect(status.uiSameCorpus.nextCaptureCommand).not.toContain('<google-sheets-url>')
    expect(status.uiSameCorpus.nextCaptureCommand).not.toContain('<microsoft-excel-web-editable-url>')
  })

  it('reuses verified public-access evidence for same-corpus UI proof commands', () => {
    const googleSheetsUrl = 'https://docs.google.com/spreadsheets/d/sameCorpusSheet/edit'
    const status = buildBiligDominanceStatus({
      input: buildFixtureInput(),
      financialCorpusStatus: completeFinancialCorpusStatus(),
      publicWorkbookCorpusStatus: completePublicWorkbookCorpusStatus(),
      stopMarkerActive: false,
      stopMarkerPath: '/repo/.agent-coordination/stop.md',
      uiSameCorpusPublicAccessCheck: sameCorpusPublicAccessCheckFixture(googleSheetsUrl),
      uiSameCorpusPublicAccessCheckPath: '.cache/ui-responsiveness/same-corpus-public-access-check.json',
    })

    expect(status.uiSameCorpus.googleSheetsUrl).toBe(googleSheetsUrl)
    expect(status.uiSameCorpus.googleSheetsUrlSource).toBe('public-access-check')
    expect(status.uiSameCorpus.microsoftExcelWebEditableUrl).toBeNull()
    expect(status.uiSameCorpus.missingInputs).toEqual([])
    expect(status.uiSameCorpus.nextPreflightCommand).toContain(googleSheetsUrl)
    expect(status.uiSameCorpus.nextPreflightCommand).not.toContain('<google-sheets-url>')
    expect(status.uiSameCorpus.nextPreflightCommand).not.toContain('<microsoft-excel-web-editable-url>')
    expect(status.uiSameCorpus.nextAuthenticatedPreflightCommand).toContain(googleSheetsUrl)
    expect(status.uiSameCorpus.nextAuthenticatedPreflightCommand).not.toContain('<microsoft-excel-web-editable-url>')
    expect(status.uiSameCorpus.nextGoogleSheetsStorageStateCommand).toContain(googleSheetsUrl)
    expect(status.uiSameCorpus.nextMicrosoftExcelWebStorageStateCommand).toContain('<microsoft-excel-web-editable-url>')
    expect(status.uiSameCorpus.nextPublicAccessCheckCommand).toContain(
      '--output .cache/ui-responsiveness/same-corpus-public-access-check.json',
    )
    expect(status.uiSameCorpus.nextCaptureCommand).toContain(googleSheetsUrl)
    expect(status.uiSameCorpus.nextCaptureCommand).not.toContain('<microsoft-excel-web-editable-url>')
    expect(status.uiSameCorpus.nextGoogleSheetsUploadInstruction).toBeNull()
    expect(status.uiSameCorpus.nextMicrosoftExcelWebUploadInstruction).toBeNull()
  })

  it('surfaces local resource guard state before same-corpus browser capture', () => {
    const status = buildBiligDominanceStatus({
      input: buildFixtureInput(),
      financialCorpusStatus: completeFinancialCorpusStatus(),
      publicWorkbookCorpusStatus: completePublicWorkbookCorpusStatus(),
      stopMarkerActive: false,
      stopMarkerPath: '/repo/.agent-coordination/stop.md',
      uiSameCorpusLocalCiResourceGuardStatus: {
        activeMarkerPaths: [
          '.agent-coordination/20260507T074946Z-codex-stop-interactive-corpus-runs.md',
          '.agent-coordination/20260508T092619Z-codex-memory-pressure-stop.md',
        ],
      },
    })

    expect(status.uiSameCorpus.browserCaptureGuard).toEqual({
      active: true,
      activeMarkerPaths: [
        '.agent-coordination/20260507T074946Z-codex-stop-interactive-corpus-runs.md',
        '.agent-coordination/20260508T092619Z-codex-memory-pressure-stop.md',
      ],
      overrideEnvVar: 'BILIG_ALLOW_LOCAL_CI_RESOURCE_GUARD',
      overridePrefix: 'BILIG_ALLOW_LOCAL_CI_RESOURCE_GUARD=1',
      nextPreflightRequiresOverride: true,
      nextCaptureRequiresOverride: true,
    })
    expect(status.uiSameCorpus.nextPreflightCommand).toBeNull()
    expect(status.uiSameCorpus.nextAuthenticatedPreflightCommand).toBeNull()
    expect(status.uiSameCorpus.nextCaptureCommand).toBeNull()
    expect(status.uiSameCorpus.nextAuthenticatedCaptureCommand).toBeNull()
    expect(status.uiSameCorpus.nextScorecardGenerateCommand).toBeNull()
    expect(status.uiSameCorpus.nextGoogleSheetsStorageStateCommand).toBeNull()
    expect(status.uiSameCorpus.nextMicrosoftExcelWebStorageStateCommand).toBeNull()
    expect(status.uiSameCorpus.blockedCommands).toEqual([
      expect.stringContaining(
        'BILIG_ALLOW_LOCAL_CI_RESOURCE_GUARD=1 pnpm ui:same-corpus:capture -- --save-storage-state .cache/ui-responsiveness/google-sheets-storage-state.json',
      ),
      expect.stringContaining(
        'BILIG_ALLOW_LOCAL_CI_RESOURCE_GUARD=1 pnpm ui:same-corpus:capture -- --save-storage-state .cache/ui-responsiveness/microsoft-excel-web-storage-state.json',
      ),
      expect.stringContaining('BILIG_ALLOW_LOCAL_CI_RESOURCE_GUARD=1 pnpm ui:same-corpus:capture -- --preflight'),
      expect.stringContaining('--google-sheets-storage-state .cache/ui-responsiveness/google-sheets-storage-state.json'),
      expect.stringContaining(
        'BILIG_ALLOW_LOCAL_CI_RESOURCE_GUARD=1 pnpm ui:same-corpus:capture -- --output .cache/ui-responsiveness/same-corpus-capture.json',
      ),
      expect.stringContaining('--google-sheets-storage-state .cache/ui-responsiveness/google-sheets-storage-state.json'),
      expect.stringContaining(
        'BILIG_ALLOW_LOCAL_CI_RESOURCE_GUARD=1 pnpm ui:browser-live:generate -- --capture .cache/ui-responsiveness/same-corpus-capture.json',
      ),
    ])
    expect(status.localCiResourceGuard).toEqual({
      active: true,
      activeMarkerPaths: [
        '.agent-coordination/20260507T074946Z-codex-stop-interactive-corpus-runs.md',
        '.agent-coordination/20260508T092619Z-codex-memory-pressure-stop.md',
      ],
      overrideEnvVar: 'BILIG_ALLOW_LOCAL_CI_RESOURCE_GUARD',
      overridePrefix: 'BILIG_ALLOW_LOCAL_CI_RESOURCE_GUARD=1',
    })
    expect(status.unmetRequirements).toContain(
      'operator/developer workflow local CI resource guard active: .agent-coordination/20260507T074946Z-codex-stop-interactive-corpus-runs.md, .agent-coordination/20260508T092619Z-codex-memory-pressure-stop.md',
    )
  })

  it('reuses verified checked-in capture URLs when same-corpus proof is not 10x', () => {
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
      requiredCaseCount: requiredUiResponsivenessSameCorpusWorkloads.length,
      tenXMeanAndP95CaseCount: 0,
      tenXRequirementSatisfied: false,
      missingRequiredWorkloads: [],
      scrollEventEvidenceCaseCount: 3,
      casesMissingScrollEventEvidence: [],
      googleSheetsUrl: 'https://docs.google.com/spreadsheets/d/sameCorpusSheet/edit',
      googleSheetsUrlSource: 'checked-in-capture',
      missingInputs: [],
    })
    expect(status.uiSameCorpus.nextPreflightCommand).toContain('https://docs.google.com/spreadsheets/d/sameCorpusSheet/edit')
    expect(status.uiSameCorpus.nextPreflightCommand).not.toContain('<microsoft-excel-web-editable-url>')
    expect(status.uiSameCorpus.nextPreflightCommand).not.toContain('<google-sheets-url>')
    expect(status.uiSameCorpus.nextCaptureCommand).toContain('https://docs.google.com/spreadsheets/d/sameCorpusSheet/edit')
    expect(status.uiSameCorpus.nextCaptureCommand).not.toContain('<microsoft-excel-web-editable-url>')
    expect(status.uiSameCorpus.nextCaptureCommand).not.toContain('<google-sheets-url>')
    expect(status.uiSameCorpus.nextGoogleSheetsUploadInstruction).toBeNull()
  })

  it('keeps asking for the Google Sheets URL when checked-in capture sources are inconsistent', () => {
    const fixtureInput = buildFixtureInput()
    const status = buildBiligDominanceStatus({
      input: {
        ...fixtureInput,
        uiResponsivenessLiveBrowserScorecard: {
          ...fixtureInput.uiResponsivenessLiveBrowserScorecard,
          sameCorpusProof: buildSameCorpusProof(
            failingSameCorpusCapture({
              firstGoogleSheetsSource: 'https://docs.google.com/spreadsheets/d/differentSameCorpusSheet/edit',
            }),
          ),
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
      tenXRequirementSatisfied: false,
      googleSheetsUrl: null,
      googleSheetsUrlSource: 'missing',
      missingInputs: ['googleSheetsUrlForUploadedSameCorpusWorkbook'],
    })
    expect(status.uiSameCorpus.nextPreflightCommand).toContain('<google-sheets-url>')
    expect(status.uiSameCorpus.nextCaptureCommand).toContain('<google-sheets-url>')
  })

  it('keeps asking for the Google Sheets URL when checked-in capture source is not verified', () => {
    const fixtureInput = buildFixtureInput()
    const checkedInProof = buildSameCorpusProof(failingSameCorpusCapture())
    const casesWithUnverifiedGoogleSheetsProof = [...checkedInProof.cases]
    const firstCase = casesWithUnverifiedGoogleSheetsProof[0]
    if (!firstCase) {
      throw new Error('Expected failing same-corpus capture fixture to include at least one case')
    }
    casesWithUnverifiedGoogleSheetsProof[0] = {
      ...firstCase,
      googleSheets: {
        ...firstCase.googleSheets,
        corpusVerification: {
          ...firstCase.googleSheets.corpusVerification,
          verified: false,
        },
      },
    }
    const status = buildBiligDominanceStatus({
      input: {
        ...fixtureInput,
        uiResponsivenessLiveBrowserScorecard: {
          ...fixtureInput.uiResponsivenessLiveBrowserScorecard,
          sameCorpusProof: {
            ...checkedInProof,
            cases: casesWithUnverifiedGoogleSheetsProof,
          },
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
      tenXRequirementSatisfied: false,
      googleSheetsUrl: null,
      googleSheetsUrlSource: 'missing',
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
      targetComplete: true,
    })
    expect(status.importExportBlockers).toEqual(
      expect.arrayContaining([
        'financial/accounting corpus cached artifacts below target: 0/5000',
        'financial/accounting corpus recorded verification cases below target: 0/5000',
      ]),
    )
    expect(status.goalStatus).toBe('active-not-achieved')
  })

  it('separates stop-marker-blocked public corpus runs from runnable next commands', () => {
    const status = buildBiligDominanceStatus({
      input: buildFixtureInput(),
      fetchPlan: incompletePublicWorkbookFetchPlan(),
      financialCorpusStatus: completeFinancialCorpusStatus(),
      publicWorkbookCorpusStatus: incompletePublicWorkbookCorpusStatus(),
      stopMarkerActive: true,
      stopMarkerPath: '.agent-coordination/stop.md',
    })

    expect(status.publicWorkbookCorpus.nextMissingVerificationPlanCommand).toBe('pnpm public-workbook-corpus:verify-missing:plan')
    expect(status.publicWorkbookCorpus.nextFetchCommand).toBeNull()
    expect(status.publicWorkbookCorpus.nextMissingVerificationCommand).toBeNull()
    expect(status.publicWorkbookCorpus.nextStaleVerificationPlanCommand).toBe('pnpm public-workbook-corpus:verify-stale:plan')
    expect(status.publicWorkbookCorpus.nextStaleVerificationCommand).toBeNull()
    expect(status.publicWorkbookCorpus.blockedCommands).toEqual([
      'BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE=1 pnpm public-workbook-corpus:discover -- --limit 10060 --allow-active-stop-marker',
      'BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE=1 pnpm public-workbook-corpus:fetch -- --limit 9906 --fetch-batch-size 6 --allow-active-stop-marker',
      'BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE=1 pnpm public-workbook-corpus:verify-missing -- --limit 1 --allow-active-stop-marker',
      'BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE=1 pnpm public-workbook-corpus:verify-stale -- --limit 1 --allow-active-stop-marker',
    ])
    expect(status.publicWorkbookCorpus.nextCorpusRunRequiresExplicitResume).toBe(true)
    expect(status.publicWorkbookCorpus.corpusRunStopMarkerOverrideEnvVar).toBe('BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE')
    expect(status.publicWorkbookCorpus.corpusRunStopMarkerOverrideFlag).toBe('--allow-active-stop-marker')
  })

  it('surfaces current and stale unsupported public corpus evidence counts', () => {
    const status = buildBiligDominanceStatus({
      input: buildFixtureInput(),
      financialCorpusStatus: completeFinancialCorpusStatus(),
      publicWorkbookCorpusStatus: {
        ...incompletePublicWorkbookCorpusStatus(),
        recordedUnsupportedCaseCount: 2_054,
        currentRecordedUnsupportedCaseCount: 58,
        staleRecordedUnsupportedCaseCount: 1_996,
        currentUnsupportedClassifications: [{ classification: 'xlsx.publicCorpus.resourceLimit:rss>1536MiB', count: 43 }],
        staleUnsupportedClassifications: [
          { classification: 'xlsx.import.warning:Some defined names were ignored during XLSX import.', count: 1_926 },
        ],
      },
      stopMarkerActive: true,
      stopMarkerPath: '.agent-coordination/stop.md',
    })

    expect(status.publicWorkbookCorpus).toMatchObject({
      recordedUnsupportedCaseCount: 2_054,
      currentRecordedUnsupportedCaseCount: 58,
      staleRecordedUnsupportedCaseCount: 1_996,
      currentUnsupportedClassifications: [{ classification: 'xlsx.publicCorpus.resourceLimit:rss>1536MiB', count: 43 }],
      staleUnsupportedClassifications: [
        { classification: 'xlsx.import.warning:Some defined names were ignored during XLSX import.', count: 1_926 },
      ],
    })
  })

  it('separates stop-marker-blocked financial corpus runs from runnable plan commands', () => {
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
      recommendedFetchBatchSize: 6,
      needsAdditionalDiscovery: false,
      targetReachableFromKnownCandidates: true,
      nextPlanCommand: 'pnpm public-workbook-corpus:discover-financial:plan',
      nextCheckCommand: 'pnpm public-workbook-corpus:discover-financial:check',
      nextFetchPlanCommand: expect.stringContaining('public-workbook-corpus:fetch-financial:plan'),
      nextFetchCommand: null,
      nextVerifyCommand: null,
      blockedCommands: [
        expect.stringContaining('public-workbook-corpus:fetch-financial'),
        expect.stringContaining('--limit 5000'),
        expect.stringContaining('public-workbook-corpus:verify-financial'),
      ],
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
          nextDiscoverCommand: null,
          blockedDiscoverCommand: expect.stringContaining("--query 'pivot table xlsx'"),
        },
      ],
    })
    expect(status.importExportBlockers).toContain('public workbook corpus missing feature witness coverage: pivots')
    expect(status.unmetRequirements).toContain('public workbook corpus missing feature witness coverage: pivots')
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
    currentRecordedUnsupportedCaseCount: 0,
    staleRecordedUnsupportedCaseCount: 0,
    currentUnsupportedClassifications: [],
    staleUnsupportedClassifications: [],
    recordedFailedCaseCount: 0,
    recordedErrorCaseCount: 0,
    recordedCoversManifest: true,
    recordedAllCasesPassed: true,
    missingManifestArtifactSample: [],
    staleRecordedVerificationSample: [],
    nextMissingVerificationCommand: null,
    nextMissingVerificationPlanCommand: null,
    blockedMissingVerificationCommand: null,
    nextStaleVerificationCommand: null,
    nextStaleVerificationPlanCommand: null,
    blockedStaleVerificationCommand: null,
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

function incompletePublicWorkbookCorpusStatus(): PublicWorkbookCorpusStatus {
  return {
    ...completePublicWorkbookCorpusStatus(),
    cachedArtifactCount: 9_900,
    scorecardCaseCount: 9_750,
    checkpointCaseCount: 9_750,
    recordedManifestArtifactCount: 9_750,
    missingManifestArtifactCount: 150,
    staleRecordedVerificationCount: 25,
    recordedPassedCaseCount: 9_750,
    recordedCoversManifest: false,
    nextMissingVerificationCommand: null,
    nextMissingVerificationPlanCommand: 'pnpm public-workbook-corpus:verify-missing:plan',
    blockedMissingVerificationCommand:
      'BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE=1 pnpm public-workbook-corpus:verify-missing -- --limit 1 --allow-active-stop-marker',
    nextStaleVerificationCommand: null,
    nextStaleVerificationPlanCommand: 'pnpm public-workbook-corpus:verify-stale:plan',
    blockedStaleVerificationCommand:
      'BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE=1 pnpm public-workbook-corpus:verify-stale -- --limit 1 --allow-active-stop-marker',
    scorecardCoversManifest: false,
    targetComplete: false,
    gaps: [
      'cached artifacts below target: 9900/10000',
      'scorecard cases do not cover manifest artifacts: 9750/9900',
      'recorded verification cases below cached artifacts: 9750/9900',
      'recorded verification cases need evidence refresh: 25',
    ],
  }
}

function incompletePublicWorkbookFetchPlan(): PublicWorkbookCorpusFetchPlan {
  return {
    targetArtifactCount: 10_000,
    cachedArtifactCount: 9_900,
    sourceCount: 10_000,
    remainingArtifactSlots: 100,
    candidateSourceCount: 40,
    candidateSourceDeficitCount: 60,
    minimumAdditionalSourceCount: 60,
    recommendedDiscoveryLimit: 10_060,
    targetReachableFromKnownCandidates: false,
    sampledCandidateSources: [],
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
    recommendedFetchBatchSize: 6,
    recommendedFetchLimit: 20,
    needsAdditionalDiscovery: false,
    targetReachableFromKnownCandidates: true,
    commands: {
      discoverPlan: null,
      discover: null,
      fetchPlan:
        'pnpm public-workbook-corpus:fetch-financial:plan -- --manifest .cache/public-workbook-corpus-financial/manifest.json --cache-dir .cache/public-workbook-corpus-financial --limit 5000',
      fetch:
        'BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE=1 pnpm public-workbook-corpus:fetch-financial -- --manifest .cache/public-workbook-corpus-financial/manifest.json --cache-dir .cache/public-workbook-corpus-financial --limit 20 --fetch-batch-size 6 --allow-active-stop-marker',
      fetchAll:
        'BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE=1 pnpm public-workbook-corpus:fetch-financial -- --manifest .cache/public-workbook-corpus-financial/manifest.json --cache-dir .cache/public-workbook-corpus-financial --limit 5000 --fetch-batch-size 6 --allow-active-stop-marker',
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
          discover: null,
        },
        blockedCommands: {
          discover:
            "BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE=1 pnpm public-workbook-corpus:discover -- --manifest .cache/public-workbook-corpus/manifest.json --cache-dir .cache/public-workbook-corpus --query 'pivot table xlsx' --limit 10000 --allow-active-stop-marker",
        },
      },
    ],
    missingWitnessCount: 1,
    missingWitnesses: [
      {
        id: 'pivots',
        label: 'pivots',
        discoveryQuery: 'pivot table xlsx',
        discoverCommand: null,
        blockedDiscoverCommand:
          "BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE=1 pnpm public-workbook-corpus:discover -- --manifest .cache/public-workbook-corpus/manifest.json --cache-dir .cache/public-workbook-corpus --query 'pivot table xlsx' --limit 10000 --allow-active-stop-marker",
        cachedCandidateCount: 0,
        cachedCandidates: [],
      },
    ],
  }
}

function sameCorpusPublicAccessCheckFixture(googleSheetsUrl: string): SameCorpusPublicAccessCheck {
  return {
    schemaVersion: 1,
    suite: 'ui-responsiveness-same-corpus-public-access-check',
    generatedAt: '2026-05-08T10:00:00.000Z',
    corpusCaseId: 'wide-mixed-250k',
    materializedCells: 250_000,
    requestedProductCount: 1,
    verifiedProductCount: 1,
    allRequestedProductsVerified: true,
    products: [
      {
        product: 'google-sheets',
        source: googleSheetsUrl,
        resolvedXlsxUrl: 'https://docs.google.com/spreadsheets/d/sameCorpusSheet/export?format=xlsx',
        byteSize: 1_440_952,
        corpusVerification: {
          verified: true,
          method: 'google-sheets-xlsx-export',
          sheetName: 'WideGrid',
          materializedCells: 250_000,
          checkedCells: [
            { address: 'A1', expected: 'metric-1', actual: 'metric-1' },
            { address: 'B1', expected: 'metric-2', actual: 'metric-2' },
            { address: 'AP977', expected: 'note-8-2', actual: 'note-8-2' },
          ],
        },
        limitations: ['Google Sheets must be shared so anyone with the link can export the native sheet as XLSX.'],
      },
    ],
    limitations: [
      'This check proves URL reachability and same-corpus workbook identity through XLSX export bytes.',
      'It is not browser timing evidence and does not satisfy the same-corpus 10x UI responsiveness requirement by itself.',
    ],
  }
}

function failingSameCorpusCapture(
  args: {
    readonly firstGoogleSheetsSource?: string
    readonly firstGoogleSheetsVerification?: SameCorpusCapture['cases'][number]['googleSheets']['corpusVerification']
  } = {},
): SameCorpusCapture {
  const defaultGoogleSheetsSource = 'https://docs.google.com/spreadsheets/d/sameCorpusSheet/edit'
  const defaultGoogleSheetsVerification = sameCorpusVerification('google-sheets-xlsx-export')
  return {
    schemaVersion: 1,
    suite: 'ui-responsiveness-same-corpus-capture',
    sampleCount: 3,
    limitations: [],
    cases: requiredUiResponsivenessSameCorpusWorkloads.map((workload, index) => ({
      id: `same-corpus-wide-mixed-250k-${workload}`,
      corpusCaseId: 'wide-mixed-250k',
      materializedCells: 250_000,
      workload,
      scenarioProof: sameCorpusScenarioProof(200, 100),
      bilig: {
        product: 'bilig',
        source: 'e2e/tests/web-shell-scroll-performance.pw.ts',
        operationResponseMsSamples: [200, 200, 200],
        postOperationFrameMsSamples: [12, 12, 12],
        ...(workload === 'scroll-vertical' || workload === 'scroll-horizontal' || workload === 'wide-sheet-navigation'
          ? { scrollEventResponseMsSamples: [200, 200, 200], scrollMovementPxSamples: [720, 720, 720] }
          : {}),
        corpusVerification: sameCorpusVerification('bilig-benchmark-state'),
        limitations: [],
      },
      googleSheets: {
        product: 'google-sheets',
        source: index === 0 ? (args.firstGoogleSheetsSource ?? defaultGoogleSheetsSource) : defaultGoogleSheetsSource,
        operationResponseMsSamples: [100, 100, 100],
        postOperationFrameMsSamples: [16, 16, 16],
        ...(workload === 'scroll-vertical' || workload === 'scroll-horizontal' || workload === 'wide-sheet-navigation'
          ? { scrollEventResponseMsSamples: [100, 100, 100], scrollMovementPxSamples: [720, 720, 720] }
          : {}),
        corpusVerification:
          index === 0 ? (args.firstGoogleSheetsVerification ?? defaultGoogleSheetsVerification) : defaultGoogleSheetsVerification,
        limitations: [],
      },
      microsoftExcelWeb: {
        product: 'microsoft-excel-web',
        source: 'https://view.officeapps.live.com/op/view.aspx?src=sameCorpusWorkbook',
        operationResponseMsSamples: [100, 100, 100],
        postOperationFrameMsSamples: [16, 16, 16],
        ...(workload === 'scroll-vertical' || workload === 'scroll-horizontal' || workload === 'wide-sheet-navigation'
          ? { scrollEventResponseMsSamples: [100, 100, 100], scrollMovementPxSamples: [720, 720, 720] }
          : {}),
        corpusVerification: sameCorpusVerification('microsoft-excel-web-source-xlsx'),
        limitations: [],
      },
    })),
  }
}

function sameCorpusScenarioProof(biligMs: number, googleMs: number) {
  const microsoftExcelWebMs = googleMs
  return {
    biligMeanMs: biligMs,
    biligP95Ms: biligMs,
    googleMeanMs: googleMs,
    googleP95Ms: googleMs,
    microsoftExcelWebMeanMs: microsoftExcelWebMs,
    microsoftExcelWebP95Ms: microsoftExcelWebMs,
    meanRatio: biligMs / googleMs,
    p95Ratio: biligMs / googleMs,
    microsoftExcelWebMeanRatio: biligMs / microsoftExcelWebMs,
    microsoftExcelWebP95Ratio: biligMs / microsoftExcelWebMs,
    screenshotProof: {
      captured: true,
      requiredProducts: ['bilig', 'google-sheets'],
      artifactPaths: ['tmp/bilig-sample-1.png', 'tmp/google-sheets-sample-1.png', 'tmp/microsoft-excel-web-sample-1.png'],
      missingProducts: [],
    },
    pixelGridProof: {
      captured: true,
      requiredProducts: ['bilig', 'google-sheets'],
      products: [
        {
          product: 'bilig',
          captured: true,
          method: 'typegpu-visible-canvas',
          viewportPixelWidth: 1440,
          viewportPixelHeight: 900,
          evidence: ['mode=typegpu-v3'],
        },
        {
          product: 'google-sheets',
          captured: true,
          method: 'google-sheets-visible-grid',
          viewportPixelWidth: 1440,
          viewportPixelHeight: 900,
          evidence: ['selector=.grid-scrollable-wrapper'],
        },
        {
          product: 'microsoft-excel-web',
          captured: true,
          method: 'excel-web-visible-grid',
          viewportPixelWidth: 1440,
          viewportPixelHeight: 900,
          evidence: ['selector=.ewr-grdcontarea-grid'],
        },
      ],
      missingProducts: [],
    },
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
