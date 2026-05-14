import { describe, expect, it } from 'vitest'

import { summarizeNumbers } from '../../packages/benchmarks/src/stats.js'
import { buildBiligDominancePromptArtifactAudit, validateBiligDominancePromptArtifactAudit } from '../bilig-dominance-audit.ts'
import { buildBiligDominanceStatus } from '../bilig-dominance-status.ts'
import type {
  UiResponsivenessSameCorpusMeasurement,
  UiResponsivenessSameCorpusProduct,
  UiResponsivenessSameCorpusProof,
} from '../gen-ui-responsiveness-live-browser-scorecard.ts'
import { buildBiligDominanceScorecard } from '../gen-bilig-dominance-scorecard.ts'
import type { PublicWorkbookCorpusStatus } from '../public-workbook-corpus-status.ts'
import { requiredUiResponsivenessSameCorpusWorkloads } from '../ui-responsiveness-same-corpus-workloads.ts'
import { buildFixtureInput } from './bilig-dominance-scorecard.fixture.ts'

const requiredUiSameCorpusWorkloadList = requiredUiResponsivenessSameCorpusWorkloads.join(', ')
const requiredUiSameCorpusInputList = 'googleSheetsUrlForUploadedSameCorpusWorkbook'

describe('bilig dominance prompt-to-artifact audit', () => {
  it('maps every objective criterion to evidence artifacts, check commands, and live blockers', () => {
    const input = buildFixtureInput()
    const status = buildBiligDominanceStatus({
      input,
      publicWorkbookCorpusStatus: publicWorkbookCorpusStatusFixture({
        targetWorkbookCount: 10_000,
        sourceCount: 10_000,
        cachedArtifactCount: 5_628,
        recordedManifestArtifactCount: 4_940,
        missingManifestArtifactCount: 688,
        scorecardCaseCount: 2_000,
        checkpointCaseCount: 4_940,
        recordedCoversManifest: false,
        scorecardCoversManifest: false,
        gaps: [
          'cached artifacts below target: 5628/10000',
          'scorecard cases do not cover manifest artifacts: 2000/5628',
          'recorded verification cases below cached artifacts: 4940/5628',
        ],
      }),
      stopMarkerActive: true,
      stopMarkerPath: '/repo/.agent-coordination/stop.md',
    })
    const audit = buildBiligDominancePromptArtifactAudit({
      scorecard: buildBiligDominanceScorecard(input),
      status,
    })

    expect(audit.completionVerdict).toMatchObject({
      goalStatus: 'active-not-achieved',
      blanketTenXClaimAllowed: false,
      allChecklistItemsPassed: false,
    })
    expect(audit.checklist.map((entry) => entry.id)).toEqual([
      'calculation-correctness',
      'recalculation-speed',
      'structural-edit-performance',
      'large-workbook-scale',
      'ui-responsiveness',
      'collaboration',
      'automation-api-extensibility',
      'import-export-compatibility',
      'auditability',
      'reliability',
      'security',
      'operator-developer-workflow',
    ])
    expect(audit.checklist.every((entry) => entry.promptRequirement.length > 0)).toBe(true)
    expect(audit.checklist.every((entry) => entry.evidenceArtifacts.length > 0)).toBe(true)
    expect(audit.checklist.every((entry) => entry.checkCommands.length > 0)).toBe(true)
    expect(audit.checklist.find((entry) => entry.id === 'import-export-compatibility')).toMatchObject({
      passed: false,
      liveBlockers: [
        'public workbook corpus cached artifacts below target: 5628/10000',
        'public workbook corpus scorecard cases below cached artifacts: 2000/5628',
        'public workbook corpus recorded verification cases below cached artifacts: 4940/5628',
      ],
      gaps: [
        'public workbook corpus cached artifacts below target: 5628/10000',
        'public workbook corpus scorecard cases below cached artifacts: 2000/5628',
        'public workbook corpus recorded verification cases below cached artifacts: 4940/5628',
      ],
      evidenceArtifacts: expect.arrayContaining(['packages/benchmarks/baselines/public-workbook-corpus-scorecard.json']),
      checkCommands: expect.arrayContaining(['pnpm public-workbook-corpus:check']),
      evidence: expect.arrayContaining([
        'live public workbook corpus cached artifacts: 5628/10000',
        'live public workbook corpus recorded verification cases: 4940/5628',
        'live public workbook corpus scorecard cases: 2000/5628',
        'live public workbook corpus recorded all cases passed: true',
      ]),
    })
    expect(audit.livePublicWorkbookCorpus).toMatchObject({
      cachedArtifactCount: 5_628,
      recordedManifestArtifactCount: 4_940,
      corpusRunStopMarkerActive: true,
      nextCorpusRunRequiresExplicitResume: true,
      nextStaleVerificationCommand: null,
    })
    expect(audit.liveUiSameCorpus).toMatchObject({
      captured: false,
      missingInputs: ['googleSheetsUrlForUploadedSameCorpusWorkbook'],
      fixture: {
        corpusCaseId: 'wide-mixed-250k',
        materializedCells: 250_000,
      },
    })
    expect(audit.liveLocalCiResourceGuard).toMatchObject({
      active: false,
      activeMarkerPaths: [],
      overrideEnvVar: 'BILIG_ALLOW_LOCAL_CI_RESOURCE_GUARD',
      overridePrefix: null,
    })
    expect(validateBiligDominancePromptArtifactAudit(audit)).toEqual([])
  })

  it('keeps UI proof gaps visible while surfacing local browser-capture guards', () => {
    const input = buildFixtureInput()
    const status = buildBiligDominanceStatus({
      input,
      publicWorkbookCorpusStatus: publicWorkbookCorpusStatusFixture(),
      stopMarkerActive: false,
      stopMarkerPath: '/repo/.agent-coordination/stop.md',
      uiSameCorpusLocalCiResourceGuardStatus: {
        activeMarkerPaths: ['.agent-coordination/20260508T092619Z-codex-memory-pressure-stop.md'],
      },
    })
    const audit = buildBiligDominancePromptArtifactAudit({
      scorecard: buildBiligDominanceScorecard(input),
      status,
    })
    const uiItem = audit.checklist.find((entry) => entry.id === 'ui-responsiveness')

    expect(uiItem).toMatchObject({
      passed: false,
      liveBlockers: [
        'same-corpus UI browser capture has not been recorded',
        `same-corpus UI proof missing required workloads: ${requiredUiSameCorpusWorkloadList}`,
        `same-corpus UI proof missing inputs: ${requiredUiSameCorpusInputList}`,
        'same-corpus UI browser capture paused by local resource guard: .agent-coordination/20260508T092619Z-codex-memory-pressure-stop.md',
      ],
      gaps: [
        'live UI browser evidence is not a same-corpus 10x proof against incumbents',
        'same-corpus UI browser capture has not been recorded',
        `same-corpus UI proof missing required workloads: ${requiredUiSameCorpusWorkloadList}`,
        `same-corpus UI proof missing inputs: ${requiredUiSameCorpusInputList}`,
        'same-corpus UI browser capture paused by local resource guard: .agent-coordination/20260508T092619Z-codex-memory-pressure-stop.md',
      ],
      evidence: expect.arrayContaining([
        'live same-corpus UI proof captured: false',
        `live same-corpus UI 10x cases: 0/${String(requiredUiResponsivenessSameCorpusWorkloads.length)}`,
        `live same-corpus UI required workloads: ${requiredUiSameCorpusWorkloadList}`,
        `live same-corpus UI missing required workloads: ${requiredUiSameCorpusWorkloadList}`,
        `live same-corpus UI missing inputs: ${requiredUiSameCorpusInputList}`,
        'live same-corpus UI browser capture guard active: true',
      ]),
    })
    expect(validateBiligDominancePromptArtifactAudit(audit)).toEqual([])
  })

  it('calls out legacy operation-only same-corpus UI captures', () => {
    const input = buildFixtureInput()
    const status = buildBiligDominanceStatus({
      input: {
        ...input,
        uiResponsivenessLiveBrowserScorecard: {
          ...input.uiResponsivenessLiveBrowserScorecard,
          sameCorpusProof: legacyOperationOnlySameCorpusProof(),
        },
      },
      publicWorkbookCorpusStatus: publicWorkbookCorpusStatusFixture(),
      stopMarkerActive: false,
      stopMarkerPath: '/repo/.agent-coordination/stop.md',
      uiSameCorpusGoogleSheetsUrl: 'https://docs.google.com/spreadsheets/d/sameCorpusSheet/edit',
    })
    const audit = buildBiligDominancePromptArtifactAudit({
      scorecard: buildBiligDominanceScorecard(input),
      status,
    })
    const uiItem = audit.checklist.find((entry) => entry.id === 'ui-responsiveness')

    expect(audit.liveUiSameCorpus).toMatchObject({
      captured: true,
      scrollEventEvidenceCaseCount: 0,
      casesMissingScrollEventEvidence: ['same-corpus-wide-mixed-250k-scroll-vertical'],
    })
    expect(uiItem).toMatchObject({
      passed: false,
      liveBlockers: expect.arrayContaining([
        `same-corpus UI proof missing required workloads: ${requiredUiResponsivenessSameCorpusWorkloads
          .filter((workload) => workload !== 'scroll-vertical')
          .join(', ')}`,
        'same-corpus UI proof missing scroll-event evidence: same-corpus-wide-mixed-250k-scroll-vertical',
        'same-corpus UI proof has 0/1 required 10x cases',
      ]),
      evidence: expect.arrayContaining([
        'live same-corpus UI scroll-event evidence cases: 0/1',
        'live same-corpus UI cases missing scroll-event evidence: same-corpus-wide-mixed-250k-scroll-vertical',
      ]),
    })
    expect(validateBiligDominancePromptArtifactAudit(audit)).toEqual([])
  })

  it('keeps operator workflow incomplete while local broad CI verification is guard-paused', () => {
    const input = buildFixtureInput()
    const status = buildBiligDominanceStatus({
      input,
      publicWorkbookCorpusStatus: publicWorkbookCorpusStatusFixture(),
      stopMarkerActive: false,
      stopMarkerPath: '/repo/.agent-coordination/stop.md',
      uiSameCorpusLocalCiResourceGuardStatus: {
        activeMarkerPaths: ['.agent-coordination/20260508T092619Z-codex-memory-pressure-stop.md'],
      },
    })
    const audit = buildBiligDominancePromptArtifactAudit({
      scorecard: buildBiligDominanceScorecard(input),
      status,
    })
    const workflowItem = audit.checklist.find((entry) => entry.id === 'operator-developer-workflow')

    expect(workflowItem).toMatchObject({
      passed: false,
      liveBlockers: ['local CI resource guard active: .agent-coordination/20260508T092619Z-codex-memory-pressure-stop.md'],
      gaps: ['local CI resource guard active: .agent-coordination/20260508T092619Z-codex-memory-pressure-stop.md'],
      evidence: expect.arrayContaining([
        'local CI resource guard active: true',
        'local CI resource guard markers: .agent-coordination/20260508T092619Z-codex-memory-pressure-stop.md',
        'local CI resource guard override env: BILIG_ALLOW_LOCAL_CI_RESOURCE_GUARD',
      ]),
    })
    expect(audit.completionVerdict.unmetRequirements).toContain(
      'operator/developer workflow local CI resource guard active: .agent-coordination/20260508T092619Z-codex-memory-pressure-stop.md',
    )
    expect(validateBiligDominancePromptArtifactAudit(audit)).toEqual([])

    const workflowIndex = audit.checklist.findIndex((entry) => entry.id === 'operator-developer-workflow')
    if (!workflowItem || workflowIndex < 0) {
      throw new Error('fixture audit is missing operator workflow checklist item')
    }
    const checklist = [...audit.checklist]
    checklist[workflowIndex] = {
      ...workflowItem,
      passed: true,
      gaps: [],
      liveBlockers: [],
    }
    expect(validateBiligDominancePromptArtifactAudit({ ...audit, checklist })).toEqual(
      expect.arrayContaining([
        'operator/developer workflow checklist passed while local CI resource guard is active',
        'local CI resource guard is active but operator/developer workflow live blockers are empty',
      ]),
    )
  })

  it('rejects hidden live corpus blockers and unverifiable blanket claims', () => {
    const input = buildFixtureInput()
    const status = buildBiligDominanceStatus({
      input,
      publicWorkbookCorpusStatus: publicWorkbookCorpusStatusFixture({
        targetWorkbookCount: 10_000,
        sourceCount: 10_000,
        cachedArtifactCount: 5_628,
        scorecardCaseCount: 2_000,
        checkpointCaseCount: 4_940,
        recordedManifestArtifactCount: 4_940,
        missingManifestArtifactCount: 688,
        recordedCoversManifest: false,
        scorecardCoversManifest: false,
        targetComplete: false,
        gaps: [
          'cached artifacts below target: 5628/10000',
          'scorecard cases do not cover manifest artifacts: 2000/5628',
          'recorded verification cases below cached artifacts: 4940/5628',
        ],
      }),
      stopMarkerActive: true,
      stopMarkerPath: '/repo/.agent-coordination/stop.md',
    })
    const audit = buildBiligDominancePromptArtifactAudit({
      scorecard: buildBiligDominanceScorecard(input),
      status,
    })
    const importExportIndex = audit.checklist.findIndex((entry) => entry.id === 'import-export-compatibility')
    const uiResponsivenessIndex = audit.checklist.findIndex((entry) => entry.id === 'ui-responsiveness')
    const checklist = [...audit.checklist]
    const importExportItem = audit.checklist[importExportIndex]
    if (!importExportItem) {
      throw new Error('fixture audit is missing import/export checklist item')
    }
    const uiResponsivenessItem = audit.checklist[uiResponsivenessIndex]
    if (!uiResponsivenessItem) {
      throw new Error('fixture audit is missing UI responsiveness checklist item')
    }
    checklist[importExportIndex] = Object.assign({}, importExportItem, {
      passed: true,
      gaps: [],
      liveBlockers: [],
    })
    checklist[uiResponsivenessIndex] = Object.assign({}, uiResponsivenessItem, {
      passed: true,
      gaps: [],
      liveBlockers: [],
    })
    const invalidAudit = {
      ...audit,
      completionVerdict: {
        ...audit.completionVerdict,
        blanketTenXClaimAllowed: true,
      },
      checklist,
    }

    expect(validateBiligDominancePromptArtifactAudit(invalidAudit)).toEqual(
      expect.arrayContaining([
        'blanket 10x claim is allowed before every checklist item passed',
        'import/export checklist passed while live public workbook corpus target is incomplete',
        'live public workbook corpus target is incomplete but import/export live blockers are empty',
        'UI responsiveness checklist passed while live same-corpus UI proof is incomplete',
        'live same-corpus UI proof is incomplete but UI responsiveness live blockers are empty',
      ]),
    )
  })

  it('rejects checklist evidence that points at missing repo artifacts or package scripts', () => {
    const input = buildFixtureInput()
    const status = buildBiligDominanceStatus({
      input,
      publicWorkbookCorpusStatus: publicWorkbookCorpusStatusFixture(),
      stopMarkerActive: true,
      stopMarkerPath: '/repo/.agent-coordination/stop.md',
    })
    const audit = buildBiligDominancePromptArtifactAudit({
      scorecard: buildBiligDominanceScorecard(input),
      status,
    })
    const checklist = [...audit.checklist]
    const firstItem = checklist[0]
    if (!firstItem) {
      throw new Error('fixture audit is missing checklist items')
    }
    checklist[0] = {
      ...firstItem,
      evidenceArtifacts: ['packages/benchmarks/baselines/does-not-exist.json'],
      checkCommands: ['pnpm missing:dominance:script'],
    }

    expect(validateBiligDominancePromptArtifactAudit({ ...audit, checklist })).toEqual(
      expect.arrayContaining([
        'calculation-correctness evidence artifact does not exist: packages/benchmarks/baselines/does-not-exist.json',
        'calculation-correctness check command references missing package script: missing:dominance:script',
      ]),
    )
  })
})

function publicWorkbookCorpusStatusFixture(overrides: Partial<PublicWorkbookCorpusStatus> = {}): PublicWorkbookCorpusStatus {
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
    ...overrides,
  }
}

function legacyOperationOnlySameCorpusProof(): UiResponsivenessSameCorpusProof {
  const bilig = sameCorpusProofMeasurement('bilig', 'bilig-benchmark-state', [200, 200, 200])
  const googleSheets = sameCorpusProofMeasurement('google-sheets', 'google-sheets-xlsx-export', [100, 100, 100])
  const microsoftExcelWeb = sameCorpusProofMeasurement('microsoft-excel-web', 'microsoft-excel-web-source-xlsx', [100, 100, 100])

  return {
    captured: true,
    evidenceKind: 'same-corpus-browser-capture',
    requiredProductCount: 2,
    requiredCaseCount: 1,
    tenXMeanAndP95CaseCount: 0,
    coveredCorpusCaseIds: ['wide-mixed-250k'],
    limitations: [],
    cases: [
      {
        id: 'same-corpus-wide-mixed-250k-scroll-vertical',
        corpusCaseId: 'wide-mixed-250k',
        materializedCells: 250_000,
        workload: 'scroll-vertical',
        sampleCount: 3,
        bilig,
        googleSheets,
        microsoftExcelWeb,
        biligToGoogleSheetsMeanRatio: 2,
        biligToGoogleSheetsP95Ratio: 2,
        biligToMicrosoftExcelWebMeanRatio: 2,
        biligToMicrosoftExcelWebP95Ratio: 2,
        tenXMeanAndP95Metric: 'operationResponseMs',
        scenarioProof: sameCorpusScenarioProof(200, 100),
        tenXMeanAndP95AgainstGoogleSheets: false,
        tenXMeanAndP95AgainstMicrosoftExcelWeb: false,
        passed: false,
      },
    ],
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

function sameCorpusProofMeasurement(
  product: UiResponsivenessSameCorpusProduct,
  method: UiResponsivenessSameCorpusMeasurement['corpusVerification']['method'],
  operationResponseMsSamples: number[],
): UiResponsivenessSameCorpusMeasurement {
  return {
    product,
    source: product === 'bilig' ? 'http://127.0.0.1:5173/?benchmarkCorpus=wide-mixed-250k' : 'https://example.com/same-corpus',
    operationResponseMs: summarizeNumbers(operationResponseMsSamples),
    postOperationFrameMs: summarizeNumbers([12, 12, 12]),
    corpusVerification: {
      verified: true,
      method,
      sheetName: 'WideGrid',
      materializedCells: 250_000,
      checkedCells:
        product === 'bilig'
          ? []
          : [
              { address: 'A1', expected: 'metric-1', actual: 'metric-1' },
              { address: 'B1', expected: 'metric-2', actual: 'metric-2' },
              { address: 'F2', expected: 'note-1-5', actual: 'note-1-5' },
            ],
    },
    limitations: [],
  }
}
