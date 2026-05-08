import { describe, expect, it } from 'vitest'

import { buildBiligDominancePromptArtifactAudit, validateBiligDominancePromptArtifactAudit } from '../bilig-dominance-audit.ts'
import { buildBiligDominanceStatus } from '../bilig-dominance-status.ts'
import { buildBiligDominanceScorecard } from '../gen-bilig-dominance-scorecard.ts'
import type { PublicWorkbookCorpusStatus } from '../public-workbook-corpus-status.ts'
import { buildFixtureInput } from './bilig-dominance-scorecard.fixture.ts'

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
        'same-corpus UI proof missing required workloads: visible-scroll-response',
        'same-corpus UI proof missing inputs: googleSheetsUrlForUploadedSameCorpusWorkbook',
        'same-corpus UI browser capture paused by local resource guard: .agent-coordination/20260508T092619Z-codex-memory-pressure-stop.md',
      ],
      gaps: [
        'live UI browser evidence is not a same-corpus 10x proof against incumbents',
        'same-corpus UI browser capture has not been recorded',
        'same-corpus UI proof missing required workloads: visible-scroll-response',
        'same-corpus UI proof missing inputs: googleSheetsUrlForUploadedSameCorpusWorkbook',
        'same-corpus UI browser capture paused by local resource guard: .agent-coordination/20260508T092619Z-codex-memory-pressure-stop.md',
      ],
      evidence: expect.arrayContaining([
        'live same-corpus UI proof captured: false',
        'live same-corpus UI 10x cases: 0/0',
        'live same-corpus UI required workloads: visible-scroll-response',
        'live same-corpus UI missing required workloads: visible-scroll-response',
        'live same-corpus UI missing inputs: googleSheetsUrlForUploadedSameCorpusWorkbook',
        'live same-corpus UI browser capture guard active: true',
      ]),
    })
    expect(validateBiligDominancePromptArtifactAudit(audit)).toEqual([])
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
    ...overrides,
  }
}
