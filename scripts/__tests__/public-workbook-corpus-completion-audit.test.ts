import { describe, expect, it } from 'vitest'

import { auditPublicWorkbookCorpusCiOfflineCachedMode } from '../public-workbook-corpus-ci-offline-audit.ts'
import {
  buildPublicWorkbookCorpusCompletionAudit,
  validatePublicWorkbookCorpusCompletionAudit,
  type PublicWorkbookCorpusAuditChecklistItem,
} from '../public-workbook-corpus-completion-audit.ts'
import { createEmptyPublicWorkbookManifest } from '../public-workbook-corpus-json.ts'
import type { PublicWorkbookCorpusStatus } from '../public-workbook-corpus-status.ts'
import type { PublicWorkbookArtifact, PublicWorkbookCorpusCase, PublicWorkbookManifest } from '../public-workbook-corpus-types.ts'

describe('public workbook corpus completion audit', () => {
  it('maps incomplete live corpus evidence to explicit unmet requirements', () => {
    const artifactA = workbookArtifact('workbook-a')
    const artifactB = workbookArtifact('workbook-b')
    const audit = buildPublicWorkbookCorpusCompletionAudit({
      generatedAt: '2026-05-08T00:00:00.000Z',
      manifest: manifestWithArtifacts([artifactA, artifactB], 3),
      recordedCases: [passedCase(artifactA, 1)],
      status: statusFixture({
        targetWorkbookCount: 3,
        sourceCount: 3,
        cachedArtifactCount: 2,
        scorecardCaseCount: 1,
        checkpointCaseCount: 1,
        recordedManifestArtifactCount: 1,
        missingManifestArtifactCount: 1,
        staleRecordedVerificationCount: 1,
        recordedPassedCaseCount: 1,
        scorecardCoversManifest: false,
        targetComplete: false,
        gaps: ['cached artifacts below target: 2/3', 'recorded verification cases below cached artifacts: 1/2'],
      }),
      stopMarkerActive: true,
    })

    expect(audit.completionVerdict).toMatchObject({
      goalStatus: 'active-not-achieved',
      allChecklistItemsPassed: false,
      targetComplete: false,
      stopMarkerActive: true,
      nextCorpusRunRequiresExplicitResume: true,
    })
    expect(audit.currentState).toMatchObject({
      targetWorkbookCount: 3,
      cachedArtifactCount: 2,
      recordedManifestArtifactCount: 1,
      missingCachedArtifactCount: 1,
      missingVerificationCount: 1,
      staleRecordedVerificationCount: 1,
      missingFeatureWitnessCount: 0,
      missingFeatureWitnesses: [],
      recordedFormulaOracleComparisonCount: 1,
    })
    expect(requirement(audit.checklist, 'download-10000-public-spreadsheets')).toMatchObject({
      passed: false,
      gaps: expect.arrayContaining(['cached artifacts below target: 2/3']),
    })
    expect(requirement(audit.checklist, 'financial-accounting-workpapers-5000')).toMatchObject({
      passed: false,
      gaps: expect.arrayContaining([
        'financial/accounting cached artifacts below target: 0/3',
        'financial/accounting recorded verification cases below target: 0/3',
      ]),
    })
    expect(requirement(audit.checklist, 'scorecard-all-10000')).toMatchObject({
      passed: false,
      gaps: expect.arrayContaining(['scorecard cases do not cover manifest artifacts: 1/2']),
      evidence: expect.arrayContaining(['stale recorded verification cases: 1']),
    })
    expect(validatePublicWorkbookCorpusCompletionAudit(audit)).toEqual([])
  })

  it('marks the goal achieved when public corpus evidence is complete and HyperFormula parity is folded in', () => {
    const artifactA = financialWorkbookArtifact('workbook-a')
    const artifactB = financialWorkbookArtifact('workbook-b')
    const audit = buildPublicWorkbookCorpusCompletionAudit({
      generatedAt: '2026-05-08T00:00:00.000Z',
      hyperformulaSecondaryCorpus: hyperFormulaSecondaryCorpusFixture(),
      manifest: manifestWithArtifacts([artifactA, artifactB], 2),
      recordedCases: [passedCase(artifactA, 1), passedCase(artifactB, 2)],
      status: statusFixture({
        targetWorkbookCount: 2,
        sourceCount: 2,
        cachedArtifactCount: 2,
        scorecardCaseCount: 2,
        checkpointCaseCount: 0,
        recordedManifestArtifactCount: 2,
        missingManifestArtifactCount: 0,
        recordedPassedCaseCount: 2,
        scorecardCoversManifest: true,
        targetComplete: true,
        gaps: [],
      }),
      stopMarkerActive: false,
    })

    expect(audit.completionVerdict).toMatchObject({
      goalStatus: 'achieved',
      targetComplete: true,
      allChecklistItemsPassed: true,
    })
    expect(requirement(audit.checklist, 'hyperformula-secondary-corpus')).toMatchObject({
      passed: true,
      gaps: [],
    })
    expect(requirement(audit.checklist, 'financial-accounting-workpapers-5000')).toMatchObject({
      passed: true,
      gaps: [],
      evidence: expect.arrayContaining([
        'financial/accounting workbook target: 2',
        'financial/accounting cached artifacts: 2/2',
        'financial/accounting recorded verification cases: 2/2',
      ]),
    })
    expect(validatePublicWorkbookCorpusCompletionAudit(audit)).toEqual([])
    expect(validatePublicWorkbookCorpusCompletionAudit(audit, { requireComplete: true })).toEqual([])
  })

  it('fails the download target when cached spreadsheets do not prioritize xlsx files', () => {
    const artifactA = workbookArtifactWithExtension('workbook-a', 'xls')
    const artifactB = workbookArtifactWithExtension('workbook-b', 'xlsm')
    const audit = buildPublicWorkbookCorpusCompletionAudit({
      generatedAt: '2026-05-08T00:00:00.000Z',
      hyperformulaSecondaryCorpus: hyperFormulaSecondaryCorpusFixture(),
      manifest: manifestWithArtifacts([artifactA, artifactB], 2),
      recordedCases: [passedCase(artifactA, 1), passedCase(artifactB, 1)],
      status: statusFixture({
        targetWorkbookCount: 2,
        sourceCount: 2,
        cachedArtifactCount: 2,
        scorecardCaseCount: 2,
        checkpointCaseCount: 0,
        recordedManifestArtifactCount: 2,
        missingManifestArtifactCount: 0,
        recordedPassedCaseCount: 2,
        scorecardCoversManifest: true,
        targetComplete: true,
        gaps: [],
      }),
      stopMarkerActive: false,
    })

    expect(requirement(audit.checklist, 'download-10000-public-spreadsheets')).toMatchObject({
      passed: false,
      gaps: expect.arrayContaining(['no .xlsx artifacts cached', '.xlsx artifacts are not the majority: 0/2']),
      evidence: expect.arrayContaining(['.xlsx cached artifacts: 0', 'non-.xlsx cached artifacts: 2']),
    })
    expect(validatePublicWorkbookCorpusCompletionAudit(audit)).toEqual([])
  })

  it('fails the financial workbook lane when a recorded financial case is non-passing', () => {
    const artifactA = financialWorkbookArtifact('workbook-a')
    const artifactB = financialWorkbookArtifact('workbook-b')
    const audit = buildPublicWorkbookCorpusCompletionAudit({
      generatedAt: '2026-05-08T00:00:00.000Z',
      hyperformulaSecondaryCorpus: hyperFormulaSecondaryCorpusFixture(),
      manifest: manifestWithArtifacts([artifactA, artifactB], 2),
      recordedCases: [passedCase(artifactA, 1), caseWithCacheIntegrityFailure(artifactB)],
      status: statusFixture({
        targetWorkbookCount: 2,
        sourceCount: 2,
        cachedArtifactCount: 2,
        scorecardCaseCount: 2,
        checkpointCaseCount: 0,
        recordedManifestArtifactCount: 2,
        missingManifestArtifactCount: 0,
        recordedPassedCaseCount: 1,
        recordedFailedCaseCount: 1,
        scorecardCoversManifest: true,
        targetComplete: true,
        gaps: [],
      }),
      stopMarkerActive: false,
    })

    expect(requirement(audit.checklist, 'financial-accounting-workpapers-5000')).toMatchObject({
      passed: false,
      gaps: expect.arrayContaining(['financial/accounting non-passing recorded cases: 1']),
      evidence: expect.arrayContaining([
        'financial/accounting cached artifacts: 2/2',
        'financial/accounting recorded verification cases: 2/2',
        'financial/accounting non-passing recorded cases: 1',
      ]),
    })
    expect(validatePublicWorkbookCorpusCompletionAudit(audit)).toEqual([])
  })

  it('treats resource-limited unsupported cases as evidenced metadata exceptions', () => {
    const artifactA = workbookArtifact('workbook-a')
    const artifactB = workbookArtifact('workbook-b')
    const audit = buildPublicWorkbookCorpusCompletionAudit({
      generatedAt: '2026-05-08T00:00:00.000Z',
      hyperformulaSecondaryCorpus: hyperFormulaSecondaryCorpusFixture(),
      manifest: manifestWithArtifacts([artifactA, artifactB], 2),
      recordedCases: [passedCase(artifactA, 1), resourceLimitedUnsupportedCase(artifactB)],
      status: statusFixture({
        targetWorkbookCount: 2,
        sourceCount: 2,
        cachedArtifactCount: 2,
        scorecardCaseCount: 2,
        checkpointCaseCount: 0,
        recordedManifestArtifactCount: 2,
        missingManifestArtifactCount: 0,
        recordedPassedCaseCount: 1,
        recordedUnsupportedCaseCount: 1,
        scorecardCoversManifest: true,
        targetComplete: true,
        gaps: [],
      }),
      stopMarkerActive: false,
    })

    expect(requirement(audit.checklist, 'source-license-hash-metadata-manifest')).toMatchObject({
      passed: true,
      gaps: [],
      evidence: expect.arrayContaining(['resource-limited unsupported cases with metadata unavailable: 1']),
    })
    expect(requirement(audit.checklist, 'unsupported-features-evidence')).toMatchObject({
      passed: true,
    })
    expect(validatePublicWorkbookCorpusCompletionAudit(audit)).toEqual([])
  })

  it('accepts offline cached corpus gates only when the scripts are safe and wired into CI', () => {
    const result = auditPublicWorkbookCorpusCiOfflineCachedMode({
      scripts: offlineCiPackageScripts(),
      ciSource: ciSourceFor([
        'public-workbook-corpus:check:offline',
        'public-workbook-corpus:resume-plan:check',
        'public-workbook-corpus:resource-limit:check',
        'public-workbook-corpus:feature-witness:check',
        'public-workbook-corpus:discover-financial:check',
        'public-workbook-corpus:resume-financial:check',
        'public-workbook-corpus:completion-audit:check',
        'test:correctness:corpus',
      ]),
    })

    expect(result).toMatchObject({ passed: true, gaps: [] })
    expect(result.evidence).toEqual(
      expect.arrayContaining([
        'package script public-workbook-corpus:check:offline: bun scripts/public-workbook-corpus.ts check --skip-manifest-check',
        'CI invokes package script: public-workbook-corpus:check:offline',
        'CI invokes package script: public-workbook-corpus:resource-limit:check',
        'CI invokes package script: public-workbook-corpus:feature-witness:check',
        'CI invokes package script: public-workbook-corpus:discover-financial:check',
        'CI invokes package script: public-workbook-corpus:resume-financial:check',
        'CI invokes package script: test:correctness:corpus',
      ]),
    )
  })

  it('accepts safe direct CI gates for offline cached corpus checks', () => {
    const result = auditPublicWorkbookCorpusCiOfflineCachedMode({
      scripts: offlineCiPackageScripts(),
      ciSource: [
        "bunScript('public workbook corpus offline scorecard check', 'scripts/public-workbook-corpus.ts', 'check', '--skip-manifest-check')",
        "bunScript('public workbook corpus resume plan check', 'scripts/public-workbook-corpus-resume-plan.ts', '--check')",
        "bunScript('public workbook corpus resource-limit plan check', 'scripts/public-workbook-corpus-resource-limit-plan.ts', '--check')",
        "bunScript('public workbook corpus feature-witness plan check', 'scripts/public-workbook-corpus-feature-witness-plan.ts', '--check')",
        "bunScript('financial public workbook corpus plan check', 'scripts/public-workbook-corpus-financial-plan.ts', '--check')",
        "directPackageScript('financial public workbook corpus resume check', 'public-workbook-corpus:resume-financial:check')",
        "bunScript('public workbook corpus completion audit check', 'scripts/public-workbook-corpus-completion-audit.ts', '--check')",
        "pnpm('correctness public workbook corpus', 'test:correctness:corpus')",
      ].join('\n'),
    })

    expect(result).toMatchObject({ passed: true, gaps: [] })
    expect(result.evidence).toEqual(
      expect.arrayContaining([
        'CI invokes equivalent direct gate: public-workbook-corpus:check:offline',
        'CI invokes equivalent direct gate: public-workbook-corpus:resume-plan:check',
        'CI invokes equivalent direct gate: public-workbook-corpus:resource-limit:check',
        'CI invokes equivalent direct gate: public-workbook-corpus:feature-witness:check',
        'CI invokes equivalent direct gate: public-workbook-corpus:discover-financial:check',
        'CI invokes package script: public-workbook-corpus:resume-financial:check',
        'CI invokes equivalent direct gate: public-workbook-corpus:completion-audit:check',
        'CI invokes package script: test:correctness:corpus',
      ]),
    )
  })

  it('rejects named offline corpus gates when they mutate the corpus or are not wired into CI', () => {
    const result = auditPublicWorkbookCorpusCiOfflineCachedMode({
      scripts: offlineCiPackageScripts([
        ['public-workbook-corpus:check:offline', 'bun scripts/public-workbook-corpus.ts fetch'],
        ['test:correctness:corpus', 'tsx scripts/run-vitest.ts --run scripts/__tests__/public-workbook-corpus.test.ts'],
      ]),
      ciSource: ciSourceFor(['public-workbook-corpus:resume-plan:check', 'test:correctness:corpus']),
    })

    expect(result.passed).toBe(false)
    expect(result.gaps).toEqual(
      expect.arrayContaining([
        'package script public-workbook-corpus:check:offline missing required tokens: check, --skip-manifest-check',
        'package script public-workbook-corpus:check:offline uses CI-unsafe corpus tokens: fetch',
        expect.stringContaining('package script test:correctness:corpus missing required coverage files:'),
        'CI does not invoke package script or equivalent direct gate: public-workbook-corpus:check:offline',
        'CI does not invoke package script or equivalent direct gate: public-workbook-corpus:resource-limit:check',
        'CI does not invoke package script or equivalent direct gate: public-workbook-corpus:feature-witness:check',
        'CI does not invoke package script or equivalent direct gate: public-workbook-corpus:discover-financial:check',
        'CI does not invoke package script or equivalent direct gate: public-workbook-corpus:resume-financial:check',
        'CI does not invoke package script or equivalent direct gate: public-workbook-corpus:completion-audit:check',
      ]),
    )
  })

  it('fails completion when manifest cache paths are not hash-addressed', () => {
    const artifact = { ...workbookArtifact('workbook-a'), cachePath: 'files/not-hash-addressed.xlsx' }
    const audit = buildPublicWorkbookCorpusCompletionAudit({
      generatedAt: '2026-05-08T00:00:00.000Z',
      hyperformulaSecondaryCorpus: hyperFormulaSecondaryCorpusFixture(),
      manifest: manifestWithArtifacts([artifact], 1),
      recordedCases: [passedCase(artifact, 1)],
      status: statusFixture({
        targetWorkbookCount: 1,
        sourceCount: 1,
        cachedArtifactCount: 1,
        scorecardCaseCount: 1,
        checkpointCaseCount: 0,
        recordedManifestArtifactCount: 1,
        missingManifestArtifactCount: 0,
        recordedPassedCaseCount: 1,
        scorecardCoversManifest: true,
        targetComplete: true,
        gaps: [],
      }),
      stopMarkerActive: false,
    })

    expect(requirement(audit.checklist, 'source-license-hash-metadata-manifest')).toMatchObject({
      passed: false,
      gaps: expect.arrayContaining(['artifact cache paths not hash-addressed: 1']),
      evidence: expect.arrayContaining(['hash-addressed artifact cache paths: 0/1']),
    })
    expect(validatePublicWorkbookCorpusCompletionAudit(audit)).toEqual([])
  })

  it('fails completion when manifest artifacts are duplicated by hash or structure fingerprint', () => {
    const artifactA = workbookArtifact('workbook-a')
    const duplicateHashArtifact = workbookArtifact('workbook-c')
    const duplicateFingerprintArtifact = {
      ...workbookArtifact('workbook-b'),
      workbookFingerprint: artifactA.workbookFingerprint,
    }
    const audit = buildPublicWorkbookCorpusCompletionAudit({
      generatedAt: '2026-05-08T00:00:00.000Z',
      hyperformulaSecondaryCorpus: hyperFormulaSecondaryCorpusFixture(),
      manifest: manifestWithArtifacts([artifactA, duplicateHashArtifact, duplicateFingerprintArtifact], 3),
      recordedCases: [passedCase(artifactA, 1), passedCase(duplicateHashArtifact, 1), passedCase(duplicateFingerprintArtifact, 1)],
      status: statusFixture({
        targetWorkbookCount: 3,
        sourceCount: 3,
        cachedArtifactCount: 3,
        scorecardCaseCount: 3,
        checkpointCaseCount: 0,
        recordedManifestArtifactCount: 3,
        missingManifestArtifactCount: 0,
        recordedPassedCaseCount: 3,
        scorecardCoversManifest: true,
        targetComplete: true,
        gaps: [],
      }),
      stopMarkerActive: false,
    })

    expect(requirement(audit.checklist, 'hash-and-structure-dedupe')).toMatchObject({
      passed: false,
      gaps: expect.arrayContaining(['duplicate artifact hashes: 1', 'duplicate workbook structure fingerprints: 1']),
      evidence: expect.arrayContaining(['unique hashes: 2/3', 'unique workbook fingerprints: 2/3']),
    })
    expect(validatePublicWorkbookCorpusCompletionAudit(audit)).toEqual([])
  })

  it('fails completion when recorded cases omit source/license/hash evidence', () => {
    const artifact = workbookArtifact('workbook-a')
    const audit = buildPublicWorkbookCorpusCompletionAudit({
      generatedAt: '2026-05-08T00:00:00.000Z',
      hyperformulaSecondaryCorpus: hyperFormulaSecondaryCorpusFixture(),
      manifest: manifestWithArtifacts([artifact], 1),
      recordedCases: [caseWithoutProvenanceEvidence(artifact)],
      status: statusFixture({
        targetWorkbookCount: 1,
        sourceCount: 1,
        cachedArtifactCount: 1,
        scorecardCaseCount: 1,
        checkpointCaseCount: 0,
        recordedManifestArtifactCount: 1,
        missingManifestArtifactCount: 0,
        recordedPassedCaseCount: 1,
        scorecardCoversManifest: true,
        targetComplete: true,
        gaps: [],
      }),
      stopMarkerActive: false,
    })

    expect(requirement(audit.checklist, 'source-license-hash-metadata-manifest')).toMatchObject({
      passed: false,
      gaps: expect.arrayContaining(['recorded cases missing source/license/hash evidence: 1']),
      evidence: expect.arrayContaining(['recorded source/license/hash evidence cases: 0/1']),
    })
    expect(validatePublicWorkbookCorpusCompletionAudit(audit)).toEqual([])
  })

  it('fails completion when verification recorded a cache integrity failure', () => {
    const artifact = workbookArtifact('workbook-a')
    const audit = buildPublicWorkbookCorpusCompletionAudit({
      generatedAt: '2026-05-08T00:00:00.000Z',
      hyperformulaSecondaryCorpus: hyperFormulaSecondaryCorpusFixture(),
      manifest: manifestWithArtifacts([artifact], 1),
      recordedCases: [caseWithCacheIntegrityFailure(artifact)],
      status: statusFixture({
        targetWorkbookCount: 1,
        sourceCount: 1,
        cachedArtifactCount: 1,
        scorecardCaseCount: 1,
        checkpointCaseCount: 0,
        recordedManifestArtifactCount: 1,
        missingManifestArtifactCount: 0,
        recordedPassedCaseCount: 0,
        recordedFailedCaseCount: 1,
        scorecardCoversManifest: true,
        targetComplete: true,
        gaps: [],
      }),
      stopMarkerActive: false,
    })

    expect(requirement(audit.checklist, 'source-license-hash-metadata-manifest')).toMatchObject({
      passed: false,
      gaps: expect.arrayContaining(['recorded cache integrity failures: 1']),
      evidence: expect.arrayContaining(['recorded cache integrity failures: 1']),
    })
    expect(validatePublicWorkbookCorpusCompletionAudit(audit)).toEqual([])
  })

  it('fails require-complete mode until every mapped objective requirement is satisfied', () => {
    const artifact = workbookArtifact('workbook-a')
    const audit = buildPublicWorkbookCorpusCompletionAudit({
      generatedAt: '2026-05-08T00:00:00.000Z',
      manifest: manifestWithArtifacts([artifact], 1),
      recordedCases: [passedCase(artifact, 1)],
      status: statusFixture({
        targetWorkbookCount: 1,
        sourceCount: 1,
        cachedArtifactCount: 1,
        scorecardCaseCount: 1,
        checkpointCaseCount: 0,
        recordedManifestArtifactCount: 1,
        missingManifestArtifactCount: 0,
        recordedPassedCaseCount: 1,
        scorecardCoversManifest: true,
        targetComplete: true,
        gaps: [],
      }),
      stopMarkerActive: false,
    })

    expect(validatePublicWorkbookCorpusCompletionAudit(audit, { requireComplete: true })).toEqual(
      expect.arrayContaining([expect.stringContaining('public workbook corpus goal is not achieved')]),
    )
  })

  it('fails completion when scorecard evidence still contains failed or error cases', () => {
    const artifactA = workbookArtifact('workbook-a')
    const artifactB = workbookArtifact('workbook-b')
    const audit = buildPublicWorkbookCorpusCompletionAudit({
      generatedAt: '2026-05-08T00:00:00.000Z',
      hyperformulaSecondaryCorpus: hyperFormulaSecondaryCorpusFixture(),
      manifest: manifestWithArtifacts([artifactA, artifactB], 2),
      recordedCases: [passedCase(artifactA, 1), passedCase(artifactB, 1)],
      status: statusFixture({
        targetWorkbookCount: 2,
        sourceCount: 2,
        cachedArtifactCount: 2,
        scorecardCaseCount: 2,
        checkpointCaseCount: 0,
        recordedManifestArtifactCount: 2,
        missingManifestArtifactCount: 0,
        recordedPassedCaseCount: 0,
        recordedFailedCaseCount: 1,
        recordedErrorCaseCount: 1,
        scorecardCoversManifest: true,
        targetComplete: true,
        gaps: [],
      }),
      stopMarkerActive: false,
    })

    expect(requirement(audit.checklist, 'scorecard-all-10000')).toMatchObject({
      passed: false,
      gaps: expect.arrayContaining(['failed scorecard cases: 1', 'error scorecard cases: 1']),
      evidence: expect.arrayContaining(['scorecard failed cases: 1', 'scorecard error cases: 1']),
    })
    expect(validatePublicWorkbookCorpusCompletionAudit(audit)).toEqual([])
  })

  it('fails completion when no supported workbook recorded a successful round-trip', () => {
    const artifact = workbookArtifact('workbook-a')
    const audit = buildPublicWorkbookCorpusCompletionAudit({
      generatedAt: '2026-05-08T00:00:00.000Z',
      hyperformulaSecondaryCorpus: hyperFormulaSecondaryCorpusFixture(),
      manifest: manifestWithArtifacts([artifact], 1),
      recordedCases: [resourceLimitedUnsupportedCase(artifact)],
      status: statusFixture({
        targetWorkbookCount: 1,
        sourceCount: 1,
        cachedArtifactCount: 1,
        scorecardCaseCount: 1,
        checkpointCaseCount: 0,
        recordedManifestArtifactCount: 1,
        missingManifestArtifactCount: 0,
        recordedUnsupportedCaseCount: 1,
        scorecardCoversManifest: true,
        targetComplete: true,
        gaps: [],
      }),
      stopMarkerActive: false,
    })

    expect(requirement(audit.checklist, 'roundtrip-supported-workbooks')).toMatchObject({
      passed: false,
      gaps: expect.arrayContaining(['no supported round-trip successes recorded']),
      evidence: expect.arrayContaining(['supported round-trip passed cases: 0']),
    })
    expect(validatePublicWorkbookCorpusCompletionAudit(audit)).toEqual([])
  })

  it('fails completion when recorded feature evidence is internally inconsistent', () => {
    const artifact = workbookArtifact('workbook-a')
    const audit = buildPublicWorkbookCorpusCompletionAudit({
      generatedAt: '2026-05-08T00:00:00.000Z',
      hyperformulaSecondaryCorpus: hyperFormulaSecondaryCorpusFixture(),
      manifest: manifestWithArtifacts([artifact], 1),
      recordedCases: [caseWithInvalidFeatureEvidence(artifact)],
      status: statusFixture({
        targetWorkbookCount: 1,
        sourceCount: 1,
        cachedArtifactCount: 1,
        scorecardCaseCount: 1,
        checkpointCaseCount: 0,
        recordedManifestArtifactCount: 1,
        missingManifestArtifactCount: 0,
        recordedPassedCaseCount: 1,
        scorecardCoversManifest: true,
        targetComplete: true,
        gaps: [],
      }),
      stopMarkerActive: false,
    })

    expect(requirement(audit.checklist, 'validate-workbook-features')).toMatchObject({
      passed: false,
      gaps: expect.arrayContaining(['recorded cases with incomplete feature validation evidence: 1']),
    })
    expect(validatePublicWorkbookCorpusCompletionAudit(audit)).toEqual([])
  })

  it('fails completion when recorded sheet dimensions omit explicit used-range evidence', () => {
    const artifact = workbookArtifact('workbook-a')
    const audit = buildPublicWorkbookCorpusCompletionAudit({
      generatedAt: '2026-05-08T00:00:00.000Z',
      hyperformulaSecondaryCorpus: hyperFormulaSecondaryCorpusFixture(),
      manifest: manifestWithArtifacts([artifact], 1),
      recordedCases: [caseWithoutUsedRangeEvidence(artifact)],
      status: statusFixture({
        targetWorkbookCount: 1,
        sourceCount: 1,
        cachedArtifactCount: 1,
        scorecardCaseCount: 1,
        checkpointCaseCount: 0,
        recordedManifestArtifactCount: 1,
        missingManifestArtifactCount: 0,
        recordedPassedCaseCount: 1,
        scorecardCoversManifest: true,
        targetComplete: true,
        gaps: [],
      }),
      stopMarkerActive: false,
    })

    expect(requirement(audit.checklist, 'validate-workbook-features')).toMatchObject({
      passed: false,
      gaps: expect.arrayContaining(['recorded cases missing explicit used-range evidence: 1']),
    })
    expect(validatePublicWorkbookCorpusCompletionAudit(audit)).toEqual([])
  })

  it('fails completion when a named feature family has no recorded witness', () => {
    const artifact = workbookArtifact('workbook-a')
    const audit = buildPublicWorkbookCorpusCompletionAudit({
      generatedAt: '2026-05-08T00:00:00.000Z',
      hyperformulaSecondaryCorpus: hyperFormulaSecondaryCorpusFixture(),
      manifest: manifestWithArtifacts([artifact], 1),
      recordedCases: [caseWithoutChartWitness(artifact)],
      status: statusFixture({
        targetWorkbookCount: 1,
        sourceCount: 1,
        cachedArtifactCount: 1,
        scorecardCaseCount: 1,
        checkpointCaseCount: 0,
        recordedManifestArtifactCount: 1,
        missingManifestArtifactCount: 0,
        recordedPassedCaseCount: 1,
        scorecardCoversManifest: true,
        targetComplete: true,
        gaps: [],
      }),
      stopMarkerActive: false,
    })

    expect(requirement(audit.checklist, 'validate-workbook-features')).toMatchObject({
      passed: false,
      gaps: expect.arrayContaining(['no recorded charts witness in corpus evidence']),
      evidence: expect.arrayContaining(['charts witnessed cases: 0; total recorded count: 0']),
    })
    expect(audit.currentState).toMatchObject({
      missingFeatureWitnessCount: 1,
      missingFeatureWitnesses: ['charts'],
    })
    expect(validatePublicWorkbookCorpusCompletionAudit(audit)).toEqual([])
  })

  it('separates stale unsupported scorecard records from current unsupported evidence', () => {
    const staleUnsupportedArtifact = workbookArtifact('workbook-a')
    const currentUnsupportedArtifact = workbookArtifact('workbook-b')
    const audit = buildPublicWorkbookCorpusCompletionAudit({
      generatedAt: '2026-05-08T00:00:00.000Z',
      manifest: manifestWithArtifacts([staleUnsupportedArtifact, currentUnsupportedArtifact], 2),
      recordedCases: [staleUnsupportedCase(staleUnsupportedArtifact), resourceLimitedUnsupportedCase(currentUnsupportedArtifact)],
      status: statusFixture({
        targetWorkbookCount: 2,
        sourceCount: 2,
        cachedArtifactCount: 2,
        scorecardCaseCount: 2,
        checkpointCaseCount: 0,
        recordedManifestArtifactCount: 2,
        missingManifestArtifactCount: 0,
        staleRecordedVerificationCount: 1,
        recordedPassedCaseCount: 0,
        recordedUnsupportedCaseCount: 2,
        scorecardCoversManifest: true,
        targetComplete: false,
        gaps: ['recorded verification cases need evidence refresh: 1'],
      }),
      stopMarkerActive: true,
    })

    expect(audit.currentState).toMatchObject({
      recordedUnsupportedCaseCount: 2,
      staleRecordedUnsupportedCaseCount: 1,
      currentRecordedUnsupportedCaseCount: 1,
      currentUnsupportedClassifications: [{ classification: 'xlsx.publicCorpus.resourceLimit:rss>1536MiB', count: 1 }],
      staleUnsupportedClassifications: [{ classification: 'xlsx.publicCorpus.resourceLimit:rss>1536MiB', count: 1 }],
    })
    expect(requirement(audit.checklist, 'scorecard-all-10000')).toMatchObject({
      evidence: expect.arrayContaining([
        'scorecard unsupported cases: 2',
        'stale unsupported cases: 1',
        'current unsupported cases: 1',
        'current unsupported classifications: xlsx.publicCorpus.resourceLimit:rss>1536MiB=1',
        'stale unsupported classifications: xlsx.publicCorpus.resourceLimit:rss>1536MiB=1',
      ]),
      gaps: expect.arrayContaining(['recorded verification cases need evidence refresh: 1']),
    })
    expect(validatePublicWorkbookCorpusCompletionAudit(audit)).toEqual([])
  })
})

function requirement(
  checklist: readonly PublicWorkbookCorpusAuditChecklistItem[],
  id: PublicWorkbookCorpusAuditChecklistItem['id'],
): PublicWorkbookCorpusAuditChecklistItem {
  const item = checklist.find((entry) => entry.id === id)
  if (!item) {
    throw new Error(`Missing checklist item: ${id}`)
  }
  return item
}

function offlineCiPackageScripts(overrides: readonly (readonly [string, string])[] = []): ReadonlyMap<string, string> {
  const scripts = new Map<string, string>([
    ['public-workbook-corpus:check:offline', 'bun scripts/public-workbook-corpus.ts check --skip-manifest-check'],
    ['public-workbook-corpus:resume-plan:check', 'bun scripts/public-workbook-corpus-resume-plan.ts --check'],
    ['public-workbook-corpus:resource-limit:check', 'bun scripts/public-workbook-corpus-resource-limit-plan.ts --check'],
    ['public-workbook-corpus:feature-witness:check', 'bun scripts/public-workbook-corpus-feature-witness-plan.ts --check'],
    ['public-workbook-corpus:discover-financial:check', 'bun scripts/public-workbook-corpus-financial-plan.ts --check'],
    [
      'public-workbook-corpus:resume-financial:check',
      'bun scripts/public-workbook-corpus-resume-plan.ts --check --manifest .cache/public-workbook-corpus-financial/manifest.json --cache-dir .cache/public-workbook-corpus-financial --scorecard .cache/public-workbook-corpus-financial/scorecard.json --verify-checkpoint .cache/public-workbook-corpus-financial/verification-checkpoint.json --fetch-limit 5000 --fetch-batch-size 6',
    ],
    ['public-workbook-corpus:completion-audit:check', 'bun scripts/public-workbook-corpus-completion-audit.ts --check'],
    [
      'test:correctness:corpus',
      [
        'tsx scripts/run-vitest.ts --run',
        'scripts/__tests__/public-workbook-corpus.test.ts',
        'scripts/__tests__/public-workbook-corpus-cli.test.ts',
        'scripts/__tests__/public-workbook-corpus-completion-audit.test.ts',
        'scripts/__tests__/public-workbook-corpus-feature-witness-plan.test.ts',
        'scripts/__tests__/public-workbook-corpus-financial-plan.test.ts',
        'scripts/__tests__/public-workbook-corpus-links.test.ts',
        'scripts/__tests__/public-workbook-corpus-resource-limit-plan.test.ts',
        'scripts/__tests__/public-workbook-corpus-verify-checkpoint.test.ts',
        'scripts/__tests__/public-workbook-corpus-workbook.test.ts',
        'packages/excel-import/src/__tests__/excel-import.test.ts',
        'packages/excel-import/src/__tests__/xlsx-export-large-simple.test.ts',
      ].join(' '),
    ],
  ])
  for (const [name, command] of overrides) {
    scripts.set(name, command)
  }
  return scripts
}

function ciSourceFor(scriptNames: readonly string[]): string {
  return scriptNames.map((scriptName) => `pnpm('fixture ${scriptName}', '${scriptName}')`).join('\n')
}

function manifestWithArtifacts(artifacts: readonly PublicWorkbookArtifact[], targetWorkbookCount: number): PublicWorkbookManifest {
  return {
    ...createEmptyPublicWorkbookManifest('2026-05-08T00:00:00.000Z', targetWorkbookCount),
    sources: artifacts.map((entry) => ({
      id: entry.sourceId,
      kind: 'direct-url',
      sourceUrl: entry.sourceUrl,
      downloadUrl: entry.downloadUrl,
      fileName: entry.fileName,
      discoveredAt: '2026-05-08T00:00:00.000Z',
      license: entry.license,
    })),
    artifacts,
  }
}

function workbookArtifact(id: string): PublicWorkbookArtifact {
  const hashNibble = id.endsWith('b') ? 'b' : 'a'
  const sha256 = hashNibble.repeat(64)
  return {
    id,
    sourceId: `source-${id}`,
    sourceUrl: `https://example.com/${id}.xlsx`,
    downloadUrl: `https://example.com/${id}.xlsx`,
    fileName: `${id}.xlsx`,
    cachePath: `files/${sha256}.xlsx`,
    sha256,
    byteSize: 1024,
    workbookFingerprint: `${id}-fingerprint`,
    fetchedAt: '2026-05-08T00:00:00.000Z',
    license: {
      spdxId: 'CC-BY-4.0',
      title: 'Creative Commons Attribution 4.0 International',
      evidenceUrl: 'https://creativecommons.org/licenses/by/4.0/',
    },
  }
}

function financialWorkbookArtifact(id: string): PublicWorkbookArtifact {
  return {
    ...workbookArtifact(id),
    topicEvidence: ['financial:fileName'],
  }
}

function workbookArtifactWithExtension(id: string, extension: string): PublicWorkbookArtifact {
  const artifact = workbookArtifact(id)
  return {
    ...artifact,
    sourceUrl: `https://example.com/${id}.${extension}`,
    downloadUrl: `https://example.com/${id}.${extension}`,
    fileName: `${id}.${extension}`,
    cachePath: `files/${artifact.sha256}.${extension}`,
  }
}

function passedCase(artifact: PublicWorkbookArtifact, formulaOracleComparisons: number): PublicWorkbookCorpusCase {
  return {
    id: artifact.id,
    sourceId: artifact.sourceId,
    sourceUrl: artifact.sourceUrl,
    fileName: artifact.fileName,
    sha256: artifact.sha256,
    byteSize: artifact.byteSize,
    license: artifact.license,
    status: 'passed',
    passed: true,
    featureCounts: {
      sheetCount: 1,
      cellCount: 2,
      formulaCellCount: formulaOracleComparisons > 0 ? 1 : 0,
      valueCellCount: 1,
      definedNameCount: 1,
      tableCount: 1,
      chartCount: 1,
      pivotCount: 1,
      mergeCount: 1,
      styleRangeCount: 1,
      conditionalFormatCount: 1,
      dataValidationCount: 0,
      macroPayloadCount: 0,
      warningCount: 0,
    },
    workbookMetadata: {
      workbookName: artifact.fileName,
      sheetNames: ['Sheet1'],
      dimensions: [
        {
          sheetName: 'Sheet1',
          rowCount: 1,
          columnCount: 2,
          nonEmptyCellCount: 2,
          usedRange: { startRow: 0, startColumn: 0, endRow: 0, endColumn: 1 },
        },
      ],
    },
    validation: {
      importPassed: true,
      formulaOraclePassed: true,
      formulaOracleComparisons,
      formulaOracleMismatches: [],
      roundTripPassed: true,
      structuralSmokePassed: true,
    },
    unsupportedFeatureClassifications: [],
    evidence: [`source=${artifact.sourceUrl}`, `license=${artifact.license.title}`, `sha256=${artifact.sha256}`],
  }
}

function resourceLimitedUnsupportedCase(artifact: PublicWorkbookArtifact): PublicWorkbookCorpusCase {
  return {
    ...passedCase(artifact, 0),
    status: 'unsupported',
    validation: {
      importPassed: false,
      formulaOraclePassed: true,
      formulaOracleComparisons: 0,
      formulaOracleMismatches: [],
      roundTripPassed: true,
      structuralSmokePassed: null,
    },
    featureCounts: {
      sheetCount: 0,
      cellCount: 0,
      formulaCellCount: 0,
      valueCellCount: 0,
      definedNameCount: 0,
      tableCount: 0,
      chartCount: 0,
      pivotCount: 0,
      mergeCount: 0,
      styleRangeCount: 0,
      conditionalFormatCount: 0,
      dataValidationCount: 0,
      macroPayloadCount: 0,
      warningCount: 0,
    },
    workbookMetadata: {
      workbookName: artifact.fileName,
      sheetNames: [],
      dimensions: [],
    },
    unsupportedFeatureClassifications: ['xlsx.publicCorpus.resourceLimit:rss>1536MiB'],
    evidence: [
      `source=${artifact.sourceUrl}`,
      `license=${artifact.license.title}`,
      `sha256=${artifact.sha256}`,
      'Public corpus verification RSS limit exceeded: 1.53 GiB > 1.50 GiB',
      'The workbook was isolated in a subprocess so the corpus verification run could continue.',
    ],
  }
}

function staleUnsupportedCase(artifact: PublicWorkbookArtifact): PublicWorkbookCorpusCase {
  return {
    ...resourceLimitedUnsupportedCase(artifact),
    workbookMetadata: {
      workbookName: artifact.fileName,
      sheetNames: ['Sheet1'],
      dimensions: [{ sheetName: 'Sheet1', rowCount: 1, columnCount: 1, nonEmptyCellCount: 1 }],
    },
  }
}

function caseWithInvalidFeatureEvidence(artifact: PublicWorkbookArtifact): PublicWorkbookCorpusCase {
  const base = passedCase(artifact, 1)
  return {
    ...base,
    featureCounts: {
      ...base.featureCounts,
      cellCount: 3,
      dataValidationCount: -1,
    },
    workbookMetadata: {
      workbookName: artifact.fileName,
      sheetNames: ['Sheet1'],
      dimensions: [
        {
          sheetName: 'Sheet1',
          rowCount: 1,
          columnCount: 2,
          nonEmptyCellCount: 2,
          usedRange: { startRow: 0, startColumn: 0, endRow: 0, endColumn: 1 },
        },
      ],
    },
  }
}

function caseWithoutUsedRangeEvidence(artifact: PublicWorkbookArtifact): PublicWorkbookCorpusCase {
  return {
    ...passedCase(artifact, 1),
    workbookMetadata: {
      workbookName: artifact.fileName,
      sheetNames: ['Sheet1'],
      dimensions: [{ sheetName: 'Sheet1', rowCount: 1, columnCount: 2, nonEmptyCellCount: 2 }],
    },
  }
}

function caseWithoutChartWitness(artifact: PublicWorkbookArtifact): PublicWorkbookCorpusCase {
  const base = passedCase(artifact, 1)
  return {
    ...base,
    featureCounts: {
      ...base.featureCounts,
      chartCount: 0,
    },
  }
}

function caseWithoutProvenanceEvidence(artifact: PublicWorkbookArtifact): PublicWorkbookCorpusCase {
  return {
    ...passedCase(artifact, 1),
    evidence: [],
  }
}

function caseWithCacheIntegrityFailure(artifact: PublicWorkbookArtifact): PublicWorkbookCorpusCase {
  return {
    ...passedCase(artifact, 1),
    status: 'failed',
    passed: false,
    validation: {
      importPassed: false,
      formulaOraclePassed: false,
      formulaOracleComparisons: 0,
      formulaOracleMismatches: [],
      roundTripPassed: false,
      structuralSmokePassed: null,
    },
    evidence: [
      `source=${artifact.sourceUrl}`,
      `license=${artifact.license.title}`,
      `sha256=${artifact.sha256}`,
      `Cached workbook hash mismatch: expected ${artifact.sha256}, received ${'0'.repeat(64)}`,
    ],
  }
}

function statusFixture(input: {
  readonly targetWorkbookCount: number
  readonly sourceCount: number
  readonly cachedArtifactCount: number
  readonly scorecardCaseCount: number
  readonly checkpointCaseCount: number
  readonly recordedManifestArtifactCount: number
  readonly missingManifestArtifactCount: number
  readonly staleRecordedVerificationCount?: number
  readonly recordedPassedCaseCount: number
  readonly recordedUnsupportedCaseCount?: number
  readonly recordedFailedCaseCount?: number
  readonly recordedErrorCaseCount?: number
  readonly scorecardCoversManifest: boolean
  readonly targetComplete: boolean
  readonly gaps: readonly string[]
}): PublicWorkbookCorpusStatus {
  return {
    targetWorkbookCount: input.targetWorkbookCount,
    sourceCount: input.sourceCount,
    cachedArtifactCount: input.cachedArtifactCount,
    scorecardCaseCount: input.scorecardCaseCount,
    checkpointCaseCount: input.checkpointCaseCount,
    recordedManifestArtifactCount: input.recordedManifestArtifactCount,
    missingManifestArtifactCount: input.missingManifestArtifactCount,
    staleRecordedVerificationCount: input.staleRecordedVerificationCount ?? 0,
    recordedPassedCaseCount: input.recordedPassedCaseCount,
    recordedUnsupportedCaseCount: input.recordedUnsupportedCaseCount ?? 0,
    recordedFailedCaseCount: input.recordedFailedCaseCount ?? 0,
    recordedErrorCaseCount: input.recordedErrorCaseCount ?? 0,
    recordedCoversManifest: input.recordedManifestArtifactCount >= input.cachedArtifactCount,
    recordedAllCasesPassed: true,
    missingManifestArtifactSample: [],
    staleRecordedVerificationSample: [],
    nextMissingVerificationCommand:
      input.missingManifestArtifactCount > 0 ? 'pnpm public-workbook-corpus:verify-missing -- --limit 1' : null,
    nextMissingVerificationPlanCommand: input.missingManifestArtifactCount > 0 ? 'pnpm public-workbook-corpus:verify-missing:plan' : null,
    nextStaleVerificationCommand:
      (input.staleRecordedVerificationCount ?? 0) > 0 ? 'pnpm public-workbook-corpus:verify-stale -- --limit 1' : null,
    nextStaleVerificationPlanCommand:
      (input.staleRecordedVerificationCount ?? 0) > 0 ? 'pnpm public-workbook-corpus:verify-stale:plan' : null,
    scorecardCoversManifest: input.scorecardCoversManifest,
    targetComplete: input.targetComplete,
    gaps: input.gaps,
  }
}

function hyperFormulaSecondaryCorpusFixture() {
  return {
    artifact: 'packages/benchmarks/baselines/workpaper-vs-hyperformula.json',
    artifactPresent: true,
    suite: 'workpaper-vs-hyperformula',
    resultCount: 2,
    comparableCount: 2,
    workpaperWins: 2,
    hyperformulaWins: 0,
    comparableVerificationEquivalentCount: 2,
    allComparableVerificationEquivalent: true,
    parseError: null,
  }
}
