import { describe, expect, it } from 'vitest'

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
      recordedFormulaOracleComparisonCount: 1,
    })
    expect(requirement(audit.checklist, 'download-10000-public-spreadsheets')).toMatchObject({
      passed: false,
      gaps: expect.arrayContaining(['cached artifacts below target: 2/3']),
    })
    expect(requirement(audit.checklist, 'scorecard-all-10000')).toMatchObject({
      passed: false,
      gaps: expect.arrayContaining(['scorecard cases do not cover manifest artifacts: 1/2']),
    })
    expect(validatePublicWorkbookCorpusCompletionAudit(audit)).toEqual([])
  })

  it('marks the goal achieved when public corpus evidence is complete and HyperFormula parity is folded in', () => {
    const artifactA = workbookArtifact('workbook-a')
    const artifactB = workbookArtifact('workbook-b')
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
    expect(validatePublicWorkbookCorpusCompletionAudit(audit)).toEqual([])
    expect(validatePublicWorkbookCorpusCompletionAudit(audit, { requireComplete: true })).toEqual([])
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
  return {
    id,
    sourceId: `source-${id}`,
    sourceUrl: `https://example.com/${id}.xlsx`,
    downloadUrl: `https://example.com/${id}.xlsx`,
    fileName: `${id}.xlsx`,
    cachePath: `files/${id}.xlsx`,
    sha256: hashNibble.repeat(64),
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
      tableCount: 0,
      chartCount: 0,
      pivotCount: 0,
      mergeCount: 0,
      styleRangeCount: 1,
      conditionalFormatCount: 0,
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

function statusFixture(input: {
  readonly targetWorkbookCount: number
  readonly sourceCount: number
  readonly cachedArtifactCount: number
  readonly scorecardCaseCount: number
  readonly checkpointCaseCount: number
  readonly recordedManifestArtifactCount: number
  readonly missingManifestArtifactCount: number
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
    recordedPassedCaseCount: input.recordedPassedCaseCount,
    recordedUnsupportedCaseCount: input.recordedUnsupportedCaseCount ?? 0,
    recordedFailedCaseCount: input.recordedFailedCaseCount ?? 0,
    recordedErrorCaseCount: input.recordedErrorCaseCount ?? 0,
    recordedCoversManifest: input.recordedManifestArtifactCount >= input.cachedArtifactCount,
    recordedAllCasesPassed: true,
    missingManifestArtifactSample: [],
    nextMissingVerificationCommand:
      input.missingManifestArtifactCount > 0 ? 'pnpm public-workbook-corpus:verify-missing -- --limit 1' : null,
    nextMissingVerificationPlanCommand: input.missingManifestArtifactCount > 0 ? 'pnpm public-workbook-corpus:verify-missing:plan' : null,
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
