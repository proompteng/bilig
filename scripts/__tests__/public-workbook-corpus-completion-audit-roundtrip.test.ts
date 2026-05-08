import { describe, expect, it } from 'vitest'

import {
  buildPublicWorkbookCorpusCompletionAudit,
  validatePublicWorkbookCorpusCompletionAudit,
  type PublicWorkbookCorpusAuditChecklistItem,
  type PublicWorkbookCorpusSecondaryFormulaCorpusStatus,
} from '../public-workbook-corpus-completion-audit.ts'
import { createEmptyPublicWorkbookManifest } from '../public-workbook-corpus-json.ts'
import type { PublicWorkbookCorpusStatus } from '../public-workbook-corpus-status.ts'
import type { PublicWorkbookArtifact, PublicWorkbookCorpusCase, PublicWorkbookManifest } from '../public-workbook-corpus-types.ts'

describe('public workbook corpus completion audit round-trip evidence', () => {
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

  it('does not count explicit round-trip skip evidence as a supported round-trip success', () => {
    const artifact = workbookArtifact('workbook-a')
    const audit = buildPublicWorkbookCorpusCompletionAudit({
      generatedAt: '2026-05-08T00:00:00.000Z',
      hyperformulaSecondaryCorpus: hyperFormulaSecondaryCorpusFixture(),
      manifest: manifestWithArtifacts([artifact], 1),
      recordedCases: [skippedRoundTripCase(artifact)],
      status: statusFixture({
        targetWorkbookCount: 1,
        sourceCount: 1,
        cachedArtifactCount: 1,
        scorecardCaseCount: 1,
        checkpointCaseCount: 0,
        recordedManifestArtifactCount: 1,
        missingManifestArtifactCount: 0,
        recordedPassedCaseCount: 1,
        recordedUnsupportedCaseCount: 0,
        scorecardCoversManifest: true,
        targetComplete: true,
        gaps: [],
      }),
      stopMarkerActive: false,
    })

    expect(requirement(audit.checklist, 'roundtrip-supported-workbooks')).toMatchObject({
      passed: false,
      gaps: expect.arrayContaining(['no supported round-trip successes recorded']),
      evidence: expect.arrayContaining(['supported round-trip passed cases: 0', 'round-trip skipped cases: 1']),
    })
    expect(validatePublicWorkbookCorpusCompletionAudit(audit)).toEqual([])
  })

  it('does not treat explicit round-trip skip evidence as a round-trip failure', () => {
    const artifactA = workbookArtifact('workbook-a')
    const artifactB = workbookArtifact('workbook-b')
    const audit = buildPublicWorkbookCorpusCompletionAudit({
      generatedAt: '2026-05-08T00:00:00.000Z',
      hyperformulaSecondaryCorpus: hyperFormulaSecondaryCorpusFixture(),
      manifest: manifestWithArtifacts([artifactA, artifactB], 2),
      recordedCases: [baseSupportedCase(artifactA), skippedRoundTripFailureCase(artifactB)],
      status: statusFixture({
        targetWorkbookCount: 2,
        sourceCount: 2,
        cachedArtifactCount: 2,
        scorecardCaseCount: 2,
        checkpointCaseCount: 0,
        recordedManifestArtifactCount: 2,
        missingManifestArtifactCount: 0,
        recordedPassedCaseCount: 2,
        recordedUnsupportedCaseCount: 0,
        scorecardCoversManifest: true,
        targetComplete: true,
        gaps: [],
      }),
      stopMarkerActive: false,
    })

    expect(requirement(audit.checklist, 'roundtrip-supported-workbooks')).toMatchObject({
      passed: true,
      gaps: [],
      evidence: expect.arrayContaining([
        'supported round-trip passed cases: 1',
        'round-trip skipped cases: 1',
        'round-trip failures among recorded cases: 0',
      ]),
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

function resourceLimitedUnsupportedCase(artifact: PublicWorkbookArtifact): PublicWorkbookCorpusCase {
  return {
    id: artifact.id,
    sourceId: artifact.sourceId,
    sourceUrl: artifact.sourceUrl,
    fileName: artifact.fileName,
    sha256: artifact.sha256,
    byteSize: artifact.byteSize,
    license: artifact.license,
    status: 'unsupported',
    passed: true,
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
    validation: {
      importPassed: false,
      formulaOraclePassed: true,
      formulaOracleComparisons: 0,
      formulaOracleMismatches: [],
      roundTripPassed: true,
      structuralSmokePassed: null,
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

function skippedRoundTripCase(artifact: PublicWorkbookArtifact): PublicWorkbookCorpusCase {
  return {
    ...baseSupportedCase(artifact),
    evidence: [
      `source=${artifact.sourceUrl}`,
      `license=${artifact.license.title}`,
      `sha256=${artifact.sha256}`,
      'Round-trip projection skipped because external workbook links are not recalculated during XLSX import.',
    ],
  }
}

function skippedRoundTripFailureCase(artifact: PublicWorkbookArtifact): PublicWorkbookCorpusCase {
  const baseCase = skippedRoundTripCase(artifact)
  return {
    ...baseCase,
    validation: {
      ...baseCase.validation,
      roundTripPassed: false,
    },
  }
}

function baseSupportedCase(artifact: PublicWorkbookArtifact): PublicWorkbookCorpusCase {
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
      cellCount: 1,
      formulaCellCount: 0,
      valueCellCount: 1,
      definedNameCount: 0,
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
          columnCount: 1,
          nonEmptyCellCount: 1,
          usedRange: { startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 },
        },
      ],
    },
    validation: {
      importPassed: true,
      formulaOraclePassed: true,
      formulaOracleComparisons: 0,
      formulaOracleMismatches: [],
      roundTripPassed: true,
      structuralSmokePassed: true,
    },
    unsupportedFeatureClassifications: [],
    evidence: [`source=${artifact.sourceUrl}`, `license=${artifact.license.title}`, `sha256=${artifact.sha256}`],
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
  readonly recordedPassedCaseCount?: number
  readonly recordedUnsupportedCaseCount: number
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
    staleRecordedVerificationCount: 0,
    recordedPassedCaseCount: input.recordedPassedCaseCount ?? 0,
    recordedUnsupportedCaseCount: input.recordedUnsupportedCaseCount,
    recordedFailedCaseCount: 0,
    recordedErrorCaseCount: 0,
    recordedCoversManifest: input.recordedManifestArtifactCount >= input.cachedArtifactCount,
    recordedAllCasesPassed: true,
    missingManifestArtifactSample: [],
    staleRecordedVerificationSample: [],
    nextMissingVerificationCommand: null,
    nextMissingVerificationPlanCommand: null,
    blockedMissingVerificationCommand: null,
    nextStaleVerificationCommand: null,
    nextStaleVerificationPlanCommand: null,
    blockedStaleVerificationCommand: null,
    scorecardCoversManifest: input.scorecardCoversManifest,
    targetComplete: input.targetComplete,
    gaps: input.gaps,
  }
}

function hyperFormulaSecondaryCorpusFixture(): PublicWorkbookCorpusSecondaryFormulaCorpusStatus {
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
