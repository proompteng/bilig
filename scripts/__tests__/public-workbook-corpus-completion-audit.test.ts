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

  it('keeps the goal active when the public corpus is complete but HyperFormula parity is still separate', () => {
    const artifactA = workbookArtifact('workbook-a')
    const artifactB = workbookArtifact('workbook-b')
    const audit = buildPublicWorkbookCorpusCompletionAudit({
      generatedAt: '2026-05-08T00:00:00.000Z',
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
      goalStatus: 'active-not-achieved',
      targetComplete: true,
      allChecklistItemsPassed: false,
    })
    expect(requirement(audit.checklist, 'hyperformula-secondary-corpus')).toMatchObject({
      passed: false,
      gaps: ['HyperFormula parity evidence remains a separate lane instead of being folded into this corpus reporting system'],
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
      dimensions: [{ sheetName: 'Sheet1', rowCount: 1, columnCount: 2, nonEmptyCellCount: 2 }],
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

function statusFixture(input: {
  readonly targetWorkbookCount: number
  readonly sourceCount: number
  readonly cachedArtifactCount: number
  readonly scorecardCaseCount: number
  readonly checkpointCaseCount: number
  readonly recordedManifestArtifactCount: number
  readonly missingManifestArtifactCount: number
  readonly recordedPassedCaseCount: number
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
    recordedUnsupportedCaseCount: 0,
    recordedFailedCaseCount: 0,
    recordedErrorCaseCount: 0,
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
