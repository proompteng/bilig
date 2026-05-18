import { spawnSync } from 'node:child_process'
import { mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

import { createEmptyPublicWorkbookManifest } from '../public-workbook-corpus.ts'
import {
  publicWorkbookFormulaOracleCacheClassifierEvidence,
  publicWorkbookImportWarningClassifierEvidence,
  publicWorkbookPivotClassifierEvidence,
  publicWorkbookResourceLimitClassifierEvidence,
} from '../public-workbook-corpus-evidence.ts'
import { buildPublicWorkbookCorpusStatus } from '../public-workbook-corpus-status.ts'
import { writePublicWorkbookCorpusVerificationCheckpoint } from '../public-workbook-corpus-verify-checkpoint.ts'
import type {
  PublicWorkbookArtifact,
  PublicWorkbookCorpusCase,
  PublicWorkbookCorpusScorecard,
  PublicWorkbookManifest,
} from '../public-workbook-corpus-types.ts'

describe('public workbook corpus evidence refresh reasons', () => {
  it('reports every stale reason for old import-warning unsupported evidence', () => {
    const artifactA = workbookArtifact('workbook-a')
    const artifactB = workbookArtifact('workbook-b')
    const manifest = manifestWithArtifacts([artifactA, artifactB])
    const staleCase = importWarningUnsupportedCase(artifactA, {
      hasCurrentClassifierEvidence: false,
      hasUsedRangeEvidence: false,
    })
    const currentCase = importWarningUnsupportedCase(artifactB, {
      hasCurrentClassifierEvidence: true,
      hasUsedRangeEvidence: true,
    })

    const status = buildPublicWorkbookCorpusStatus({
      manifest,
      scorecard: emptyScorecard(),
      checkpointCases: [staleCase, currentCase],
    })

    expect(status.staleRecordedVerificationCount).toBe(1)
    expect(status.recordedUnsupportedCaseCount).toBe(2)
    expect(status.currentRecordedUnsupportedCaseCount).toBe(1)
    expect(status.staleRecordedUnsupportedCaseCount).toBe(1)
    expect(status.currentUnsupportedClassifications).toEqual([
      { classification: 'xlsx.import.warning:Some defined names were ignored during XLSX import.', count: 1 },
    ])
    expect(status.staleUnsupportedClassifications).toEqual([
      { classification: 'xlsx.import.warning:Some defined names were ignored during XLSX import.', count: 1 },
    ])
    expect(status.staleRecordedVerificationSample).toEqual([
      expect.objectContaining({
        id: artifactA.id,
        reason: 'missing-used-range-evidence',
        reasons: ['missing-used-range-evidence', 'missing-import-warning-classifier-evidence'],
      }),
    ])
  })

  it('marks pre-full-precision classifier evidence stale after importer precision support changes', () => {
    const artifact = workbookArtifact('workbook-a')
    const status = buildPublicWorkbookCorpusStatus({
      manifest: manifestWithArtifacts([artifact]),
      scorecard: emptyScorecard(),
      checkpointCases: [
        {
          ...importWarningUnsupportedCase(artifact, {
            hasCurrentClassifierEvidence: false,
            hasUsedRangeEvidence: true,
          }),
          evidence: [
            `source=${artifact.sourceUrl}`,
            `license=${artifact.license.title}`,
            `sha256=${artifact.sha256}`,
            'import-warning-classifier=2026-05-08',
          ],
        },
      ],
    })

    expect(status.staleRecordedVerificationCount).toBe(1)
    expect(status.staleRecordedVerificationSample[0]?.reasons).toEqual(['missing-import-warning-classifier-evidence'])
  })

  it('marks old raw pivot classifier evidence stale after external-cache pivot detection changes', () => {
    const artifact = workbookArtifact('workbook-a')
    const status = buildPublicWorkbookCorpusStatus({
      manifest: manifestWithArtifacts([artifact]),
      scorecard: emptyScorecard(),
      checkpointCases: [
        pivotUnsupportedCase(artifact, {
          hasCurrentClassifierEvidence: false,
          hasUsedRangeEvidence: true,
        }),
      ],
    })

    expect(status.staleRecordedVerificationCount).toBe(1)
    expect(status.staleRecordedVerificationSample[0]?.reasons).toEqual(['missing-pivot-classifier-evidence'])
  })

  it('marks pre-resource-limit classifier evidence stale after verifier behavior changes', () => {
    const staleArtifact = workbookArtifact('workbook-a')
    const currentArtifact = workbookArtifact('workbook-b')
    const status = buildPublicWorkbookCorpusStatus({
      manifest: manifestWithArtifacts([staleArtifact, currentArtifact]),
      scorecard: emptyScorecard(),
      checkpointCases: [
        resourceLimitUnsupportedCase(staleArtifact, { hasCurrentClassifierEvidence: false }),
        resourceLimitUnsupportedCase(currentArtifact, { hasCurrentClassifierEvidence: true }),
      ],
    })

    expect(status.staleRecordedVerificationCount).toBe(1)
    expect(status.currentRecordedUnsupportedCaseCount).toBe(1)
    expect(status.staleRecordedUnsupportedCaseCount).toBe(1)
    expect(status.staleRecordedVerificationSample[0]?.reasons).toEqual(['missing-resource-limit-classifier-evidence'])
  })

  it('marks isolated-worker-only resource-limit evidence stale after native streaming scanner changes', () => {
    const artifact = workbookArtifact('workbook-a')
    const status = buildPublicWorkbookCorpusStatus({
      manifest: manifestWithArtifacts([artifact]),
      scorecard: emptyScorecard(),
      checkpointCases: [
        {
          ...resourceLimitUnsupportedCase(artifact, { hasCurrentClassifierEvidence: false }),
          evidence: [
            `source=${artifact.sourceUrl}`,
            `license=${artifact.license.title}`,
            `sha256=${artifact.sha256}`,
            'resource-limit-classifier=2026-05-08-isolated-worker-footprint-aware',
          ],
        },
      ],
    })

    expect(status.staleRecordedVerificationCount).toBe(1)
    expect(status.staleRecordedVerificationSample[0]?.reasons).toEqual(['missing-resource-limit-classifier-evidence'])
  })

  it('marks formula-oracle cache classifier evidence stale after independent recalculation support changes', () => {
    const staleArtifact = workbookArtifact('workbook-a')
    const currentArtifact = workbookArtifact('workbook-b')
    const status = buildPublicWorkbookCorpusStatus({
      manifest: manifestWithArtifacts([staleArtifact, currentArtifact]),
      scorecard: emptyScorecard(),
      checkpointCases: [
        formulaOracleCacheUnsupportedCase(staleArtifact, { hasCurrentClassifierEvidence: false }),
        formulaOracleCacheUnsupportedCase(currentArtifact, { hasCurrentClassifierEvidence: true }),
      ],
    })

    expect(status.staleRecordedVerificationCount).toBe(1)
    expect(status.currentRecordedUnsupportedCaseCount).toBe(1)
    expect(status.staleRecordedUnsupportedCaseCount).toBe(1)
    expect(status.staleRecordedVerificationSample[0]?.reasons).toEqual(['missing-formula-oracle-cache-classifier-evidence'])
  })

  it('keeps all stale reasons in verify-stale dry-run output', () => {
    const artifactA = workbookArtifact('workbook-a')
    const artifactB = workbookArtifact('workbook-b')
    const dir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-evidence-refresh-'))
    const manifestPath = join(dir, 'manifest.json')
    const scorecardPath = join(dir, 'scorecard.json')
    const checkpointPath = join(dir, 'verification-checkpoint.json')
    const manifest = manifestWithArtifacts([artifactA, artifactB])
    writeFileSync(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`)
    writeFileSync(scorecardPath, `${JSON.stringify(emptyScorecard(), null, 2)}\n`)
    writePublicWorkbookCorpusVerificationCheckpoint({
      path: checkpointPath,
      manifest,
      casesById: new Map([
        [
          artifactA.id,
          importWarningUnsupportedCase(artifactA, {
            hasCurrentClassifierEvidence: false,
            hasUsedRangeEvidence: false,
          }),
        ],
        [
          artifactB.id,
          importWarningUnsupportedCase(artifactB, {
            hasCurrentClassifierEvidence: true,
            hasUsedRangeEvidence: true,
          }),
        ],
      ]),
      generatedAt: '2026-05-08T00:00:00.000Z',
    })

    const result = spawnSync(
      'bun',
      [
        corpusScriptPath(),
        'verify-stale',
        '--dry-run',
        '--manifest',
        manifestPath,
        '--scorecard',
        scorecardPath,
        '--verify-checkpoint',
        checkpointPath,
        '--cache-dir',
        dir,
      ],
      { encoding: 'utf8' },
    )
    const plan = JSON.parse(result.stdout) as unknown

    expect(result.status).toBe(0)
    expect(plan).toMatchObject({
      totalStaleArtifactCount: 1,
      selectedArtifactCount: 1,
      artifacts: [
        {
          id: artifactA.id,
          reason: 'missing-used-range-evidence',
          reasons: ['missing-used-range-evidence', 'missing-import-warning-classifier-evidence'],
        },
      ],
    })
  })
})

function corpusScriptPath(): string {
  return join(dirname(fileURLToPath(import.meta.url)), '../public-workbook-corpus.ts')
}

function manifestWithArtifacts(artifacts: readonly PublicWorkbookArtifact[]): PublicWorkbookManifest {
  return {
    ...createEmptyPublicWorkbookManifest('2026-05-08T00:00:00.000Z'),
    sources: artifacts.map((entry) => ({
      id: entry.sourceId,
      kind: 'direct-url',
      sourceUrl: entry.sourceUrl,
      downloadUrl: entry.sourceUrl,
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

function importWarningUnsupportedCase(
  artifact: PublicWorkbookArtifact,
  options: {
    readonly hasCurrentClassifierEvidence: boolean
    readonly hasUsedRangeEvidence: boolean
  },
): PublicWorkbookCorpusCase {
  const dimensions: PublicWorkbookCorpusCase['workbookMetadata']['dimensions'] = [
    {
      sheetName: 'Sheet1',
      rowCount: 1,
      columnCount: 1,
      nonEmptyCellCount: 1,
      ...(options.hasUsedRangeEvidence ? { usedRange: { startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 } } : {}),
    },
  ]
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
      sheetCount: 1,
      cellCount: 1,
      formulaCellCount: 0,
      valueCellCount: 1,
      definedNameCount: 1,
      tableCount: 0,
      chartCount: 0,
      pivotCount: 0,
      mergeCount: 0,
      styleRangeCount: 0,
      conditionalFormatCount: 0,
      dataValidationCount: 0,
      macroPayloadCount: 0,
      warningCount: 1,
    },
    workbookMetadata: {
      workbookName: artifact.fileName,
      sheetNames: ['Sheet1'],
      dimensions,
    },
    validation: {
      importPassed: true,
      formulaOraclePassed: true,
      formulaOracleComparisons: 0,
      formulaOracleMismatches: [],
      roundTripPassed: true,
      structuralSmokePassed: true,
    },
    unsupportedFeatureClassifications: ['xlsx.import.warning:Some defined names were ignored during XLSX import.'],
    evidence: [
      `source=${artifact.sourceUrl}`,
      `license=${artifact.license.title}`,
      `sha256=${artifact.sha256}`,
      ...(options.hasCurrentClassifierEvidence ? [publicWorkbookImportWarningClassifierEvidence] : []),
    ],
  }
}

function pivotUnsupportedCase(
  artifact: PublicWorkbookArtifact,
  options: {
    readonly hasCurrentClassifierEvidence: boolean
    readonly hasUsedRangeEvidence: boolean
  },
): PublicWorkbookCorpusCase {
  const dimensions: PublicWorkbookCorpusCase['workbookMetadata']['dimensions'] = [
    {
      sheetName: 'Sheet1',
      rowCount: 1,
      columnCount: 1,
      nonEmptyCellCount: 1,
      ...(options.hasUsedRangeEvidence ? { usedRange: { startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 } } : {}),
    },
  ]
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
      sheetCount: 1,
      cellCount: 1,
      formulaCellCount: 0,
      valueCellCount: 1,
      definedNameCount: 0,
      tableCount: 0,
      chartCount: 0,
      pivotCount: 1,
      mergeCount: 0,
      styleRangeCount: 0,
      conditionalFormatCount: 0,
      dataValidationCount: 0,
      macroPayloadCount: 0,
      warningCount: 0,
    },
    workbookMetadata: {
      workbookName: artifact.fileName,
      sheetNames: ['Sheet1'],
      dimensions,
    },
    validation: {
      importPassed: true,
      formulaOraclePassed: true,
      formulaOracleComparisons: 0,
      formulaOracleMismatches: [],
      roundTripPassed: true,
      structuralSmokePassed: true,
    },
    unsupportedFeatureClassifications: ['xlsx.pivots.rawPartNotSemanticallyImported'],
    evidence: [
      `source=${artifact.sourceUrl}`,
      `license=${artifact.license.title}`,
      `sha256=${artifact.sha256}`,
      ...(options.hasCurrentClassifierEvidence ? [publicWorkbookPivotClassifierEvidence] : []),
    ],
  }
}

function resourceLimitUnsupportedCase(
  artifact: PublicWorkbookArtifact,
  options: {
    readonly hasCurrentClassifierEvidence: boolean
  },
): PublicWorkbookCorpusCase {
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
      ...(options.hasCurrentClassifierEvidence ? [publicWorkbookResourceLimitClassifierEvidence] : []),
    ],
  }
}

function formulaOracleCacheUnsupportedCase(
  artifact: PublicWorkbookArtifact,
  options: {
    readonly hasCurrentClassifierEvidence: boolean
  },
): PublicWorkbookCorpusCase {
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
      sheetCount: 1,
      cellCount: 2,
      formulaCellCount: 1,
      valueCellCount: 2,
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
      formulaOracleComparisons: 1,
      formulaOracleMismatches: [],
      roundTripPassed: true,
      structuralSmokePassed: true,
    },
    unsupportedFeatureClassifications: ['xlsx.publicCorpus.formulaOracleCache:independentRecalcMatched'],
    evidence: [
      `source=${artifact.sourceUrl}`,
      `license=${artifact.license.title}`,
      `sha256=${artifact.sha256}`,
      ...(options.hasCurrentClassifierEvidence ? [publicWorkbookFormulaOracleCacheClassifierEvidence] : []),
    ],
  }
}

function emptyScorecard(): PublicWorkbookCorpusScorecard {
  return {
    schemaVersion: 1,
    suite: 'public-workbook-corpus',
    generatedAt: '2026-05-08T00:00:00.000Z',
    summary: {
      targetWorkbookCount: 10_000,
      sourceCount: 0,
      cachedWorkbookCount: 0,
      importedWorkbookCount: 0,
      passedWorkbookCount: 0,
      failedWorkbookCount: 0,
      errorWorkbookCount: 0,
      unsupportedWorkbookCount: 0,
      formulaOracleComparisonCount: 0,
      formulaOracleMatchCount: 0,
      structuralSmokeRunCount: 0,
      allCachedWorkbooksPassed: true,
      remainingToTarget: 10_000,
    },
    cases: [],
  }
}
