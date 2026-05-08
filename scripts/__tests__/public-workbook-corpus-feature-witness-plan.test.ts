import { mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { describe, expect, it } from 'vitest'

import { createEmptyPublicWorkbookManifest } from '../public-workbook-corpus-json.ts'
import {
  buildPublicWorkbookCorpusFeatureWitnessPlan,
  readPublicWorkbookCorpusFeatureWitnessCases,
  validatePublicWorkbookCorpusFeatureWitnessPlan,
} from '../public-workbook-corpus-feature-witness-plan.ts'
import type {
  PublicWorkbookArtifact,
  PublicWorkbookCorpusCase,
  PublicWorkbookFeatureCounts,
  PublicWorkbookManifest,
  PublicWorkbookSource,
} from '../public-workbook-corpus-types.ts'

describe('public workbook corpus feature witness plan', () => {
  it('reports missing pivot witnesses with a guarded targeted discovery command', () => {
    const pivotArtifact = artifactForCase(caseWithFeatures({ pivotCount: 0 }), {
      fileName: 'visitor-visas-granted-pivot-table.xlsx',
      id: 'pivot-artifact-a',
      sourceUrl: 'https://data.gov.au/data/dataset/visitor-visas-granted-pivot-table',
    })
    const plan = buildPublicWorkbookCorpusFeatureWitnessPlan({
      artifacts: [pivotArtifact],
      cacheDir: '/repo/.cache/public-workbook-corpus',
      cases: [caseWithFeatures({ pivotCount: 0 })],
      discoveryLimit: 10_000,
      displayRootDir: '/repo',
      generatedAt: '2026-05-08T08:00:00.000Z',
      manifestPath: '/repo/.cache/public-workbook-corpus/manifest.json',
      stopMarkerActive: true,
      stopMarkerPath: '/repo/.agent-coordination/stop.md',
      verifyCheckpointPath: '/repo/.cache/public-workbook-corpus/verification-checkpoint.json',
    })
    const pivotCoverage = plan.coverage.find((entry) => entry.id === 'pivots')

    expect(plan).toMatchObject({
      schemaVersion: 1,
      mode: 'feature-witness-plan',
      missingWitnessCount: 1,
      missingWitnesses: [
        {
          id: 'pivots',
          label: 'pivots',
          discoveryQuery: 'pivot table xlsx',
          discoverCommand: null,
          blockedDiscoverCommand: expect.stringContaining("--query 'pivot table xlsx'"),
          cachedCandidateCount: 1,
          cachedCandidates: [
            {
              artifactId: 'pivot-artifact-a',
              fileName: 'visitor-visas-granted-pivot-table.xlsx',
              byteSize: 1024,
              sourceUrl: 'https://data.gov.au/data/dataset/visitor-visas-granted-pivot-table',
              verifyArtifactCommand: null,
              blockedVerifyArtifactCommand: expect.stringContaining('--artifact-id pivot-artifact-a'),
            },
          ],
        },
      ],
      recordedCaseCount: 1,
      stopMarker: {
        active: true,
        path: '.agent-coordination/stop.md',
      },
    })
    expect(pivotCoverage).toMatchObject({
      discoveryQuery: 'pivot table xlsx',
      needsWitness: true,
      totalCount: 0,
      witnessCaseCount: 0,
      commands: {
        discover: null,
      },
      blockedCommands: {
        discover: expect.stringContaining("--query 'pivot table xlsx'"),
      },
    })
    expect(pivotCoverage?.cachedCandidateCount).toBe(1)
    expect(pivotCoverage?.cachedCandidates[0]?.blockedVerifyArtifactCommand).toContain('BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE=1')
    expect(pivotCoverage?.cachedCandidates[0]?.blockedVerifyArtifactCommand).toContain('public-workbook-corpus:verify-artifact')
    expect(pivotCoverage?.cachedCandidates[0]?.blockedVerifyArtifactCommand).toContain('--update-verify-checkpoint')
    expect(pivotCoverage?.cachedCandidates[0]?.blockedVerifyArtifactCommand).toContain('--allow-active-stop-marker')
    expect(pivotCoverage?.blockedCommands.discover).toContain('BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE=1')
    expect(pivotCoverage?.blockedCommands.discover).toContain('public-workbook-corpus:discover')
    expect(pivotCoverage?.blockedCommands.discover).toContain("--query 'pivot table xlsx'")
    expect(pivotCoverage?.blockedCommands.discover).toContain('--allow-active-stop-marker')
    expect(JSON.stringify(plan)).not.toContain('/repo/')
    expect(validatePublicWorkbookCorpusFeatureWitnessPlan(plan)).toEqual([])
  })

  it('deduplicates scorecard and checkpoint evidence through the current manifest', () => {
    const dir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-feature-witness-'))
    const manifestPath = join(dir, 'manifest.json')
    const scorecardPath = join(dir, 'scorecard.json')
    const checkpointPath = join(dir, 'verification-checkpoint.json')
    const workbookCase = caseWithFeatures({ pivotCount: 1 })
    writeFileSync(manifestPath, `${JSON.stringify(manifestWithArtifacts([artifactForCase(workbookCase)]), null, 2)}\n`)
    writeFileSync(
      scorecardPath,
      `${JSON.stringify({ schemaVersion: 1, suite: 'public-workbook-corpus', cases: [workbookCase] }, null, 2)}\n`,
    )
    writeFileSync(
      checkpointPath,
      `${JSON.stringify(
        {
          schemaVersion: 1,
          suite: 'public-workbook-corpus-verification-checkpoint',
          generatedAt: '2026-05-08T10:00:00.000Z',
          cases: [workbookCase],
        },
        null,
        2,
      )}\n`,
    )

    const cases = readPublicWorkbookCorpusFeatureWitnessCases({ manifestPath, scorecardPath, verifyCheckpointPath: checkpointPath })
    const plan = buildPublicWorkbookCorpusFeatureWitnessPlan({
      cacheDir: join(dir, 'cache'),
      cases,
      discoveryLimit: 10_000,
      generatedAt: '2026-05-08T10:00:00.000Z',
      manifestPath,
      stopMarkerActive: false,
      stopMarkerPath: join(dir, 'stop.md'),
    })
    const pivotCoverage = plan.coverage.find((entry) => entry.id === 'pivots')

    expect(cases.map((entry) => entry.id)).toEqual(['artifact-a'])
    expect(plan.recordedCaseCount).toBe(1)
    expect(pivotCoverage).toMatchObject({
      totalCount: 1,
      witnessCaseCount: 1,
      needsWitness: false,
    })
  })
})

function caseWithFeatures(featureCounts: Partial<PublicWorkbookFeatureCounts>): PublicWorkbookCorpusCase {
  return {
    id: 'artifact-a',
    sourceId: 'source-a',
    sourceUrl: 'https://example.com/source-a.xlsx',
    fileName: 'source-a.xlsx',
    sha256: 'a'.repeat(64),
    byteSize: 1024,
    license: {
      spdxId: 'CC-BY-4.0',
      title: 'Creative Commons Attribution 4.0 International',
      evidenceUrl: 'https://creativecommons.org/licenses/by/4.0/',
    },
    status: 'passed',
    passed: true,
    featureCounts: {
      sheetCount: 1,
      cellCount: 9,
      formulaCellCount: 1,
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
      ...featureCounts,
    },
    workbookMetadata: {
      workbookName: 'source-a',
      sheetNames: ['Sheet1'],
      dimensions: [
        {
          sheetName: 'Sheet1',
          rowCount: 3,
          columnCount: 3,
          nonEmptyCellCount: 9,
          usedRange: { startRow: 0, startColumn: 0, endRow: 2, endColumn: 2 },
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
    unsupportedFeatureClassifications: [],
    evidence: [
      'source=https://example.com/source-a.xlsx',
      'license=Creative Commons Attribution 4.0 International',
      `sha256=${'a'.repeat(64)}`,
    ],
  }
}

function artifactForCase(
  entry: PublicWorkbookCorpusCase,
  overrides: Partial<Pick<PublicWorkbookArtifact, 'byteSize' | 'cachePath' | 'fileName' | 'id' | 'sourceUrl'>> = {},
): PublicWorkbookArtifact {
  return {
    id: overrides.id ?? entry.id,
    sourceId: entry.sourceId,
    sourceUrl: overrides.sourceUrl ?? entry.sourceUrl,
    downloadUrl: overrides.sourceUrl ?? entry.sourceUrl,
    cachePath: overrides.cachePath ?? `.cache/public-workbook-corpus/${overrides.id ?? entry.id}.xlsx`,
    fileName: overrides.fileName ?? entry.fileName,
    sha256: entry.sha256,
    byteSize: overrides.byteSize ?? entry.byteSize,
    license: entry.license,
    fetchedAt: '2026-05-08T10:00:00.000Z',
    workbookFingerprint: `${overrides.id ?? entry.id}-fingerprint`,
  }
}

function sourceForArtifact(entry: PublicWorkbookArtifact): PublicWorkbookSource {
  return {
    id: entry.sourceId,
    kind: 'direct-url',
    sourceUrl: entry.sourceUrl,
    downloadUrl: entry.downloadUrl,
    fileName: entry.fileName,
    discoveredAt: '2026-05-08T10:00:00.000Z',
    license: entry.license,
  }
}

function manifestWithArtifacts(artifacts: readonly PublicWorkbookArtifact[]): PublicWorkbookManifest {
  return {
    ...createEmptyPublicWorkbookManifest('2026-05-08T10:00:00.000Z', artifacts.length),
    sources: artifacts.map(sourceForArtifact),
    artifacts,
  }
}
