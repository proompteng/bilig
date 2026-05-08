import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

import { describe, expect, it } from 'vitest'

import {
  buildPublicWorkbookCorpusResourceLimitPlan,
  validatePublicWorkbookCorpusResourceLimitPlan,
} from '../public-workbook-corpus-resource-limit-plan.ts'
import { publicWorkbookResourceLimitClassifierEvidence } from '../public-workbook-corpus-evidence.ts'
import { asRecord, createEmptyPublicWorkbookManifest } from '../public-workbook-corpus-json.ts'
import type { PublicWorkbookArtifact, PublicWorkbookCorpusCase, PublicWorkbookManifest } from '../public-workbook-corpus-types.ts'

describe('public workbook corpus resource-limit plan', () => {
  it('splits current resource-limit blockers from stale unsupported scorecard evidence', () => {
    const currentArtifact = workbookArtifact('workbook-current', 2_000)
    const staleArtifact = workbookArtifact('workbook-stale', 1_000)
    const plan = buildPublicWorkbookCorpusResourceLimitPlan({
      cacheDir: '/repo/.cache/public-workbook-corpus',
      displayRootDir: '/repo',
      generatedAt: '2026-05-08T10:00:00.000Z',
      manifest: manifestWithArtifacts([currentArtifact, staleArtifact]),
      manifestPath: '/repo/.cache/public-workbook-corpus/manifest.json',
      recordedCases: [resourceLimitedCase(currentArtifact), staleResourceLimitedCase(staleArtifact)],
      sampleLimit: 10,
      scorecardPath: '/repo/packages/benchmarks/baselines/public-workbook-corpus-scorecard.json',
      stopMarkerActive: true,
      stopMarkerPath: '/repo/.agent-coordination/stop.md',
      verifyCheckpointPath: '/repo/.cache/public-workbook-corpus/verification-checkpoint.json',
      verifyMaxRssMiB: 1536,
    })

    expect(plan).toMatchObject({
      schemaVersion: 1,
      mode: 'resource-limit-plan',
      stopMarker: {
        active: true,
        path: '.agent-coordination/stop.md',
      },
      currentState: {
        manifestArtifactCount: 2,
        recordedCaseCount: 2,
        resourceLimitCaseCount: 2,
        currentResourceLimitCaseCount: 1,
        staleResourceLimitCaseCount: 1,
        currentClassifications: [{ classification: 'xlsx.publicCorpus.resourceLimit:rss>1536MiB', count: 1 }],
        staleClassifications: [{ classification: 'xlsx.publicCorpus.resourceLimit:rss>1536MiB', count: 1 }],
      },
    })
    expect(plan.currentSamples).toHaveLength(1)
    expect(plan.currentSamples[0]).toMatchObject({
      id: 'workbook-current',
      byteSize: 2_000,
      classifications: ['xlsx.publicCorpus.resourceLimit:rss>1536MiB'],
      rssEvidence: ['Public corpus verification RSS limit exceeded: 1.60 GiB > 1.50 GiB'],
    })
    expect(plan.currentSamples[0]?.probeCommand).toContain('public-workbook-corpus:verify-artifact')
    expect(plan.currentSamples[0]?.probeCommand).toContain('--manifest .cache/public-workbook-corpus/manifest.json')
    expect(plan.currentSamples[0]?.probeCommand).not.toContain('/repo/')
    expect(plan.currentSamples[0]?.probeCommand).not.toContain('--allow-active-stop-marker')
    expect(plan.currentSamples[0]?.checkpointRefreshCommand).toContain('BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE=1')
    expect(plan.currentSamples[0]?.checkpointRefreshCommand).toContain('--allow-active-stop-marker')
    expect(validatePublicWorkbookCorpusResourceLimitPlan(plan)).toEqual([])
  })

  it('rejects inconsistent resource-limit plan counts and unsafe commands', () => {
    const artifact = workbookArtifact('workbook-current', 2_000)
    const plan = buildPublicWorkbookCorpusResourceLimitPlan({
      cacheDir: '/repo/.cache/public-workbook-corpus',
      displayRootDir: '/repo',
      generatedAt: '2026-05-08T10:00:00.000Z',
      manifest: manifestWithArtifacts([artifact]),
      manifestPath: '/repo/.cache/public-workbook-corpus/manifest.json',
      recordedCases: [resourceLimitedCase(artifact)],
      sampleLimit: 10,
      scorecardPath: '/repo/packages/benchmarks/baselines/public-workbook-corpus-scorecard.json',
      stopMarkerActive: true,
      stopMarkerPath: '/repo/.agent-coordination/stop.md',
      verifyCheckpointPath: '/repo/.cache/public-workbook-corpus/verification-checkpoint.json',
      verifyMaxRssMiB: 1536,
    })
    const invalidPlan = {
      ...plan,
      currentState: {
        ...plan.currentState,
        staleResourceLimitCaseCount: 1,
      },
      currentSamples: [
        {
          ...plan.currentSamples[0],
          classifications: [],
          probeCommand: `${plan.currentSamples[0].probeCommand} --update-verify-checkpoint`,
        },
      ],
    }

    expect(validatePublicWorkbookCorpusResourceLimitPlan(invalidPlan)).toEqual(
      expect.arrayContaining([
        'current and stale resource-limit counts do not add up to total resource-limit cases',
        'current sample is missing resource-limit classifications: workbook-current',
        'current probe command mutates the verification checkpoint: workbook-current',
      ]),
    )
  })

  it('exposes a package script for non-mutating resource-limit planning', () => {
    const packageJson = asRecord(JSON.parse(readFileSync(packageJsonPath(), 'utf8')))
    const scripts = asRecord(packageJson['scripts'])

    expect(scripts['public-workbook-corpus:resource-limit:plan']).toBe('bun scripts/public-workbook-corpus-resource-limit-plan.ts')
    expect(scripts['public-workbook-corpus:resource-limit:check']).toBe('bun scripts/public-workbook-corpus-resource-limit-plan.ts --check')
  })
})

function manifestWithArtifacts(artifacts: readonly PublicWorkbookArtifact[]): PublicWorkbookManifest {
  return {
    ...createEmptyPublicWorkbookManifest('2026-05-08T10:00:00.000Z', artifacts.length),
    sources: artifacts.map((artifact) => ({
      id: artifact.sourceId,
      kind: 'direct-url',
      sourceUrl: artifact.sourceUrl,
      downloadUrl: artifact.downloadUrl,
      fileName: artifact.fileName,
      discoveredAt: '2026-05-08T10:00:00.000Z',
      license: artifact.license,
    })),
    artifacts,
  }
}

function workbookArtifact(id: string, byteSize: number): PublicWorkbookArtifact {
  const sha256 = id.includes('stale') ? 'b'.repeat(64) : 'a'.repeat(64)
  return {
    id,
    sourceId: `source-${id}`,
    sourceUrl: `https://example.com/${id}.xlsx`,
    downloadUrl: `https://example.com/${id}.xlsx`,
    fileName: `${id}.xlsx`,
    cachePath: `files/${sha256}.xlsx`,
    sha256,
    byteSize,
    workbookFingerprint: `${id}-fingerprint`,
    fetchedAt: '2026-05-08T10:00:00.000Z',
    license: {
      spdxId: 'CC-BY-4.0',
      title: 'Creative Commons Attribution 4.0 International',
      evidenceUrl: 'https://creativecommons.org/licenses/by/4.0/',
    },
  }
}

function resourceLimitedCase(artifact: PublicWorkbookArtifact): PublicWorkbookCorpusCase {
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
      publicWorkbookResourceLimitClassifierEvidence,
      'Public corpus verification RSS limit exceeded: 1.60 GiB > 1.50 GiB',
    ],
  }
}

function staleResourceLimitedCase(artifact: PublicWorkbookArtifact): PublicWorkbookCorpusCase {
  return {
    ...resourceLimitedCase(artifact),
    evidence: resourceLimitedCase(artifact).evidence.filter((entry) => entry !== publicWorkbookResourceLimitClassifierEvidence),
    workbookMetadata: {
      workbookName: artifact.fileName,
      sheetNames: ['Sheet1'],
      dimensions: [{ sheetName: 'Sheet1', rowCount: 1, columnCount: 1, nonEmptyCellCount: 1 }],
    },
  }
}

function packageJsonPath(): string {
  return join(dirname(fileURLToPath(import.meta.url)), '../../package.json')
}
