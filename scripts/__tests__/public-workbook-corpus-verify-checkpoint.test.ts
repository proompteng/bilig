import { mkdtempSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'

import {
  buildPublicWorkbookCorpusScorecard,
  createEmptyPublicWorkbookManifest,
  validatePublicWorkbookCorpusScorecard,
} from '../public-workbook-corpus.ts'
import {
  readReusablePublicWorkbookCorpusCases,
  writePublicWorkbookCorpusVerificationCheckpoint,
} from '../public-workbook-corpus-verify-checkpoint.ts'
import type { PublicWorkbookArtifact, PublicWorkbookCorpusCase, PublicWorkbookManifest } from '../public-workbook-corpus-types.ts'

describe('public workbook corpus verification checkpoints', () => {
  it('reuses matching passed cases instead of touching missing cache files', async () => {
    const manifest = manifestWithArtifacts([workbookArtifact('workbook-a'), workbookArtifact('workbook-b')])
    const reusableCases = manifest.artifacts.map((entry) => passedCase(entry, true))

    const scorecard = await buildPublicWorkbookCorpusScorecard({
      manifest,
      cacheDir: mkdtempSync(join(tmpdir(), 'public-workbook-corpus-reuse-')),
      generatedAt: '2026-05-07T01:00:00.000Z',
      reusableCases,
    })

    expect(scorecard.summary).toMatchObject({
      cachedWorkbookCount: 2,
      importedWorkbookCount: 2,
      allCachedWorkbooksPassed: true,
    })
    expect(scorecard.cases.map((entry) => entry.id)).toEqual(['workbook-a', 'workbook-b'])
    validatePublicWorkbookCorpusScorecard(scorecard)
  })

  it('reverifies reused cases that are missing newly required structural smoke evidence', async () => {
    const artifact = workbookArtifact('workbook-a')
    const manifest = manifestWithArtifacts([artifact])

    const scorecard = await buildPublicWorkbookCorpusScorecard({
      manifest,
      cacheDir: mkdtempSync(join(tmpdir(), 'public-workbook-corpus-reverify-')),
      generatedAt: '2026-05-07T01:00:00.000Z',
      structuralSmokeSampleLimit: 1,
      reusableCases: [passedCase(artifact, null)],
    })

    expect(scorecard.summary).toMatchObject({
      cachedWorkbookCount: 1,
      importedWorkbookCount: 0,
      errorWorkbookCount: 1,
      allCachedWorkbooksPassed: false,
    })
    expect(scorecard.cases[0]?.evidence).toContain('Missing cached workbook file: files/workbook-a.xlsx')
  })

  it('writes checkpoint cases in manifest order for deterministic resume', () => {
    const artifactA = workbookArtifact('workbook-a')
    const artifactB = workbookArtifact('workbook-b')
    const checkpointPath = join(mkdtempSync(join(tmpdir(), 'public-workbook-corpus-checkpoint-')), 'checkpoint.json')

    writePublicWorkbookCorpusVerificationCheckpoint({
      path: checkpointPath,
      manifest: manifestWithArtifacts([artifactA, artifactB]),
      casesById: new Map([
        [artifactB.id, passedCase(artifactB, true)],
        [artifactA.id, passedCase(artifactA, true)],
      ]),
      generatedAt: '2026-05-07T01:00:00.000Z',
    })

    expect(readReusablePublicWorkbookCorpusCases([checkpointPath]).map((entry) => entry.id)).toEqual(['workbook-a', 'workbook-b'])
  })
})

function manifestWithArtifacts(artifacts: readonly PublicWorkbookArtifact[]): PublicWorkbookManifest {
  return {
    ...createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z'),
    sources: artifacts.map((entry) => ({
      id: entry.sourceId,
      kind: 'direct-url',
      sourceUrl: entry.sourceUrl,
      downloadUrl: entry.sourceUrl,
      fileName: entry.fileName,
      discoveredAt: '2026-05-07T00:00:00.000Z',
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
    fetchedAt: '2026-05-07T00:00:00.000Z',
    license: {
      spdxId: 'CC-BY-4.0',
      title: 'Creative Commons Attribution 4.0 International',
      evidenceUrl: 'https://creativecommons.org/licenses/by/4.0/',
    },
  }
}

function passedCase(artifact: PublicWorkbookArtifact, structuralSmokePassed: boolean | null): PublicWorkbookCorpusCase {
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
      styleRangeCount: 0,
      conditionalFormatCount: 0,
      dataValidationCount: 0,
      macroPayloadCount: 0,
      warningCount: 0,
    },
    workbookMetadata: {
      workbookName: artifact.fileName,
      sheetNames: ['Sheet1'],
      dimensions: [{ sheetName: 'Sheet1', rowCount: 1, columnCount: 1, nonEmptyCellCount: 1 }],
    },
    validation: {
      importPassed: true,
      formulaOraclePassed: true,
      formulaOracleComparisons: 0,
      formulaOracleMismatches: [],
      roundTripPassed: true,
      structuralSmokePassed,
    },
    unsupportedFeatureClassifications: [],
    evidence: [`source=${artifact.sourceUrl}`, `license=${artifact.license.title}`, `sha256=${artifact.sha256}`],
  }
}
