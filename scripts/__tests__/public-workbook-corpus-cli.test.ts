import { spawnSync } from 'node:child_process'
import { mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

import { buildPublicWorkbookCorpusScorecard, createEmptyPublicWorkbookManifest } from '../public-workbook-corpus.ts'
import type { PublicWorkbookArtifact, PublicWorkbookCorpusCase, PublicWorkbookManifest } from '../public-workbook-corpus-types.ts'

describe('public workbook corpus CLI resource guards', () => {
  it('refuses in-process verification unless explicitly enabled for debugging', () => {
    const env = { ...process.env }
    delete env.BILIG_ALLOW_IN_PROCESS_PUBLIC_CORPUS_VERIFY

    const result = spawnSync('bun', [corpusScriptPath(), 'verify', '--in-process'], {
      encoding: 'utf8',
      env,
    })

    expect(result.status).not.toBe(0)
    expect(result.stderr).toContain('--in-process is disabled for public corpus CLI runs')
  })

  it('refuses in-process fingerprinting unless explicitly enabled for debugging', () => {
    const env = { ...process.env }
    delete env.BILIG_ALLOW_IN_PROCESS_PUBLIC_CORPUS_FINGERPRINT

    const result = spawnSync('bun', [corpusScriptPath(), 'fetch', '--in-process-fingerprint'], {
      encoding: 'utf8',
      env,
    })

    expect(result.status).not.toBe(0)
    expect(result.stderr).toContain('--in-process-fingerprint is disabled for public corpus CLI runs')
  })

  it('refuses parallel verification unless explicitly enabled for a sized host', () => {
    const env = { ...process.env }
    delete env.BILIG_ALLOW_PARALLEL_PUBLIC_CORPUS_VERIFY

    const result = spawnSync('bun', [corpusScriptPath(), 'verify', '--verify-concurrency', '4'], {
      encoding: 'utf8',
      env,
    })

    expect(result.status).not.toBe(0)
    expect(result.stderr).toContain('--verify-concurrency greater than 1 is disabled for public corpus CLI runs')
  })

  it('fails check when the local manifest has cached artifacts missing from the scorecard', async () => {
    const artifactA = workbookArtifact('workbook-a')
    const artifactB = workbookArtifact('workbook-b')
    const dir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-cli-check-'))
    const manifestPath = join(dir, 'manifest.json')
    const scorecardPath = join(dir, 'scorecard.json')
    const scorecard = await buildPublicWorkbookCorpusScorecard({
      manifest: manifestWithArtifacts([artifactA]),
      cacheDir: dir,
      generatedAt: '2026-05-07T01:00:00.000Z',
      reusableCases: [passedCase(artifactA)],
    })
    writeFileSync(manifestPath, `${JSON.stringify(manifestWithArtifacts([artifactA, artifactB]), null, 2)}\n`)
    writeFileSync(scorecardPath, `${JSON.stringify(scorecard, null, 2)}\n`)

    const result = spawnSync('bun', [corpusScriptPath(), 'check', '--manifest', manifestPath, '--scorecard', scorecardPath], {
      encoding: 'utf8',
    })

    expect(result.status).not.toBe(0)
    expect(result.stderr).toContain('Public workbook corpus scorecard source count does not match the manifest')
  })
})

function corpusScriptPath(): string {
  return join(dirname(fileURLToPath(import.meta.url)), '../public-workbook-corpus.ts')
}

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

function passedCase(artifact: PublicWorkbookArtifact): PublicWorkbookCorpusCase {
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
      structuralSmokePassed: true,
    },
    unsupportedFeatureClassifications: [],
    evidence: [`source=${artifact.sourceUrl}`, `license=${artifact.license.title}`, `sha256=${artifact.sha256}`],
  }
}
