import { mkdtempSync, writeFileSync } from 'node:fs'
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
  upsertPublicWorkbookCorpusVerificationCheckpoint,
  writePublicWorkbookCorpusVerificationCheckpoint,
} from '../public-workbook-corpus-verify-checkpoint.ts'
import { listMissingPublicWorkbookArtifacts, listStalePublicWorkbookArtifacts } from '../public-workbook-corpus-missing.ts'
import { validatePublicWorkbookCorpusScorecardManifestCoverage } from '../public-workbook-corpus-scorecard.ts'
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

  it('upserts a focused artifact verification into an existing checkpoint', () => {
    const artifactA = workbookArtifact('workbook-a')
    const artifactB = workbookArtifact('workbook-b')
    const checkpointPath = join(mkdtempSync(join(tmpdir(), 'public-workbook-corpus-checkpoint-upsert-')), 'checkpoint.json')
    const staleFailure = failedCase(artifactB)
    const freshPass = passedCase(artifactB, true)

    writePublicWorkbookCorpusVerificationCheckpoint({
      path: checkpointPath,
      manifest: manifestWithArtifacts([artifactA, artifactB]),
      casesById: new Map([
        [artifactA.id, passedCase(artifactA, true)],
        [artifactB.id, staleFailure],
      ]),
      generatedAt: '2026-05-07T01:00:00.000Z',
    })

    upsertPublicWorkbookCorpusVerificationCheckpoint({
      path: checkpointPath,
      manifest: manifestWithArtifacts([artifactA, artifactB]),
      verifiedCase: freshPass,
      generatedAt: '2026-05-07T02:00:00.000Z',
    })

    const cases = readReusablePublicWorkbookCorpusCases([checkpointPath])
    expect(cases.map((entry) => [entry.id, entry.status, entry.passed])).toEqual([
      ['workbook-a', 'passed', true],
      ['workbook-b', 'passed', true],
    ])
  })

  it('normalizes legacy RSS-limit checkpoint errors as unsupported resource cases', () => {
    const artifact = workbookArtifact('workbook-a')
    const checkpointPath = join(mkdtempSync(join(tmpdir(), 'public-workbook-corpus-checkpoint-legacy-rss-')), 'checkpoint.json')

    writePublicWorkbookCorpusVerificationCheckpoint({
      path: checkpointPath,
      manifest: manifestWithArtifacts([artifact]),
      casesById: new Map([
        [
          artifact.id,
          {
            ...failedCase(artifact),
            status: 'error',
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
              ...artifactBaseEvidence(artifact),
              'Verification subprocess exceeded RSS limit: 1.52 GiB > 1.50 GiB',
              'The workbook was isolated in a subprocess so the corpus verification run could continue.',
            ],
          },
        ],
      ]),
      generatedAt: '2026-05-07T01:00:00.000Z',
    })

    expect(readReusablePublicWorkbookCorpusCases([checkpointPath])[0]).toMatchObject({
      id: artifact.id,
      status: 'unsupported',
      passed: true,
      validation: {
        importPassed: false,
        formulaOraclePassed: true,
        formulaOracleComparisons: 0,
        formulaOracleMismatches: [],
        roundTripPassed: true,
        structuralSmokePassed: null,
      },
      unsupportedFeatureClassifications: ['xlsx.publicCorpus.resourceLimit:rss>1536MiB'],
      evidence: expect.arrayContaining(['Public corpus verification RSS limit exceeded: 1.52 GiB > 1.50 GiB']),
    })
  })

  it('reads failed scorecards so resume can reuse their passed cases', () => {
    const artifactA = workbookArtifact('workbook-a')
    const artifactB = workbookArtifact('workbook-b')
    const scorecardPath = join(mkdtempSync(join(tmpdir(), 'public-workbook-corpus-failed-scorecard-')), 'scorecard.json')
    const passed = passedCase(artifactA, true)
    const failed = failedCase(artifactB)

    writeFileSync(
      scorecardPath,
      `${JSON.stringify(
        {
          schemaVersion: 1,
          suite: 'public-workbook-corpus',
          generatedAt: '2026-05-07T01:00:00.000Z',
          summary: {
            targetWorkbookCount: 10_000,
            sourceCount: 2,
            cachedWorkbookCount: 2,
            importedWorkbookCount: 2,
            passedWorkbookCount: 1,
            failedWorkbookCount: 1,
            errorWorkbookCount: 0,
            unsupportedWorkbookCount: 0,
            formulaOracleComparisonCount: 1,
            formulaOracleMatchCount: 0,
            structuralSmokeRunCount: 2,
            allCachedWorkbooksPassed: false,
            remainingToTarget: 9_998,
          },
          cases: [passed, failed],
        },
        null,
        2,
      )}\n`,
    )

    const cases = readReusablePublicWorkbookCorpusCases([scorecardPath])

    expect(cases.map((entry) => [entry.id, entry.passed])).toEqual([
      ['workbook-a', true],
      ['workbook-b', false],
    ])
  })

  it('rejects scorecards that do not cover the current manifest artifacts', async () => {
    const artifactA = workbookArtifact('workbook-a')
    const artifactB = workbookArtifact('workbook-b')
    const scorecard = await buildPublicWorkbookCorpusScorecard({
      manifest: manifestWithArtifacts([artifactA]),
      cacheDir: mkdtempSync(join(tmpdir(), 'public-workbook-corpus-coverage-')),
      generatedAt: '2026-05-07T01:00:00.000Z',
      reusableCases: [passedCase(artifactA, true)],
    })

    expect(() =>
      validatePublicWorkbookCorpusScorecardManifestCoverage({
        scorecard,
        manifest: manifestWithArtifacts([artifactA, artifactB]),
      }),
    ).toThrow('Public workbook corpus scorecard source count does not match the manifest')
  })

  it('prioritizes smaller missing and stale verification artifacts', () => {
    const large = workbookArtifact('workbook-large', { byteSize: 5_000, fileName: 'large.xlsx' })
    const matched = workbookArtifact('workbook-matched', { byteSize: 1, fileName: 'already-recorded.xlsx' })
    const smallLater = workbookArtifact('workbook-small-later', { byteSize: 10, fileName: 'z-small.xlsx' })
    const smallFirst = workbookArtifact('workbook-small-first', { byteSize: 10, fileName: 'a-small.xlsx' })
    const manifest = manifestWithArtifacts([large, matched, smallLater, smallFirst])

    expect(listMissingPublicWorkbookArtifacts({ manifest, cases: [passedCase(matched, true)] }).map((entry) => entry.id)).toEqual([
      'workbook-small-first',
      'workbook-small-later',
      'workbook-large',
    ])
    expect(
      listStalePublicWorkbookArtifacts({
        manifest,
        cases: [passedCase(large, true), passedCase(smallLater, true), passedCase(smallFirst, true)],
      }).map((entry) => entry.id),
    ).toEqual(['workbook-small-first', 'workbook-small-later', 'workbook-large'])
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

function workbookArtifact(
  id: string,
  overrides: Partial<Pick<PublicWorkbookArtifact, 'byteSize' | 'fileName'>> = {},
): PublicWorkbookArtifact {
  const hashNibble = id.endsWith('b') ? 'b' : 'a'
  return {
    id,
    sourceId: `source-${id}`,
    sourceUrl: `https://example.com/${id}.xlsx`,
    downloadUrl: `https://example.com/${id}.xlsx`,
    fileName: overrides.fileName ?? `${id}.xlsx`,
    cachePath: `files/${id}.xlsx`,
    sha256: hashNibble.repeat(64),
    byteSize: overrides.byteSize ?? 1024,
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
    evidence: artifactBaseEvidence(artifact),
  }
}

function failedCase(artifact: PublicWorkbookArtifact): PublicWorkbookCorpusCase {
  const base = passedCase(artifact, true)
  return {
    ...base,
    status: 'failed',
    passed: false,
    validation: {
      ...base.validation,
      formulaOraclePassed: false,
      formulaOracleComparisons: 1,
      formulaOracleMismatches: ['Sheet1!A1 expected 1 got error:3'],
    },
    evidence: [...base.evidence, 'Sheet1!A1 expected 1 got error:3'],
  }
}

function artifactBaseEvidence(artifact: PublicWorkbookArtifact): string[] {
  return [`source=${artifact.sourceUrl}`, `license=${artifact.license.title}`, `sha256=${artifact.sha256}`]
}
