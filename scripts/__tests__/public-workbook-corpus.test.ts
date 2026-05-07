import { EventEmitter } from 'node:events'
import { mkdtempSync, mkdirSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it, vi } from 'vitest'
import * as XLSX from 'xlsx'

import type { WorkbookSnapshot } from '../../packages/protocol/src/types.js'
import {
  buildPublicWorkbookCorpusScorecard,
  createEmptyPublicWorkbookManifest,
  fetchPublicWorkbookArtifacts,
  discoverCkanWorkbookSources,
  sha256Hex,
  validatePublicWorkbookCorpusScorecard,
  validatePublicWorkbookManifest,
  type PublicWorkbookManifest,
  type PublicWorkbookSource,
} from '../public-workbook-corpus.ts'
import { roundTripSemanticsDigest } from '../public-workbook-corpus-roundtrip.ts'

const spawnMock = vi.hoisted(() => vi.fn())

vi.mock('node:child_process', () => ({
  spawn: spawnMock,
}))

describe('public workbook corpus', () => {
  afterEach(() => {
    vi.useRealTimers()
    vi.clearAllMocks()
    vi.unstubAllGlobals()
  })

  it('validates source license evidence before a workbook can enter the corpus', () => {
    const manifest = createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z')
    const withMissingLicense: PublicWorkbookManifest = {
      ...manifest,
      sources: [
        {
          id: 'source-missing-license',
          kind: 'direct-url',
          sourceUrl: 'https://example.com/workbook.xlsx',
          downloadUrl: 'https://example.com/workbook.xlsx',
          fileName: 'workbook.xlsx',
          discoveredAt: '2026-05-07T00:00:00.000Z',
          license: {
            spdxId: null,
            title: '',
            evidenceUrl: null,
          },
        },
      ],
    }

    expect(() => validatePublicWorkbookManifest(withMissingLicense)).toThrow(
      'Public workbook source source-missing-license is missing usable license evidence',
    )
  })

  it('supports focused corpus slices with a custom target workbook count', async () => {
    const manifest = createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z', 5_000)
    const cacheDir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-target-'))

    validatePublicWorkbookManifest(manifest)
    const scorecard = await buildPublicWorkbookCorpusScorecard({
      manifest,
      cacheDir,
      generatedAt: '2026-05-07T01:00:00.000Z',
    })

    expect(scorecard.summary).toMatchObject({
      targetWorkbookCount: 5_000,
      cachedWorkbookCount: 0,
      allCachedWorkbooksPassed: true,
      remainingToTarget: 5_000,
    })
    validatePublicWorkbookCorpusScorecard(scorecard)
  })

  it('builds an offline scorecard from cached spreadsheet artifacts', async () => {
    const cacheDir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-test-'))
    mkdirSync(join(cacheDir, 'files'), { recursive: true })
    const workbookBytes = buildWorkbookBytes()
    const sha256 = await sha256Hex(workbookBytes)
    writeFileSync(join(cacheDir, 'files', `${sha256}.xlsx`), workbookBytes)

    const manifest: PublicWorkbookManifest = {
      ...createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z'),
      sources: [
        {
          id: 'source-1',
          kind: 'direct-url',
          sourceUrl: 'https://example.com/public-budget.xlsx',
          downloadUrl: 'https://example.com/public-budget.xlsx',
          fileName: 'public-budget.xlsx',
          discoveredAt: '2026-05-07T00:00:00.000Z',
          license: {
            spdxId: 'CC-BY-4.0',
            title: 'Creative Commons Attribution 4.0 International',
            evidenceUrl: 'https://creativecommons.org/licenses/by/4.0/',
          },
        },
      ],
      artifacts: [
        {
          id: `workbook-${sha256.slice(0, 16)}`,
          sourceId: 'source-1',
          sourceUrl: 'https://example.com/public-budget.xlsx',
          downloadUrl: 'https://example.com/public-budget.xlsx',
          fileName: 'public-budget.xlsx',
          cachePath: `files/${sha256}.xlsx`,
          sha256,
          byteSize: workbookBytes.byteLength,
          workbookFingerprint: 'test-fingerprint',
          fetchedAt: '2026-05-07T00:00:00.000Z',
          license: {
            spdxId: 'CC-BY-4.0',
            title: 'Creative Commons Attribution 4.0 International',
            evidenceUrl: 'https://creativecommons.org/licenses/by/4.0/',
          },
        },
      ],
    }

    const scorecard = await buildPublicWorkbookCorpusScorecard({
      manifest,
      cacheDir,
      generatedAt: '2026-05-07T01:00:00.000Z',
      structuralSmokeSampleLimit: 1,
    })

    expect(scorecard).toMatchObject({
      schemaVersion: 1,
      suite: 'public-workbook-corpus',
      generatedAt: '2026-05-07T01:00:00.000Z',
      summary: {
        targetWorkbookCount: 10_000,
        sourceCount: 1,
        cachedWorkbookCount: 1,
        importedWorkbookCount: 1,
        formulaOracleComparisonCount: 1,
        formulaOracleMatchCount: 1,
        structuralSmokeRunCount: 1,
        allCachedWorkbooksPassed: true,
        remainingToTarget: 9_999,
      },
    })
    expect(scorecard.cases).toHaveLength(1)
    expect(scorecard.cases[0]).toMatchObject({
      id: `workbook-${sha256.slice(0, 16)}`,
      status: 'passed',
      passed: true,
      featureCounts: {
        sheetCount: 2,
        formulaCellCount: 1,
        definedNameCount: 1,
        mergeCount: 1,
      },
      validation: {
        importPassed: true,
        formulaOraclePassed: true,
        roundTripPassed: true,
        structuralSmokePassed: true,
      },
    })
    expect(scorecard.cases[0]?.unsupportedFeatureClassifications).toEqual([])
    validatePublicWorkbookCorpusScorecard(scorecard)
  })

  it('classifies oversized workbook verification as an explicit unsupported resource case', async () => {
    const cacheDir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-resource-limit-'))
    mkdirSync(join(cacheDir, 'files'), { recursive: true })
    const workbookBytes = buildWorkbookBytes()
    const sha256 = await sha256Hex(workbookBytes)
    writeFileSync(join(cacheDir, 'files', `${sha256}.xlsx`), workbookBytes)
    const license = {
      spdxId: 'CC-BY-4.0',
      title: 'Creative Commons Attribution 4.0 International',
      evidenceUrl: 'https://creativecommons.org/licenses/by/4.0/',
    }
    const manifest: PublicWorkbookManifest = {
      ...createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z'),
      sources: [
        {
          id: 'source-resource-limit',
          kind: 'direct-url',
          sourceUrl: 'https://example.com/resource-limit.xlsx',
          downloadUrl: 'https://example.com/resource-limit.xlsx',
          fileName: 'resource-limit.xlsx',
          discoveredAt: '2026-05-07T00:00:00.000Z',
          license,
        },
      ],
      artifacts: [
        {
          id: `workbook-${sha256.slice(0, 16)}`,
          sourceId: 'source-resource-limit',
          sourceUrl: 'https://example.com/resource-limit.xlsx',
          downloadUrl: 'https://example.com/resource-limit.xlsx',
          fileName: 'resource-limit.xlsx',
          cachePath: `files/${sha256}.xlsx`,
          sha256,
          byteSize: workbookBytes.byteLength,
          workbookFingerprint: 'resource-limit-fingerprint',
          fetchedAt: '2026-05-07T00:00:00.000Z',
          license,
        },
      ],
    }

    const scorecard = await buildPublicWorkbookCorpusScorecard({
      manifest,
      cacheDir,
      generatedAt: '2026-05-07T01:00:00.000Z',
      verifyMaxCellCount: 2,
    })

    expect(scorecard.summary.allCachedWorkbooksPassed).toBe(true)
    expect(scorecard.summary.importedWorkbookCount).toBe(0)
    expect(scorecard.cases[0]).toMatchObject({
      status: 'unsupported',
      passed: true,
      validation: { importPassed: false },
      unsupportedFeatureClassifications: ['xlsx.publicCorpus.resourceLimit:cellCount>2'],
    })
    expect(scorecard.cases[0]?.evidence).toEqual(expect.arrayContaining(['Public corpus verification cell-count limit exceeded: 10 > 2']))
  })

  it('passes corpus round-trip validation for sheet names with trailing spaces', async () => {
    const cacheDir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-trailing-space-'))
    mkdirSync(join(cacheDir, 'files'), { recursive: true })
    const workbookBytes = buildTrailingSpaceWorkbookBytes()
    const sha256 = await sha256Hex(workbookBytes)
    writeFileSync(join(cacheDir, 'files', `${sha256}.xlsx`), workbookBytes)

    const license = {
      spdxId: 'CC-BY-4.0',
      title: 'Creative Commons Attribution 4.0 International',
      evidenceUrl: 'https://creativecommons.org/licenses/by/4.0/',
    }
    const manifest: PublicWorkbookManifest = {
      ...createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z'),
      sources: [
        {
          id: 'source-trailing-space',
          kind: 'direct-url',
          sourceUrl: 'https://example.com/trailing-space.xlsx',
          downloadUrl: 'https://example.com/trailing-space.xlsx',
          fileName: 'trailing-space.xlsx',
          discoveredAt: '2026-05-07T00:00:00.000Z',
          license,
        },
      ],
      artifacts: [
        {
          id: `workbook-${sha256.slice(0, 16)}`,
          sourceId: 'source-trailing-space',
          sourceUrl: 'https://example.com/trailing-space.xlsx',
          downloadUrl: 'https://example.com/trailing-space.xlsx',
          fileName: 'trailing-space.xlsx',
          cachePath: `files/${sha256}.xlsx`,
          sha256,
          byteSize: workbookBytes.byteLength,
          workbookFingerprint: 'trailing-space-fingerprint',
          fetchedAt: '2026-05-07T00:00:00.000Z',
          license,
        },
      ],
    }

    const scorecard = await buildPublicWorkbookCorpusScorecard({
      manifest,
      cacheDir,
      generatedAt: '2026-05-07T01:00:00.000Z',
      structuralSmokeSampleLimit: 1,
    })

    expect(scorecard.summary.allCachedWorkbooksPassed).toBe(true)
    expect(scorecard.cases[0]).toMatchObject({
      status: 'passed',
      validation: {
        roundTripPassed: true,
        structuralSmokePassed: true,
      },
    })
  })

  it('compares populated-cell styles when round trips shrink blank style ranges', () => {
    const broadStyleRange = buildInteriorStyledSnapshot('A1', 'C3')
    const populatedOnlyStyleRange = buildInteriorStyledSnapshot('B2', 'B2')
    const differentPopulatedStyle = buildInteriorStyledSnapshot('B2', 'B2', '#00ccff')

    expect(roundTripSemanticsDigest(broadStyleRange)).toBe(roundTripSemanticsDigest(populatedOnlyStyleRange))
    expect(roundTripSemanticsDigest(broadStyleRange)).not.toBe(roundTripSemanticsDigest(differentPopulatedStyle))
  })

  it('runs structural smoke against a mutable sheet when the first sheet is protected', async () => {
    const cacheDir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-protected-first-'))
    mkdirSync(join(cacheDir, 'files'), { recursive: true })
    const workbookBytes = buildProtectedFirstWorkbookBytes()
    const sha256 = await sha256Hex(workbookBytes)
    writeFileSync(join(cacheDir, 'files', `${sha256}.xlsx`), workbookBytes)

    const license = {
      spdxId: 'CC-BY-4.0',
      title: 'Creative Commons Attribution 4.0 International',
      evidenceUrl: 'https://creativecommons.org/licenses/by/4.0/',
    }
    const manifest: PublicWorkbookManifest = {
      ...createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z'),
      sources: [
        {
          id: 'source-protected-first',
          kind: 'direct-url',
          sourceUrl: 'https://example.com/protected-first.xlsx',
          downloadUrl: 'https://example.com/protected-first.xlsx',
          fileName: 'protected-first.xlsx',
          discoveredAt: '2026-05-07T00:00:00.000Z',
          license,
        },
      ],
      artifacts: [
        {
          id: `workbook-${sha256.slice(0, 16)}`,
          sourceId: 'source-protected-first',
          sourceUrl: 'https://example.com/protected-first.xlsx',
          downloadUrl: 'https://example.com/protected-first.xlsx',
          fileName: 'protected-first.xlsx',
          cachePath: `files/${sha256}.xlsx`,
          sha256,
          byteSize: workbookBytes.byteLength,
          workbookFingerprint: 'protected-first-fingerprint',
          fetchedAt: '2026-05-07T00:00:00.000Z',
          license,
        },
      ],
    }

    const scorecard = await buildPublicWorkbookCorpusScorecard({
      manifest,
      cacheDir,
      generatedAt: '2026-05-07T01:00:00.000Z',
      structuralSmokeSampleLimit: 1,
    })

    expect(scorecard.summary.allCachedWorkbooksPassed).toBe(true)
    expect(scorecard.cases[0]).toMatchObject({
      status: 'passed',
      validation: {
        roundTripPassed: true,
        structuralSmokePassed: true,
      },
    })
  })

  it('runs structural smoke on an editable sheet when the first sheet is protected', async () => {
    const cacheDir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-protected-first-sheet-'))
    mkdirSync(join(cacheDir, 'files'), { recursive: true })
    const workbookBytes = buildProtectedFirstSheetWorkbookBytes()
    const sha256 = await sha256Hex(workbookBytes)
    writeFileSync(join(cacheDir, 'files', `${sha256}.xlsx`), workbookBytes)

    const license = {
      spdxId: 'CC-BY-4.0',
      title: 'Creative Commons Attribution 4.0 International',
      evidenceUrl: 'https://creativecommons.org/licenses/by/4.0/',
    }
    const manifest: PublicWorkbookManifest = {
      ...createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z'),
      sources: [
        {
          id: 'source-protected-first-sheet',
          kind: 'direct-url',
          sourceUrl: 'https://example.com/protected-first-sheet.xlsx',
          downloadUrl: 'https://example.com/protected-first-sheet.xlsx',
          fileName: 'protected-first-sheet.xlsx',
          discoveredAt: '2026-05-07T00:00:00.000Z',
          license,
        },
      ],
      artifacts: [
        {
          id: `workbook-${sha256.slice(0, 16)}`,
          sourceId: 'source-protected-first-sheet',
          sourceUrl: 'https://example.com/protected-first-sheet.xlsx',
          downloadUrl: 'https://example.com/protected-first-sheet.xlsx',
          fileName: 'protected-first-sheet.xlsx',
          cachePath: `files/${sha256}.xlsx`,
          sha256,
          byteSize: workbookBytes.byteLength,
          workbookFingerprint: 'protected-first-sheet-fingerprint',
          fetchedAt: '2026-05-07T00:00:00.000Z',
          license,
        },
      ],
    }

    const scorecard = await buildPublicWorkbookCorpusScorecard({
      manifest,
      cacheDir,
      generatedAt: '2026-05-07T01:00:00.000Z',
      structuralSmokeSampleLimit: 1,
    })

    expect(scorecard.summary.allCachedWorkbooksPassed).toBe(true)
    expect(scorecard.cases[0]).toMatchObject({
      status: 'passed',
      validation: {
        roundTripPassed: true,
        structuralSmokePassed: true,
      },
      unsupportedFeatureClassifications: [],
    })
  })

  it('rejects stale scorecards without all cached workbook cases', async () => {
    const scorecard = await buildPublicWorkbookCorpusScorecard({
      manifest: createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z'),
      cacheDir: mkdtempSync(join(tmpdir(), 'public-workbook-corpus-empty-')),
      generatedAt: '2026-05-07T01:00:00.000Z',
    })

    expect(() =>
      validatePublicWorkbookCorpusScorecard({
        ...scorecard,
        summary: {
          ...scorecard.summary,
          cachedWorkbookCount: 1,
        },
      }),
    ).toThrow('Public workbook corpus scorecard case count does not match cached workbook count')
  })

  it('rejects scorecards when any cached workbook fails verification', async () => {
    const cacheDir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-missing-'))
    const missingHash = 'a'.repeat(64)
    const license = {
      spdxId: 'CC-BY-4.0',
      title: 'Creative Commons Attribution 4.0 International',
      evidenceUrl: 'https://creativecommons.org/licenses/by/4.0/',
    }
    const manifest: PublicWorkbookManifest = {
      ...createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z'),
      sources: [
        {
          id: 'source-missing-file',
          kind: 'direct-url',
          sourceUrl: 'https://example.com/missing.xlsx',
          downloadUrl: 'https://example.com/missing.xlsx',
          fileName: 'missing.xlsx',
          discoveredAt: '2026-05-07T00:00:00.000Z',
          license,
        },
      ],
      artifacts: [
        {
          id: 'workbook-missing-file',
          sourceId: 'source-missing-file',
          sourceUrl: 'https://example.com/missing.xlsx',
          downloadUrl: 'https://example.com/missing.xlsx',
          fileName: 'missing.xlsx',
          cachePath: 'files/missing.xlsx',
          sha256: missingHash,
          byteSize: 1024,
          workbookFingerprint: 'missing-file-fingerprint',
          fetchedAt: '2026-05-07T00:00:00.000Z',
          license,
        },
      ],
    }

    const scorecard = await buildPublicWorkbookCorpusScorecard({
      manifest,
      cacheDir,
      generatedAt: '2026-05-07T01:00:00.000Z',
    })

    expect(scorecard.summary).toMatchObject({
      cachedWorkbookCount: 1,
      importedWorkbookCount: 0,
      errorWorkbookCount: 1,
      formulaOracleMatchCount: 0,
      allCachedWorkbooksPassed: false,
    })
    expect(scorecard.cases[0]?.validation.formulaOracleMismatches).toEqual([])
    expect(() => validatePublicWorkbookCorpusScorecard(scorecard)).toThrow(
      'Public workbook corpus scorecard has cached workbooks that did not pass',
    )
  })

  it('fetches public workbook artifacts in bounded batches', async () => {
    const cacheDir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-fetch-'))
    const workbookA = buildWorkbookBytes('SummaryA')
    const workbookB = buildWorkbookBytes('SummaryB')
    const fetchMock = vi.fn(async (input: string | URL | Request) => {
      const url = typeof input === 'string' ? input : input instanceof URL ? input.href : input.url
      const sourceIndex = Number(/workbook-(\d+)\.xlsx/u.exec(url)?.[1] ?? '0')
      const bytes = sourceIndex === 1 ? workbookB : workbookA
      return new Response(bytes, {
        headers: {
          'content-length': String(bytes.byteLength),
        },
      })
    })
    vi.stubGlobal('fetch', fetchMock)

    const license = {
      spdxId: 'CC-BY-4.0',
      title: 'Creative Commons Attribution 4.0 International',
      evidenceUrl: 'https://creativecommons.org/licenses/by/4.0/',
    }
    const sources: PublicWorkbookSource[] = Array.from({ length: 50 }, (_, index) => ({
      id: `source-${String(index)}`,
      kind: 'direct-url',
      sourceUrl: `https://example.com/workbook-${String(index)}.xlsx`,
      downloadUrl: `https://example.com/workbook-${String(index)}.xlsx`,
      fileName: `workbook-${String(index)}.xlsx`,
      discoveredAt: '2026-05-07T00:00:00.000Z',
      license,
    }))
    const manifest: PublicWorkbookManifest = {
      ...createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z'),
      sources,
    }

    const fetched = await fetchPublicWorkbookArtifacts({
      manifest,
      cacheDir,
      limit: 2,
      fetchedAt: '2026-05-07T01:00:00.000Z',
    })

    expect(fetched.artifacts).toHaveLength(2)
    expect(fetchMock).toHaveBeenCalledTimes(24)
  })

  it('dedupes candidate download URLs before fetching artifacts', async () => {
    const cacheDir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-fetch-dedupe-'))
    const workbookBytes = buildWorkbookBytes()
    const fetchMock = vi.fn(async () => new Response(workbookBytes, { headers: { 'content-length': String(workbookBytes.byteLength) } }))
    vi.stubGlobal('fetch', fetchMock)

    const license = {
      spdxId: 'CC-BY-4.0',
      title: 'Creative Commons Attribution 4.0 International',
      evidenceUrl: 'https://creativecommons.org/licenses/by/4.0/',
    }
    const manifest: PublicWorkbookManifest = {
      ...createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z'),
      sources: [
        {
          id: 'source-1',
          kind: 'direct-url',
          sourceUrl: 'https://example.com/dataset-1',
          downloadUrl: 'https://example.com/shared-budget.xlsx',
          fileName: 'shared-budget.xlsx',
          discoveredAt: '2026-05-07T00:00:00.000Z',
          license,
        },
        {
          id: 'source-2',
          kind: 'direct-url',
          sourceUrl: 'https://example.com/dataset-2',
          downloadUrl: 'https://example.com/shared-budget.xlsx',
          fileName: 'shared-budget.xlsx',
          discoveredAt: '2026-05-07T00:00:00.000Z',
          license,
        },
      ],
    }

    const fetched = await fetchPublicWorkbookArtifacts({
      manifest,
      cacheDir,
      limit: 2,
      fetchedAt: '2026-05-07T01:00:00.000Z',
    })

    expect(fetchMock).toHaveBeenCalledTimes(1)
    expect(fetched.artifacts).toHaveLength(1)
  })

  it('checkpoints the manifest when fetched artifacts are committed', async () => {
    const cacheDir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-fetch-checkpoint-'))
    const workbookBytes = buildWorkbookBytes()
    vi.stubGlobal(
      'fetch',
      vi.fn(async () => new Response(workbookBytes, { headers: { 'content-length': String(workbookBytes.byteLength) } })),
    )

    const license = {
      spdxId: 'CC-BY-4.0',
      title: 'Creative Commons Attribution 4.0 International',
      evidenceUrl: 'https://creativecommons.org/licenses/by/4.0/',
    }
    const manifest: PublicWorkbookManifest = {
      ...createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z'),
      sources: [
        {
          id: 'source-checkpoint',
          kind: 'direct-url',
          sourceUrl: 'https://example.com/checkpoint.xlsx',
          downloadUrl: 'https://example.com/checkpoint.xlsx',
          fileName: 'checkpoint.xlsx',
          discoveredAt: '2026-05-07T00:00:00.000Z',
          license,
        },
      ],
    }
    const checkpoints: number[] = []

    await fetchPublicWorkbookArtifacts({
      manifest,
      cacheDir,
      limit: 1,
      fetchedAt: '2026-05-07T01:00:00.000Z',
      onArtifactsCommitted: (checkpointManifest) => checkpoints.push(checkpointManifest.artifacts.length),
    })

    expect(checkpoints).toEqual([1])
  })

  it('times out stalled workbook response bodies during fetch', async () => {
    vi.useFakeTimers()

    const cacheDir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-stalled-body-'))
    const fetchMock = vi.fn(
      async () =>
        new Response(
          new ReadableStream<Uint8Array>({
            pull: () => new Promise<void>(() => undefined),
          }),
          {
            headers: {
              'content-length': '32',
            },
          },
        ),
    )
    vi.stubGlobal('fetch', fetchMock)

    const license = {
      spdxId: 'CC-BY-4.0',
      title: 'Creative Commons Attribution 4.0 International',
      evidenceUrl: 'https://creativecommons.org/licenses/by/4.0/',
    }
    const manifest: PublicWorkbookManifest = {
      ...createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z'),
      sources: [
        {
          id: 'source-stalled-body',
          kind: 'direct-url',
          sourceUrl: 'https://example.com/stalled.xlsx',
          downloadUrl: 'https://example.com/stalled.xlsx',
          fileName: 'stalled.xlsx',
          discoveredAt: '2026-05-07T00:00:00.000Z',
          license,
        },
      ],
    }

    const fetchedPromise = fetchPublicWorkbookArtifacts({
      manifest,
      cacheDir,
      limit: 1,
      fetchedAt: '2026-05-07T01:00:00.000Z',
      downloadTimeoutMs: 5,
    })

    await vi.advanceTimersByTimeAsync(5)
    const fetched = await fetchedPromise

    expect(fetchMock).toHaveBeenCalledTimes(1)
    expect(fetched.artifacts).toHaveLength(0)
  })

  it('skips malformed CKAN resource URLs during workbook discovery', async () => {
    const fetchMock = vi.fn(async () =>
      Response.json({
        result: {
          results: [
            {
              id: 'dataset-1',
              name: 'dataset-1',
              license_id: 'CC-BY-4.0',
              license_title: 'Creative Commons Attribution 4.0 International',
              license_url: 'https://creativecommons.org/licenses/by/4.0/',
              resources: [
                {
                  id: 'bad-url',
                  name: '',
                  url: 'http:// https://example.com/not-a-url.xlsx',
                },
                {
                  id: 'good-url',
                  name: 'workbook.xlsx',
                  url: 'https://example.com/workbook.xlsx',
                },
              ],
            },
          ],
        },
      }),
    )
    vi.stubGlobal('fetch', fetchMock)

    const manifest = await discoverCkanWorkbookSources({
      manifest: createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z'),
      portalBases: ['https://example-ckan.test/api/3/action'],
      query: 'xlsx',
      limit: 10,
      rowsPerRequest: 10,
      discoveredAt: '2026-05-07T01:00:00.000Z',
    })

    expect(manifest.sources.map((source) => source.resourceId)).toEqual(['good-url'])
    expect(manifest.sources[0]).toMatchObject({
      downloadUrl: 'https://example.com/workbook.xlsx',
      fileName: 'workbook.xlsx',
    })
  })

  it('resolves relative CKAN resource URLs during workbook discovery', async () => {
    const fetchMock = vi.fn(async () =>
      Response.json({
        result: {
          results: [
            {
              id: 'dataset-relative',
              name: 'dataset-relative',
              license_id: 'CC-BY-4.0',
              license_title: 'Creative Commons Attribution 4.0 International',
              license_url: 'https://creativecommons.org/licenses/by/4.0/',
              resources: [
                {
                  id: 'relative-url',
                  name: '',
                  url: '/data/dataset/dataset-relative/resource/relative-url/download/output.xlsx',
                },
              ],
            },
          ],
        },
      }),
    )
    vi.stubGlobal('fetch', fetchMock)

    const manifest = await discoverCkanWorkbookSources({
      manifest: createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z'),
      portalBases: ['https://example-ckan.test/data/api/3/action'],
      query: 'xlsx',
      limit: 10,
      rowsPerRequest: 10,
      discoveredAt: '2026-05-07T01:00:00.000Z',
    })

    expect(manifest.sources).toHaveLength(1)
    expect(manifest.sources[0]).toMatchObject({
      downloadUrl: 'https://example-ckan.test/data/dataset/dataset-relative/resource/relative-url/download/output.xlsx',
      fileName: 'output.xlsx',
      resourceId: 'relative-url',
    })
  })

  it('filters CKAN discovery to financial workbook topic evidence when requested', async () => {
    const fetchMock = vi.fn(async () =>
      Response.json({
        result: {
          results: [
            {
              id: 'dataset-budget',
              name: 'state-budget',
              title: 'State Budget Financial Tables',
              license_id: 'CC-BY-4.0',
              license_title: 'Creative Commons Attribution 4.0 International',
              license_url: 'https://creativecommons.org/licenses/by/4.0/',
              resources: [{ id: 'budget-resource', name: 'tables.xlsx', url: 'https://example.com/tables.xlsx' }],
            },
            {
              id: 'dataset-population',
              name: 'population',
              title: 'Population estimates',
              license_id: 'CC-BY-4.0',
              license_title: 'Creative Commons Attribution 4.0 International',
              license_url: 'https://creativecommons.org/licenses/by/4.0/',
              resources: [{ id: 'population-resource', name: 'population.xlsx', url: 'https://example.com/population.xlsx' }],
            },
          ],
        },
      }),
    )
    vi.stubGlobal('fetch', fetchMock)

    const manifest = await discoverCkanWorkbookSources({
      manifest: createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z', 5_000),
      portalBases: ['https://example-ckan.test/api/3/action'],
      query: 'budget',
      limit: 5_000,
      rowsPerRequest: 10,
      discoveredAt: '2026-05-07T01:00:00.000Z',
      requiredTopic: 'financial-workpapers',
    })

    expect(manifest.sources.map((source) => source.resourceId)).toEqual(['budget-resource'])
    expect(manifest.sources[0]?.topicEvidence).toEqual(expect.arrayContaining(['budget:dataset.title']))
  })

  it('surfaces isolated verification subprocess failures with stderr evidence', async () => {
    const fixture = createIsolatedVerificationFixture()
    const child = createMockChildProcess()
    spawnMock.mockImplementationOnce(() => child)

    const scorecardPromise = buildPublicWorkbookCorpusScorecard({
      manifest: fixture.manifest,
      cacheDir: fixture.cacheDir,
      manifestPath: fixture.manifestPath,
      generatedAt: '2026-05-07T01:00:00.000Z',
      isolatedVerification: true,
      verifyConcurrency: 1,
      verifyTimeoutMs: 1_000,
    })

    child.stderr.emit('data', 'Error: Cannot find module "./public-workbook-corpus.ts"\n')
    child.emit('close', 1, null)

    const scorecard = await scorecardPromise

    expect(spawnMock).toHaveBeenCalledTimes(1)
    expect(scorecard.cases).toHaveLength(1)
    expect(scorecard.cases[0]).toMatchObject({
      status: 'error',
      passed: false,
    })
    expect(scorecard.cases[0]?.evidence).toEqual(
      expect.arrayContaining(['Verification subprocess exited with code 1', 'Error: Cannot find module "./public-workbook-corpus.ts"']),
    )
  })

  it('reports isolated verification timeouts in the evidence trail', async () => {
    vi.useFakeTimers()

    const fixture = createIsolatedVerificationFixture()
    const child = createMockChildProcess()
    spawnMock.mockImplementationOnce(() => child)

    const scorecardPromise = buildPublicWorkbookCorpusScorecard({
      manifest: fixture.manifest,
      cacheDir: fixture.cacheDir,
      manifestPath: fixture.manifestPath,
      generatedAt: '2026-05-07T01:00:00.000Z',
      isolatedVerification: true,
      verifyConcurrency: 1,
      verifyTimeoutMs: 5,
    })

    await vi.advanceTimersByTimeAsync(5)
    const scorecard = await scorecardPromise

    expect(spawnMock).toHaveBeenCalledTimes(1)
    expect(spawnMock.mock.calls[0]?.[1]).toEqual(expect.arrayContaining(['verify-artifact-worker', '--verify-max-rss-mb', '4096']))
    expect(child.kill).toHaveBeenCalledWith('SIGTERM')
    expect(scorecard.cases[0]?.status).toBe('error')
    expect(scorecard.cases[0]?.evidence).toEqual(
      expect.arrayContaining([
        'Verification timed out after 5ms',
        'The workbook was isolated in a subprocess so the corpus verification run could continue.',
      ]),
    )
  })

  it('kills isolated verification workers that exceed the RSS limit', async () => {
    vi.useFakeTimers()

    const fixture = createIsolatedVerificationFixture()
    const verificationChild = createMockChildProcess(12_345)
    const psChild = createMockChildProcess()
    spawnMock.mockImplementation((command: string) => (command === '/bin/ps' ? psChild : verificationChild))

    const scorecardPromise = buildPublicWorkbookCorpusScorecard({
      manifest: fixture.manifest,
      cacheDir: fixture.cacheDir,
      manifestPath: fixture.manifestPath,
      generatedAt: '2026-05-07T01:00:00.000Z',
      isolatedVerification: true,
      verifyConcurrency: 1,
      verifyTimeoutMs: 60_000,
      verifyMaxRssBytes: 1024 * 1024,
      verifyRssCheckIntervalMs: 100,
    })

    await vi.advanceTimersByTimeAsync(100)
    psChild.stdout.emit('data', '2048\n')
    psChild.emit('close', 0, null)

    const scorecard = await scorecardPromise

    expect(spawnMock.mock.calls[0]?.[1]).toEqual(expect.arrayContaining(['verify-artifact-worker', '--verify-max-rss-mb', '1']))
    expect(verificationChild.kill).toHaveBeenCalledWith('SIGTERM')
    expect(scorecard.cases[0]?.status).toBe('error')
    expect(scorecard.cases[0]?.evidence).toEqual(
      expect.arrayContaining([
        'Verification subprocess exceeded RSS limit: 2.0 MiB > 1.0 MiB',
        'The workbook was isolated in a subprocess so the corpus verification run could continue.',
      ]),
    )
  })

  it('caps high isolated verification RSS overrides', async () => {
    vi.useFakeTimers()

    const fixture = createIsolatedVerificationFixture()
    const child = createMockChildProcess()
    spawnMock.mockImplementationOnce(() => child)

    const scorecardPromise = buildPublicWorkbookCorpusScorecard({
      manifest: fixture.manifest,
      cacheDir: fixture.cacheDir,
      manifestPath: fixture.manifestPath,
      generatedAt: '2026-05-07T01:00:00.000Z',
      isolatedVerification: true,
      verifyConcurrency: 1,
      verifyTimeoutMs: 5,
      verifyMaxRssBytes: 12 * 1024 * 1024 * 1024,
    })

    await vi.advanceTimersByTimeAsync(5)
    await scorecardPromise

    expect(spawnMock.mock.calls[0]?.[1]).toEqual(expect.arrayContaining(['verify-artifact-worker', '--verify-max-rss-mb', '4096']))
    expect(child.kill).toHaveBeenCalledWith('SIGTERM')
  })
})

function buildWorkbookBytes(summarySheetName = 'Summary'): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const summary = XLSX.utils.aoa_to_sheet([
    ['Metric', 'Value'],
    ['Revenue', 12],
    ['Cost', 5],
    ['Profit', null],
  ])
  summary.B4 = { t: 'n', f: 'B2-B3', v: 7 }
  summary['!ref'] = 'A1:B4'
  summary['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }]
  const assumptions = XLSX.utils.aoa_to_sheet([['TaxRate'], [0.21]])
  XLSX.utils.book_append_sheet(workbook, summary, summarySheetName)
  XLSX.utils.book_append_sheet(workbook, assumptions, 'Assumptions')
  workbook.Workbook = {
    Names: [{ Name: 'ProfitCell', Ref: `${summarySheetName}!$B$4` }],
  }
  return XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' })
}

function buildTrailingSpaceWorkbookBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const sheetName = 'Table 2.1.2  '
  const sheet = XLSX.utils.aoa_to_sheet([
    ['Header', 'Value'],
    ['Amount', 12],
  ])
  sheet['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }]
  XLSX.utils.book_append_sheet(workbook, sheet, sheetName)
  return XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' })
}

function buildInteriorStyledSnapshot(startAddress: string, endAddress: string, backgroundColor = '#ffcc00'): WorkbookSnapshot {
  return {
    version: 1,
    workbook: {
      name: 'interior-styled-range',
      metadata: {
        styles: [
          {
            id: 'accent-style',
            fill: { backgroundColor },
            font: { bold: true },
          },
        ],
      },
    },
    sheets: [
      {
        id: 1,
        name: 'Styled',
        order: 0,
        cells: [{ address: 'B2', value: 'Interior' }],
        metadata: {
          styleRanges: [
            {
              range: { sheetName: 'Styled', startAddress, endAddress },
              styleId: 'accent-style',
            },
          ],
        },
      },
    ],
  }
}

function buildProtectedFirstWorkbookBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const protectedSheet = XLSX.utils.aoa_to_sheet([
    ['Locked', 'Value'],
    ['Amount', 12],
  ])
  protectedSheet['!protect'] = {}
  const mutableSheet = XLSX.utils.aoa_to_sheet([
    ['Open', 'Value'],
    ['Amount', 7],
  ])
  XLSX.utils.book_append_sheet(workbook, protectedSheet, 'Locked')
  XLSX.utils.book_append_sheet(workbook, mutableSheet, 'Open')
  return XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' })
}

function buildProtectedFirstSheetWorkbookBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const locked = XLSX.utils.aoa_to_sheet([['Locked'], [1]])
  locked['!protect'] = {}
  const open = XLSX.utils.aoa_to_sheet([['Open'], [2]])
  XLSX.utils.book_append_sheet(workbook, locked, 'Locked')
  XLSX.utils.book_append_sheet(workbook, open, 'Open')
  return XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' })
}

function createIsolatedVerificationFixture(): {
  readonly cacheDir: string
  readonly manifestPath: string
  readonly manifest: PublicWorkbookManifest
} {
  const cacheDir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-isolated-'))
  const manifestPath = join(cacheDir, 'manifest.json')
  const manifest = {
    ...createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z'),
    sources: [
      {
        id: 'source-isolated-verification',
        kind: 'direct-url',
        sourceUrl: 'https://example.com/public-budget.xlsx',
        downloadUrl: 'https://example.com/public-budget.xlsx',
        fileName: 'public-budget.xlsx',
        discoveredAt: '2026-05-07T00:00:00.000Z',
        license: {
          spdxId: 'CC-BY-4.0',
          title: 'Creative Commons Attribution 4.0 International',
          evidenceUrl: 'https://creativecommons.org/licenses/by/4.0/',
        },
      },
    ],
    artifacts: [
      {
        id: 'workbook-isolated-verification',
        sourceId: 'source-isolated-verification',
        sourceUrl: 'https://example.com/public-budget.xlsx',
        downloadUrl: 'https://example.com/public-budget.xlsx',
        fileName: 'public-budget.xlsx',
        cachePath: 'files/public-budget.xlsx',
        sha256: 'a'.repeat(64),
        byteSize: 1_024,
        workbookFingerprint: 'isolated-verification-fingerprint',
        fetchedAt: '2026-05-07T00:00:00.000Z',
        license: {
          spdxId: 'CC-BY-4.0',
          title: 'Creative Commons Attribution 4.0 International',
          evidenceUrl: 'https://creativecommons.org/licenses/by/4.0/',
        },
      },
    ],
  } satisfies PublicWorkbookManifest

  return { cacheDir, manifestPath, manifest }
}

class MockProcessStream extends EventEmitter {
  readonly setEncoding = vi.fn()
}

class MockChildProcess extends EventEmitter {
  readonly stdout = new MockProcessStream()
  readonly stderr = new MockProcessStream()
  readonly kill = vi.fn()

  constructor(readonly pid?: number) {
    super()
  }
}

function createMockChildProcess(pid?: number): MockChildProcess {
  return new MockChildProcess(pid)
}
