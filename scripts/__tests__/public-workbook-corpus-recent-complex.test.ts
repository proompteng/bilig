import { mkdtempSync, readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

import { describe, expect, it } from 'vitest'

import { defaultRecentComplexCkanPortalBases } from '../public-workbook-corpus-discovery.ts'
import { createEmptyPublicWorkbookManifest } from '../public-workbook-corpus-json.ts'
import { defaultRecentComplexGithubRepositoryQueries } from '../public-workbook-corpus-github.ts'
import {
  buildPublicWorkbookCorpusRecentComplexSummary,
  hasRecentWorkbookEvidence,
  recentComplexityScore,
  selectRecentComplexWorkbookCandidates,
  validatePublicWorkbookCorpusRecentComplexSummary,
} from '../public-workbook-corpus-recent-complex.ts'
import { defaultRecentComplexWorkbookQueries } from '../public-workbook-corpus-topics.ts'
import type { PublicWorkbookArtifact, PublicWorkbookCorpusCase, PublicWorkbookCorpusScorecard } from '../public-workbook-corpus-types.ts'

describe('public workbook recent complex headless corpus gate', () => {
  it('selects only passing 2025-2026 complex formula workbooks', () => {
    const complex = passedCase(recentArtifact('recent-complex'), {
      cellCount: 5_000,
      formulaCellCount: 80,
      sheetCount: 4,
      tableCount: 2,
    })
    const simple = passedCase(recentArtifact('recent-simple'), {
      cellCount: 50,
      formulaCellCount: 1,
      sheetCount: 1,
      tableCount: 0,
    })
    const old = passedCase(artifact('old-complex'), {
      cellCount: 5_000,
      formulaCellCount: 80,
      sheetCount: 4,
      tableCount: 2,
    })
    const scorecard = scorecardFor([complex, simple, old])

    const selected = selectRecentComplexWorkbookCandidates({
      cacheDir: '/tmp/cache',
      manifestArtifacts: [recentArtifact('recent-complex'), recentArtifact('recent-simple'), artifact('old-complex')],
      scorecard,
      minComplexityScore: 5,
      minFormulaCells: 10,
    })

    expect(selected.map((candidate) => candidate.artifact.id)).toEqual(['recent-complex'])
    expect(hasRecentWorkbookEvidence(selected[0]?.artifact ?? {})).toBe(true)
    expect(recentComplexityScore(complex)).toBeGreaterThanOrEqual(5)
  })

  it('does not count volume-only datasets as complex formula workbooks by default', () => {
    const formula = passedCase(recentArtifact('recent-formula-complex'), {
      cellCount: 5_000,
      formulaCellCount: 1,
      sheetCount: 4,
    })
    const volumeOnly = passedCase(recentArtifact('recent-volume-only'), {
      cellCount: 100_000,
      formulaCellCount: 0,
      sheetCount: 8,
    })
    const selected = selectRecentComplexWorkbookCandidates({
      cacheDir: '/tmp/cache',
      manifestArtifacts: [recentArtifact('recent-formula-complex'), recentArtifact('recent-volume-only')],
      scorecard: scorecardFor([formula, volumeOnly]),
    })

    expect(recentComplexityScore(volumeOnly)).toBeGreaterThanOrEqual(5)
    expect(selected.map((candidate) => candidate.artifact.id)).toEqual(['recent-formula-complex'])
  })

  it('requires formula oracle comparisons before selecting recent formula workbooks', () => {
    const comparedArtifact = recentArtifact('recent-compared-formula')
    const cacheOnlyArtifact = recentArtifact('recent-cache-only-formula')
    const compared = passedCase(comparedArtifact, {
      cellCount: 5_000,
      formulaCellCount: 80,
      sheetCount: 4,
    })
    const cacheOnly = {
      ...passedCase(cacheOnlyArtifact, {
        cellCount: 5_000,
        formulaCellCount: 80,
        sheetCount: 4,
      }),
      validation: {
        ...passedCase(cacheOnlyArtifact).validation,
        formulaOracleComparisons: 0,
      },
    }

    const selected = selectRecentComplexWorkbookCandidates({
      cacheDir: '/tmp/cache',
      manifestArtifacts: [comparedArtifact, cacheOnlyArtifact],
      scorecard: scorecardFor([compared, cacheOnly]),
    })

    expect(selected.map((candidate) => candidate.artifact.id)).toEqual(['recent-compared-formula'])
  })

  it('checks the headless result against the same recent complex artifact set', () => {
    const dir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-recent-complex-'))
    const cacheDir = join(dir, 'cache')
    const manifestPath = join(dir, 'manifest.json')
    const scorecardPath = join(dir, 'scorecard.json')
    const headlessScorecardPath = join(dir, 'headless-scorecard.json')
    const selectedArtifact = recentArtifact('recent-complex')
    const manifest = {
      ...createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z', 1),
      artifacts: [selectedArtifact],
      sources: [sourceFor(selectedArtifact)],
    }
    writeFileSync(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`)
    writeFileSync(scorecardPath, `${JSON.stringify(scorecardFor([passedCase(selectedArtifact)]), null, 2)}\n`)
    writeFileSync(
      headlessScorecardPath,
      `${JSON.stringify(
        {
          summary: {
            totalFiles: 1,
            filesProcessed: 1,
            ok: 1,
            failedTimeouts: 0,
            failedErrors: 0,
            formulaCells: 80,
            comparableFormulaCells: 80,
            matchingFormulaCells: 80,
            mismatchedFormulaCells: 0,
            skippedFormulaCells: 0,
            matchRate: 1,
            elapsedMs: 100,
          },
          files: [
            {
              path: join(cacheDir, selectedArtifact.cachePath),
              fileName: selectedArtifact.fileName,
              status: 'ok',
              formulaCells: 80,
              comparableFormulaCells: 80,
              matchingFormulaCells: 80,
              mismatchedFormulaCells: 0,
              skippedFormulaCells: 0,
              matchRate: 1,
              elapsedMs: 100,
            },
          ],
          mismatches: [],
          skippedByReason: {
            'missing-cached-result': 0,
            'stale-cached-result': 0,
            'stale-cached-name-error': 0,
            'unsupported-cached-result-type': 0,
            'volatile-or-environment-dependent-formula': 0,
          },
        },
        null,
        2,
      )}\n`,
    )

    const summary = buildPublicWorkbookCorpusRecentComplexSummary({
      cacheDir,
      childTimeoutMs: 31_000,
      corpusRunStopMarkerPath: join(dir, 'missing-stop.md'),
      generatedAt: '2026-05-08T00:00:00.000Z',
      headlessScorecardPath,
      manifestPath,
      maxFileBytes: 50 * 1024 * 1024,
      minComplexityScore: 5,
      minFormulaCells: 10,
      scorecardPath,
      targetWorkbookCount: 1,
      timeoutMs: 30_000,
    })

    expect(summary).toMatchObject({
      targetWorkbookCount: 1,
      manifestTargetWorkbookCount: 1,
      publicPassingRecentComplexCount: 1,
      endToEndPassingWorkbookCount: 1,
      remainingToTarget: 0,
      allSelectedHeadlessWorkbooksPassed: true,
    })
    expect(validatePublicWorkbookCorpusRecentComplexSummary(summary, { requireTarget: true })).toEqual([])
  })

  it('does not count a selected workbook as end-to-end passing without comparable headless formulas', () => {
    const dir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-recent-complex-no-comparable-'))
    const cacheDir = join(dir, 'cache')
    const manifestPath = join(dir, 'manifest.json')
    const scorecardPath = join(dir, 'scorecard.json')
    const headlessScorecardPath = join(dir, 'headless-scorecard.json')
    const selectedArtifact = recentArtifact('recent-complex')
    const manifest = {
      ...createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z', 1),
      artifacts: [selectedArtifact],
      sources: [sourceFor(selectedArtifact)],
    }
    writeFileSync(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`)
    writeFileSync(scorecardPath, `${JSON.stringify(scorecardFor([passedCase(selectedArtifact)]), null, 2)}\n`)
    writeFileSync(
      headlessScorecardPath,
      `${JSON.stringify(
        {
          files: [
            {
              comparableFormulaCells: 0,
              matchRate: 1,
              mismatchedFormulaCells: 0,
              path: join(cacheDir, selectedArtifact.cachePath),
              status: 'ok',
            },
          ],
          mismatches: [],
          schemaVersion: 1,
          summary: {
            failedErrors: 0,
            failedTimeouts: 0,
            mismatchedFormulaCells: 0,
            ok: 1,
            totalFiles: 1,
          },
        },
        null,
        2,
      )}\n`,
    )

    const summary = buildPublicWorkbookCorpusRecentComplexSummary({
      cacheDir,
      childTimeoutMs: 31_000,
      corpusRunStopMarkerPath: join(dir, 'missing-stop.md'),
      generatedAt: '2026-05-08T00:00:00.000Z',
      headlessScorecardPath,
      manifestPath,
      maxFileBytes: 50 * 1024 * 1024,
      minComplexityScore: 5,
      minFormulaCells: 10,
      scorecardPath,
      targetWorkbookCount: 1,
      timeoutMs: 30_000,
    })

    expect(summary.endToEndPassingWorkbookCount).toBe(0)
    expect(summary.sampleMissingHeadlessArtifactIds).toEqual([selectedArtifact.id])
    expect(validatePublicWorkbookCorpusRecentComplexSummary(summary, { requireTarget: true })).toEqual([
      'one or more selected recent complex workbooks did not pass headless verification',
      'end-to-end recent complex headless target not met: 0/1',
    ])
  })

  it('exposes package scripts for the recent complex corpus lane', () => {
    const scripts = readPackageScripts()

    expect(scripts['public-workbook-corpus:recent-complex:plan']).toBe('bun scripts/public-workbook-corpus-recent-complex.ts plan')
    expect(scripts['public-workbook-corpus:retarget-recent-complex']).toContain('public-workbook-corpus.ts retarget')
    expect(scripts['public-workbook-corpus:discover-recent-complex']).toContain('discover-recent-complex-ckan')
    expect(scripts['public-workbook-corpus:discover-recent-complex-hdx']).toContain('https://data.humdata.org/api/3/action')
    expect(scripts['public-workbook-corpus:discover-recent-complex-github']).toContain('discover-recent-complex-github')
    expect(scripts['public-workbook-corpus:discover-recent-complex-zenodo']).toContain('discover-recent-complex-zenodo')
    expect(scripts['public-workbook-corpus:discover-recent-complex-figshare']).toContain('discover-recent-complex-figshare')
    expect(scripts['public-workbook-corpus:fetch-recent-complex']).toContain('--fetch-batch-size 2')
    expect(scripts['public-workbook-corpus:fetch-recent-complex']).toContain('--limit 5000')
    expect(scripts['public-workbook-corpus:headless-recent-complex']).toContain('public-workbook-corpus-recent-complex.ts headless')
    expect(scripts['public-workbook-corpus:check-recent-complex']).toContain('--require-target')
  })

  it('prioritizes model and template discovery before broad recent workbook queries', () => {
    const modelIndex = defaultRecentComplexWorkbookQueries.indexOf('2026 model xlsx')
    const templateIndex = defaultRecentComplexWorkbookQueries.indexOf('2026 template xlsx')
    const broadBudgetIndex = defaultRecentComplexWorkbookQueries.indexOf('2026 budget')
    const broadAccountingIndex = defaultRecentComplexWorkbookQueries.indexOf('2025 accounting')

    expect(modelIndex).toBeGreaterThanOrEqual(0)
    expect(templateIndex).toBeGreaterThanOrEqual(0)
    expect(modelIndex).toBeLessThan(broadBudgetIndex)
    expect(templateIndex).toBeLessThan(broadAccountingIndex)
  })

  it('includes productive Canadian CKAN portals in recent complex discovery', () => {
    expect(defaultRecentComplexCkanPortalBases).toEqual(
      expect.arrayContaining([
        'https://data.ontario.ca/api/3/action',
        'https://open.alberta.ca/api/3/action',
        'https://catalogue.data.gov.bc.ca/api/3/action',
      ]),
    )
  })

  it('includes targeted financial-model repository queries for formula-heavy recent workbooks', () => {
    expect(defaultRecentComplexGithubRepositoryQueries).toEqual(
      expect.arrayContaining([
        '3 statement model excel license:mit',
        'excel dashboard 2025 license:mit',
        'financial analysis excel license:mit',
        'financial model 2026 excel license:mit',
        'bond valuation excel license:mit',
        'financial modeling course excel license:mit',
        'investment banking excel model license:mit',
        'm&a valuation excel license:mit',
        'portfolio optimization excel license:mit',
        'real estate financial model excel license:mit',
        'saas financial model excel license:mit',
        'startup valuation excel license:mit',
        'actuarial excel model license:mit',
      ]),
    )
  })

  it('reports when the manifest target must be expanded before fetching more artifacts', () => {
    const dir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-recent-complex-retarget-'))
    const cacheDir = join(dir, 'cache')
    const manifestPath = join(dir, 'manifest.json')
    const scorecardPath = join(dir, 'scorecard.json')
    const headlessScorecardPath = join(dir, 'headless-scorecard.json')
    const selectedArtifact = recentArtifact('recent-complex')
    const manifest = {
      ...createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z', 1),
      artifacts: [selectedArtifact],
      sources: [sourceFor(selectedArtifact)],
    }
    writeFileSync(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`)
    writeFileSync(scorecardPath, `${JSON.stringify(scorecardFor([passedCase(selectedArtifact)]), null, 2)}\n`)

    const summary = buildPublicWorkbookCorpusRecentComplexSummary({
      cacheDir,
      childTimeoutMs: 31_000,
      corpusRunStopMarkerPath: join(dir, 'missing-stop.md'),
      generatedAt: '2026-05-08T00:00:00.000Z',
      headlessScorecardPath,
      manifestPath,
      maxFileBytes: 50 * 1024 * 1024,
      minComplexityScore: 5,
      minFormulaCells: 10,
      scorecardPath,
      targetWorkbookCount: 2,
      timeoutMs: 30_000,
    })

    expect(summary.manifestTargetWorkbookCount).toBe(1)
    expect(summary.commands.retarget).toContain('public-workbook-corpus.ts retarget')
    expect(summary.commands.retarget).toContain('--target-workbook-count 2')
    expect(summary.commands.discoverHdx).toContain('https://data.humdata.org/api/3/action')
    expect(summary.commands.discoverHdx).toContain('--query 2025')
    expect(summary.commands.discoverHdx).toContain('--query 2026')
    expect(summary.commands.discoverGithub).toContain('discover-recent-complex-github')
    expect(summary.commands.discoverGithub).toContain('--skip-code-search')
    expect(summary.commands.discoverGithub).toContain('--max-repository-pages-per-query 3')
    expect(summary.commands.discoverZenodo).toContain('discover-recent-complex-zenodo')
    expect(summary.commands.discoverZenodo).toContain('--max-pages-per-query 20')
    expect(summary.commands.discoverFigshare).toContain('discover-recent-complex-figshare')
    expect(summary.commands.discoverFigshare).toContain('--max-pages-per-query 20')
    expect(validatePublicWorkbookCorpusRecentComplexSummary(summary)).toContain(
      'manifest target workbook count is below the recent complex target',
    )
  })

  it('does not suggest retargeting or fetching below the current manifest target', () => {
    const dir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-recent-complex-safe-retarget-'))
    const cacheDir = join(dir, 'cache')
    const manifestPath = join(dir, 'manifest.json')
    const scorecardPath = join(dir, 'scorecard.json')
    const headlessScorecardPath = join(dir, 'headless-scorecard.json')
    const artifactA = recentArtifact('recent-complex-a')
    const artifactB = recentArtifact('recent-complex-b')
    const manifest = {
      ...createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z', 3),
      artifacts: [artifactA, artifactB],
      sources: [sourceFor(artifactA), sourceFor(artifactB)],
    }
    writeFileSync(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`)
    writeFileSync(scorecardPath, `${JSON.stringify(scorecardFor([passedCase(artifactA), passedCase(artifactB)]), null, 2)}\n`)

    const summary = buildPublicWorkbookCorpusRecentComplexSummary({
      cacheDir,
      childTimeoutMs: 31_000,
      corpusRunStopMarkerPath: join(dir, 'missing-stop.md'),
      generatedAt: '2026-05-08T00:00:00.000Z',
      headlessScorecardPath,
      manifestPath,
      maxFileBytes: 50 * 1024 * 1024,
      minComplexityScore: 5,
      minFormulaCells: 10,
      scorecardPath,
      targetWorkbookCount: 1,
      timeoutMs: 30_000,
    })

    expect(summary.recommendedManifestTargetWorkbookCount).toBe(3)
    expect(summary.recommendedFetchArtifactLimit).toBe(3)
    expect(summary.recommendedDiscoverySourceLimit).toBe(7)
    expect(summary.commands.retarget).toContain('--target-workbook-count 3')
    expect(summary.commands.retarget).not.toContain('--target-workbook-count 1')
    expect(summary.commands.discover).toContain('--limit 7')
    expect(summary.commands.discoverGithub).toContain('--limit 7')
    expect(summary.commands.fetch).toContain('--limit 3')
    expect(summary.commands.fetch).not.toContain('--limit 1')
  })

  it('keeps discovery commands actionable when existing sources exceed the old default limit', () => {
    const dir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-recent-complex-discovery-limit-'))
    const cacheDir = join(dir, 'cache')
    const manifestPath = join(dir, 'manifest.json')
    const scorecardPath = join(dir, 'scorecard.json')
    const headlessScorecardPath = join(dir, 'headless-scorecard.json')
    const sources = Array.from({ length: 6 }, (_, index) => sourceFor(recentArtifact(`source-${String(index)}`)))
    writeFileSync(
      manifestPath,
      `${JSON.stringify(
        {
          ...createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z', 500),
          artifacts: [],
          sources,
        },
        null,
        2,
      )}\n`,
    )
    writeFileSync(scorecardPath, `${JSON.stringify(scorecardFor([]), null, 2)}\n`)

    const summary = buildPublicWorkbookCorpusRecentComplexSummary({
      cacheDir,
      childTimeoutMs: 31_000,
      corpusRunStopMarkerPath: join(dir, 'missing-stop.md'),
      generatedAt: '2026-05-08T00:00:00.000Z',
      headlessScorecardPath,
      manifestPath,
      maxFileBytes: 50 * 1024 * 1024,
      minComplexityScore: 5,
      minFormulaCells: 10,
      scorecardPath,
      targetWorkbookCount: 1,
      timeoutMs: 30_000,
    })

    expect(summary.manifestSourceCount).toBe(6)
    expect(summary.recommendedDiscoverySourceLimit).toBe(11)
    expect(summary.commands.discover).toContain('--limit 11')
    expect(summary.commands.discoverHdx).toContain('--limit 11')
    expect(summary.commands.discoverGithub).toContain('--limit 11')
    expect(summary.commands.discoverZenodo).toContain('--limit 11')
    expect(summary.commands.discoverFigshare).toContain('--limit 11')
  })
})

function readPackageScripts(): Record<string, string> {
  const parsed: unknown = JSON.parse(readFileSync(join(dirname(fileURLToPath(import.meta.url)), '../../package.json'), 'utf8'))
  if (!isRecord(parsed) || !isRecord(parsed['scripts'])) {
    throw new Error('expected package.json scripts')
  }
  const scripts: Record<string, string> = {}
  for (const [key, value] of Object.entries(parsed['scripts'])) {
    if (typeof value === 'string') {
      scripts[key] = value
    }
  }
  return scripts
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function sourceFor(workbookArtifact: PublicWorkbookArtifact) {
  return {
    id: workbookArtifact.sourceId,
    kind: 'ckan-resource' as const,
    sourceUrl: workbookArtifact.sourceUrl,
    downloadUrl: workbookArtifact.downloadUrl,
    fileName: workbookArtifact.fileName,
    discoveredAt: '2026-05-07T00:00:00.000Z',
    license: workbookArtifact.license,
    topicEvidence: workbookArtifact.topicEvidence,
  }
}

function recentArtifact(id: string): PublicWorkbookArtifact {
  return {
    ...artifact(id),
    topicEvidence: ['recent-2025:downloadUrl'],
  }
}

function artifact(id: string): PublicWorkbookArtifact {
  return {
    id,
    sourceId: `${id}-source`,
    sourceUrl: `https://example.com/${id}`,
    downloadUrl: `https://example.com/${id}.xlsx`,
    fileName: `${id}.xlsx`,
    cachePath: `files/${id}.xlsx`,
    sha256: hashFor(id),
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

function hashFor(id: string): string {
  return Array.from(id)
    .map((char) => char.codePointAt(0)?.toString(16).slice(-1) ?? '0')
    .join('')
    .padEnd(64, 'a')
    .slice(0, 64)
}

function passedCase(
  workbookArtifact: PublicWorkbookArtifact,
  counts: Partial<PublicWorkbookCorpusCase['featureCounts']> = {},
): PublicWorkbookCorpusCase {
  return {
    id: workbookArtifact.id,
    sourceId: workbookArtifact.sourceId,
    sourceUrl: workbookArtifact.sourceUrl,
    fileName: workbookArtifact.fileName,
    sha256: workbookArtifact.sha256,
    byteSize: workbookArtifact.byteSize,
    license: workbookArtifact.license,
    status: 'passed',
    passed: true,
    featureCounts: {
      sheetCount: counts.sheetCount ?? 4,
      cellCount: counts.cellCount ?? 5_000,
      formulaCellCount: counts.formulaCellCount ?? 80,
      valueCellCount: counts.valueCellCount ?? 4_000,
      definedNameCount: counts.definedNameCount ?? 0,
      tableCount: counts.tableCount ?? 2,
      chartCount: counts.chartCount ?? 0,
      pivotCount: counts.pivotCount ?? 0,
      mergeCount: counts.mergeCount ?? 0,
      styleRangeCount: counts.styleRangeCount ?? 0,
      conditionalFormatCount: counts.conditionalFormatCount ?? 0,
      dataValidationCount: counts.dataValidationCount ?? 0,
      macroPayloadCount: counts.macroPayloadCount ?? 0,
      warningCount: counts.warningCount ?? 0,
    },
    workbookMetadata: { workbookName: workbookArtifact.fileName, sheetNames: ['Sheet1'], dimensions: [] },
    validation: {
      importPassed: true,
      formulaOraclePassed: true,
      formulaOracleComparisons: counts.formulaCellCount ?? 80,
      formulaOracleMismatches: [],
      roundTripPassed: true,
      structuralSmokePassed: true,
    },
    unsupportedFeatureClassifications: [],
    evidence: [],
  }
}

function scorecardFor(cases: readonly PublicWorkbookCorpusCase[]): PublicWorkbookCorpusScorecard {
  return {
    schemaVersion: 1,
    suite: 'public-workbook-corpus',
    generatedAt: '2026-05-08T00:00:00.000Z',
    summary: {
      targetWorkbookCount: cases.length,
      sourceCount: cases.length,
      cachedWorkbookCount: cases.length,
      importedWorkbookCount: cases.length,
      passedWorkbookCount: cases.length,
      failedWorkbookCount: 0,
      errorWorkbookCount: 0,
      unsupportedWorkbookCount: 0,
      formulaOracleComparisonCount: cases.reduce((sum, entry) => sum + entry.validation.formulaOracleComparisons, 0),
      formulaOracleMatchCount: cases.reduce((sum, entry) => sum + entry.validation.formulaOracleComparisons, 0),
      structuralSmokeRunCount: cases.length,
      allCachedWorkbooksPassed: true,
      remainingToTarget: 0,
    },
    cases,
  }
}
