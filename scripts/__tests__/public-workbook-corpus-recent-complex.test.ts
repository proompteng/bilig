import { mkdtempSync, readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

import { describe, expect, it } from 'vitest'

import { createEmptyPublicWorkbookManifest } from '../public-workbook-corpus-json.ts'
import {
  buildPublicWorkbookCorpusRecentComplexSummary,
  hasRecentWorkbookEvidence,
  recentComplexityScore,
  selectRecentComplexWorkbookCandidates,
  validatePublicWorkbookCorpusRecentComplexSummary,
} from '../public-workbook-corpus-recent-complex.ts'
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
      publicPassingRecentComplexCount: 1,
      endToEndPassingWorkbookCount: 1,
      remainingToTarget: 0,
      allSelectedHeadlessWorkbooksPassed: true,
    })
    expect(validatePublicWorkbookCorpusRecentComplexSummary(summary, { requireTarget: true })).toEqual([])
  })

  it('exposes package scripts for the recent complex corpus lane', () => {
    const scripts = readPackageScripts()

    expect(scripts['public-workbook-corpus:recent-complex:plan']).toBe('bun scripts/public-workbook-corpus-recent-complex.ts plan')
    expect(scripts['public-workbook-corpus:discover-recent-complex']).toContain('discover-recent-complex-ckan')
    expect(scripts['public-workbook-corpus:fetch-recent-complex']).toContain('--fetch-batch-size 2')
    expect(scripts['public-workbook-corpus:headless-recent-complex']).toContain('public-workbook-corpus-recent-complex.ts headless')
    expect(scripts['public-workbook-corpus:check-recent-complex']).toContain('--require-target')
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
