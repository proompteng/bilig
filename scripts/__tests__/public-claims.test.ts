import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'

import { buildPublicClaimCheckReport, collectPublicClaimFiles, findBroadGoogleSheetsTenXClaims, rootDir } from '../check-public-claims.ts'

const tempRoots: string[] = []

afterEach(() => {
  for (const root of tempRoots.splice(0)) {
    rmSync(root, { recursive: true, force: true })
  }
})

describe('public claim check', () => {
  it('finds broad Google Sheets 10x claims', () => {
    expect(findBroadGoogleSheetsTenXClaims('Bilig is 10x faster than Google Sheets for every workbook.', 'README.md')).toEqual([
      {
        path: 'README.md',
        line: 1,
        column: 10,
        match: '10x faster than Google Sheets',
        text: 'Bilig is 10x faster than Google Sheets for every workbook.',
      },
    ])

    expect(findBroadGoogleSheetsTenXClaims('Bilig beats Google Sheets by 10x on UI.', 'docs/index.html')).toHaveLength(1)
    expect(findBroadGoogleSheetsTenXClaims('No blanket 10x claim is allowed yet.', 'README.md')).toEqual([])
  })

  it('blocks public broad claims', () => {
    const repoRoot = makeTempRepo({
      'README.md': 'Bilig is 10x better than Google Sheets.',
      'docs/index.html': '<h1>Scoped compatibility evidence</h1>',
    })

    const report = buildPublicClaimCheckReport({ repoRoot })

    expect(report.violations).toHaveLength(1)
    expect(report.violations[0]).toMatchObject({
      path: 'README.md',
      match: '10x better than Google Sheets',
    })
  })

  it('blocks broad claims without trusting caller-provided contract state', () => {
    const repoRoot = makeTempRepo({
      'README.md': 'Bilig is 10x faster than Google Sheets.',
    })

    const report = buildPublicClaimCheckReport({ repoRoot })

    expect(report.violations).toHaveLength(1)
  })

  it('scans public docs while excluding internal planning docs', () => {
    const repoRoot = makeTempRepo({
      'README.md': 'Root readme',
      'docs/index.html': '<main>Public site</main>',
      'docs/public-api.md': '# Public API',
      'internal/plans/workbook-view-platform-10x-production-plan-2026-04-29.md': '# Internal plan',
      'packages/workpaper/README.md': '# Headless',
    })

    expect(collectPublicClaimFiles(repoRoot)).toEqual([
      'README.md',
      'docs/index.html',
      'docs/public-api.md',
      'packages/workpaper/README.md',
    ])
  })

  it('passes for the current checked-in public surfaces', () => {
    expect(buildPublicClaimCheckReport({ repoRoot: rootDir }).violations).toEqual([])
  })
})

function makeTempRepo(files: Record<string, string>): string {
  const repoRoot = mkdtempSync(join(tmpdir(), 'bilig-public-claims-'))
  tempRoots.push(repoRoot)
  for (const [repoPath, source] of Object.entries(files)) {
    const path = join(repoRoot, repoPath)
    mkdirSync(dirname(path), { recursive: true })
    writeFileSync(path, source)
  }
  return repoRoot
}
