import { existsSync, mkdtempSync, readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { spawnSync } from 'node:child_process'

import { describe, expect, it } from 'vitest'

import { asRecord, createEmptyPublicWorkbookManifest } from '../public-workbook-corpus-json.ts'

describe('public workbook financial corpus plan CLI', () => {
  it('plans the financial corpus lane without creating cache files or starting network work', () => {
    const dir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-financial-plan-'))
    const manifestPath = join(dir, 'manifest.json')
    const cacheDir = join(dir, 'cache')
    const result = spawnSync(
      'bun',
      [financialPlanScriptPath(), '--manifest', manifestPath, '--cache-dir', cacheDir, '--target-workbook-count', '5', '--limit', '5'],
      {
        encoding: 'utf8',
      },
    )
    const plan: unknown = JSON.parse(result.stdout)

    expect(result.status).toBe(0)
    expect(existsSync(manifestPath)).toBe(false)
    expect(existsSync(cacheDir)).toBe(false)
    expect(plan).toMatchObject({
      mode: 'plan',
      corpus: 'financial-accounting-workpapers',
      manifestExists: false,
      targetWorkbookCount: 5,
      sourceCount: 0,
      targetArtifactCount: 5,
      cachedArtifactCount: 0,
      remainingArtifactSlots: 5,
      candidateSourceCount: 0,
      candidateSourceDeficitCount: 5,
      minimumAdditionalSourceCount: 5,
      recommendedDiscoveryLimit: 5,
      targetReachableFromKnownCandidates: false,
      commands: {
        discoverPlan: expect.stringContaining('public-workbook-corpus:discover-financial:plan'),
        discover: expect.stringContaining('public-workbook-corpus:discover-financial'),
        fetchPlan: expect.stringContaining('public-workbook-corpus:fetch-financial:plan'),
        fetch: expect.stringContaining('public-workbook-corpus:fetch-financial'),
        verify: expect.stringContaining('public-workbook-corpus:verify-financial'),
        check: expect.stringContaining('public-workbook-corpus:check-financial'),
      },
      sampledCandidateSources: [],
    })
  })

  it('exposes non-mutating package scripts for the financial corpus lane', () => {
    const packageJson = asRecord(JSON.parse(readFileSync(packageJsonPath(), 'utf8')))
    const scripts = asRecord(packageJson['scripts'])

    expect(scripts['public-workbook-corpus:discover-financial:plan']).toBe('bun scripts/public-workbook-corpus-financial-plan.ts')
    expect(scripts['public-workbook-corpus:fetch-financial:plan']).toBe(
      'bun scripts/public-workbook-corpus.ts fetch --dry-run --manifest .cache/public-workbook-corpus-financial/manifest.json --cache-dir .cache/public-workbook-corpus-financial --limit 5000 --sample-limit 20',
    )
  })

  it('lets package-script arguments override baked-in corpus paths', () => {
    const dir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-financial-override-'))
    const missingManifestPath = join(dir, 'missing-manifest.json')
    const manifestPath = join(dir, 'manifest.json')
    writeFileSync(manifestPath, `${JSON.stringify(createEmptyPublicWorkbookManifest(undefined, 5), null, 2)}\n`)

    const result = spawnSync(
      'bun',
      [corpusScriptPath(), 'fetch', '--dry-run', '--manifest', missingManifestPath, '--manifest', manifestPath, '--limit', '5'],
      {
        encoding: 'utf8',
      },
    )
    const plan: unknown = JSON.parse(result.stdout)

    expect(result.status).toBe(0)
    expect(plan).toMatchObject({
      targetArtifactCount: 5,
      sourceCount: 0,
      cachedArtifactCount: 0,
    })
  })
})

function financialPlanScriptPath(): string {
  return join(dirname(fileURLToPath(import.meta.url)), '../public-workbook-corpus-financial-plan.ts')
}

function corpusScriptPath(): string {
  return join(dirname(fileURLToPath(import.meta.url)), '../public-workbook-corpus.ts')
}

function packageJsonPath(): string {
  return join(dirname(fileURLToPath(import.meta.url)), '../../package.json')
}
