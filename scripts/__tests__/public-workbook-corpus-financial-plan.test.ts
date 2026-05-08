import { existsSync, mkdtempSync, readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { spawnSync } from 'node:child_process'

import { describe, expect, it } from 'vitest'

import { asRecord, createEmptyPublicWorkbookManifest } from '../public-workbook-corpus-json.ts'
import {
  validatePublicWorkbookCorpusFinancialPlan,
  type PublicWorkbookCorpusFinancialPlan,
} from '../public-workbook-corpus-financial-plan.ts'
import type { PublicWorkbookSource } from '../public-workbook-corpus-types.ts'

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
      schemaVersion: 1,
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
      needsAdditionalDiscovery: true,
      recommendedFetchLimit: 5,
      recommendedFetchTrancheSize: 20,
      recommendedDiscoveryLimit: 5,
      targetReachableFromKnownCandidates: false,
      commands: {
        discoverPlan: expect.stringContaining('public-workbook-corpus:discover-financial:plan'),
        discover: expect.stringContaining('public-workbook-corpus:discover-financial'),
        fetchPlan: expect.stringContaining('public-workbook-corpus:fetch-financial:plan'),
        fetch: expect.stringContaining('public-workbook-corpus:fetch-financial'),
        resumePlan: expect.stringContaining('public-workbook-corpus:resume-financial:plan'),
        resumeCheck: expect.stringContaining('public-workbook-corpus:resume-financial:check'),
        verify: expect.stringContaining('public-workbook-corpus:verify-financial'),
        check: expect.stringContaining('public-workbook-corpus:check-financial'),
      },
      sampledCandidateSources: [],
    })
  })

  it('checks the financial corpus lane plan without creating cache files or starting network work', () => {
    const dir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-financial-check-'))
    const manifestPath = join(dir, 'manifest.json')
    const cacheDir = join(dir, 'cache')
    const inactiveStopMarkerPath = join(dir, 'missing-stop.md')
    const result = spawnSync(
      'bun',
      [
        financialPlanScriptPath(),
        '--check',
        '--manifest',
        manifestPath,
        '--cache-dir',
        cacheDir,
        '--corpus-run-stop-marker',
        inactiveStopMarkerPath,
        '--target-workbook-count',
        '5',
        '--limit',
        '5',
      ],
      {
        encoding: 'utf8',
      },
    )
    const check = JSON.parse(result.stdout)

    expect(result.status).toBe(0)
    expect(existsSync(manifestPath)).toBe(false)
    expect(existsSync(cacheDir)).toBe(false)
    expect(check).toMatchObject({
      mode: 'check',
      schemaVersion: 1,
      corpus: 'financial-accounting-workpapers',
      targetWorkbookCount: 5,
      sourceCount: 0,
      cachedArtifactCount: 0,
      remainingArtifactSlots: 5,
      candidateSourceCount: 0,
      needsAdditionalDiscovery: true,
      recommendedFetchLimit: 5,
      stopMarker: {
        active: false,
        requiresExplicitResume: false,
      },
      nextCommands: {
        discover: expect.stringContaining('public-workbook-corpus:discover-financial'),
        fetch: expect.stringContaining('public-workbook-corpus:fetch-financial'),
        fetchPlan: expect.stringContaining('public-workbook-corpus:fetch-financial:plan'),
        resumeCheck: expect.stringContaining('public-workbook-corpus:resume-financial:check'),
        verify: expect.stringContaining('public-workbook-corpus:verify-financial'),
        check: expect.stringContaining('public-workbook-corpus:check-financial'),
      },
      blockedCommands: {},
    })
  })

  it('keeps stop-marker blocked mutating commands out of financial check next steps', () => {
    const dir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-financial-check-paused-'))
    const manifestPath = join(dir, 'manifest.json')
    const cacheDir = join(dir, 'cache')
    const stopMarkerPath = join(dir, 'stop.md')
    writeFileSync(stopMarkerPath, '# stop\n')

    const result = spawnSync(
      'bun',
      [
        financialPlanScriptPath(),
        '--check',
        '--manifest',
        manifestPath,
        '--cache-dir',
        cacheDir,
        '--corpus-run-stop-marker',
        stopMarkerPath,
        '--target-workbook-count',
        '5',
        '--limit',
        '5',
      ],
      {
        encoding: 'utf8',
      },
    )
    const check = JSON.parse(result.stdout)

    expect(result.status).toBe(0)
    expect(check).toMatchObject({
      mode: 'check',
      stopMarker: {
        active: true,
        requiresExplicitResume: true,
        overrideFlag: '--allow-active-stop-marker',
        overrideEnvVar: 'BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE',
      },
      nextCommands: {
        fetchPlan: expect.stringContaining('public-workbook-corpus:fetch-financial:plan'),
        resumeCheck: expect.stringContaining('public-workbook-corpus:resume-financial:check'),
        check: expect.stringContaining('public-workbook-corpus:check-financial'),
      },
      blockedCommands: {
        discover: expect.stringContaining('public-workbook-corpus:discover-financial'),
        fetch: expect.stringContaining('public-workbook-corpus:fetch-financial'),
        fetchAll: expect.stringContaining('public-workbook-corpus:fetch-financial'),
        verify: expect.stringContaining('public-workbook-corpus:verify-financial'),
      },
    })
    expect(check.nextCommands).not.toHaveProperty('discover')
    expect(check.nextCommands).not.toHaveProperty('fetch')
    expect(check.nextCommands).not.toHaveProperty('verify')
    expect(check.blockedCommands.discover).toContain('BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE=1')
    expect(check.blockedCommands.discover).toContain('--allow-active-stop-marker')
    expect(check.blockedCommands.fetch).toContain('BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE=1')
    expect(check.blockedCommands.fetch).toContain('--allow-active-stop-marker')
    expect(check.blockedCommands.verify).toContain('BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE=1')
    expect(check.blockedCommands.verify).toContain('--allow-active-stop-marker')
  })

  it('omits mutating discovery commands when known financial candidates can fill the target', () => {
    const dir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-financial-plan-no-discovery-'))
    const manifestPath = join(dir, 'manifest.json')
    const cacheDir = join(dir, 'cache')
    const manifest = {
      ...createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z', 2),
      sources: [financialSource('source-a'), financialSource('source-b')],
    }
    writeFileSync(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`)

    const result = spawnSync(
      'bun',
      [financialPlanScriptPath(), '--manifest', manifestPath, '--cache-dir', cacheDir, '--target-workbook-count', '2', '--limit', '2'],
      {
        encoding: 'utf8',
      },
    )
    const plan = asRecord(JSON.parse(result.stdout))
    const commands = asRecord(plan['commands'])

    expect(result.status).toBe(0)
    expect(plan).toMatchObject({
      candidateSourceCount: 2,
      candidateSourceDeficitCount: 0,
      needsAdditionalDiscovery: false,
      targetReachableFromKnownCandidates: true,
    })
    expect(commands['discoverPlan']).toBeNull()
    expect(commands['discover']).toBeNull()
    expect(commands['fetch']).toEqual(expect.stringContaining('public-workbook-corpus:fetch-financial'))
  })

  it('recommends bounded financial fetch tranches before the full target fetch', () => {
    const dir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-financial-plan-tranche-'))
    const manifestPath = join(dir, 'manifest.json')
    const cacheDir = join(dir, 'cache')
    const manifest = {
      ...createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z', 50),
      sources: Array.from({ length: 50 }, (_, index) => financialSource(`source-${String(index + 1)}`)),
    }
    writeFileSync(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`)

    const result = spawnSync(
      'bun',
      [
        financialPlanScriptPath(),
        '--manifest',
        manifestPath,
        '--cache-dir',
        cacheDir,
        '--target-workbook-count',
        '50',
        '--limit',
        '50',
        '--fetch-tranche-size',
        '7',
      ],
      {
        encoding: 'utf8',
      },
    )
    const plan = asRecord(JSON.parse(result.stdout))
    const commands = asRecord(plan['commands'])

    expect(result.status).toBe(0)
    expect(plan).toMatchObject({
      cachedArtifactCount: 0,
      recommendedFetchLimit: 7,
      recommendedFetchTrancheSize: 7,
      remainingArtifactSlots: 50,
    })
    expect(commands['fetch']).toEqual(expect.stringContaining('--limit 7'))
    expect(commands['fetchAll']).toEqual(expect.stringContaining('--limit 50'))
  })

  it('exposes non-mutating package scripts for the financial corpus lane', () => {
    const packageJson = asRecord(JSON.parse(readFileSync(packageJsonPath(), 'utf8')))
    const scripts = asRecord(packageJson['scripts'])

    expect(scripts['public-workbook-corpus:discover-financial:plan']).toBe('bun scripts/public-workbook-corpus-financial-plan.ts')
    expect(scripts['public-workbook-corpus:discover-financial:check']).toBe('bun scripts/public-workbook-corpus-financial-plan.ts --check')
    expect(scripts['public-workbook-corpus:fetch-financial:plan']).toBe(
      'bun scripts/public-workbook-corpus.ts fetch --dry-run --manifest .cache/public-workbook-corpus-financial/manifest.json --cache-dir .cache/public-workbook-corpus-financial --limit 5000 --sample-limit 20',
    )
    expect(scripts['public-workbook-corpus:resume-financial:plan']).toBe(
      'bun scripts/public-workbook-corpus-resume-plan.ts --manifest .cache/public-workbook-corpus-financial/manifest.json --cache-dir .cache/public-workbook-corpus-financial --scorecard .cache/public-workbook-corpus-financial/scorecard.json --verify-checkpoint .cache/public-workbook-corpus-financial/verification-checkpoint.json --fetch-limit 5000 --fetch-batch-size 6',
    )
    expect(scripts['public-workbook-corpus:resume-financial:check']).toBe(
      'bun scripts/public-workbook-corpus-resume-plan.ts --check --manifest .cache/public-workbook-corpus-financial/manifest.json --cache-dir .cache/public-workbook-corpus-financial --scorecard .cache/public-workbook-corpus-financial/scorecard.json --verify-checkpoint .cache/public-workbook-corpus-financial/verification-checkpoint.json --fetch-limit 5000 --fetch-batch-size 6',
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

  it('marks mutating financial commands with explicit stop-marker overrides when paused', () => {
    const dir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-financial-paused-'))
    const manifestPath = join(dir, 'manifest.json')
    const stopMarkerPath = join(dir, 'stop.md')
    writeFileSync(stopMarkerPath, '# stop\n')

    const result = spawnSync(
      'bun',
      [
        financialPlanScriptPath(),
        '--manifest',
        manifestPath,
        '--corpus-run-stop-marker',
        stopMarkerPath,
        '--target-workbook-count',
        '5',
        '--limit',
        '5',
      ],
      {
        encoding: 'utf8',
      },
    )
    const plan = asRecord(JSON.parse(result.stdout))
    const stopMarker = asRecord(plan['stopMarker'])
    const commands = asRecord(plan['commands'])

    expect(result.status).toBe(0)
    expect(stopMarker).toMatchObject({
      active: true,
      overrideFlag: '--allow-active-stop-marker',
      overrideEnvVar: 'BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE',
    })
    expect(commands['discover']).toContain('BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE=1')
    expect(commands['discover']).toContain('--allow-active-stop-marker')
    expect(commands['fetch']).toContain('BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE=1')
    expect(commands['fetch']).toContain('--allow-active-stop-marker')
    expect(commands['fetchAll']).toContain('BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE=1')
    expect(commands['fetchAll']).toContain('--allow-active-stop-marker')
    expect(commands['resumePlan']).toContain('public-workbook-corpus:resume-financial:plan')
    expect(commands['resumePlan']).not.toContain('--allow-active-stop-marker')
    expect(commands['resumeCheck']).toContain('public-workbook-corpus:resume-financial:check')
    expect(commands['resumeCheck']).not.toContain('--allow-active-stop-marker')
    expect(commands['verify']).toContain('BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE=1')
    expect(commands['verify']).toContain('--allow-active-stop-marker')
    expect(commands['fetchPlan']).not.toContain('--allow-active-stop-marker')
    expect(commands['check']).not.toContain('--allow-active-stop-marker')
    expect(validatePublicWorkbookCorpusFinancialPlan(parseFinancialPlan(result.stdout))).toEqual([])
  })

  it('rejects inconsistent financial plan counts and unsafe command overrides', () => {
    const plan: PublicWorkbookCorpusFinancialPlan = {
      schemaVersion: 1,
      mode: 'plan',
      corpus: 'financial-accounting-workpapers',
      generatedAt: '2026-05-08T10:00:00.000Z',
      manifestExists: false,
      targetWorkbookCount: 5,
      manifestPath: '.cache/public-workbook-corpus-financial/manifest.json',
      cacheDir: '.cache/public-workbook-corpus-financial',
      scorecardPath: '.cache/public-workbook-corpus-financial/scorecard.json',
      verifyCheckpointPath: '.cache/public-workbook-corpus-financial/verification-checkpoint.json',
      stopMarker: {
        active: true,
        path: '.agent-coordination/stop.md',
        overrideFlag: '--allow-active-stop-marker',
        overrideEnvVar: 'BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE',
      },
      sourceCount: 0,
      targetArtifactCount: 5,
      cachedArtifactCount: 0,
      remainingArtifactSlots: 4,
      candidateSourceCount: 0,
      candidateSourceDeficitCount: 5,
      minimumAdditionalSourceCount: 5,
      recommendedDiscoveryLimit: 5,
      recommendedFetchTrancheSize: 20,
      recommendedFetchLimit: null,
      needsAdditionalDiscovery: true,
      targetReachableFromKnownCandidates: false,
      commands: {
        discoverPlan: 'pnpm public-workbook-corpus:discover-financial:plan -- --allow-active-stop-marker',
        discover: 'pnpm public-workbook-corpus:discover-financial',
        fetchPlan: 'pnpm public-workbook-corpus:fetch-financial:plan',
        fetch: null,
        fetchAll: 'pnpm public-workbook-corpus:fetch-financial',
        resumePlan: 'pnpm public-workbook-corpus:resume-financial:plan',
        resumeCheck: 'pnpm public-workbook-corpus:resume-financial:check',
        verify: 'pnpm public-workbook-corpus:verify-financial',
        check: 'pnpm public-workbook-corpus:check-financial',
      },
      sampledCandidateSources: [],
    }

    expect(validatePublicWorkbookCorpusFinancialPlan(plan)).toEqual(
      expect.arrayContaining([
        'remaining artifact slots do not match target and cached artifact counts',
        'recommended fetch limit is missing while artifacts remain',
        'bounded fetch command is missing while artifacts remain',
        expect.stringContaining('mutating command is missing stop-marker override'),
        expect.stringContaining('non-mutating command unexpectedly bypasses stop marker'),
      ]),
    )
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

function parseFinancialPlan(json: string): PublicWorkbookCorpusFinancialPlan {
  const value: unknown = JSON.parse(json)
  if (!isFinancialPlan(value)) {
    throw new Error('Expected financial corpus plan JSON')
  }
  return value
}

function isFinancialPlan(value: unknown): value is PublicWorkbookCorpusFinancialPlan {
  const record = asRecord(value)
  const stopMarker = asRecord(record['stopMarker'])
  const commands = asRecord(record['commands'])
  return (
    record['schemaVersion'] === 1 &&
    record['mode'] === 'plan' &&
    record['corpus'] === 'financial-accounting-workpapers' &&
    typeof record['generatedAt'] === 'string' &&
    typeof record['manifestExists'] === 'boolean' &&
    typeof record['targetWorkbookCount'] === 'number' &&
    typeof record['manifestPath'] === 'string' &&
    typeof record['cacheDir'] === 'string' &&
    typeof record['scorecardPath'] === 'string' &&
    typeof record['verifyCheckpointPath'] === 'string' &&
    typeof stopMarker['active'] === 'boolean' &&
    typeof stopMarker['path'] === 'string' &&
    typeof stopMarker['overrideFlag'] === 'string' &&
    typeof stopMarker['overrideEnvVar'] === 'string' &&
    typeof record['sourceCount'] === 'number' &&
    typeof record['targetArtifactCount'] === 'number' &&
    typeof record['cachedArtifactCount'] === 'number' &&
    typeof record['remainingArtifactSlots'] === 'number' &&
    typeof record['candidateSourceCount'] === 'number' &&
    typeof record['candidateSourceDeficitCount'] === 'number' &&
    typeof record['minimumAdditionalSourceCount'] === 'number' &&
    typeof record['recommendedDiscoveryLimit'] === 'number' &&
    typeof record['recommendedFetchTrancheSize'] === 'number' &&
    (typeof record['recommendedFetchLimit'] === 'number' || record['recommendedFetchLimit'] === null) &&
    typeof record['needsAdditionalDiscovery'] === 'boolean' &&
    typeof record['targetReachableFromKnownCandidates'] === 'boolean' &&
    (typeof commands['discoverPlan'] === 'string' || commands['discoverPlan'] === null) &&
    (typeof commands['discover'] === 'string' || commands['discover'] === null) &&
    typeof commands['fetchPlan'] === 'string' &&
    (typeof commands['fetch'] === 'string' || commands['fetch'] === null) &&
    typeof commands['fetchAll'] === 'string' &&
    typeof commands['resumePlan'] === 'string' &&
    typeof commands['resumeCheck'] === 'string' &&
    typeof commands['verify'] === 'string' &&
    typeof commands['check'] === 'string' &&
    Array.isArray(record['sampledCandidateSources'])
  )
}

function financialSource(id: string): PublicWorkbookSource {
  return {
    id,
    kind: 'direct-url',
    sourceUrl: `https://example.com/${id}.xlsx`,
    downloadUrl: `https://example.com/${id}.xlsx`,
    fileName: `${id}.xlsx`,
    discoveredAt: '2026-05-07T00:00:00.000Z',
    license: {
      spdxId: 'CC-BY-4.0',
      title: 'Creative Commons Attribution 4.0 International',
      evidenceUrl: 'https://creativecommons.org/licenses/by/4.0/',
    },
    topicEvidence: ['financial:fileName'],
  }
}
