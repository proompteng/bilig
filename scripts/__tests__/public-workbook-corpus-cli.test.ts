import { spawnSync } from 'node:child_process'
import { mkdtempSync, readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

import { buildPublicWorkbookCorpusScorecard, createEmptyPublicWorkbookManifest } from '../public-workbook-corpus.ts'
import { buildPublicWorkbookCorpusResumePlan, validatePublicWorkbookCorpusResumePlan } from '../public-workbook-corpus-resume-plan.ts'
import {
  readReusablePublicWorkbookCorpusCases,
  writePublicWorkbookCorpusVerificationCheckpoint,
} from '../public-workbook-corpus-verify-checkpoint.ts'
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

  it('refuses parallel fetch unless explicitly enabled for a sized host', () => {
    const env = { ...process.env }
    delete env.BILIG_ALLOW_PARALLEL_PUBLIC_CORPUS_FETCH

    const result = spawnSync('bun', [corpusScriptPath(), 'fetch', '--fetch-concurrency', '4'], {
      encoding: 'utf8',
      env,
    })

    expect(result.status).not.toBe(0)
    expect(result.stderr).toContain('--fetch-concurrency greater than 1 is disabled for public corpus CLI runs')
  })

  it('refuses broad corpus runs while a stop marker is active', () => {
    const dir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-cli-stop-marker-'))
    const stopMarkerPath = join(dir, 'stop.md')
    writeFileSync(stopMarkerPath, '# stop\n')
    const env = { ...process.env }
    delete env.BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE

    const result = spawnSync('bun', [corpusScriptPath(), 'fetch', '--corpus-run-stop-marker', stopMarkerPath], {
      encoding: 'utf8',
      env,
    })

    expect(result.status).not.toBe(0)
    expect(result.stderr).toContain('public-workbook-corpus fetch is disabled while the public corpus stop marker is active')
    expect(result.stderr).toContain('--allow-active-stop-marker')
    expect(result.stderr).toContain('BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE=1')
  })

  it('requires both stop-marker override flag and environment variable', () => {
    const dir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-cli-stop-marker-opt-in-'))
    const stopMarkerPath = join(dir, 'stop.md')
    writeFileSync(stopMarkerPath, '# stop\n')
    const env = { ...process.env }
    delete env.BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE

    const result = spawnSync(
      'bun',
      [corpusScriptPath(), 'fetch', '--corpus-run-stop-marker', stopMarkerPath, '--allow-active-stop-marker'],
      {
        encoding: 'utf8',
        env,
      },
    )

    expect(result.status).not.toBe(0)
    expect(result.stderr).toContain('public-workbook-corpus fetch is disabled while the public corpus stop marker is active')
  })

  it('plans fetch progress without starting network work', () => {
    const artifactA = workbookArtifact('workbook-a')
    const artifactB = workbookArtifact('workbook-b')
    const dir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-cli-fetch-plan-'))
    const manifestPath = join(dir, 'manifest.json')
    const manifest = {
      ...manifestWithArtifacts([artifactA, artifactB]),
      artifacts: [artifactA],
    }
    writeFileSync(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`)

    const result = spawnSync('bun', [corpusScriptPath(), 'fetch', '--dry-run', '--manifest', manifestPath, '--limit', '2'], {
      encoding: 'utf8',
    })
    const plan: unknown = JSON.parse(result.stdout)

    expect(result.status).toBe(0)
    expect(plan).toMatchObject({
      targetArtifactCount: 2,
      cachedArtifactCount: 1,
      sourceCount: 2,
      remainingArtifactSlots: 1,
      candidateSourceCount: 1,
      candidateSourceDeficitCount: 0,
      minimumAdditionalSourceCount: 0,
      recommendedDiscoveryLimit: 2,
      recommendedDiscoveryPlanCommand: 'pnpm public-workbook-corpus:discover:plan -- --limit 2',
      recommendedDiscoveryCommand: expect.stringContaining('pnpm public-workbook-corpus:discover --'),
      targetReachableFromKnownCandidates: true,
      sampledCandidateSources: [
        {
          id: artifactB.sourceId,
          fileName: artifactB.fileName,
          sourceUrl: artifactB.sourceUrl,
          downloadUrl: artifactB.downloadUrl,
        },
      ],
    })
  })

  it('plans source discovery when known candidates cannot reach the artifact target', () => {
    const artifactA = workbookArtifact('workbook-a')
    const artifactB = workbookArtifact('workbook-b')
    const dir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-cli-discover-plan-'))
    const manifestPath = join(dir, 'manifest.json')
    const manifest = {
      ...manifestWithArtifacts([artifactA, artifactB]),
      targetWorkbookCount: 4,
      artifacts: [artifactA],
    }
    writeFileSync(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`)

    const result = spawnSync('bun', [corpusScriptPath(), 'discover-plan', '--manifest', manifestPath, '--limit', '4'], {
      encoding: 'utf8',
    })
    const plan: unknown = JSON.parse(result.stdout)

    expect(result.status).toBe(0)
    expect(plan).toMatchObject({
      sourceCount: 2,
      targetArtifactCount: 4,
      cachedArtifactCount: 1,
      remainingArtifactSlots: 3,
      candidateSourceCount: 1,
      candidateSourceDeficitCount: 2,
      minimumAdditionalSourceCount: 2,
      recommendedDiscoveryLimit: 4,
      recommendedDiscoveryCommand: expect.stringContaining('pnpm public-workbook-corpus:discover --'),
      targetReachableFromKnownCandidates: false,
    })
  })

  it('builds a stop-marker-aware bounded resume plan for the remaining corpus evidence', () => {
    const plan = buildPublicWorkbookCorpusResumePlan({
      cacheDir: '/repo/.cache/public-workbook-corpus',
      fetchBatchSize: 6,
      fetchLimit: 10_000,
      fetchPlan: {
        candidateSourceCount: 2_389,
        candidateSourceDeficitCount: 1_983,
        recommendedDiscoveryLimit: 11_983,
        remainingArtifactSlots: 4_372,
        targetReachableFromKnownCandidates: false,
      },
      generatedAt: '2026-05-07T08:00:00.000Z',
      manifestPath: '/repo/.cache/public-workbook-corpus/manifest.json',
      scorecardPath: '/repo/packages/benchmarks/baselines/public-workbook-corpus-scorecard.json',
      status: {
        targetWorkbookCount: 10_000,
        cachedArtifactCount: 5_628,
        recordedManifestArtifactCount: 4_940,
        missingManifestArtifactCount: 688,
        recordedAllCasesPassed: true,
      },
      stopMarkerActive: true,
      stopMarkerPath: '/repo/.agent-coordination/stop.md',
      verifyBatchSize: 20,
      verifyCheckpointPath: '/repo/.cache/public-workbook-corpus/verification-checkpoint.json',
    })

    expect(plan).toMatchObject({
      schemaVersion: 1,
      stopMarker: {
        active: true,
        requiresExplicitResume: true,
        overrideFlag: '--allow-active-stop-marker',
        overrideEnvVar: 'BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE',
      },
      currentState: {
        cachedArtifactCount: 5_628,
        recordedManifestArtifactCount: 4_940,
        missingCachedArtifactCount: 4_372,
        missingVerificationCount: 688,
      },
      phases: {
        verifyMissingCachedArtifacts: {
          status: 'blocked-by-stop-marker',
          totalWorkItems: 688,
          batchSize: 20,
          batchCount: 35,
          commands: expect.arrayContaining([
            expect.stringContaining('public-workbook-corpus:verify-missing:plan'),
            expect.stringContaining('BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE=1 pnpm public-workbook-corpus:verify-missing'),
          ]),
        },
        discoverAdditionalSources: {
          status: 'blocked-by-stop-marker',
          totalWorkItems: 1_983,
          batchCount: 1,
          commands: expect.arrayContaining([expect.stringContaining('--limit 11983')]),
        },
        fetchAdditionalArtifacts: {
          status: 'blocked-by-stop-marker',
          totalWorkItems: 4_372,
          batchSize: 6,
          batchCount: 729,
        },
        finalEvidenceRefresh: {
          status: 'blocked-by-stop-marker',
          commands: expect.arrayContaining([
            expect.stringContaining('public-workbook-corpus:verify'),
            'pnpm public-workbook-corpus:completion-audit:check -- --require-complete',
            'pnpm dominance:generate',
            'pnpm dominance:check',
          ]),
        },
      },
    })
    expect(validatePublicWorkbookCorpusResumePlan(plan)).toEqual([])
  })

  it('checks a resume plan from the checked-in scorecard when the local manifest cache is absent', async () => {
    const artifactA = workbookArtifact('workbook-a')
    const artifactB = workbookArtifact('workbook-b')
    const dir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-resume-plan-scorecard-'))
    const scorecardPath = join(dir, 'scorecard.json')
    const missingManifestPath = join(dir, 'missing-manifest.json')
    const checkpointPath = join(dir, 'verification-checkpoint.json')
    const manifest: PublicWorkbookManifest = {
      ...manifestWithArtifacts([artifactA, artifactB]),
      targetWorkbookCount: 2,
    }
    const scorecard = await buildPublicWorkbookCorpusScorecard({
      manifest,
      cacheDir: dir,
      generatedAt: '2026-05-07T01:00:00.000Z',
      reusableCases: [passedCase(artifactA), passedCase(artifactB)],
    })
    writeFileSync(scorecardPath, `${JSON.stringify(scorecard, null, 2)}\n`)

    const result = spawnSync(
      'bun',
      [
        resumePlanScriptPath(),
        '--check',
        '--manifest',
        missingManifestPath,
        '--scorecard',
        scorecardPath,
        '--verify-checkpoint',
        checkpointPath,
        '--cache-dir',
        dir,
        '--fetch-limit',
        '2',
      ],
      { encoding: 'utf8' },
    )
    const output: unknown = JSON.parse(result.stdout)

    expect(result.status).toBe(0)
    expect(output).toMatchObject({
      mode: 'check',
      currentState: {
        targetWorkbookCount: 2,
        cachedArtifactCount: 2,
        recordedManifestArtifactCount: 2,
        missingCachedArtifactCount: 0,
        missingVerificationCount: 0,
      },
      phases: {
        discoverAdditionalSources: { status: 'not-needed' },
        fetchAdditionalArtifacts: { status: 'not-needed' },
      },
    })
  })

  it('rejects unsafe resume plans that hide active stop-marker overrides', () => {
    const plan = buildPublicWorkbookCorpusResumePlan({
      cacheDir: '/repo/.cache/public-workbook-corpus',
      fetchBatchSize: 6,
      fetchLimit: 10_000,
      fetchPlan: {
        candidateSourceCount: 1,
        candidateSourceDeficitCount: 0,
        recommendedDiscoveryLimit: 10_000,
        remainingArtifactSlots: 1,
        targetReachableFromKnownCandidates: true,
      },
      generatedAt: '2026-05-07T08:00:00.000Z',
      manifestPath: '/repo/.cache/public-workbook-corpus/manifest.json',
      scorecardPath: '/repo/packages/benchmarks/baselines/public-workbook-corpus-scorecard.json',
      status: {
        targetWorkbookCount: 10_000,
        cachedArtifactCount: 9_999,
        recordedManifestArtifactCount: 9_998,
        missingManifestArtifactCount: 1,
        recordedAllCasesPassed: true,
      },
      stopMarkerActive: true,
      stopMarkerPath: '/repo/.agent-coordination/stop.md',
      verifyBatchSize: 20,
      verifyCheckpointPath: '/repo/.cache/public-workbook-corpus/verification-checkpoint.json',
    })
    const invalidPlan = {
      ...plan,
      phases: {
        ...plan.phases,
        fetchAdditionalArtifacts: Object.assign({}, plan.phases.fetchAdditionalArtifacts, {
          commands: ['pnpm public-workbook-corpus:fetch -- --limit 10000'],
        }),
      },
    }

    expect(validatePublicWorkbookCorpusResumePlan(invalidPlan)).toEqual(
      expect.arrayContaining([
        'fetchAdditionalArtifacts mutating command is missing the explicit stop-marker override: pnpm public-workbook-corpus:fetch -- --limit 10000',
      ]),
    )
  })

  it('rejects stale resume plans with impossible current-state counts', () => {
    const plan = buildPublicWorkbookCorpusResumePlan({
      cacheDir: '/repo/.cache/public-workbook-corpus',
      fetchBatchSize: 6,
      fetchLimit: 10_000,
      fetchPlan: {
        candidateSourceCount: 6,
        candidateSourceDeficitCount: 0,
        recommendedDiscoveryLimit: 10_000,
        remainingArtifactSlots: 6,
        targetReachableFromKnownCandidates: true,
      },
      generatedAt: '2026-05-07T08:00:00.000Z',
      manifestPath: '/repo/.cache/public-workbook-corpus/manifest.json',
      scorecardPath: '/repo/packages/benchmarks/baselines/public-workbook-corpus-scorecard.json',
      status: {
        targetWorkbookCount: 10_000,
        cachedArtifactCount: 9_994,
        recordedManifestArtifactCount: 9_994,
        missingManifestArtifactCount: 0,
        recordedAllCasesPassed: true,
      },
      stopMarkerActive: false,
      stopMarkerPath: '/repo/.agent-coordination/stop.md',
      verifyBatchSize: 20,
      verifyCheckpointPath: '/repo/.cache/public-workbook-corpus/verification-checkpoint.json',
    })
    const invalidPlan = {
      ...plan,
      currentState: {
        ...plan.currentState,
        missingCachedArtifactCount: 7,
        missingVerificationCount: -1,
        recordedManifestArtifactCount: 9_995,
      },
    }

    expect(validatePublicWorkbookCorpusResumePlan(invalidPlan)).toEqual(
      expect.arrayContaining([
        'missing verification count must be non-negative',
        'missing cached artifact count is 7, expected 6',
        'recorded verification count exceeds cached artifact count',
      ]),
    )
  })

  it('refuses large verify-missing tranches unless explicitly enabled', () => {
    const env = { ...process.env }
    delete env.BILIG_ALLOW_LARGE_PUBLIC_CORPUS_VERIFY_MISSING

    const result = spawnSync('bun', [corpusScriptPath(), 'verify-missing', '--limit', '21'], {
      encoding: 'utf8',
      env,
    })

    expect(result.status).not.toBe(0)
    expect(result.stderr).toContain('--limit above 20 is disabled for public corpus verify-missing runs')
  })

  it('refuses high-RSS verification on interactive corpus runs', () => {
    const env = { ...process.env }
    delete env.BILIG_ALLOW_HIGH_RSS_PUBLIC_CORPUS_VERIFY

    const result = spawnSync('bun', [corpusScriptPath(), 'verify', '--verify-max-rss-mb', '4096'], {
      encoding: 'utf8',
      env,
    })

    expect(result.status).not.toBe(0)
    expect(result.stderr).toContain('RSS limits above 1536 MiB are disabled because')
  })

  it('refuses inherited high-RSS environment for interactive corpus runs', () => {
    const env = { ...process.env, BILIG_ALLOW_HIGH_RSS_PUBLIC_CORPUS_VERIFY: '1' }

    const result = spawnSync('bun', [corpusScriptPath(), 'verify', '--verify-max-rss-mb', '4096'], {
      encoding: 'utf8',
      env,
    })

    expect(result.status).not.toBe(0)
    expect(result.stderr).toContain('RSS limits above 1536 MiB are disabled because')
  })

  it('fails check when the local manifest has cached artifacts missing from the scorecard', async () => {
    const artifactA = workbookArtifact('workbook-a')
    const artifactB = workbookArtifact('workbook-b')
    const dir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-cli-check-'))
    const manifestPath = join(dir, 'manifest.json')
    const scorecardPath = join(dir, 'scorecard.json')
    const checkpointPath = join(dir, 'verification-checkpoint.json')
    const scorecard = await buildPublicWorkbookCorpusScorecard({
      manifest: manifestWithArtifacts([artifactA]),
      cacheDir: dir,
      generatedAt: '2026-05-07T01:00:00.000Z',
      reusableCases: [passedCase(artifactA)],
    })
    writeFileSync(manifestPath, `${JSON.stringify(manifestWithArtifacts([artifactA, artifactB]), null, 2)}\n`)
    writeFileSync(scorecardPath, `${JSON.stringify(scorecard, null, 2)}\n`)

    const result = spawnSync(
      'bun',
      [
        corpusScriptPath(),
        'check',
        '--manifest',
        manifestPath,
        '--scorecard',
        scorecardPath,
        '--verify-checkpoint',
        checkpointPath,
        '--cache-dir',
        dir,
        '--corpus-run-stop-marker',
        join(dir, 'inactive-stop-marker.md'),
      ],
      {
        encoding: 'utf8',
      },
    )

    expect(result.status).not.toBe(0)
    expect(result.stderr).toContain('Public workbook corpus verification incomplete')
    expect(result.stderr).toContain('recorded verification cases below cached artifacts: 1/2')
    expect(result.stderr).toContain('next command: pnpm public-workbook-corpus:verify-missing')
  })

  it('passes check when checkpoint evidence covers a stale scorecard', async () => {
    const artifactA = workbookArtifact('workbook-a')
    const artifactB = workbookArtifact('workbook-b')
    const dir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-cli-check-checkpoint-'))
    const manifestPath = join(dir, 'manifest.json')
    const scorecardPath = join(dir, 'scorecard.json')
    const checkpointPath = join(dir, 'verification-checkpoint.json')
    const fullManifest = manifestWithArtifacts([artifactA, artifactB])
    const scorecard = await buildPublicWorkbookCorpusScorecard({
      manifest: manifestWithArtifacts([artifactA]),
      cacheDir: dir,
      generatedAt: '2026-05-07T01:00:00.000Z',
      reusableCases: [passedCase(artifactA)],
    })
    writeFileSync(manifestPath, `${JSON.stringify(fullManifest, null, 2)}\n`)
    writeFileSync(scorecardPath, `${JSON.stringify(scorecard, null, 2)}\n`)
    writePublicWorkbookCorpusVerificationCheckpoint({
      path: checkpointPath,
      manifest: fullManifest,
      casesById: new Map([
        [artifactA.id, passedCase(artifactA)],
        [artifactB.id, passedCase(artifactB)],
      ]),
      generatedAt: '2026-05-07T01:30:00.000Z',
    })

    const result = spawnSync(
      'bun',
      [corpusScriptPath(), 'check', '--manifest', manifestPath, '--scorecard', scorecardPath, '--verify-checkpoint', checkpointPath],
      { encoding: 'utf8' },
    )

    expect(result.status).toBe(0)
    expect(result.stdout).toContain('Checked public workbook corpus with 2/2 recorded cached workbooks')
  })

  it('reports manifest, scorecard, and checkpoint progress without starting corpus jobs', async () => {
    const artifactA = workbookArtifact('workbook-a')
    const artifactB = workbookArtifact('workbook-b')
    const dir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-cli-status-'))
    const manifestPath = join(dir, 'manifest.json')
    const scorecardPath = join(dir, 'scorecard.json')
    const checkpointPath = join(dir, 'verification-checkpoint.json')
    const fullManifest = manifestWithArtifacts([artifactA, artifactB])
    const scorecard = await buildPublicWorkbookCorpusScorecard({
      manifest: manifestWithArtifacts([artifactA]),
      cacheDir: dir,
      generatedAt: '2026-05-07T01:00:00.000Z',
      reusableCases: [passedCase(artifactA)],
    })
    writeFileSync(manifestPath, `${JSON.stringify(fullManifest, null, 2)}\n`)
    writeFileSync(scorecardPath, `${JSON.stringify(scorecard, null, 2)}\n`)
    writePublicWorkbookCorpusVerificationCheckpoint({
      path: checkpointPath,
      manifest: fullManifest,
      casesById: new Map([
        [artifactA.id, passedCase(artifactA)],
        [artifactB.id, passedCase(artifactB)],
      ]),
      generatedAt: '2026-05-07T01:30:00.000Z',
    })

    const result = spawnSync(
      'bun',
      [corpusScriptPath(), 'status', '--manifest', manifestPath, '--scorecard', scorecardPath, '--verify-checkpoint', checkpointPath],
      { encoding: 'utf8' },
    )
    const status: unknown = JSON.parse(result.stdout)

    expect(result.status).toBe(0)
    expect(status).toMatchObject({
      targetWorkbookCount: 10000,
      sourceCount: 2,
      cachedArtifactCount: 2,
      scorecardCaseCount: 1,
      checkpointCaseCount: 2,
      recordedManifestArtifactCount: 2,
      missingManifestArtifactCount: 0,
      recordedPassedCaseCount: 2,
      recordedUnsupportedCaseCount: 0,
      recordedFailedCaseCount: 0,
      recordedErrorCaseCount: 0,
      missingManifestArtifactSample: [],
      nextMissingVerificationCommand: null,
      nextMissingVerificationPlanCommand: null,
      scorecardCoversManifest: false,
      targetComplete: false,
      gaps: expect.arrayContaining(['cached artifacts below target: 2/10000', 'scorecard cases do not cover manifest artifacts: 1/2']),
    })
  })

  it('reports a bounded sample of cached artifacts missing verification evidence', async () => {
    const artifactA = workbookArtifact('workbook-a')
    const artifactB = workbookArtifact('workbook-b')
    const dir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-cli-status-missing-'))
    const manifestPath = join(dir, 'manifest.json')
    const scorecardPath = join(dir, 'missing-scorecard.json')
    const checkpointPath = join(dir, 'verification-checkpoint.json')
    const fullManifest = manifestWithArtifacts([artifactA, artifactB])
    writeFileSync(manifestPath, `${JSON.stringify(fullManifest, null, 2)}\n`)
    writePublicWorkbookCorpusVerificationCheckpoint({
      path: checkpointPath,
      manifest: fullManifest,
      casesById: new Map([[artifactA.id, passedCase(artifactA)]]),
      generatedAt: '2026-05-07T01:30:00.000Z',
    })

    const result = spawnSync(
      'bun',
      [
        corpusScriptPath(),
        'status',
        '--manifest',
        manifestPath,
        '--scorecard',
        scorecardPath,
        '--verify-checkpoint',
        checkpointPath,
        '--cache-dir',
        dir,
      ],
      {
        encoding: 'utf8',
      },
    )
    const status: unknown = JSON.parse(result.stdout)

    expect(result.status).toBe(0)
    expect(status).toMatchObject({
      cachedArtifactCount: 2,
      scorecardCaseCount: 0,
      checkpointCaseCount: 1,
      recordedManifestArtifactCount: 1,
      missingManifestArtifactCount: 1,
      missingManifestArtifactSample: [
        {
          id: artifactB.id,
          fileName: artifactB.fileName,
          byteSize: artifactB.byteSize,
          sourceUrl: artifactB.sourceUrl,
        },
      ],
      nextMissingVerificationCommand: expect.stringContaining('public-workbook-corpus:verify-missing'),
      nextMissingVerificationPlanCommand: expect.stringContaining('public-workbook-corpus:verify-missing:plan'),
    })
    expect(status).toMatchObject({
      nextMissingVerificationCommand: expect.stringContaining(`--manifest ${manifestPath}`),
      nextMissingVerificationPlanCommand: expect.stringContaining(`--verify-checkpoint ${checkpointPath}`),
    })
    expect(status).toMatchObject({
      nextMissingVerificationCommand: expect.stringContaining(`--cache-dir ${dir}`),
      nextMissingVerificationPlanCommand: expect.stringContaining(`--cache-dir ${dir}`),
    })
    expect(status).toMatchObject({
      nextMissingVerificationCommand: expect.stringContaining('--limit 1'),
    })
  })

  it('guards status suggested verification commands while a stop marker is active', async () => {
    const artifactA = workbookArtifact('workbook-a')
    const artifactB = workbookArtifact('workbook-b')
    const dir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-cli-status-missing-stopped-'))
    const manifestPath = join(dir, 'manifest.json')
    const scorecardPath = join(dir, 'missing-scorecard.json')
    const checkpointPath = join(dir, 'verification-checkpoint.json')
    const stopMarkerPath = join(dir, 'stop.md')
    const fullManifest = manifestWithArtifacts([artifactA, artifactB])
    writeFileSync(manifestPath, `${JSON.stringify(fullManifest, null, 2)}\n`)
    writeFileSync(stopMarkerPath, '# paused\n')
    writePublicWorkbookCorpusVerificationCheckpoint({
      path: checkpointPath,
      manifest: fullManifest,
      casesById: new Map([[artifactA.id, passedCase(artifactA)]]),
      generatedAt: '2026-05-07T01:30:00.000Z',
    })

    const result = spawnSync(
      'bun',
      [
        corpusScriptPath(),
        'status',
        '--manifest',
        manifestPath,
        '--scorecard',
        scorecardPath,
        '--verify-checkpoint',
        checkpointPath,
        '--cache-dir',
        dir,
        '--corpus-run-stop-marker',
        stopMarkerPath,
      ],
      {
        encoding: 'utf8',
      },
    )
    const status: unknown = JSON.parse(result.stdout)

    expect(result.status).toBe(0)
    expect(status).toMatchObject({
      nextMissingVerificationCommand: expect.stringContaining('BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE=1'),
      nextMissingVerificationPlanCommand: expect.not.stringContaining('BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE=1'),
    })
    expect(status).toMatchObject({
      nextMissingVerificationCommand: expect.stringContaining('--allow-active-stop-marker'),
      nextMissingVerificationPlanCommand: expect.stringContaining('public-workbook-corpus:verify-missing:plan'),
    })
  })

  it('verifies a bounded missing slice into the checkpoint without rewriting the scorecard', async () => {
    const artifactA = workbookArtifact('workbook-a')
    const artifactB = workbookArtifact('workbook-b')
    const dir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-cli-verify-missing-'))
    const manifestPath = join(dir, 'manifest.json')
    const scorecardPath = join(dir, 'scorecard.json')
    const checkpointPath = join(dir, 'verification-checkpoint.json')
    const fullManifest = manifestWithArtifacts([artifactA, artifactB])
    const scorecard = await buildPublicWorkbookCorpusScorecard({
      manifest: manifestWithArtifacts([artifactA]),
      cacheDir: dir,
      generatedAt: '2026-05-07T01:00:00.000Z',
      reusableCases: [passedCase(artifactA)],
    })
    writeFileSync(manifestPath, `${JSON.stringify(fullManifest, null, 2)}\n`)
    writeFileSync(scorecardPath, `${JSON.stringify(scorecard, null, 2)}\n`)

    const result = spawnSync(
      'bun',
      [
        corpusScriptPath(),
        'verify-missing',
        '--manifest',
        manifestPath,
        '--scorecard',
        scorecardPath,
        '--verify-checkpoint',
        checkpointPath,
        '--cache-dir',
        dir,
        '--corpus-run-stop-marker',
        join(dir, 'inactive-stop-marker.md'),
        '--limit',
        '1',
        '--in-process',
      ],
      {
        encoding: 'utf8',
        env: { ...process.env, BILIG_ALLOW_IN_PROCESS_PUBLIC_CORPUS_VERIFY: '1' },
      },
    )
    const checkpointCases = readReusablePublicWorkbookCorpusCases([checkpointPath])

    expect(result.status).toBe(0)
    expect(result.stdout).toContain('Verified 1 missing public workbook cases')
    expect(checkpointCases.map((entry) => [entry.id, entry.status])).toEqual([
      [artifactA.id, 'passed'],
      [artifactB.id, 'error'],
    ])
    expect(checkpointCases[1]?.evidence).toEqual(expect.arrayContaining([`Missing cached workbook file: ${artifactB.cachePath}`]))
  })

  it('lists a bounded missing slice without starting verification', async () => {
    const artifactA = workbookArtifact('workbook-a')
    const artifactB = workbookArtifact('workbook-b')
    const dir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-cli-verify-missing-plan-'))
    const manifestPath = join(dir, 'manifest.json')
    const scorecardPath = join(dir, 'scorecard.json')
    const checkpointPath = join(dir, 'verification-checkpoint.json')
    const fullManifest = manifestWithArtifacts([artifactA, artifactB])
    const scorecard = await buildPublicWorkbookCorpusScorecard({
      manifest: manifestWithArtifacts([artifactA]),
      cacheDir: dir,
      generatedAt: '2026-05-07T01:00:00.000Z',
      reusableCases: [passedCase(artifactA)],
    })
    writeFileSync(manifestPath, `${JSON.stringify(fullManifest, null, 2)}\n`)
    writeFileSync(scorecardPath, `${JSON.stringify(scorecard, null, 2)}\n`)

    const result = spawnSync(
      'bun',
      [
        corpusScriptPath(),
        'verify-missing',
        '--dry-run',
        '--manifest',
        manifestPath,
        '--scorecard',
        scorecardPath,
        '--verify-checkpoint',
        checkpointPath,
        '--limit',
        '1',
      ],
      {
        encoding: 'utf8',
      },
    )
    const planned: unknown = JSON.parse(result.stdout)

    expect(result.status).toBe(0)
    expect(planned).toMatchObject({
      totalMissingArtifactCount: 1,
      selectedArtifactCount: 1,
      artifacts: [
        {
          id: artifactB.id,
          fileName: artifactB.fileName,
          cachePath: artifactB.cachePath,
        },
      ],
    })
    expect(readReusablePublicWorkbookCorpusCases([checkpointPath])).toEqual([])
  })

  it('refreshes the checked-in scorecard from existing checkpoint cases without verification workers', async () => {
    const artifactA = workbookArtifact('workbook-a')
    const artifactB = workbookArtifact('workbook-b')
    const dir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-cli-refresh-scorecard-'))
    const manifestPath = join(dir, 'manifest.json')
    const scorecardPath = join(dir, 'scorecard.json')
    const checkpointPath = join(dir, 'verification-checkpoint.json')
    const manifest = manifestWithArtifacts([artifactA, artifactB])
    const staleScorecard = await buildPublicWorkbookCorpusScorecard({
      manifest: manifestWithArtifacts([artifactA]),
      cacheDir: dir,
      generatedAt: '2026-05-07T01:00:00.000Z',
      reusableCases: [passedCase(artifactA)],
    })
    writeFileSync(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`)
    writeFileSync(scorecardPath, `${JSON.stringify(staleScorecard, null, 2)}\n`)
    writePublicWorkbookCorpusVerificationCheckpoint({
      path: checkpointPath,
      manifest,
      casesById: new Map([
        [artifactA.id, passedCase(artifactA)],
        [artifactB.id, passedCase(artifactB)],
      ]),
      generatedAt: '2026-05-07T02:00:00.000Z',
    })

    const checkBefore = spawnSync(
      'bun',
      [
        corpusScriptPath(),
        'refresh-scorecard-from-checkpoint',
        '--check',
        '--manifest',
        manifestPath,
        '--scorecard',
        scorecardPath,
        '--verify-checkpoint',
        checkpointPath,
      ],
      { encoding: 'utf8' },
    )
    expect(checkBefore.status).not.toBe(0)
    expect(checkBefore.stderr).toContain('Public workbook corpus scorecard is stale: 2 checkpoint-backed cases are available')

    const writeResult = spawnSync(
      'bun',
      [
        corpusScriptPath(),
        'refresh-scorecard-from-checkpoint',
        '--manifest',
        manifestPath,
        '--scorecard',
        scorecardPath,
        '--verify-checkpoint',
        checkpointPath,
      ],
      { encoding: 'utf8' },
    )
    const summary: unknown = JSON.parse(writeResult.stdout)
    const refreshed: unknown = JSON.parse(readFileSync(scorecardPath, 'utf8'))

    expect(writeResult.status).toBe(0)
    expect(summary).toMatchObject({
      mode: 'write',
      cachedWorkbookCount: 2,
      passedWorkbookCount: 2,
      remainingToTarget: 9_998,
    })
    expect(refreshed).toMatchObject({
      summary: {
        cachedWorkbookCount: 2,
        passedWorkbookCount: 2,
        sourceCount: 2,
      },
      cases: [{ id: artifactA.id }, { id: artifactB.id }],
    })

    const checkAfter = spawnSync(
      'bun',
      [
        corpusScriptPath(),
        'refresh-scorecard-from-checkpoint',
        '--check',
        '--manifest',
        manifestPath,
        '--scorecard',
        scorecardPath,
        '--verify-checkpoint',
        checkpointPath,
      ],
      { encoding: 'utf8' },
    )
    expect(checkAfter.status).toBe(0)
  })

  it('exposes safe verify-missing package scripts', () => {
    const packageJson = readPackageJson()

    expect(packageJson.scripts?.['public-workbook-corpus:link-plan']).toBe('bun scripts/public-workbook-corpus.ts link-plan')
    expect(packageJson.scripts?.['public-workbook-corpus:link-plan:check']).toContain(
      'bun scripts/public-workbook-corpus.ts link-plan --source-url',
    )
    expect(packageJson.scripts?.['public-workbook-corpus:add-link']).toBe('bun scripts/public-workbook-corpus.ts add-link')
    expect(packageJson.scripts?.['public-workbook-corpus:add-link:check']).toContain(
      'bun scripts/public-workbook-corpus.ts add-link --dry-run',
    )
    expect(packageJson.scripts?.['public-workbook-corpus:fetch:plan']).toBe(
      'bun scripts/public-workbook-corpus.ts fetch --dry-run --sample-limit 20',
    )
    expect(packageJson.scripts?.['public-workbook-corpus:fetch-source']).toBe('bun scripts/public-workbook-corpus.ts fetch-source')
    expect(packageJson.scripts?.['public-workbook-corpus:verify-artifact']).toBe('bun scripts/public-workbook-corpus.ts verify-artifact')
    expect(packageJson.scripts?.['public-workbook-corpus:discover:plan']).toBe('bun scripts/public-workbook-corpus.ts discover-plan')
    expect(packageJson.scripts?.['public-workbook-corpus:verify-missing']).toBe('bun scripts/public-workbook-corpus.ts verify-missing')
    expect(packageJson.scripts?.['public-workbook-corpus:verify-missing:plan']).toBe(
      'bun scripts/public-workbook-corpus.ts verify-missing --dry-run --limit 20',
    )
    expect(packageJson.scripts?.['public-workbook-corpus:refresh-scorecard-from-checkpoint']).toBe(
      'bun scripts/public-workbook-corpus.ts refresh-scorecard-from-checkpoint',
    )
    expect(packageJson.scripts?.['public-workbook-corpus:refresh-scorecard-from-checkpoint:check']).toBe(
      'bun scripts/public-workbook-corpus.ts refresh-scorecard-from-checkpoint --check',
    )
    expect(packageJson.scripts?.['public-workbook-corpus:resume-plan']).toBe('bun scripts/public-workbook-corpus-resume-plan.ts')
    expect(packageJson.scripts?.['public-workbook-corpus:resume-plan:check']).toBe(
      'bun scripts/public-workbook-corpus-resume-plan.ts --check',
    )
  })
})

function readPackageJson(): { readonly scripts?: Record<string, string> } {
  const parsed: unknown = JSON.parse(readFileSync(join(dirname(fileURLToPath(import.meta.url)), '../..', 'package.json'), 'utf8'))
  if (
    typeof parsed !== 'object' ||
    parsed === null ||
    (Reflect.get(parsed, 'scripts') !== undefined &&
      (typeof Reflect.get(parsed, 'scripts') !== 'object' || Reflect.get(parsed, 'scripts') === null))
  ) {
    throw new Error('package.json scripts payload was not an object')
  }
  const scripts = Reflect.get(parsed, 'scripts')
  if (!scripts) {
    return {}
  }
  return { scripts: Object.fromEntries(Object.entries(scripts).filter((entry): entry is [string, string] => typeof entry[1] === 'string')) }
}

function corpusScriptPath(): string {
  return join(dirname(fileURLToPath(import.meta.url)), '../public-workbook-corpus.ts')
}

function resumePlanScriptPath(): string {
  return join(dirname(fileURLToPath(import.meta.url)), '../public-workbook-corpus-resume-plan.ts')
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
