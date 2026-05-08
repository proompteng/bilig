import { spawnSync } from 'node:child_process'
import { mkdtempSync, readFileSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

import { buildPublicWorkbookCorpusScorecard, createEmptyPublicWorkbookManifest } from '../public-workbook-corpus.ts'
import { formatPublicCorpusStopMarkerPathForMessage } from '../public-workbook-corpus-cli.ts'
import { asRecord } from '../public-workbook-corpus-json.ts'
import { buildPublicWorkbookCorpusResumePlan, validatePublicWorkbookCorpusResumePlan } from '../public-workbook-corpus-resume-plan.ts'
import { publicWorkbookImportWarningClassifierEvidence } from '../public-workbook-corpus-evidence.ts'
import { buildPublicWorkbookCorpusStatus } from '../public-workbook-corpus-status.ts'
import {
  readReusablePublicWorkbookCorpusCases,
  writePublicWorkbookCorpusVerificationCheckpoint,
} from '../public-workbook-corpus-verify-checkpoint.ts'
import type { PublicWorkbookArtifact, PublicWorkbookCorpusCase, PublicWorkbookManifest } from '../public-workbook-corpus-types.ts'

describe('public workbook corpus CLI resource guards', () => {
  it('refuses to reinitialize an existing manifest unless forced', () => {
    const dir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-cli-init-existing-'))
    const manifestPath = join(dir, 'manifest.json')
    const existingManifest = createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z', 12_345)
    writeFileSync(manifestPath, `${JSON.stringify(existingManifest, null, 2)}\n`)

    const result = spawnSync('bun', [corpusScriptPath(), 'init', '--manifest', manifestPath, '--cache-dir', dir], {
      encoding: 'utf8',
    })
    const manifestAfterBlockedInit: unknown = JSON.parse(readFileSync(manifestPath, 'utf8'))

    expect(result.status).not.toBe(0)
    expect(result.stderr).toContain('Public workbook corpus manifest already exists')
    expect(result.stderr).toContain('pass --force to reinitialize it')
    expect(manifestAfterBlockedInit).toMatchObject({
      generatedAt: '2026-05-07T00:00:00.000Z',
      targetWorkbookCount: 12_345,
    })
  })

  it('allows explicit forced manifest reinitialization', () => {
    const dir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-cli-init-force-'))
    const manifestPath = join(dir, 'manifest.json')
    writeFileSync(manifestPath, `${JSON.stringify(createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z', 12_345), null, 2)}\n`)

    const result = spawnSync(
      'bun',
      [corpusScriptPath(), 'init', '--manifest', manifestPath, '--cache-dir', dir, '--target-workbook-count', '9', '--force'],
      {
        encoding: 'utf8',
      },
    )
    const manifestAfterForcedInit: unknown = JSON.parse(readFileSync(manifestPath, 'utf8'))

    expect(result.status).toBe(0)
    expect(manifestAfterForcedInit).toMatchObject({
      targetWorkbookCount: 9,
      sources: [],
      artifacts: [],
    })
  })

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

  it('validates the fetch RSS guard limit before running a mutating fetch', () => {
    const dir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-cli-fetch-rss-'))
    const manifestPath = join(dir, 'manifest.json')
    writeFileSync(manifestPath, `${JSON.stringify(createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z'), null, 2)}\n`)

    const result = spawnSync(
      'bun',
      [
        corpusScriptPath(),
        'fetch',
        '--manifest',
        manifestPath,
        '--cache-dir',
        dir,
        '--fetch-max-rss-mb',
        '0',
        '--corpus-run-stop-marker',
        join(dir, 'inactive-stop-marker.md'),
      ],
      {
        encoding: 'utf8',
      },
    )

    expect(result.status).not.toBe(0)
    expect(result.stderr).toContain('Expected --fetch-max-rss-mb to be a positive number of MiB')
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

  it('refuses broad corpus discovery while a stop marker is active', () => {
    const dir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-cli-discover-stop-marker-'))
    const stopMarkerPath = join(dir, 'stop.md')
    writeFileSync(stopMarkerPath, '# stop\n')
    const env = { ...process.env }
    delete env.BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE

    const genericDiscovery = spawnSync(
      'bun',
      [corpusScriptPath(), 'discover-ckan', '--corpus-run-stop-marker', stopMarkerPath, '--ckan-base', 'https://example.invalid/api'],
      {
        encoding: 'utf8',
        env,
      },
    )
    const financialDiscovery = spawnSync(
      'bun',
      [
        corpusScriptPath(),
        'discover-financial-ckan',
        '--corpus-run-stop-marker',
        stopMarkerPath,
        '--ckan-base',
        'https://example.invalid/api',
      ],
      {
        encoding: 'utf8',
        env,
      },
    )

    expect(genericDiscovery.status).not.toBe(0)
    expect(genericDiscovery.stderr).toContain('public-workbook-corpus discover is disabled while the public corpus stop marker is active')
    expect(financialDiscovery.status).not.toBe(0)
    expect(financialDiscovery.stderr).toContain(
      'public-workbook-corpus discover-financial is disabled while the public corpus stop marker is active',
    )
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

  it('refuses checkpoint-updating artifact verification while a stop marker is active', () => {
    const dir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-cli-verify-artifact-stop-marker-'))
    const stopMarkerPath = join(dir, 'stop.md')
    writeFileSync(stopMarkerPath, '# stop\n')
    const env = { ...process.env }
    delete env.BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE

    const result = spawnSync(
      'bun',
      [
        corpusScriptPath(),
        'verify-artifact',
        '--artifact-id',
        'workbook-abc123',
        '--update-verify-checkpoint',
        '--corpus-run-stop-marker',
        stopMarkerPath,
      ],
      {
        encoding: 'utf8',
        env,
      },
    )

    expect(result.status).not.toBe(0)
    expect(result.stderr).toContain('public-workbook-corpus verify-artifact is disabled while the public corpus stop marker is active')
  })

  it('formats repo-local stop marker paths without exposing the checkout root', () => {
    expect(formatPublicCorpusStopMarkerPathForMessage('/repo/.agent-coordination/stop.md', '/repo')).toBe('.agent-coordination/stop.md')
    expect(formatPublicCorpusStopMarkerPathForMessage('/tmp/stop.md', '/repo')).toBe('/tmp/stop.md')
  })

  it('plans fetch progress without starting network work', () => {
    const artifactA = workbookArtifact('workbook-a')
    const artifactB = workbookArtifact('workbook-b')
    const dir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-cli-fetch-plan-'))
    const manifestPath = join(dir, 'manifest.json')
    const inactiveStopMarkerPath = join(dir, 'missing-stop.md')
    const manifest = {
      ...manifestWithArtifacts([artifactA, artifactB]),
      artifacts: [artifactA],
    }
    writeFileSync(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`)

    const result = spawnSync(
      'bun',
      [
        corpusScriptPath(),
        'fetch',
        '--dry-run',
        '--manifest',
        manifestPath,
        '--limit',
        '2',
        '--corpus-run-stop-marker',
        inactiveStopMarkerPath,
      ],
      {
        encoding: 'utf8',
      },
    )
    const plan: unknown = JSON.parse(result.stdout)

    expect(result.status).toBe(0)
    expect(plan).toMatchObject({
      stopMarker: {
        active: false,
        requiresExplicitResume: false,
      },
      targetArtifactCount: 2,
      cachedArtifactCount: 1,
      sourceCount: 2,
      remainingArtifactSlots: 1,
      candidateSourceCount: 1,
      candidateSourceDeficitCount: 0,
      minimumAdditionalSourceCount: 0,
      recommendedDiscoveryLimit: 2,
      recommendedDiscoveryPlanCommand: null,
      recommendedDiscoveryCommand: null,
      recommendedFetchCommand: expect.stringContaining('public-workbook-corpus:fetch'),
      blockedCommands: {},
      targetReachableFromKnownCandidates: true,
      sampledCandidateSources: [
        {
          id: artifactB.sourceId,
          fileName: artifactB.fileName,
          sourceUrl: artifactB.sourceUrl,
          downloadUrl: artifactB.downloadUrl,
          license: artifactB.license,
          topicEvidence: [`test:${artifactB.id}`],
        },
      ],
    })
    expect(asRecord(plan)['recommendedFetchCommand']).toContain('--limit 2')
  })

  it('omits mutating discovery commands when known candidates can reach the artifact target', () => {
    const artifactA = workbookArtifact('workbook-a')
    const artifactB = workbookArtifact('workbook-b')
    const dir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-cli-discover-not-needed-'))
    const manifestPath = join(dir, 'manifest.json')
    const manifest = {
      ...manifestWithArtifacts([artifactA, artifactB]),
      targetWorkbookCount: 2,
      artifacts: [artifactA],
    }
    writeFileSync(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`)

    const result = spawnSync('bun', [corpusScriptPath(), 'discover-plan', '--manifest', manifestPath, '--limit', '2'], {
      encoding: 'utf8',
    })
    const plan: unknown = JSON.parse(result.stdout)

    expect(result.status).toBe(0)
    expect(plan).toMatchObject({
      candidateSourceDeficitCount: 0,
      recommendedDiscoveryCommand: null,
      targetReachableFromKnownCandidates: true,
    })
  })

  it('plans source discovery when known candidates cannot reach the artifact target', () => {
    const artifactA = workbookArtifact('workbook-a')
    const artifactB = workbookArtifact('workbook-b')
    const dir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-cli-discover-plan-'))
    const manifestPath = join(dir, 'manifest.json')
    const inactiveStopMarkerPath = join(dir, 'missing-stop.md')
    const manifest = {
      ...manifestWithArtifacts([artifactA, artifactB]),
      targetWorkbookCount: 4,
      artifacts: [artifactA],
    }
    writeFileSync(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`)

    const result = spawnSync(
      'bun',
      [corpusScriptPath(), 'discover-plan', '--manifest', manifestPath, '--limit', '4', '--corpus-run-stop-marker', inactiveStopMarkerPath],
      {
        encoding: 'utf8',
      },
    )
    const plan: unknown = JSON.parse(result.stdout)

    expect(result.status).toBe(0)
    expect(plan).toMatchObject({
      stopMarker: {
        active: false,
        requiresExplicitResume: false,
      },
      sourceCount: 2,
      targetArtifactCount: 4,
      cachedArtifactCount: 1,
      remainingArtifactSlots: 3,
      candidateSourceCount: 1,
      candidateSourceDeficitCount: 2,
      minimumAdditionalSourceCount: 2,
      recommendedDiscoveryLimit: 4,
      recommendedDiscoveryCommand: expect.stringContaining('pnpm public-workbook-corpus:discover --'),
      blockedCommands: {},
      targetReachableFromKnownCandidates: false,
    })
  })

  it('keeps mutating discovery commands blocked in fetch plans while the stop marker is active', () => {
    const artifactA = workbookArtifact('workbook-a')
    const artifactB = workbookArtifact('workbook-b')
    const dir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-cli-fetch-plan-paused-'))
    const manifestPath = join(dir, 'manifest.json')
    const stopMarkerPath = join(dir, 'stop.md')
    const manifest = {
      ...manifestWithArtifacts([artifactA, artifactB]),
      targetWorkbookCount: 4,
      artifacts: [artifactA],
    }
    writeFileSync(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`)
    writeFileSync(stopMarkerPath, '# stop\n')

    const result = spawnSync(
      'bun',
      [corpusScriptPath(), 'fetch', '--dry-run', '--manifest', manifestPath, '--limit', '4', '--corpus-run-stop-marker', stopMarkerPath],
      {
        encoding: 'utf8',
      },
    )
    const plan = asRecord(JSON.parse(result.stdout))
    const blockedCommands = asRecord(plan['blockedCommands'])

    expect(result.status).toBe(0)
    expect(plan).toMatchObject({
      stopMarker: {
        active: true,
        requiresExplicitResume: true,
        overrideFlag: '--allow-active-stop-marker',
        overrideEnvVar: 'BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE',
      },
      recommendedDiscoveryPlanCommand: 'pnpm public-workbook-corpus:discover:plan -- --limit 4',
      recommendedDiscoveryCommand: null,
      targetReachableFromKnownCandidates: false,
    })
    expect(blockedCommands['discover']).toContain('BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE=1')
    expect(blockedCommands['discover']).toContain('public-workbook-corpus:discover')
    expect(blockedCommands['discover']).toContain('--allow-active-stop-marker')
  })

  it('keeps mutating fetch commands blocked in fetch plans while the stop marker is active', () => {
    const artifactA = workbookArtifact('workbook-a')
    const artifactB = workbookArtifact('workbook-b')
    const dir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-cli-fetch-plan-fetch-paused-'))
    const manifestPath = join(dir, 'manifest.json')
    const stopMarkerPath = join(dir, 'stop.md')
    const manifest = {
      ...manifestWithArtifacts([artifactA, artifactB]),
      targetWorkbookCount: 2,
      artifacts: [artifactA],
    }
    writeFileSync(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`)
    writeFileSync(stopMarkerPath, '# stop\n')

    const result = spawnSync(
      'bun',
      [corpusScriptPath(), 'fetch', '--dry-run', '--manifest', manifestPath, '--limit', '2', '--corpus-run-stop-marker', stopMarkerPath],
      {
        encoding: 'utf8',
      },
    )
    const plan = asRecord(JSON.parse(result.stdout))
    const blockedCommands = asRecord(plan['blockedCommands'])

    expect(result.status).toBe(0)
    expect(plan).toMatchObject({
      stopMarker: {
        active: true,
        requiresExplicitResume: true,
        overrideFlag: '--allow-active-stop-marker',
        overrideEnvVar: 'BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE',
      },
      recommendedDiscoveryCommand: null,
      recommendedFetchCommand: null,
      targetReachableFromKnownCandidates: true,
    })
    expect(blockedCommands['fetch']).toContain('BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE=1')
    expect(blockedCommands['fetch']).toContain('public-workbook-corpus:fetch')
    expect(blockedCommands['fetch']).toContain('--limit 2')
    expect(blockedCommands['fetch']).toContain('--allow-active-stop-marker')
  })

  it('preserves financial fetch script and batch size in stopped dry-run plans', () => {
    const artifactA = workbookArtifact('workbook-a')
    const artifactB = workbookArtifact('workbook-b')
    const dir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-cli-financial-fetch-plan-paused-'))
    const manifestPath = join(dir, 'manifest.json')
    const stopMarkerPath = join(dir, 'stop.md')
    const manifest = {
      ...manifestWithArtifacts([artifactA, artifactB]),
      targetWorkbookCount: 2,
      artifacts: [artifactA],
    }
    writeFileSync(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`)
    writeFileSync(stopMarkerPath, '# stop\n')

    const result = spawnSync(
      'bun',
      [
        corpusScriptPath(),
        'fetch',
        '--dry-run',
        '--manifest',
        manifestPath,
        '--limit',
        '2',
        '--fetch-script-name',
        'public-workbook-corpus:fetch-financial',
        '--fetch-batch-size',
        '6',
        '--corpus-run-stop-marker',
        stopMarkerPath,
      ],
      {
        encoding: 'utf8',
      },
    )
    const plan = asRecord(JSON.parse(result.stdout))
    const blockedCommands = asRecord(plan['blockedCommands'])

    expect(result.status).toBe(0)
    expect(blockedCommands['fetch']).toContain('BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE=1')
    expect(blockedCommands['fetch']).toContain('public-workbook-corpus:fetch-financial')
    expect(blockedCommands['fetch']).toContain('--fetch-batch-size 6')
    expect(blockedCommands['fetch']).toContain('--allow-active-stop-marker')
  })

  it('builds a stop-marker-aware bounded resume plan for the remaining corpus evidence', () => {
    const plan = buildPublicWorkbookCorpusResumePlan(
      resumePlanArgs({
        fetchPlan: {
          candidateSourceCount: 2_389,
          candidateSourceDeficitCount: 1_983,
          recommendedDiscoveryLimit: 11_983,
          remainingArtifactSlots: 4_372,
          targetReachableFromKnownCandidates: false,
        },
        status: {
          targetWorkbookCount: 10_000,
          cachedArtifactCount: 5_628,
          recordedManifestArtifactCount: 4_940,
          missingManifestArtifactCount: 688,
          recordedAllCasesPassed: true,
        },
        staleRecordedVerificationCount: 4_897,
      }),
    )

    expect(plan).toMatchObject({
      schemaVersion: 1,
      stopMarker: {
        active: true,
        path: '.agent-coordination/stop.md',
        requiresExplicitResume: true,
        overrideFlag: '--allow-active-stop-marker',
        overrideEnvVar: 'BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE',
      },
      currentState: {
        cachedArtifactCount: 5_628,
        recordedManifestArtifactCount: 4_940,
        missingCachedArtifactCount: 4_372,
        missingVerificationCount: 688,
        staleRecordedVerificationCount: 4_897,
      },
      phases: {
        verifyMissingCachedArtifacts: {
          status: 'blocked-by-stop-marker',
          totalWorkItems: 688,
          batchSize: 1,
          batchCount: 688,
          commands: [expect.stringContaining('public-workbook-corpus:verify-missing:plan')],
          blockedCommands: expect.arrayContaining([
            expect.stringContaining('BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE=1 pnpm public-workbook-corpus:verify-missing'),
            expect.stringContaining('--limit 1'),
          ]),
        },
        refreshStaleRecordedEvidence: {
          status: 'blocked-by-stop-marker',
          totalWorkItems: 4_897,
          batchSize: 1,
          batchCount: 4_897,
          commands: [expect.stringContaining('public-workbook-corpus:verify-stale:plan')],
          blockedCommands: expect.arrayContaining([
            expect.stringContaining('BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE=1 pnpm public-workbook-corpus:verify-stale'),
            expect.stringContaining('--limit 1'),
          ]),
        },
        discoverAdditionalSources: {
          status: 'blocked-by-stop-marker',
          totalWorkItems: 1_983,
          batchCount: 1,
          commands: [expect.stringContaining('public-workbook-corpus:discover:plan')],
          blockedCommands: [expect.stringContaining('--limit 11983')],
        },
        fetchAdditionalArtifacts: {
          status: 'blocked-by-stop-marker',
          totalWorkItems: 4_372,
          batchSize: 6,
          batchCount: 729,
          commands: [expect.stringContaining('public-workbook-corpus:fetch:plan')],
          blockedCommands: expect.arrayContaining([
            expect.stringContaining('BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE=1 pnpm public-workbook-corpus:fetch'),
            expect.stringContaining('--limit 5634'),
          ]),
        },
        finalEvidenceRefresh: {
          status: 'blocked-by-stop-marker',
          commands: expect.arrayContaining([
            'pnpm public-workbook-corpus:completion-audit:check -- --require-complete',
            'pnpm dominance:generate',
            'pnpm dominance:check',
          ]),
          blockedCommands: [expect.stringContaining('public-workbook-corpus:verify')],
        },
      },
    })
    expect(JSON.stringify(plan)).not.toContain('/repo/')
    expect(plan.phases.verifyMissingCachedArtifacts.commands[0]).toContain('--manifest .cache/public-workbook-corpus/manifest.json')
    expect(plan.phases.verifyMissingCachedArtifacts.commands[0]).toContain(
      '--scorecard packages/benchmarks/baselines/public-workbook-corpus-scorecard.json',
    )
    expect(plan.phases.fetchAdditionalArtifacts.blockedCommands[0]).toContain('--limit 5634')
    expect(plan.phases.fetchAdditionalArtifacts.blockedCommands[0]).toContain('--fetch-batch-size 6')
    expect(validatePublicWorkbookCorpusResumePlan(plan)).toEqual([])
  })

  it('preserves financial fetch scripts in stopped resume plans', () => {
    const artifactA = workbookArtifact('workbook-a')
    const artifactB = workbookArtifact('workbook-b')
    const dir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-financial-resume-plan-paused-'))
    const manifestPath = join(dir, 'manifest.json')
    const stopMarkerPath = join(dir, 'stop.md')
    const manifest = {
      ...manifestWithArtifacts([artifactA, artifactB]),
      targetWorkbookCount: 2,
      artifacts: [artifactA],
    }
    writeFileSync(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`)
    writeFileSync(stopMarkerPath, '# stop\n')

    const result = spawnSync(
      'bun',
      [
        resumePlanScriptPath(),
        '--manifest',
        manifestPath,
        '--cache-dir',
        dir,
        '--fetch-limit',
        '2',
        '--fetch-batch-size',
        '6',
        '--fetch-plan-script-name',
        'public-workbook-corpus:fetch-financial:plan',
        '--fetch-script-name',
        'public-workbook-corpus:fetch-financial',
        '--corpus-run-stop-marker',
        stopMarkerPath,
      ],
      {
        encoding: 'utf8',
      },
    )
    const plan = asRecord(JSON.parse(result.stdout))
    const phases = asRecord(plan['phases'])
    const fetchPhase = asRecord(phases['fetchAdditionalArtifacts'])
    const commands = stringArrayField(fetchPhase, 'commands')
    const blockedCommands = stringArrayField(fetchPhase, 'blockedCommands')

    expect(result.status).toBe(0)
    expect(commands[0]).toContain('public-workbook-corpus:fetch-financial:plan')
    expect(blockedCommands[0]).toContain('public-workbook-corpus:fetch-financial')
    expect(blockedCommands[0]).toContain('--fetch-batch-size 6')
  })

  it('uses repo-relative paths in status suggested commands for paths inside the checkout', () => {
    const artifactA = workbookArtifact('workbook-a')
    const artifactB = workbookArtifact('workbook-b')
    const status = buildPublicWorkbookCorpusStatus({
      manifest: manifestWithArtifacts([artifactA, artifactB]),
      scorecard: null,
      checkpointCases: [passedCase(artifactA)],
      commandPaths: {
        manifestPath: '/repo/.cache/public-workbook-corpus/manifest.json',
        scorecardPath: '/repo/packages/benchmarks/baselines/public-workbook-corpus-scorecard.json',
        verifyCheckpointPath: '/repo/.cache/public-workbook-corpus/verification-checkpoint.json',
        cacheDir: '/repo/.cache/public-workbook-corpus',
        displayRootDir: '/repo',
        stopMarkerActive: true,
      },
    })

    expect(status.nextMissingVerificationCommand).toBeNull()
    expect(status.blockedMissingVerificationCommand).toContain('--manifest .cache/public-workbook-corpus/manifest.json')
    expect(status.blockedMissingVerificationCommand).toContain(
      '--scorecard packages/benchmarks/baselines/public-workbook-corpus-scorecard.json',
    )
    expect(status.blockedMissingVerificationCommand).not.toContain('/repo/')
    expect(status.nextMissingVerificationPlanCommand).not.toContain('/repo/')
  })

  it('does not report scorecard health gaps before any artifacts are cached', () => {
    const status = buildPublicWorkbookCorpusStatus({
      manifest: {
        ...createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z', 5_000),
        sources: [
          {
            id: 'source-financial-a',
            kind: 'direct-url',
            sourceUrl: 'https://example.com/financial-a.xlsx',
            downloadUrl: 'https://example.com/financial-a.xlsx',
            fileName: 'financial-a.xlsx',
            discoveredAt: '2026-05-07T00:00:00.000Z',
            license: {
              spdxId: 'CC-BY-4.0',
              title: 'Creative Commons Attribution 4.0 International',
              evidenceUrl: 'https://creativecommons.org/licenses/by/4.0/',
            },
            topicEvidence: ['accounting:test'],
          },
        ],
      },
      scorecard: null,
      checkpointCases: [],
    })

    expect(status.gaps).toEqual(expect.arrayContaining(['cached artifacts below target: 0/5000']))
    expect(status.gaps).not.toContain('scorecard is missing or has non-passing cached workbooks')
    expect(status.gaps).not.toContain('scorecard is missing for cached workbooks')
    expect(status.gaps).not.toContain('scorecard has non-passing cached workbooks')
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
        staleRecordedVerificationCount: 0,
      },
      phases: {
        refreshStaleRecordedEvidence: { status: 'not-needed', totalWorkItems: 0 },
        discoverAdditionalSources: { status: 'not-needed' },
        fetchAdditionalArtifacts: { status: 'not-needed' },
      },
    })
  })

  it('rejects unsafe resume plans that hide active stop-marker overrides', () => {
    const plan = buildPublicWorkbookCorpusResumePlan(
      resumePlanArgs({
        fetchPlan: {
          candidateSourceCount: 1,
          candidateSourceDeficitCount: 0,
          recommendedDiscoveryLimit: 10_000,
          remainingArtifactSlots: 1,
          targetReachableFromKnownCandidates: true,
        },
        status: {
          targetWorkbookCount: 10_000,
          cachedArtifactCount: 9_999,
          recordedManifestArtifactCount: 9_998,
          missingManifestArtifactCount: 1,
          recordedAllCasesPassed: true,
        },
      }),
    )
    const invalidPlan = {
      ...plan,
      phases: {
        ...plan.phases,
        fetchAdditionalArtifacts: Object.assign({}, plan.phases.fetchAdditionalArtifacts, {
          commands: ['pnpm public-workbook-corpus:fetch -- --limit 10006'],
        }),
      },
    }

    expect(validatePublicWorkbookCorpusResumePlan(invalidPlan)).toEqual(
      expect.arrayContaining([
        'fetchAdditionalArtifacts mutating command is runnable while stop marker is active: pnpm public-workbook-corpus:fetch -- --limit 10006',
        'fetchAdditionalArtifacts mutating command limit 10006 exceeds one fetch tranche ending at 10005',
      ]),
    )
  })

  it('rejects stale-evidence resume plans that hide active stop-marker overrides', () => {
    const plan = buildPublicWorkbookCorpusResumePlan(
      resumePlanArgs({
        staleRecordedVerificationCount: 12,
      }),
    )
    const invalidPlan = {
      ...plan,
      phases: {
        ...plan.phases,
        refreshStaleRecordedEvidence: Object.assign({}, plan.phases.refreshStaleRecordedEvidence, {
          commands: ['pnpm public-workbook-corpus:verify-stale -- --limit 20'],
        }),
      },
    }

    expect(validatePublicWorkbookCorpusResumePlan(invalidPlan)).toEqual(
      expect.arrayContaining([
        'refreshStaleRecordedEvidence mutating command is runnable while stop marker is active: pnpm public-workbook-corpus:verify-stale -- --limit 20',
      ]),
    )
  })

  it('rejects stale resume plans with impossible current-state counts', () => {
    const plan = buildPublicWorkbookCorpusResumePlan(
      resumePlanArgs({
        fetchPlan: {
          candidateSourceCount: 6,
          candidateSourceDeficitCount: 0,
          recommendedDiscoveryLimit: 10_000,
          remainingArtifactSlots: 6,
          targetReachableFromKnownCandidates: true,
        },
        status: {
          targetWorkbookCount: 10_000,
          cachedArtifactCount: 9_994,
          recordedManifestArtifactCount: 9_994,
          missingManifestArtifactCount: 0,
          recordedAllCasesPassed: true,
        },
        stopMarkerActive: false,
      }),
    )
    const invalidPlan = {
      ...plan,
      currentState: {
        ...plan.currentState,
        missingCachedArtifactCount: 7,
        missingVerificationCount: -1,
        recordedManifestArtifactCount: 9_995,
        staleRecordedVerificationCount: -1,
      },
    }

    expect(validatePublicWorkbookCorpusResumePlan(invalidPlan)).toEqual(
      expect.arrayContaining([
        'missing verification count must be non-negative',
        'stale recorded verification count must be non-negative',
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

  it('refuses large verify-stale tranches unless explicitly enabled', () => {
    const env = { ...process.env }
    delete env.BILIG_ALLOW_LARGE_PUBLIC_CORPUS_VERIFY_STALE

    const result = spawnSync('bun', [corpusScriptPath(), 'verify-stale', '--limit', '21'], {
      encoding: 'utf8',
      env,
    })

    expect(result.status).not.toBe(0)
    expect(result.stderr).toContain('--limit above 20 is disabled for public corpus verify-stale runs')
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
        [artifactA.id, passedCaseWithUsedRange(artifactA)],
        [artifactB.id, passedCaseWithUsedRange(artifactB)],
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
        [artifactA.id, passedCaseWithUsedRange(artifactA)],
        [artifactB.id, passedCaseWithUsedRange(artifactB)],
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
      staleRecordedVerificationCount: 0,
      recordedPassedCaseCount: 2,
      recordedUnsupportedCaseCount: 0,
      recordedFailedCaseCount: 0,
      recordedErrorCaseCount: 0,
      missingManifestArtifactSample: [],
      staleRecordedVerificationSample: [],
      nextMissingVerificationCommand: null,
      nextMissingVerificationPlanCommand: null,
      blockedMissingVerificationCommand: null,
      nextStaleVerificationCommand: null,
      nextStaleVerificationPlanCommand: null,
      blockedStaleVerificationCommand: null,
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
    const inactiveStopMarkerPath = join(dir, 'not-stopped.md')
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
        '--corpus-run-stop-marker',
        inactiveStopMarkerPath,
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

  it('reports stale verification evidence with bounded next-step commands', async () => {
    const artifactA = workbookArtifact('workbook-a')
    const artifactB = workbookArtifact('workbook-b')
    const dir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-cli-status-stale-'))
    const manifestPath = join(dir, 'manifest.json')
    const scorecardPath = join(dir, 'missing-scorecard.json')
    const checkpointPath = join(dir, 'verification-checkpoint.json')
    const inactiveStopMarkerPath = join(dir, 'not-stopped.md')
    const fullManifest = manifestWithArtifacts([artifactA, artifactB])
    writeFileSync(manifestPath, `${JSON.stringify(fullManifest, null, 2)}\n`)
    writePublicWorkbookCorpusVerificationCheckpoint({
      path: checkpointPath,
      manifest: fullManifest,
      casesById: new Map([
        [artifactA.id, passedCase(artifactA)],
        [artifactB.id, passedCaseWithUsedRange(artifactB)],
      ]),
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
        inactiveStopMarkerPath,
      ],
      {
        encoding: 'utf8',
      },
    )
    const status: unknown = JSON.parse(result.stdout)

    expect(result.status).toBe(0)
    expect(status).toMatchObject({
      recordedManifestArtifactCount: 2,
      missingManifestArtifactCount: 0,
      staleRecordedVerificationCount: 1,
      staleRecordedVerificationSample: [
        {
          id: artifactA.id,
          fileName: artifactA.fileName,
          byteSize: artifactA.byteSize,
          sourceUrl: artifactA.sourceUrl,
          reason: 'missing-used-range-evidence',
        },
      ],
      nextStaleVerificationCommand: expect.stringContaining('public-workbook-corpus:verify-stale'),
      nextStaleVerificationPlanCommand: expect.stringContaining('public-workbook-corpus:verify-stale:plan'),
      gaps: expect.arrayContaining(['recorded verification cases need evidence refresh: 1']),
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
      nextMissingVerificationCommand: null,
      nextMissingVerificationPlanCommand: expect.not.stringContaining('BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE=1'),
      blockedMissingVerificationCommand: expect.stringContaining('BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE=1'),
      nextStaleVerificationCommand: null,
      nextStaleVerificationPlanCommand: expect.not.stringContaining('BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE=1'),
      blockedStaleVerificationCommand: expect.stringContaining('BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE=1'),
    })
    expect(status).toMatchObject({
      nextMissingVerificationPlanCommand: expect.stringContaining('public-workbook-corpus:verify-missing:plan'),
      blockedMissingVerificationCommand: expect.stringContaining('--allow-active-stop-marker'),
      nextStaleVerificationPlanCommand: expect.stringContaining('public-workbook-corpus:verify-stale:plan'),
      blockedStaleVerificationCommand: expect.stringContaining('--allow-active-stop-marker'),
    })
  })

  it('verifies a bounded missing slice into the checkpoint without duplicating scorecard cases', async () => {
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
    const recordedCases = readReusablePublicWorkbookCorpusCases([scorecardPath, checkpointPath])

    expect(result.status).toBe(0)
    expect(result.stdout).toContain('Verified 1 missing public workbook cases')
    expect(checkpointCases.map((entry) => [entry.id, entry.status])).toEqual([[artifactB.id, 'error']])
    expect(recordedCases.map((entry) => [entry.id, entry.status])).toEqual([
      [artifactA.id, 'passed'],
      [artifactB.id, 'error'],
    ])
    expect(checkpointCases[0]?.evidence).toEqual(expect.arrayContaining([`Missing cached workbook file: ${artifactB.cachePath}`]))
  })

  it('lists a bounded missing slice without starting verification', async () => {
    const artifactA = workbookArtifact('workbook-a')
    const artifactB = workbookArtifact('workbook-b')
    const artifactC = { ...workbookArtifact('workbook-c'), sha256: 'c'.repeat(64) }
    const dir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-cli-verify-missing-plan-'))
    const manifestPath = join(dir, 'manifest.json')
    const scorecardPath = join(dir, 'scorecard.json')
    const checkpointPath = join(dir, 'verification-checkpoint.json')
    const stopMarkerPath = join(dir, 'stop.md')
    const fullManifest = manifestWithArtifacts([artifactA, artifactB, artifactC])
    const scorecard = await buildPublicWorkbookCorpusScorecard({
      manifest: manifestWithArtifacts([artifactA]),
      cacheDir: dir,
      generatedAt: '2026-05-07T01:00:00.000Z',
      reusableCases: [passedCase(artifactA)],
    })
    writeFileSync(manifestPath, `${JSON.stringify(fullManifest, null, 2)}\n`)
    writeFileSync(scorecardPath, `${JSON.stringify(scorecard, null, 2)}\n`)
    writeFileSync(stopMarkerPath, '# paused\n')

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
        '--corpus-run-stop-marker',
        stopMarkerPath,
        '--limit',
        '20',
      ],
      {
        encoding: 'utf8',
      },
    )
    const planned: unknown = JSON.parse(result.stdout)

    expect(result.status).toBe(0)
    expect(planned).toMatchObject({
      totalMissingArtifactCount: 2,
      selectedArtifactCount: 1,
      artifacts: [
        {
          id: artifactB.id,
          fileName: artifactB.fileName,
          cachePath: artifactB.cachePath,
        },
      ],
      nextVerificationCommand: null,
      blockedVerificationCommand: expect.stringContaining('public-workbook-corpus:verify-missing'),
      artifactVerificationCommand: expect.stringContaining('public-workbook-corpus:verify-artifact'),
    })
    const blockedVerificationCommand = readPlanCommand(planned, 'blockedVerificationCommand')
    const artifactVerificationCommand = readPlanCommand(planned, 'artifactVerificationCommand')
    expect(blockedVerificationCommand).toContain('--limit 1')
    expect(blockedVerificationCommand).toContain('BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE=1')
    expect(blockedVerificationCommand).toContain('--allow-active-stop-marker')
    expect(artifactVerificationCommand).toContain(`--artifact-id ${artifactB.id}`)
    expect(artifactVerificationCommand).not.toContain('BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE=1')
    expect(artifactVerificationCommand).not.toContain('--allow-active-stop-marker')
    expect(artifactVerificationCommand).not.toContain('--update-verify-checkpoint')
    expect(readReusablePublicWorkbookCorpusCases([checkpointPath])).toEqual([])
  })

  it('lists a bounded stale-evidence slice without starting verification', async () => {
    const artifactA = workbookArtifact('workbook-a')
    const artifactB = workbookArtifact('workbook-b')
    const artifactC = { ...workbookArtifact('workbook-c'), sha256: 'c'.repeat(64) }
    const dir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-cli-verify-stale-plan-'))
    const manifestPath = join(dir, 'manifest.json')
    const scorecardPath = join(dir, 'scorecard.json')
    const checkpointPath = join(dir, 'verification-checkpoint.json')
    const stopMarkerPath = join(dir, 'stop.md')
    const fullManifest = manifestWithArtifacts([artifactA, artifactB, artifactC])
    writeFileSync(manifestPath, `${JSON.stringify(fullManifest, null, 2)}\n`)
    writeFileSync(stopMarkerPath, '# paused\n')
    writePublicWorkbookCorpusVerificationCheckpoint({
      path: checkpointPath,
      manifest: fullManifest,
      casesById: new Map([
        [artifactA.id, passedCase(artifactA)],
        [artifactB.id, passedCase(artifactB)],
        [artifactC.id, passedCaseWithUsedRange(artifactC)],
      ]),
      generatedAt: '2026-05-07T01:30:00.000Z',
    })

    const result = spawnSync(
      'bun',
      [
        corpusScriptPath(),
        'verify-stale',
        '--dry-run',
        '--manifest',
        manifestPath,
        '--scorecard',
        scorecardPath,
        '--verify-checkpoint',
        checkpointPath,
        '--corpus-run-stop-marker',
        stopMarkerPath,
        '--limit',
        '20',
      ],
      {
        encoding: 'utf8',
      },
    )
    const planned: unknown = JSON.parse(result.stdout)

    expect(result.status).toBe(0)
    expect(planned).toMatchObject({
      totalStaleArtifactCount: 2,
      selectedArtifactCount: 1,
      artifacts: [
        {
          id: artifactA.id,
          fileName: artifactA.fileName,
          cachePath: artifactA.cachePath,
          reason: 'missing-used-range-evidence',
        },
      ],
      nextVerificationCommand: null,
      blockedVerificationCommand: expect.stringContaining('public-workbook-corpus:verify-stale'),
      artifactVerificationCommand: expect.stringContaining('public-workbook-corpus:verify-artifact'),
    })
    const blockedVerificationCommand = readPlanCommand(planned, 'blockedVerificationCommand')
    const artifactVerificationCommand = readPlanCommand(planned, 'artifactVerificationCommand')
    expect(blockedVerificationCommand).toContain('--limit 1')
    expect(blockedVerificationCommand).toContain('BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE=1')
    expect(blockedVerificationCommand).toContain('--allow-active-stop-marker')
    expect(artifactVerificationCommand).toContain(`--artifact-id ${artifactA.id}`)
    expect(artifactVerificationCommand).not.toContain('BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE=1')
    expect(artifactVerificationCommand).not.toContain('--allow-active-stop-marker')
    expect(artifactVerificationCommand).not.toContain('--update-verify-checkpoint')
  })

  it('lists stale import-warning classifier evidence in verify-stale plans', async () => {
    const artifactA = workbookArtifact('workbook-a')
    const artifactB = workbookArtifact('workbook-b')
    const dir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-cli-verify-stale-warning-plan-'))
    const manifestPath = join(dir, 'manifest.json')
    const scorecardPath = join(dir, 'scorecard.json')
    const checkpointPath = join(dir, 'verification-checkpoint.json')
    const stopMarkerPath = join(dir, 'stop.md')
    const fullManifest = manifestWithArtifacts([artifactA, artifactB])
    writeFileSync(manifestPath, `${JSON.stringify(fullManifest, null, 2)}\n`)
    writeFileSync(stopMarkerPath, '# paused\n')
    writePublicWorkbookCorpusVerificationCheckpoint({
      path: checkpointPath,
      manifest: fullManifest,
      casesById: new Map([
        [artifactA.id, importWarningUnsupportedCaseWithUsedRange(artifactA, false)],
        [artifactB.id, importWarningUnsupportedCaseWithUsedRange(artifactB, true)],
      ]),
      generatedAt: '2026-05-07T01:30:00.000Z',
    })

    const result = spawnSync(
      'bun',
      [
        corpusScriptPath(),
        'verify-stale',
        '--dry-run',
        '--manifest',
        manifestPath,
        '--scorecard',
        scorecardPath,
        '--verify-checkpoint',
        checkpointPath,
        '--corpus-run-stop-marker',
        stopMarkerPath,
        '--limit',
        '20',
      ],
      {
        encoding: 'utf8',
      },
    )
    const planned: unknown = JSON.parse(result.stdout)

    expect(result.status).toBe(0)
    expect(planned).toMatchObject({
      totalStaleArtifactCount: 1,
      selectedArtifactCount: 1,
      artifacts: [
        {
          id: artifactA.id,
          reason: 'missing-import-warning-classifier-evidence',
        },
      ],
    })
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
    expect(packageJson.scripts?.['public-workbook-corpus:verify-stale']).toBe('bun scripts/public-workbook-corpus.ts verify-stale')
    expect(packageJson.scripts?.['public-workbook-corpus:verify-stale:plan']).toBe(
      'bun scripts/public-workbook-corpus.ts verify-stale --dry-run --limit 20',
    )
    expect(packageJson.scripts?.['public-workbook-corpus:verify-financial']).toContain(
      '--verify-checkpoint .cache/public-workbook-corpus-financial/verification-checkpoint.json',
    )
    expect(packageJson.scripts?.['public-workbook-corpus:check-financial']).toContain(
      '--manifest .cache/public-workbook-corpus-financial/manifest.json',
    )
    expect(packageJson.scripts?.['public-workbook-corpus:check-financial']).toContain('--cache-dir .cache/public-workbook-corpus-financial')
    expect(packageJson.scripts?.['public-workbook-corpus:check-financial']).toContain(
      '--verify-checkpoint .cache/public-workbook-corpus-financial/verification-checkpoint.json',
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
    expect(packageJson.scripts?.['public-workbook-corpus:feature-witness:plan']).toBe(
      'bun scripts/public-workbook-corpus-feature-witness-plan.ts',
    )
    expect(packageJson.scripts?.['public-workbook-corpus:feature-witness:check']).toBe(
      'bun scripts/public-workbook-corpus-feature-witness-plan.ts --check',
    )
  })
})

type ResumePlanArgs = Parameters<typeof buildPublicWorkbookCorpusResumePlan>[0]

function resumePlanArgs(overrides: Partial<ResumePlanArgs> = {}): ResumePlanArgs {
  return {
    cacheDir: '/repo/.cache/public-workbook-corpus',
    fetchBatchSize: 6,
    fetchLimit: 10_000,
    fetchPlan: {
      candidateSourceCount: 1,
      candidateSourceDeficitCount: 0,
      recommendedDiscoveryLimit: 10_000,
      remainingArtifactSlots: 0,
      targetReachableFromKnownCandidates: true,
    },
    generatedAt: '2026-05-07T08:00:00.000Z',
    displayRootDir: '/repo',
    manifestPath: '/repo/.cache/public-workbook-corpus/manifest.json',
    scorecardPath: '/repo/packages/benchmarks/baselines/public-workbook-corpus-scorecard.json',
    status: {
      targetWorkbookCount: 10_000,
      cachedArtifactCount: 10_000,
      recordedManifestArtifactCount: 10_000,
      missingManifestArtifactCount: 0,
      recordedAllCasesPassed: true,
    },
    stopMarkerActive: true,
    stopMarkerPath: '/repo/.agent-coordination/stop.md',
    verifyBatchSize: 20,
    verifyCheckpointPath: '/repo/.cache/public-workbook-corpus/verification-checkpoint.json',
    ...overrides,
  }
}

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

function readPlanCommand(plan: unknown, key: string): string {
  if (typeof plan !== 'object' || plan === null) {
    throw new Error('verify slice plan was not an object')
  }
  const command = Reflect.get(plan, key)
  if (typeof command !== 'string') {
    throw new Error(`verify slice plan did not include ${key}`)
  }
  return command
}

function stringArrayField(value: Record<string, unknown>, key: string): readonly string[] {
  const field = value[key]
  if (!Array.isArray(field) || !field.every((entry) => typeof entry === 'string')) {
    throw new Error(`expected ${key} to be a string array`)
  }
  return field
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
      topicEvidence: [`test:${entry.id}`],
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

function passedCaseWithUsedRange(artifact: PublicWorkbookArtifact): PublicWorkbookCorpusCase {
  return {
    ...passedCase(artifact),
    workbookMetadata: {
      workbookName: artifact.fileName,
      sheetNames: ['Sheet1'],
      dimensions: [
        {
          sheetName: 'Sheet1',
          rowCount: 1,
          columnCount: 1,
          nonEmptyCellCount: 1,
          usedRange: { startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 },
        },
      ],
    },
  }
}

function importWarningUnsupportedCaseWithUsedRange(
  artifact: PublicWorkbookArtifact,
  hasCurrentClassifierEvidence: boolean,
): PublicWorkbookCorpusCase {
  const base = passedCaseWithUsedRange(artifact)
  return {
    ...base,
    status: 'unsupported',
    featureCounts: {
      ...base.featureCounts,
      warningCount: 1,
    },
    unsupportedFeatureClassifications: ['xlsx.import.warning:Some defined names were ignored during XLSX import.'],
    evidence: [...base.evidence, ...(hasCurrentClassifierEvidence ? [publicWorkbookImportWarningClassifierEvidence] : [])],
  }
}
