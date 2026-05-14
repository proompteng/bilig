import { spawnSync } from 'node:child_process'
import { mkdtempSync, readFileSync, writeFileSync } from 'node:fs'
import { createServer } from 'node:http'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { afterEach, describe, expect, it, vi } from 'vitest'
import * as XLSX from 'xlsx'

import {
  createEmptyPublicWorkbookManifest,
  fetchPublicWorkbookArtifacts,
  formatPublicWorkbookCorpusVerifyArtifactCommand,
  parsePublicWorkbookManifestJson,
  validatePublicWorkbookManifest,
} from '../public-workbook-corpus.ts'
import { asRecord } from '../public-workbook-corpus-json.ts'
import { addPublicWorkbookLinkSource } from '../public-workbook-corpus-links.ts'
import type { PublicWorkbookArtifact, PublicWorkbookManifest, PublicWorkbookSource } from '../public-workbook-corpus-types.ts'

const license = {
  licenseSpdxId: 'CC-BY-4.0',
  licenseTitle: 'Creative Commons Attribution 4.0 International',
  licenseUrl: 'https://creativecommons.org/licenses/by/4.0/',
}

describe('public workbook corpus shared links', () => {
  afterEach(() => {
    vi.unstubAllGlobals()
  })

  it('normalizes a Google Sheets shared link into a licensed direct corpus source', () => {
    const manifest = createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z')

    const result = addPublicWorkbookLinkSource({
      manifest,
      sourceUrl: 'https://docs.google.com/spreadsheets/d/abc123SharedSheet/edit?usp=sharing',
      discoveredAt: '2026-05-07T01:00:00.000Z',
      ...license,
    })

    expect(result.added).toBe(true)
    expect(result.source).toMatchObject({
      kind: 'direct-url',
      sourceUrl: 'https://docs.google.com/spreadsheets/d/abc123SharedSheet/edit?usp=sharing',
      downloadUrl: 'https://docs.google.com/spreadsheets/d/abc123SharedSheet/export?format=xlsx',
      fileName: 'google-sheet-abc123SharedSheet.xlsx',
      license: {
        spdxId: 'CC-BY-4.0',
        title: license.licenseTitle,
        evidenceUrl: license.licenseUrl,
      },
    })
    expect(result.manifest.sources).toHaveLength(1)
    validatePublicWorkbookManifest(result.manifest)
  })

  it('dedupes repeated shared links with the same license evidence', () => {
    const manifest = createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z')
    const first = addPublicWorkbookLinkSource({
      manifest,
      sourceUrl: 'https://example.com/shared-budget.xlsx',
      discoveredAt: '2026-05-07T01:00:00.000Z',
      ...license,
    })
    const second = addPublicWorkbookLinkSource({
      manifest: first.manifest,
      sourceUrl: 'https://example.com/shared-budget.xlsx',
      discoveredAt: '2026-05-07T02:00:00.000Z',
      ...license,
    })

    expect(second.added).toBe(false)
    expect(second.source.id).toBe(first.source.id)
    expect(second.manifest.sources).toHaveLength(1)
  })

  it('requires public license evidence before adding shared links to the proof corpus', () => {
    const manifest = createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z')

    expect(() =>
      addPublicWorkbookLinkSource({
        manifest,
        sourceUrl: 'https://example.com/shared-budget.xlsx',
        licenseTitle: '',
        licenseUrl: '',
      }),
    ).toThrow('Shared workbook corpus sources require usable public license evidence')
  })

  it('prints a bounded fetch-source command during dry-run link intake', () => {
    const dir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-add-link-'))
    const manifestPath = join(dir, 'manifest.json')
    const inactiveStopMarkerPath = join(dir, 'not-stopped.md')
    writeFileSync(manifestPath, `${JSON.stringify(createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z'), null, 2)}\n`)

    const result = spawnSync(
      'bun',
      [
        corpusScriptPath(),
        'add-link',
        '--dry-run',
        '--manifest',
        manifestPath,
        '--cache-dir',
        dir,
        '--corpus-run-stop-marker',
        inactiveStopMarkerPath,
        '--source-url',
        'https://docs.google.com/spreadsheets/d/abc123SharedSheet/edit?usp=sharing',
        '--license-title',
        license.licenseTitle,
        '--license-url',
        license.licenseUrl,
        '--license-spdx',
        license.licenseSpdxId,
      ],
      {
        encoding: 'utf8',
      },
    )
    const dryRun: unknown = JSON.parse(result.stdout)
    const storedManifestJson: unknown = JSON.parse(readFileSync(manifestPath, 'utf8'))
    const storedManifest = parsePublicWorkbookManifestJson(storedManifestJson)

    expect(result.status).toBe(0)
    expect(dryRun).toMatchObject({
      mode: 'dry-run',
      added: true,
      sourceCountBefore: 0,
      sourceCountAfter: 1,
      source: {
        downloadUrl: 'https://docs.google.com/spreadsheets/d/abc123SharedSheet/export?format=xlsx',
      },
      nextFetchSourceCommand: expect.stringContaining('public-workbook-corpus:fetch-source'),
    })
    expect(storedManifest.sources).toEqual([])
  })

  it('plans the shared-link source lifecycle without mutating the manifest', () => {
    const dir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-link-plan-'))
    const manifestPath = join(dir, 'manifest.json')
    const scorecardPath = join(dir, 'scorecard.json')
    const checkpointPath = join(dir, 'verification-checkpoint.json')
    const inactiveStopMarkerPath = join(dir, 'not-stopped.md')
    writeFileSync(manifestPath, `${JSON.stringify(createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z'), null, 2)}\n`)

    const result = spawnSync(
      'bun',
      [
        corpusScriptPath(),
        'link-plan',
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
        '--source-url',
        'https://docs.google.com/spreadsheets/d/abc123SharedSheet/edit?usp=sharing',
        '--license-title',
        license.licenseTitle,
        '--license-url',
        license.licenseUrl,
        '--license-spdx',
        license.licenseSpdxId,
      ],
      {
        encoding: 'utf8',
      },
    )
    const plan: unknown = JSON.parse(result.stdout)
    const storedManifestJson: unknown = JSON.parse(readFileSync(manifestPath, 'utf8'))
    const storedManifest = parsePublicWorkbookManifestJson(storedManifestJson)

    expect(result.status).toBe(0)
    expect(plan).toMatchObject({
      mode: 'plan',
      sourceAlreadyKnown: false,
      artifactIds: [],
      recordedCaseIds: [],
      unverifiedArtifactIds: [],
      commands: {
        addLink: expect.stringContaining('public-workbook-corpus:add-link'),
        fetchSource: expect.stringContaining('public-workbook-corpus:fetch-source'),
        verifyArtifacts: [],
        status: expect.stringContaining('public-workbook-corpus:status'),
      },
    })
    expect(storedManifest.sources).toEqual([])
  })

  it('plans checkpoint verification for an already cached shared-link artifact', () => {
    const dir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-link-plan-cached-'))
    const manifestPath = join(dir, 'manifest.json')
    const inactiveStopMarkerPath = join(dir, 'not-stopped.md')
    const source = directSource('source-cached', 'https://example.com/cached.xlsx')
    const artifact = workbookArtifact(source)
    const manifest: PublicWorkbookManifest = {
      ...createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z'),
      sources: [source],
      artifacts: [artifact],
    }
    writeFileSync(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`)

    const result = spawnSync(
      'bun',
      [
        corpusScriptPath(),
        'link-plan',
        '--manifest',
        manifestPath,
        '--cache-dir',
        dir,
        '--corpus-run-stop-marker',
        inactiveStopMarkerPath,
        '--source-url',
        source.sourceUrl,
        '--license-title',
        license.licenseTitle,
        '--license-url',
        license.licenseUrl,
        '--license-spdx',
        license.licenseSpdxId,
      ],
      {
        encoding: 'utf8',
      },
    )
    const plan: unknown = JSON.parse(result.stdout)

    expect(result.status).toBe(0)
    expect(plan).toMatchObject({
      mode: 'plan',
      sourceAlreadyKnown: true,
      artifactIds: [artifact.id],
      recordedCaseIds: [],
      unverifiedArtifactIds: [artifact.id],
      commands: {
        verifyArtifacts: [expect.stringContaining(`--artifact-id ${artifact.id}`)],
      },
    })
  })

  it('marks generated fetch-source commands with the explicit stop-marker override when corpus runs are paused', () => {
    const dir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-add-link-paused-'))
    const manifestPath = join(dir, 'manifest.json')
    const stopMarkerPath = join(dir, 'stop.md')
    writeFileSync(manifestPath, `${JSON.stringify(createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z'), null, 2)}\n`)
    writeFileSync(stopMarkerPath, '# paused\n')

    const result = spawnSync(
      'bun',
      [
        corpusScriptPath(),
        'add-link',
        '--dry-run',
        '--manifest',
        manifestPath,
        '--cache-dir',
        dir,
        '--corpus-run-stop-marker',
        stopMarkerPath,
        '--source-url',
        'https://docs.google.com/spreadsheets/d/abc123SharedSheet/edit?usp=sharing',
        '--license-title',
        license.licenseTitle,
        '--license-url',
        license.licenseUrl,
        '--license-spdx',
        license.licenseSpdxId,
      ],
      {
        encoding: 'utf8',
      },
    )
    const dryRun: unknown = JSON.parse(result.stdout)

    expect(result.status).toBe(0)
    expect(dryRun).toMatchObject({
      nextFetchSourceCommand: null,
      blockedFetchSourceCommand: expect.stringContaining('BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE=1'),
    })
    expect(dryRun).toMatchObject({
      blockedFetchSourceCommand: expect.stringContaining('--allow-active-stop-marker'),
    })
  })

  it('keeps paused link-plan fetch and verification commands in blocked commands', () => {
    const dir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-link-plan-paused-'))
    const manifestPath = join(dir, 'manifest.json')
    const stopMarkerPath = join(dir, 'stop.md')
    const source = directSource('source-cached', 'https://example.com/cached.xlsx')
    const artifact = workbookArtifact(source)
    const manifest: PublicWorkbookManifest = {
      ...createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z'),
      sources: [source],
      artifacts: [artifact],
    }
    writeFileSync(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`)
    writeFileSync(stopMarkerPath, '# paused\n')

    const result = spawnSync(
      'bun',
      [
        corpusScriptPath(),
        'link-plan',
        '--manifest',
        manifestPath,
        '--cache-dir',
        dir,
        '--corpus-run-stop-marker',
        stopMarkerPath,
        '--source-url',
        source.sourceUrl,
        '--license-title',
        license.licenseTitle,
        '--license-url',
        license.licenseUrl,
        '--license-spdx',
        license.licenseSpdxId,
      ],
      {
        encoding: 'utf8',
      },
    )
    const plan: unknown = JSON.parse(result.stdout)

    expect(result.status).toBe(0)
    expect(plan).toMatchObject({
      commands: {
        fetchSource: null,
        verifyArtifacts: [],
      },
      blockedCommands: {
        fetchSource: expect.stringContaining('BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE=1'),
        verifyArtifacts: [expect.stringContaining(`--artifact-id ${artifact.id}`)],
      },
    })
    expect(plan).toMatchObject({
      blockedCommands: {
        fetchSource: expect.stringContaining('--allow-active-stop-marker'),
        verifyArtifacts: [expect.stringContaining('--allow-active-stop-marker')],
      },
    })
  })

  it('prints default shared-link plan commands with repo-relative corpus paths', () => {
    const inactiveStopMarkerPath = join(tmpdir(), 'public-workbook-corpus-default-link-plan-not-stopped.md')
    const result = spawnSync(
      'bun',
      [
        corpusScriptPath(),
        'link-plan',
        '--corpus-run-stop-marker',
        inactiveStopMarkerPath,
        '--source-url',
        'https://docs.google.com/spreadsheets/d/repoRelativePlanCheck/edit?usp=sharing',
        '--license-title',
        license.licenseTitle,
        '--license-url',
        license.licenseUrl,
        '--license-spdx',
        license.licenseSpdxId,
      ],
      { encoding: 'utf8' },
    )
    const plan: unknown = JSON.parse(result.stdout)

    expect(result.status).toBe(0)
    expect(result.stdout).not.toContain(repoRoot())
    expect(plan).toMatchObject({
      commands: {
        addLink: expect.stringContaining('--manifest .cache/public-workbook-corpus/manifest.json'),
        fetchSource: expect.stringContaining('--cache-dir .cache/public-workbook-corpus'),
        status: expect.stringContaining('--scorecard packages/benchmarks/baselines/public-workbook-corpus-scorecard.json'),
      },
    })
  })

  it('prints default shared-link dry-run commands with repo-relative corpus paths', () => {
    const inactiveStopMarkerPath = join(tmpdir(), 'public-workbook-corpus-default-add-link-not-stopped.md')
    const result = spawnSync(
      'bun',
      [
        corpusScriptPath(),
        'add-link',
        '--dry-run',
        '--corpus-run-stop-marker',
        inactiveStopMarkerPath,
        '--source-url',
        'https://docs.google.com/spreadsheets/d/repoRelativeDryRunCheck/edit?usp=sharing',
        '--license-title',
        license.licenseTitle,
        '--license-url',
        license.licenseUrl,
        '--license-spdx',
        license.licenseSpdxId,
      ],
      { encoding: 'utf8' },
    )
    const dryRun: unknown = JSON.parse(result.stdout)

    expect(result.status).toBe(0)
    expect(result.stdout).not.toContain(repoRoot())
    expect(dryRun).toMatchObject({
      nextFetchSourceCommand: expect.stringContaining('--cache-dir .cache/public-workbook-corpus'),
      nextPlanCommand: expect.stringContaining('--verify-checkpoint .cache/public-workbook-corpus/verification-checkpoint.json'),
    })
  })

  it('refuses bounded fetch-source while the corpus stop marker is active', () => {
    const dir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-fetch-source-paused-'))
    const stopMarkerPath = join(dir, 'stop.md')
    writeFileSync(stopMarkerPath, '# paused\n')

    const result = spawnSync(
      'bun',
      [corpusScriptPath(), 'fetch-source', '--cache-dir', dir, '--corpus-run-stop-marker', stopMarkerPath, '--source-id', 'source-a'],
      {
        encoding: 'utf8',
        env: { ...process.env, BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE: '' },
      },
    )

    expect(result.status).not.toBe(0)
    expect(result.stderr).toContain('public-workbook-corpus fetch-source is disabled while the public corpus stop marker is active')
    expect(result.stderr).toContain('--allow-active-stop-marker')
  })

  it('fetches only the explicitly selected shared-link source', async () => {
    const cacheDir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-fetch-source-'))
    const workbookBytes = buildWorkbookBytes()
    const fetchMock = vi.fn(async () => new Response(workbookBytes, { headers: { 'content-length': String(workbookBytes.byteLength) } }))
    vi.stubGlobal('fetch', fetchMock)
    const sourceA = directSource('source-a', 'https://example.com/a.xlsx')
    const sourceB = directSource('source-b', 'https://example.com/b.xlsx')
    const manifest: PublicWorkbookManifest = {
      ...createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z'),
      sources: [sourceA, sourceB],
    }

    const fetched = await fetchPublicWorkbookArtifacts({
      manifest,
      cacheDir,
      limit: 1,
      fetchedAt: '2026-05-07T01:00:00.000Z',
      sourceIds: ['source-b'],
    })

    expect(fetchMock).toHaveBeenCalledTimes(1)
    expect(fetchMock.mock.calls[0]?.[0]).toBe('https://example.com/b.xlsx')
    expect(fetched.artifacts).toHaveLength(1)
    expect(fetched.artifacts[0]?.sourceId).toBe('source-b')
  })

  it('prints fetch-source checkpoint progress when the selected source fails', async () => {
    const cacheDir = mkdtempSync(join(tmpdir(), 'public-workbook-corpus-fetch-source-progress-'))
    const manifestPath = join(cacheDir, 'manifest.json')
    const server = createServer((_request, response) => {
      response.writeHead(404, { 'content-type': 'text/plain' })
      response.end('missing')
    })
    await new Promise<void>((resolve, reject) => {
      server.once('error', reject)
      server.listen(0, '127.0.0.1', () => resolve())
    })
    try {
      const address = server.address()
      if (!address || typeof address === 'string') {
        throw new Error('Expected local HTTP server address')
      }
      const url = `http://127.0.0.1:${String(address.port)}/missing.xlsx`
      const manifest: PublicWorkbookManifest = {
        ...createEmptyPublicWorkbookManifest('2026-05-07T00:00:00.000Z'),
        sources: [directSource('source-missing', url)],
      }
      writeFileSync(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`)

      const result = spawnSync(
        'bun',
        [
          corpusScriptPath(),
          'fetch-source',
          '--manifest',
          manifestPath,
          '--cache-dir',
          cacheDir,
          '--source-id',
          'source-missing',
          '--download-timeout-ms',
          '1000',
        ],
        { encoding: 'utf8' },
      )
      const output = asRecord(JSON.parse(result.stdout) as unknown)
      const checkpointProgress = output['checkpointProgress']
      expect(Array.isArray(checkpointProgress)).toBe(true)
      const firstProgress = asRecord(checkpointProgress[0])
      const failedSourceSamples = firstProgress['failedSourceSamples']
      expect(Array.isArray(failedSourceSamples)).toBe(true)

      expect(result.status).toBe(0)
      expect(firstProgress).toMatchObject({
        artifactCount: 0,
        exhaustedSourceCount: 1,
        failedSourceCount: 1,
      })
      expect(asRecord(failedSourceSamples[0])).toMatchObject({
        sourceId: 'source-missing',
        error: 'Request timed out after 1000ms',
      })
    } finally {
      await new Promise<void>((resolve, reject) => {
        server.close((error) => {
          if (error) {
            reject(error)
            return
          }
          resolve()
        })
      })
    }
  })

  it('formats the one-artifact checkpoint verification command for fetched shared links', () => {
    expect(
      formatPublicWorkbookCorpusVerifyArtifactCommand({
        artifactId: 'workbook-abc123',
        cacheDir: '/repo/.cache/public-workbook-corpus',
        manifestPath: '/repo/.cache/public-workbook-corpus/manifest.json',
        verifyCheckpointPath: '/repo/.cache/public-workbook-corpus/verification-checkpoint.json',
      }),
    ).toBe(
      'pnpm public-workbook-corpus:verify-artifact -- --manifest /repo/.cache/public-workbook-corpus/manifest.json --cache-dir /repo/.cache/public-workbook-corpus --verify-checkpoint /repo/.cache/public-workbook-corpus/verification-checkpoint.json --artifact-id workbook-abc123 --update-verify-checkpoint',
    )
  })

  it('formats read-only one-artifact verification commands without checkpoint mutation', () => {
    expect(
      formatPublicWorkbookCorpusVerifyArtifactCommand({
        artifactId: 'workbook-abc123',
        cacheDir: '/repo/.cache/public-workbook-corpus',
        manifestPath: '/repo/.cache/public-workbook-corpus/manifest.json',
        updateVerifyCheckpoint: false,
        verifyCheckpointPath: '/repo/.cache/public-workbook-corpus/verification-checkpoint.json',
      }),
    ).toBe(
      'pnpm public-workbook-corpus:verify-artifact -- --manifest /repo/.cache/public-workbook-corpus/manifest.json --cache-dir /repo/.cache/public-workbook-corpus --verify-checkpoint /repo/.cache/public-workbook-corpus/verification-checkpoint.json --artifact-id workbook-abc123',
    )
  })

  it('marks checkpoint-updating artifact verification commands with stop-marker overrides when paused', () => {
    expect(
      formatPublicWorkbookCorpusVerifyArtifactCommand({
        artifactId: 'workbook-abc123',
        cacheDir: '/repo/.cache/public-workbook-corpus',
        manifestPath: '/repo/.cache/public-workbook-corpus/manifest.json',
        stopMarkerActive: true,
        verifyCheckpointPath: '/repo/.cache/public-workbook-corpus/verification-checkpoint.json',
      }),
    ).toBe(
      'BILIG_ALLOW_PUBLIC_CORPUS_STOP_MARKER_OVERRIDE=1 pnpm public-workbook-corpus:verify-artifact -- --manifest /repo/.cache/public-workbook-corpus/manifest.json --cache-dir /repo/.cache/public-workbook-corpus --verify-checkpoint /repo/.cache/public-workbook-corpus/verification-checkpoint.json --artifact-id workbook-abc123 --update-verify-checkpoint --allow-active-stop-marker',
    )
  })
})

function directSource(id: string, url: string): PublicWorkbookSource {
  return {
    id,
    kind: 'direct-url',
    sourceUrl: url,
    downloadUrl: url,
    fileName: url.split('/').at(-1) ?? 'workbook.xlsx',
    discoveredAt: '2026-05-07T00:00:00.000Z',
    license: {
      spdxId: license.licenseSpdxId,
      title: license.licenseTitle,
      evidenceUrl: license.licenseUrl,
    },
  }
}

function workbookArtifact(source: PublicWorkbookSource): PublicWorkbookArtifact {
  return {
    id: 'workbook-cached',
    sourceId: source.id,
    sourceUrl: source.sourceUrl,
    downloadUrl: source.downloadUrl,
    fileName: source.fileName,
    cachePath: 'files/cached.xlsx',
    sha256: 'a'.repeat(64),
    byteSize: 1024,
    workbookFingerprint: 'cached-fingerprint',
    fetchedAt: '2026-05-07T01:00:00.000Z',
    license: source.license,
  }
}

function buildWorkbookBytes(): Uint8Array {
  const workbook = XLSX.utils.book_new()
  const sheet = XLSX.utils.aoa_to_sheet([
    ['Name', 'Amount'],
    ['Revenue', 10],
  ])
  XLSX.utils.book_append_sheet(workbook, sheet, 'Summary')
  const bytes: unknown = XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' })
  if (!(bytes instanceof Uint8Array)) {
    throw new Error('Expected XLSX writer to return workbook bytes')
  }
  return bytes
}

function corpusScriptPath(): string {
  return join(dirname(fileURLToPath(import.meta.url)), '../public-workbook-corpus.ts')
}

function repoRoot(): string {
  return join(dirname(fileURLToPath(import.meta.url)), '../..')
}
