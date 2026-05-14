import { spawn } from 'node:child_process'
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

import { asRecord, readString, spreadsheetExtension, validatePublicWorkbookManifest } from './public-workbook-corpus-json.ts'
import { fetchBodyBytesWithTimeout } from './public-workbook-corpus-http.ts'
import { formatByteSize, startChildRssWatchdog, terminateChildProcess } from './public-workbook-corpus-process.ts'
import type {
  FetchCorpusArgs,
  PublicWorkbookArtifact,
  PublicWorkbookCorpusFetchCheckpointProgress,
  PublicWorkbookCorpusFetchFailureSample,
  PublicWorkbookManifest,
  PublicWorkbookSource,
  WorkbookDownloadResult,
} from './public-workbook-corpus-types.ts'
import { fingerprintWorkbookBytes, sha256HexSync } from './public-workbook-corpus-workbook.ts'

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const publicWorkbookCorpusScriptPath = fileURLToPath(new URL('./public-workbook-corpus.ts', import.meta.url))

export const defaultDownloadTimeoutMs = 60_000
export const defaultFetchBatchSize = 6
export const defaultFetchConcurrency = 1
export const defaultFetchMaxRssBytes = 1536 * 1024 * 1024
export const defaultFingerprintTimeoutMs = 180_000
export const defaultFingerprintMaxRssBytes = 1024 * 1024 * 1024
const noop = (): void => undefined

export interface PublicWorkbookCorpusFetchPlan {
  readonly targetArtifactCount: number
  readonly cachedArtifactCount: number
  readonly sourceCount: number
  readonly remainingArtifactSlots: number
  readonly candidateSourceCount: number
  readonly candidateSourceDeficitCount: number
  readonly minimumAdditionalSourceCount: number
  readonly recommendedDiscoveryLimit: number
  readonly targetReachableFromKnownCandidates: boolean
  readonly sampledCandidateSources: readonly PublicWorkbookSource[]
}

export function planPublicWorkbookCorpusFetch(args: {
  readonly manifest: PublicWorkbookManifest
  readonly limit: number
  readonly sampleLimit?: number
}): PublicWorkbookCorpusFetchPlan {
  validatePublicWorkbookManifest(args.manifest)
  const targetArtifactCount = Math.min(args.limit, args.manifest.targetWorkbookCount)
  const remainingArtifactSlots = Math.max(0, targetArtifactCount - args.manifest.artifacts.length)
  const exhaustedSourceIds = new Set(args.manifest.fetchState?.exhaustedSourceIds ?? [])
  const existingSourceIds = new Set(args.manifest.artifacts.map((artifact) => artifact.sourceId))
  const existingDownloadUrls = new Set(args.manifest.artifacts.map((artifact) => normalizeSourceUrl(artifact.downloadUrl)))
  const candidateSources = prioritizeCandidateSources(
    dedupeCandidateSources(
      args.manifest.sources.filter(
        (source) =>
          !exhaustedSourceIds.has(source.id) &&
          !existingSourceIds.has(source.id) &&
          !existingDownloadUrls.has(normalizeSourceUrl(source.downloadUrl)),
      ),
    ),
  )
  return {
    targetArtifactCount,
    cachedArtifactCount: args.manifest.artifacts.length,
    sourceCount: args.manifest.sources.length,
    remainingArtifactSlots,
    candidateSourceCount: candidateSources.length,
    candidateSourceDeficitCount: Math.max(0, remainingArtifactSlots - candidateSources.length),
    minimumAdditionalSourceCount: Math.max(0, remainingArtifactSlots - candidateSources.length),
    recommendedDiscoveryLimit: args.manifest.sources.length + Math.max(0, remainingArtifactSlots - candidateSources.length),
    targetReachableFromKnownCandidates: candidateSources.length >= remainingArtifactSlots,
    sampledCandidateSources: candidateSources.slice(0, Math.max(0, Math.trunc(args.sampleLimit ?? 20))),
  }
}

export async function fetchPublicWorkbookArtifacts(args: FetchCorpusArgs): Promise<PublicWorkbookManifest> {
  validatePublicWorkbookManifest(args.manifest)
  const fetchedAt = args.fetchedAt ?? new Date().toISOString()
  const maxBytes = args.maxBytes ?? 50 * 1024 * 1024
  const downloadTimeoutMs = args.downloadTimeoutMs ?? defaultDownloadTimeoutMs
  const fetchBatchSize = Math.max(1, Math.trunc(args.fetchBatchSize ?? defaultFetchBatchSize))
  const fetchConcurrency = Math.max(1, Math.trunc(args.fetchConcurrency ?? defaultFetchConcurrency))
  const fingerprintTimeoutMs = args.fingerprintTimeoutMs ?? defaultFingerprintTimeoutMs
  const fingerprintMaxRssBytes = args.fingerprintMaxRssBytes ?? defaultFingerprintMaxRssBytes
  const isolatedFingerprinting = args.isolatedFingerprinting === true
  const artifacts: PublicWorkbookArtifact[] = [...args.manifest.artifacts]
  const exhaustedSourceIds = new Set(args.manifest.fetchState?.exhaustedSourceIds ?? [])
  const knownHashes = new Set(artifacts.map((artifact) => artifact.sha256))
  const knownFingerprints = new Set(artifacts.map((artifact) => artifact.workbookFingerprint))
  mkdirSync(join(args.cacheDir, 'files'), { recursive: true })
  const targetArtifactCount = Math.min(args.limit, args.manifest.targetWorkbookCount)
  const remainingArtifactSlots = Math.max(0, targetArtifactCount - artifacts.length)
  if (remainingArtifactSlots === 0) {
    return {
      ...args.manifest,
      generatedAt: fetchedAt,
      artifacts,
    }
  }
  const candidateSources = planPublicWorkbookCorpusFetch({
    manifest: args.manifest,
    limit: args.limit,
    sampleLimit: Number.MAX_SAFE_INTEGER,
  }).sampledCandidateSources
  const allowedSourceIds = args.sourceIds ? new Set(args.sourceIds) : null
  const selectedCandidateSources = allowedSourceIds
    ? candidateSources.filter((source) => allowedSourceIds.has(source.id))
    : candidateSources
  await fetchArtifactsFromCandidateSources({
    candidateSources: selectedCandidateSources,
    artifacts,
    cacheDir: args.cacheDir,
    fetchedAt,
    knownFingerprints,
    knownHashes,
    downloadTimeoutMs,
    fetchBatchSize,
    fetchConcurrency,
    fingerprintTimeoutMs,
    fingerprintMaxRssBytes,
    fingerprintRssCheckIntervalMs: args.fingerprintRssCheckIntervalMs,
    isolatedFingerprinting,
    maxBytes,
    onArtifactsCommitted: args.onArtifactsCommitted,
    exhaustedSourceIds,
    targetArtifactCount,
    sourceManifest: args.manifest,
  })
  return {
    ...args.manifest,
    generatedAt: fetchedAt,
    artifacts,
    ...fetchStateForManifest(args.manifest, exhaustedSourceIds),
  }
}

async function fetchArtifactsFromCandidateSources(args: {
  readonly candidateSources: readonly PublicWorkbookSource[]
  readonly artifacts: PublicWorkbookArtifact[]
  readonly cacheDir: string
  readonly fetchedAt: string
  readonly knownFingerprints: Set<string>
  readonly knownHashes: Set<string>
  readonly downloadTimeoutMs: number
  readonly fetchBatchSize: number
  readonly fetchConcurrency: number
  readonly fingerprintTimeoutMs: number
  readonly fingerprintMaxRssBytes: number
  readonly fingerprintRssCheckIntervalMs?: number
  readonly isolatedFingerprinting: boolean
  readonly maxBytes: number
  readonly onArtifactsCommitted?: (
    manifest: PublicWorkbookManifest,
    progress: PublicWorkbookCorpusFetchCheckpointProgress,
  ) => void | Promise<void>
  readonly exhaustedSourceIds: Set<string>
  readonly sourceManifest: PublicWorkbookManifest
  readonly targetArtifactCount: number
}): Promise<void> {
  let startIndex = 0
  while (args.artifacts.length < args.targetArtifactCount && startIndex < args.candidateSources.length) {
    const remainingArtifactSlots = args.targetArtifactCount - args.artifacts.length
    const batchSize = Math.min(
      args.candidateSources.length - startIndex,
      Math.max(args.fetchConcurrency, Math.min(args.fetchBatchSize, remainingArtifactSlots * 3)),
    )
    const batch = args.candidateSources.slice(startIndex, startIndex + batchSize)
    // oxlint-disable-next-line eslint(no-await-in-loop) -- Each batch must finish and release workbook bytes before fetching the next batch.
    const downloadResults = await mapWithConcurrency(batch, args.fetchConcurrency, (source) =>
      downloadWorkbookCandidate(source, {
        downloadTimeoutMs: args.downloadTimeoutMs,
        fingerprintTimeoutMs: args.fingerprintTimeoutMs,
        fingerprintMaxRssBytes: args.fingerprintMaxRssBytes,
        fingerprintRssCheckIntervalMs: args.fingerprintRssCheckIntervalMs,
        isolatedFingerprinting: args.isolatedFingerprinting,
        maxBytes: args.maxBytes,
      }),
    )
    let committedArtifactCount = 0
    let exhaustedSourceCount = 0
    let failedSourceCount = 0
    let duplicateHashSourceCount = 0
    let duplicateFingerprintSourceCount = 0
    const failedSourceSamples: PublicWorkbookCorpusFetchFailureSample[] = []
    for (const result of downloadResults) {
      if (args.artifacts.length >= args.targetArtifactCount) {
        break
      }
      if (result.error || !result.bytes || !result.sha256 || !result.workbookFingerprint) {
        failedSourceCount += 1
        if (failedSourceSamples.length < 3) {
          failedSourceSamples.push({
            sourceId: result.source.id,
            fileName: result.source.fileName,
            error: result.error ?? 'download result was missing workbook bytes, hash, or fingerprint',
          })
        }
        if (!args.exhaustedSourceIds.has(result.source.id)) {
          args.exhaustedSourceIds.add(result.source.id)
          exhaustedSourceCount += 1
        }
        continue
      }
      if (args.knownHashes.has(result.sha256)) {
        duplicateHashSourceCount += 1
        if (!args.exhaustedSourceIds.has(result.source.id)) {
          args.exhaustedSourceIds.add(result.source.id)
          exhaustedSourceCount += 1
        }
        continue
      }
      if (args.knownFingerprints.has(result.workbookFingerprint)) {
        duplicateFingerprintSourceCount += 1
        if (!args.exhaustedSourceIds.has(result.source.id)) {
          args.exhaustedSourceIds.add(result.source.id)
          exhaustedSourceCount += 1
        }
        continue
      }
      const source = result.source
      const bytes = result.bytes
      const hash = result.sha256
      const workbookFingerprint = result.workbookFingerprint
      const cachePath = `files/${hash}.${spreadsheetExtension(source.fileName)}`
      writeFileSync(join(args.cacheDir, cachePath), bytes)
      args.knownHashes.add(hash)
      args.knownFingerprints.add(workbookFingerprint)
      args.artifacts.push({
        id: `workbook-${hash.slice(0, 16)}`,
        sourceId: source.id,
        sourceUrl: source.sourceUrl,
        downloadUrl: source.downloadUrl,
        fileName: source.fileName,
        cachePath,
        sha256: hash,
        byteSize: bytes.byteLength,
        workbookFingerprint,
        fetchedAt: args.fetchedAt,
        license: source.license,
        ...(source.topicEvidence ? { topicEvidence: source.topicEvidence } : {}),
      })
      if (!args.exhaustedSourceIds.has(source.id)) {
        args.exhaustedSourceIds.add(source.id)
        exhaustedSourceCount += 1
      }
      committedArtifactCount += 1
    }
    if (committedArtifactCount > 0 || exhaustedSourceCount > 0) {
      // oxlint-disable-next-line eslint(no-await-in-loop) -- Checkpoint each bounded batch before continuing the long corpus fetch.
      await args.onArtifactsCommitted?.(
        {
          ...args.sourceManifest,
          generatedAt: args.fetchedAt,
          artifacts: [...args.artifacts],
          ...fetchStateForManifest(args.sourceManifest, args.exhaustedSourceIds),
        },
        {
          artifactCount: args.artifacts.length,
          exhaustedSourceCount: args.exhaustedSourceIds.size,
          committedArtifactCount,
          exhaustedSourceDelta: exhaustedSourceCount,
          failedSourceCount,
          duplicateHashSourceCount,
          duplicateFingerprintSourceCount,
          failedSourceSamples,
        },
      )
    }
    releaseFetchBatchMemory()
    startIndex += batchSize
  }
}

function fetchStateForManifest(
  manifest: PublicWorkbookManifest,
  exhaustedSourceIds: ReadonlySet<string>,
): Pick<PublicWorkbookManifest, 'fetchState'> | Record<string, never> {
  const orderedExhaustedSourceIds = manifest.sources.map((source) => source.id).filter((sourceId) => exhaustedSourceIds.has(sourceId))
  return orderedExhaustedSourceIds.length > 0 ? { fetchState: { exhaustedSourceIds: orderedExhaustedSourceIds } } : {}
}

function dedupeCandidateSources(sources: readonly PublicWorkbookSource[]): PublicWorkbookSource[] {
  const seenDownloadUrls = new Set<string>()
  const deduped: PublicWorkbookSource[] = []
  for (const source of sources) {
    const downloadUrl = normalizeSourceUrl(source.downloadUrl)
    if (seenDownloadUrls.has(downloadUrl)) {
      continue
    }
    seenDownloadUrls.add(downloadUrl)
    deduped.push(source)
  }
  return deduped
}

function prioritizeCandidateSources(sources: readonly PublicWorkbookSource[]): PublicWorkbookSource[] {
  return sources
    .map((source, index) => ({ index, source }))
    .toSorted((left, right) => {
      const priorityDifference = workbookSourceFetchPriority(left.source) - workbookSourceFetchPriority(right.source)
      return priorityDifference === 0 ? left.index - right.index : priorityDifference
    })
    .map(({ source }) => source)
}

function workbookSourceFetchPriority(source: PublicWorkbookSource): number {
  const extension = spreadsheetExtension(source.fileName)
  if (extension === 'xlsx') {
    return 0
  }
  if (extension === 'xlsm') {
    return 1
  }
  return 2
}

function normalizeSourceUrl(value: string): string {
  return value.trim().toLowerCase()
}

function releaseFetchBatchMemory(): void {
  const maybeBun = Reflect.get(globalThis, 'Bun')
  const maybeBunGc = maybeBun && typeof maybeBun === 'object' ? Reflect.get(maybeBun, 'gc') : null
  if (typeof maybeBunGc === 'function') {
    try {
      Reflect.apply(maybeBunGc, maybeBun, [true])
    } catch {
      // Batch memory release is a best-effort guard; fetch correctness is owned by the checkpointed artifacts.
    }
    return
  }
  const maybeGlobalGc = Reflect.get(globalThis, 'gc')
  if (typeof maybeGlobalGc === 'function') {
    try {
      Reflect.apply(maybeGlobalGc, globalThis, [])
    } catch {
      // Ignore unavailable host GC hooks after already committing the bounded batch.
    }
  }
}

async function downloadWorkbookCandidate(
  source: PublicWorkbookSource,
  args: {
    readonly downloadTimeoutMs: number
    readonly fingerprintTimeoutMs: number
    readonly fingerprintMaxRssBytes: number
    readonly fingerprintRssCheckIntervalMs?: number
    readonly isolatedFingerprinting: boolean
    readonly maxBytes: number
  },
): Promise<WorkbookDownloadResult> {
  try {
    const bytes = await downloadWorkbookBytes(source.downloadUrl, args.maxBytes, args.downloadTimeoutMs)
    const sha256 = sha256HexSync(bytes)
    const workbookFingerprint = args.isolatedFingerprinting
      ? await fingerprintWorkbookBytesIsolated(bytes, source.fileName, args.fingerprintTimeoutMs, {
          maxRssBytes: args.fingerprintMaxRssBytes,
          rssCheckIntervalMs: args.fingerprintRssCheckIntervalMs,
        })
      : fingerprintWorkbookBytes(bytes, source.fileName)
    return {
      source,
      bytes,
      sha256,
      workbookFingerprint,
      error: null,
    }
  } catch (error) {
    return {
      source,
      bytes: null,
      sha256: null,
      workbookFingerprint: null,
      error: error instanceof Error ? error.message : String(error),
    }
  }
}

function fingerprintWorkbookBytesIsolated(
  bytes: Uint8Array,
  fileName: string,
  timeoutMs: number,
  resourceLimits: {
    readonly maxRssBytes: number
    readonly rssCheckIntervalMs?: number
  },
): Promise<string> {
  const tempDir = mkdtempSync(join(tmpdir(), 'public-workbook-fingerprint-'))
  const tempPath = join(tempDir, `workbook.${spreadsheetExtension(fileName)}`)
  writeFileSync(tempPath, bytes)
  return fingerprintWorkbookFileIsolated(tempPath, fileName, timeoutMs, resourceLimits, () => {
    rmSync(tempDir, { recursive: true, force: true })
  })
}

export function fingerprintWorkbookFileIsolated(
  filePath: string,
  fileName: string,
  timeoutMs: number,
  resourceLimits: {
    readonly maxRssBytes: number
    readonly rssCheckIntervalMs?: number
  },
  onCleanup: () => void = noop,
): Promise<string> {
  return new Promise<string>((resolvePromise, rejectPromise) => {
    const child = spawn(
      process.execPath,
      [
        publicWorkbookCorpusScriptPath,
        'fingerprint-artifact-worker',
        '--file',
        filePath,
        '--file-name',
        fileName,
        '--fingerprint-max-rss-mb',
        String(Math.ceil(resourceLimits.maxRssBytes / 1024 / 1024)),
      ],
      {
        detached: true,
        stdio: ['ignore', 'pipe', 'pipe'],
      },
    )
    let stdout = ''
    let stderr = ''
    let timer: ReturnType<typeof setTimeout>
    let settled = false
    let stopRssWatchdog = noop
    const terminateChild = (signal: 'SIGTERM' | 'SIGKILL'): void => {
      terminateChildProcess(child, signal, { processGroup: true })
    }
    const terminateChildOnParentExit = (): void => terminateChild('SIGTERM')
    const parentSignalHandlers = registerParentTerminationHandlers(terminateChildOnParentExit)
    const cleanup = (): void => {
      clearTimeout(timer)
      stopRssWatchdog()
      process.off('exit', terminateChildOnParentExit)
      for (const { signal, handler } of parentSignalHandlers) {
        process.off(signal, handler)
      }
      onCleanup()
    }
    process.once('exit', terminateChildOnParentExit)
    const finish = (value: string): void => {
      if (settled) {
        return
      }
      settled = true
      cleanup()
      resolvePromise(value)
    }
    const fail = (error: Error): void => {
      if (settled) {
        return
      }
      settled = true
      cleanup()
      rejectPromise(error)
    }
    timer = setTimeout(() => {
      terminateChild('SIGTERM')
      const forceKillTimer = setTimeout(() => terminateChild('SIGKILL'), 5_000)
      forceKillTimer.unref()
      fail(new Error(`Workbook fingerprinting timed out after ${String(timeoutMs)}ms`))
    }, timeoutMs)
    stopRssWatchdog = startChildRssWatchdog(child, {
      maxRssBytes: resourceLimits.maxRssBytes,
      intervalMs: resourceLimits.rssCheckIntervalMs,
      onLimitExceeded: (rssBytes) => {
        terminateChild('SIGTERM')
        const forceKillTimer = setTimeout(() => terminateChild('SIGKILL'), 5_000)
        forceKillTimer.unref()
        fail(
          new Error(
            `Workbook fingerprinting subprocess exceeded RSS limit: ${formatByteSize(rssBytes)} > ${formatByteSize(
              resourceLimits.maxRssBytes,
            )}`,
          ),
        )
      },
    })
    child.stdout.setEncoding('utf8')
    child.stderr.setEncoding('utf8')
    child.stdout.on('data', (chunk: string) => {
      stdout += chunk
    })
    child.stderr.on('data', (chunk: string) => {
      stderr += chunk
    })
    child.on('error', (error) => {
      fail(new Error(`Workbook fingerprinting subprocess failed to start: ${error.message}`))
    })
    child.on('close', (code, signal) => {
      if (code !== 0) {
        const details = compactProcessOutput(stderr || stdout)
        fail(
          new Error(
            `Workbook fingerprinting subprocess exited with ${signal ? `signal ${signal}` : `code ${String(code)}`}${
              details ? `: ${details}` : ''
            }`,
          ),
        )
        return
      }
      try {
        const parsed = asRecord(JSON.parse(stdout) as unknown)
        const workbookFingerprint = readString(parsed, 'workbookFingerprint')
        if (!workbookFingerprint) {
          throw new Error('Missing workbookFingerprint')
        }
        finish(workbookFingerprint)
      } catch (error) {
        fail(
          new Error(`Workbook fingerprinting subprocess returned invalid JSON: ${error instanceof Error ? error.message : String(error)}`),
        )
      }
    })
  })
}

function registerParentTerminationHandlers(onTerminate: () => void): ReadonlyArray<{
  readonly signal: NodeJS.Signals
  readonly handler: () => void
}> {
  return (['SIGHUP', 'SIGINT', 'SIGTERM'] as const).map((signal) => {
    const handler = (): void => {
      onTerminate()
      process.exit(signalExitCode(signal))
    }
    process.once(signal, handler)
    return { signal, handler }
  })
}

function signalExitCode(signal: NodeJS.Signals): number {
  if (signal === 'SIGHUP') {
    return 129
  }
  if (signal === 'SIGINT') {
    return 130
  }
  return 143
}

async function downloadWorkbookBytes(url: string, maxBytes: number, timeoutMs: number): Promise<Uint8Array> {
  const { bytes } = await fetchBodyBytesWithTimeout(
    url,
    {
      headers: {
        'user-agent': 'bilig-public-workbook-corpus/1.0',
        accept: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel,*/*',
      },
    },
    {
      timeoutMs,
      maxBytes,
      maxBytesLabel: 'Workbook',
      validateResponse: (response) => {
        if (!response.ok) {
          throw new Error(`Unable to download ${url}: HTTP ${String(response.status)}`)
        }
        const contentLength = Number(response.headers.get('content-length') ?? '0')
        if (contentLength > maxBytes) {
          throw new Error(`Workbook exceeds max byte size: ${String(contentLength)} > ${String(maxBytes)}`)
        }
      },
    },
  )
  return bytes
}

async function mapWithConcurrency<T, R>(
  items: readonly T[],
  concurrency: number,
  mapper: (item: T, index: number) => Promise<R>,
): Promise<R[]> {
  const results: R[] = []
  let nextIndex = 0
  const workerCount = Math.min(Math.max(1, concurrency), items.length)
  const runNext = async (): Promise<void> => {
    const index = nextIndex
    nextIndex += 1
    if (index >= items.length) {
      return
    }
    results[index] = await mapper(items[index], index)
    await runNext()
  }
  await Promise.all(Array.from({ length: workerCount }, () => runNext()))
  return results
}

function compactProcessOutput(value: string): string | null {
  const compacted = value.replaceAll(rootDir, '<repo>').replace(/\s+/gu, ' ').trim()
  return compacted.length > 0 ? compacted.slice(0, 1_000) : null
}
