#!/usr/bin/env bun

import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { pathToFileURL } from 'node:url'

import {
  createEmptyPublicWorkbookManifest,
  parsePublicWorkbookCorpusScorecardJson,
  parsePublicWorkbookManifestJson,
  validatePublicWorkbookManifest,
} from './public-workbook-corpus-json.ts'
import { defaultCkanPortalBases, discoverCkanWorkbookSources, discoverFinancialCkanQueries } from './public-workbook-corpus-discovery.ts'
import {
  defaultDownloadTimeoutMs,
  defaultFetchBatchSize,
  defaultFetchConcurrency,
  defaultFetchMaxRssBytes,
  defaultFingerprintMaxRssBytes,
  defaultFingerprintTimeoutMs,
  fetchPublicWorkbookArtifacts,
  planPublicWorkbookCorpusFetch,
} from './public-workbook-corpus-fetch.ts'
import { withPublicWorkbookCorpusCacheLock } from './public-workbook-corpus-lock.ts'
import { indexPublicWorkbookCorpusCases, publicWorkbookCorpusCaseMatchesArtifact } from './public-workbook-corpus-missing.ts'
import { addPublicWorkbookLinkSource } from './public-workbook-corpus-links.ts'
import { defaultSelfRssCheckIntervalMs, startSelfRssGuard } from './public-workbook-corpus-process.ts'
import { buildPublicWorkbookCorpusScorecardFromCases, validatePublicWorkbookCorpusScorecard } from './public-workbook-corpus-scorecard.ts'
import { writePublicWorkbookCorpusCheck, writePublicWorkbookCorpusStatus } from './public-workbook-corpus-status.ts'
import { defaultFinancialWorkbookQueries } from './public-workbook-corpus-topics.ts'
import {
  runPublicWorkbookCorpusVerifyMissingCommand,
  runPublicWorkbookCorpusVerifyStaleCommand,
} from './public-workbook-corpus-verify-slices.ts'
import {
  readReusablePublicWorkbookCorpusCases,
  upsertPublicWorkbookCorpusVerificationCheckpoint,
  writePublicWorkbookCorpusVerificationCheckpoint,
} from './public-workbook-corpus-verify-checkpoint.ts'
import { sha256HexSync } from './public-workbook-corpus-workbook.ts'
import {
  writeFingerprintArtifactResult,
  writeFingerprintArtifactWorkerResult,
  writeFootprintWorkerResult,
} from './public-workbook-corpus-worker-commands.ts'
import type {
  PublicWorkbookArtifact,
  PublicWorkbookCorpusCase,
  PublicWorkbookCorpusFetchCheckpointProgress,
  PublicWorkbookManifest,
} from './public-workbook-corpus-types.ts'
import {
  buildPublicWorkbookCorpusScorecard,
  capVerifyMaxRssBytes,
  defaultVerifyConcurrency,
  defaultVerifyMaxCellCount,
  defaultVerifyMaxRssBytes,
  defaultVerifyTimeoutMs,
  verifyCachedWorkbookArtifact,
  verifyCachedWorkbookArtifactIsolated,
} from './public-workbook-corpus-verify.ts'
import { formatJsonForRepo } from './scorecard-format.ts'
import {
  assertPublicCorpusRunNotStopped,
  readDebugOnlyFlagArg,
  readFetchRunArgs,
  readFlagArg,
  readMegabytesArg,
  readNumberArg,
  readRepeatedStringArg,
  readStringArg,
  readVerifyConcurrencyArg,
} from './public-workbook-corpus-cli.ts'
import {
  formatCommandPath,
  formatPublicWorkbookCorpusAddLinkCommand,
  formatPublicWorkbookCorpusDiscoverCommand,
  formatPublicWorkbookCorpusFetchSourceCommand,
  formatPublicWorkbookCorpusLinkPlanCommand,
  formatPublicWorkbookCorpusStatusCommand,
  formatPublicWorkbookCorpusVerifyArtifactCommand,
  publicWorkbookCorpusPlanStopMarker,
  splitPublicWorkbookCorpusFetchCommand,
  splitPublicWorkbookCorpusFetchSourceCommand,
  splitPublicWorkbookCorpusVerifyArtifactCommand,
  type PublicWorkbookLinkInput,
} from './public-workbook-corpus-command-format.ts'

export {
  buildPublicWorkbookCorpusScorecard,
  createEmptyPublicWorkbookManifest,
  discoverCkanWorkbookSources,
  formatPublicWorkbookCorpusVerifyArtifactCommand,
  parsePublicWorkbookCorpusScorecardJson,
  parsePublicWorkbookManifestJson,
  validatePublicWorkbookCorpusScorecard,
  validatePublicWorkbookManifest,
  fetchPublicWorkbookArtifacts,
}
export type {
  PublicWorkbookArtifact,
  PublicWorkbookCaseStatus,
  PublicWorkbookCorpusCase,
  PublicWorkbookCorpusFetchCheckpointProgress,
  PublicWorkbookCorpusScorecard,
  PublicWorkbookFeatureCounts,
  PublicWorkbookLicenseEvidence,
  PublicWorkbookManifest,
  PublicWorkbookSource,
  PublicWorkbookSourceKind,
  PublicWorkbookValidationSummary,
} from './public-workbook-corpus-types.ts'

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const defaultCacheDir = join(rootDir, '.cache', 'public-workbook-corpus')
const defaultManifestPath = join(defaultCacheDir, 'manifest.json')
const defaultScorecardPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'public-workbook-corpus-scorecard.json')
const defaultCorpusRunStopMarkerPath = join(rootDir, '.agent-coordination', '20260507T074946Z-codex-stop-interactive-corpus-runs.md')

export async function sha256Hex(bytes: Uint8Array): Promise<string> {
  return sha256HexSync(bytes)
}

function readManifest(path: string): PublicWorkbookManifest {
  const parsed: unknown = JSON.parse(readFileSync(path, 'utf8'))
  return parsePublicWorkbookManifestJson(parsed)
}

function writeJson(path: string, value: unknown, tempPrefix: string): void {
  mkdirSync(dirname(path), { recursive: true })
  writeFileSync(path, serializeJsonForRepo(value, tempPrefix))
}

function serializeJsonForRepo(value: unknown, tempPrefix: string): string {
  return formatJsonForRepo({
    rootDir,
    serializedJson: `${JSON.stringify(value, null, 2)}\n`,
    tempPrefix,
  })
}

function readOrCreateManifest(path: string, targetWorkbookCount = 10_000): PublicWorkbookManifest {
  return existsSync(path) ? readManifest(path) : createEmptyPublicWorkbookManifest(undefined, targetWorkbookCount)
}

async function main(): Promise<void> {
  const command = process.argv[2] ?? 'verify'
  const manifestPath = resolve(readStringArg('--manifest', defaultManifestPath))
  const cacheDir = resolve(readStringArg('--cache-dir', defaultCacheDir))
  const scorecardPath = resolve(readStringArg('--scorecard', defaultScorecardPath))
  const verifyCheckpointPath = resolve(readStringArg('--verify-checkpoint', join(cacheDir, 'verification-checkpoint.json')))
  const targetWorkbookCount = readNumberArg('--target-workbook-count', 10_000)
  const corpusRunStopMarkerPath = resolve(readStringArg('--corpus-run-stop-marker', defaultCorpusRunStopMarkerPath))
  if (command === 'init') {
    await withPublicWorkbookCorpusCacheLock(cacheDir, 'init', async () => {
      if (existsSync(manifestPath) && !readFlagArg('--force')) {
        throw new Error(
          `Public workbook corpus manifest already exists at ${formatCommandPath(manifestPath)}; pass --force to reinitialize it`,
        )
      }
      writeJson(manifestPath, createEmptyPublicWorkbookManifest(undefined, targetWorkbookCount), 'public-workbook-corpus-manifest')
    })
    return
  }
  if (command === 'link-plan') {
    const linkInput = readPublicWorkbookLinkInput(command)
    const manifest = readOrCreateManifest(manifestPath, targetWorkbookCount)
    const result = addPublicWorkbookLinkSourceFromInput(manifest, linkInput)
    const recordedCases = readReusablePublicWorkbookCorpusCases([scorecardPath, verifyCheckpointPath])
    const recordedCaseIds = new Set(recordedCases.map((entry) => entry.id))
    const sourceArtifacts = manifest.artifacts.filter((artifact) => artifact.sourceId === result.source.id)
    const unverifiedArtifactIds = sourceArtifacts.filter((artifact) => !recordedCaseIds.has(artifact.id)).map((artifact) => artifact.id)
    const stopMarkerActive = existsSync(corpusRunStopMarkerPath)
    const fetchSourceCommand = splitPublicWorkbookCorpusFetchSourceCommand({
      cacheDir,
      manifestPath,
      sourceId: result.source.id,
      stopMarkerActive,
    })
    const verifyArtifactCommands = unverifiedArtifactIds.map((artifactId) =>
      splitPublicWorkbookCorpusVerifyArtifactCommand({
        artifactId,
        cacheDir,
        manifestPath,
        stopMarkerActive,
        verifyCheckpointPath,
      }),
    )
    process.stdout.write(
      `${JSON.stringify(
        {
          mode: 'plan',
          sourceAlreadyKnown: !result.added,
          source: result.source,
          sourceCountBefore: manifest.sources.length,
          sourceCountAfter: result.manifest.sources.length,
          artifactIds: sourceArtifacts.map((artifact) => artifact.id),
          recordedCaseIds: sourceArtifacts.filter((artifact) => recordedCaseIds.has(artifact.id)).map((artifact) => artifact.id),
          unverifiedArtifactIds,
          commands: {
            addLink: formatPublicWorkbookCorpusAddLinkCommand({ linkInput, manifestPath }),
            fetchSource: fetchSourceCommand.command,
            verifyArtifacts: verifyArtifactCommands.flatMap((entry) => (entry.command ? [entry.command] : [])),
            status: formatPublicWorkbookCorpusStatusCommand({ cacheDir, manifestPath, scorecardPath, verifyCheckpointPath }),
          },
          blockedCommands: {
            fetchSource: fetchSourceCommand.blockedCommand,
            verifyArtifacts: verifyArtifactCommands.flatMap((entry) => (entry.blockedCommand ? [entry.blockedCommand] : [])),
          },
        },
        null,
        2,
      )}\n`,
    )
    return
  }
  if (command === 'add-link') {
    const linkInput = readPublicWorkbookLinkInput(command)
    const addSource = (manifest: PublicWorkbookManifest) => addPublicWorkbookLinkSourceFromInput(manifest, linkInput)
    if (readFlagArg('--dry-run') || readFlagArg('--list')) {
      const manifest = readOrCreateManifest(manifestPath, targetWorkbookCount)
      const result = addSource(manifest)
      const fetchSourceCommand = splitPublicWorkbookCorpusFetchSourceCommand({
        cacheDir,
        manifestPath,
        sourceId: result.source.id,
        stopMarkerActive: existsSync(corpusRunStopMarkerPath),
      })
      process.stdout.write(
        `${JSON.stringify(
          {
            mode: 'dry-run',
            added: result.added,
            sourceCountBefore: manifest.sources.length,
            sourceCountAfter: result.manifest.sources.length,
            source: result.source,
            nextFetchSourceCommand: fetchSourceCommand.command,
            blockedFetchSourceCommand: fetchSourceCommand.blockedCommand,
            nextPlanCommand: formatPublicWorkbookCorpusLinkPlanCommand({ linkInput, manifestPath, scorecardPath, verifyCheckpointPath }),
          },
          null,
          2,
        )}\n`,
      )
      return
    }
    const result = await withPublicWorkbookCorpusCacheLock(cacheDir, 'add-link', async () => {
      const added = addSource(readOrCreateManifest(manifestPath, targetWorkbookCount))
      writeJson(manifestPath, added.manifest, 'public-workbook-corpus-manifest')
      return added
    })
    console.log(`${result.added ? 'Added' : 'Reused'} public workbook source ${result.source.id}`)
    console.log(
      formatPublicWorkbookCorpusFetchSourceCommand({
        cacheDir,
        manifestPath,
        sourceId: result.source.id,
        stopMarkerActive: existsSync(corpusRunStopMarkerPath),
      }),
    )
    return
  }
  if (command === 'discover-ckan') {
    const portalBases = readRepeatedStringArg('--ckan-base')
    assertPublicCorpusRunNotStopped({
      commandName: 'public-workbook-corpus discover',
      stopMarkerPath: corpusRunStopMarkerPath,
    })
    const manifest = await withPublicWorkbookCorpusCacheLock(cacheDir, 'discover-ckan', async () => {
      const discoveredManifest = await discoverCkanWorkbookSources({
        manifest: readOrCreateManifest(manifestPath, targetWorkbookCount),
        portalBases: portalBases.length > 0 ? portalBases : defaultCkanPortalBases,
        query: readStringArg('--query', 'xlsx'),
        limit: readNumberArg('--limit', 10_000),
        rowsPerRequest: readNumberArg('--rows', 100),
        ...(readStringArg('--required-topic', '') === 'financial-workpapers' ? { requiredTopic: 'financial-workpapers' as const } : {}),
      })
      writeJson(manifestPath, discoveredManifest, 'public-workbook-corpus-manifest')
      return discoveredManifest
    })
    console.log(`Discovered ${String(manifest.sources.length)} public workbook sources`)
    return
  }
  if (command === 'discover-financial-ckan') {
    const portalBases = readRepeatedStringArg('--ckan-base')
    const queries = readRepeatedStringArg('--query')
    const limit = readNumberArg('--limit', 5_000)
    const rowsPerRequest = readNumberArg('--rows', 100)
    assertPublicCorpusRunNotStopped({
      commandName: 'public-workbook-corpus discover-financial',
      stopMarkerPath: corpusRunStopMarkerPath,
    })
    const manifest = await withPublicWorkbookCorpusCacheLock(cacheDir, 'discover-financial-ckan', async () => {
      const discoveredManifest = await discoverFinancialCkanQueries({
        manifest: readOrCreateManifest(manifestPath, targetWorkbookCount),
        portalBases: portalBases.length > 0 ? portalBases : defaultCkanPortalBases,
        queries: queries.length > 0 ? queries : defaultFinancialWorkbookQueries,
        limit,
        rowsPerRequest,
        onQueryDiscovered: (checkpointManifest) => {
          writeJson(manifestPath, checkpointManifest, 'public-workbook-corpus-manifest')
        },
      })
      writeJson(manifestPath, discoveredManifest, 'public-workbook-corpus-manifest')
      return discoveredManifest
    })
    console.log(`Discovered ${String(manifest.sources.length)} financial workbook sources`)
    return
  }
  if (command === 'fetch') {
    if (readFlagArg('--dry-run') || readFlagArg('--list')) {
      const fetchBatchSize = process.argv.includes('--fetch-batch-size') ? readNumberArg('--fetch-batch-size', defaultFetchBatchSize) : null
      const fetchScriptName = readStringArg('--fetch-script-name', 'public-workbook-corpus:fetch')
      const plan = planPublicWorkbookCorpusFetch({
        manifest: readManifest(manifestPath),
        limit: readNumberArg('--limit', 10_000),
        sampleLimit: readNumberArg('--sample-limit', 20),
      })
      const stopMarkerActive = existsSync(corpusRunStopMarkerPath)
      const needsDiscovery = !plan.targetReachableFromKnownCandidates
      const fetchCommand =
        plan.remainingArtifactSlots > 0 && plan.candidateSourceCount > 0 && plan.targetReachableFromKnownCandidates
          ? splitPublicWorkbookCorpusFetchCommand({
              cacheDir,
              fetchBatchSize,
              limit: plan.targetArtifactCount,
              manifestPath,
              scriptName: fetchScriptName,
              stopMarkerActive,
            })
          : null
      const blockedCommands = {
        ...(needsDiscovery && stopMarkerActive
          ? {
              discover: formatPublicWorkbookCorpusDiscoverCommand({
                cacheDir,
                limit: plan.recommendedDiscoveryLimit,
                manifestPath,
                stopMarkerActive: true,
              }),
            }
          : {}),
        ...(fetchCommand?.blockedCommand ? { fetch: fetchCommand.blockedCommand } : {}),
      }
      process.stdout.write(
        `${JSON.stringify(
          {
            stopMarker: publicWorkbookCorpusPlanStopMarker(stopMarkerActive),
            targetArtifactCount: plan.targetArtifactCount,
            cachedArtifactCount: plan.cachedArtifactCount,
            sourceCount: plan.sourceCount,
            remainingArtifactSlots: plan.remainingArtifactSlots,
            candidateSourceCount: plan.candidateSourceCount,
            candidateSourceDeficitCount: plan.candidateSourceDeficitCount,
            minimumAdditionalSourceCount: plan.minimumAdditionalSourceCount,
            recommendedDiscoveryLimit: plan.recommendedDiscoveryLimit,
            recommendedDiscoveryPlanCommand: plan.targetReachableFromKnownCandidates
              ? null
              : `pnpm public-workbook-corpus:discover:plan -- --limit ${String(plan.recommendedDiscoveryLimit)}`,
            recommendedDiscoveryCommand:
              !needsDiscovery || stopMarkerActive
                ? null
                : formatPublicWorkbookCorpusDiscoverCommand({
                    cacheDir,
                    limit: plan.recommendedDiscoveryLimit,
                    manifestPath,
                  }),
            recommendedFetchCommand: fetchCommand?.command ?? null,
            blockedCommands,
            targetReachableFromKnownCandidates: plan.targetReachableFromKnownCandidates,
            sampledCandidateSources: plan.sampledCandidateSources.map((source) => ({
              id: source.id,
              kind: source.kind,
              fileName: source.fileName,
              sourceUrl: source.sourceUrl,
              downloadUrl: source.downloadUrl,
              license: source.license,
              topicEvidence: source.topicEvidence ?? [],
              portal: source.portal,
              datasetId: source.datasetId,
              resourceId: source.resourceId,
            })),
          },
          null,
          2,
        )}\n`,
      )
      return
    }
    const { inProcessFingerprinting, ...fetchRunArgs } = readFetchRunArgs({
      batchSize: defaultFetchBatchSize,
      concurrency: defaultFetchConcurrency,
    })
    assertPublicCorpusRunNotStopped({
      commandName: 'public-workbook-corpus fetch',
      stopMarkerPath: corpusRunStopMarkerPath,
    })
    const stopSelfRssGuard = startSelfRssGuard(
      readMegabytesArg('--fetch-max-rss-mb', defaultFetchMaxRssBytes),
      'Public workbook corpus fetch',
    )
    let manifest: PublicWorkbookManifest
    try {
      manifest = await withPublicWorkbookCorpusCacheLock(cacheDir, 'fetch', async () => {
        const fetchedManifest = await fetchPublicWorkbookArtifacts({
          manifest: readManifest(manifestPath),
          cacheDir,
          limit: readNumberArg('--limit', 10_000),
          downloadTimeoutMs: readNumberArg('--download-timeout-ms', defaultDownloadTimeoutMs),
          ...fetchRunArgs,
          fingerprintTimeoutMs: readNumberArg('--fingerprint-timeout-ms', defaultFingerprintTimeoutMs),
          fingerprintMaxRssBytes: readMegabytesArg('--fingerprint-max-rss-mb', defaultFingerprintMaxRssBytes),
          isolatedFingerprinting: !inProcessFingerprinting,
          maxBytes: readNumberArg('--max-bytes', 50 * 1024 * 1024),
          onArtifactsCommitted: (checkpointManifest, progress) => {
            writeJson(manifestPath, checkpointManifest, 'public-workbook-corpus-manifest')
            console.error(formatFetchCheckpointProgress(progress))
          },
        })
        writeJson(manifestPath, fetchedManifest, 'public-workbook-corpus-manifest')
        return fetchedManifest
      })
    } finally {
      stopSelfRssGuard()
    }
    console.log(`Cached ${String(manifest.artifacts.length)} public workbook artifacts`)
    return
  }
  if (command === 'fetch-source') {
    const sourceId = readStringArg('--source-id', '')
    if (!sourceId) {
      throw new Error('Expected --source-id for fetch-source')
    }
    assertPublicCorpusRunNotStopped({
      commandName: 'public-workbook-corpus fetch-source',
      stopMarkerPath: corpusRunStopMarkerPath,
    })
    const { inProcessFingerprinting, ...fetchRunArgs } = readFetchRunArgs({
      batchSize: 1,
      concurrency: 1,
    })
    const stopSelfRssGuard = startSelfRssGuard(
      readMegabytesArg('--fetch-max-rss-mb', defaultFetchMaxRssBytes),
      'Public workbook corpus fetch-source',
    )
    let result: {
      readonly artifactCountBefore: number
      readonly artifactCountAfter: number
      readonly checkpointProgress: readonly PublicWorkbookCorpusFetchCheckpointProgress[]
      readonly fetchedArtifactIds: readonly string[]
    }
    try {
      result = await withPublicWorkbookCorpusCacheLock(cacheDir, 'fetch-source', async () => {
        const manifest = readManifest(manifestPath)
        if (!manifest.sources.some((source) => source.id === sourceId)) {
          throw new Error(`Manifest does not contain public workbook source ${sourceId}`)
        }
        const beforeArtifactIds = new Set(manifest.artifacts.map((artifact) => artifact.id))
        const checkpointProgress: PublicWorkbookCorpusFetchCheckpointProgress[] = []
        const fetchedManifest = await fetchPublicWorkbookArtifacts({
          manifest,
          cacheDir,
          limit: manifest.artifacts.length + 1,
          downloadTimeoutMs: readNumberArg('--download-timeout-ms', defaultDownloadTimeoutMs),
          ...fetchRunArgs,
          fingerprintTimeoutMs: readNumberArg('--fingerprint-timeout-ms', defaultFingerprintTimeoutMs),
          fingerprintMaxRssBytes: readMegabytesArg('--fingerprint-max-rss-mb', defaultFingerprintMaxRssBytes),
          isolatedFingerprinting: !inProcessFingerprinting,
          maxBytes: readNumberArg('--max-bytes', 50 * 1024 * 1024),
          sourceIds: [sourceId],
          onArtifactsCommitted: (checkpointManifest, progress) => {
            writeJson(manifestPath, checkpointManifest, 'public-workbook-corpus-manifest')
            checkpointProgress.push(progress)
          },
        })
        writeJson(manifestPath, fetchedManifest, 'public-workbook-corpus-manifest')
        return {
          artifactCountBefore: manifest.artifacts.length,
          artifactCountAfter: fetchedManifest.artifacts.length,
          checkpointProgress,
          fetchedArtifactIds: fetchedManifest.artifacts
            .filter((artifact) => !beforeArtifactIds.has(artifact.id))
            .map((artifact) => artifact.id),
        }
      })
    } finally {
      stopSelfRssGuard()
    }
    process.stdout.write(
      `${JSON.stringify(
        {
          sourceId,
          ...result,
          nextVerifyArtifactCommands: result.fetchedArtifactIds.flatMap((artifactId) => {
            const verifyCommand = splitPublicWorkbookCorpusVerifyArtifactCommand({
              artifactId,
              cacheDir,
              manifestPath,
              stopMarkerActive: existsSync(corpusRunStopMarkerPath),
              verifyCheckpointPath,
            })
            return verifyCommand.command ? [verifyCommand.command] : []
          }),
          blockedVerifyArtifactCommands: result.fetchedArtifactIds.flatMap((artifactId) => {
            const verifyCommand = splitPublicWorkbookCorpusVerifyArtifactCommand({
              artifactId,
              cacheDir,
              manifestPath,
              stopMarkerActive: existsSync(corpusRunStopMarkerPath),
              verifyCheckpointPath,
            })
            return verifyCommand.blockedCommand ? [verifyCommand.blockedCommand] : []
          }),
        },
        null,
        2,
      )}\n`,
    )
    return
  }
  if (command === 'discover-plan') {
    const plan = planPublicWorkbookCorpusFetch({
      manifest: readOrCreateManifest(manifestPath, targetWorkbookCount),
      limit: readNumberArg('--limit', targetWorkbookCount),
      sampleLimit: 0,
    })
    const stopMarkerActive = existsSync(corpusRunStopMarkerPath)
    const needsDiscovery = !plan.targetReachableFromKnownCandidates
    process.stdout.write(
      `${JSON.stringify(
        {
          stopMarker: publicWorkbookCorpusPlanStopMarker(stopMarkerActive),
          sourceCount: plan.sourceCount,
          targetArtifactCount: plan.targetArtifactCount,
          cachedArtifactCount: plan.cachedArtifactCount,
          remainingArtifactSlots: plan.remainingArtifactSlots,
          candidateSourceCount: plan.candidateSourceCount,
          candidateSourceDeficitCount: plan.candidateSourceDeficitCount,
          minimumAdditionalSourceCount: plan.minimumAdditionalSourceCount,
          recommendedDiscoveryLimit: plan.recommendedDiscoveryLimit,
          recommendedDiscoveryCommand:
            !needsDiscovery || stopMarkerActive
              ? null
              : formatPublicWorkbookCorpusDiscoverCommand({
                  cacheDir,
                  limit: plan.recommendedDiscoveryLimit,
                  manifestPath,
                }),
          blockedCommands:
            needsDiscovery && stopMarkerActive
              ? {
                  discover: formatPublicWorkbookCorpusDiscoverCommand({
                    cacheDir,
                    limit: plan.recommendedDiscoveryLimit,
                    manifestPath,
                    stopMarkerActive: true,
                  }),
                }
              : {},
          targetReachableFromKnownCandidates: plan.targetReachableFromKnownCandidates,
        },
        null,
        2,
      )}\n`,
    )
    return
  }
  if (command === 'fingerprint-artifact') {
    await writeFingerprintArtifactResult({
      filePath: readStringArg('--file', ''),
      fileName: readStringArg('--file-name', 'workbook.xlsx'),
      fingerprintTimeoutMs: readNumberArg('--fingerprint-timeout-ms', defaultFingerprintTimeoutMs),
      fingerprintMaxRssBytes: readMegabytesArg('--fingerprint-max-rss-mb', defaultFingerprintMaxRssBytes),
    })
    return
  }
  if (command === 'fingerprint-artifact-worker') {
    writeFingerprintArtifactWorkerResult({
      filePath: readStringArg('--file', ''),
      fileName: readStringArg('--file-name', 'workbook.xlsx'),
      fingerprintMaxRssBytes: readMegabytesArg('--fingerprint-max-rss-mb', defaultFingerprintMaxRssBytes),
    })
    return
  }
  if (command === 'footprint-worker') {
    writeFootprintWorkerResult({
      fileName: readStringArg('--file-name', 'workbook.xlsx'),
      verifyMaxRssBytes: capVerifyMaxRssBytes(readMegabytesArg('--verify-max-rss-mb', defaultVerifyMaxRssBytes)),
    })
    return
  }
  if (command === 'verify-artifact') {
    const artifactId = readStringArg('--artifact-id', '')
    if (!artifactId) {
      throw new Error('Expected --artifact-id for verify-artifact')
    }
    const updateVerifyCheckpoint = readFlagArg('--update-verify-checkpoint')
    if (updateVerifyCheckpoint) {
      assertPublicCorpusRunNotStopped({
        commandName: 'public-workbook-corpus verify-artifact',
        stopMarkerPath: corpusRunStopMarkerPath,
      })
    }
    const manifest = readManifest(manifestPath)
    const artifact = manifest.artifacts.find((entry) => entry.id === artifactId)
    if (!artifact) {
      throw new Error(`Manifest does not contain public workbook artifact ${artifactId}`)
    }
    const result = await verifyCachedWorkbookArtifactIsolated({
      artifact,
      cacheDir,
      manifestPath,
      runStructuralSmoke: readFlagArg('--structural-smoke'),
      timeoutMs: readNumberArg('--verify-timeout-ms', defaultVerifyTimeoutMs),
      maxRssBytes: capVerifyMaxRssBytes(readMegabytesArg('--verify-max-rss-mb', defaultVerifyMaxRssBytes)),
      maxCellCount: readNumberArg('--verify-max-cells', defaultVerifyMaxCellCount),
      rssCheckIntervalMs: 250,
    })
    if (updateVerifyCheckpoint) {
      await withPublicWorkbookCorpusCacheLock(cacheDir, 'verify-artifact-checkpoint', async () => {
        upsertPublicWorkbookCorpusVerificationCheckpoint({
          path: verifyCheckpointPath,
          manifest,
          verifiedCase: result,
        })
      })
    }
    process.stdout.write(`${JSON.stringify(result)}\n`)
    return
  }
  if (command === 'verify-artifact-worker') {
    const verifyMaxRssBytes = capVerifyMaxRssBytes(readMegabytesArg('--verify-max-rss-mb', defaultVerifyMaxRssBytes))
    const stopSelfRssGuard = startSelfRssGuard(verifyMaxRssBytes, 'Workbook verification worker')
    const artifactId = readStringArg('--artifact-id', '')
    try {
      if (!artifactId) {
        throw new Error('Expected --artifact-id for verify-artifact-worker')
      }
      const manifest = readManifest(manifestPath)
      const artifact = manifest.artifacts.find((entry) => entry.id === artifactId)
      if (!artifact) {
        throw new Error(`Manifest does not contain public workbook artifact ${artifactId}`)
      }
      const result = await verifyCachedWorkbookArtifact(
        artifact,
        cacheDir,
        readFlagArg('--structural-smoke'),
        readNumberArg('--verify-max-cells', defaultVerifyMaxCellCount),
        {
          timeoutMs: readNumberArg('--verify-timeout-ms', defaultVerifyTimeoutMs),
          maxRssBytes: verifyMaxRssBytes,
          rssCheckIntervalMs: defaultSelfRssCheckIntervalMs,
          onPhase: (phase) => {
            process.stderr.write(`bilig-public-workbook-verify-phase=${phase}\n`)
          },
        },
      )
      process.stdout.write(`${JSON.stringify(result)}\n`)
    } finally {
      stopSelfRssGuard()
    }
    return
  }
  if (command === 'verify') {
    const inProcessVerification = readDebugOnlyFlagArg(
      '--in-process',
      'BILIG_ALLOW_IN_PROCESS_PUBLIC_CORPUS_VERIFY',
      'it can retain workbook verification memory across large corpus runs',
    )
    const verifyConcurrency = readVerifyConcurrencyArg(defaultVerifyConcurrency)
    const verifyMaxRssBytes = capVerifyMaxRssBytes(readMegabytesArg('--verify-max-rss-mb', defaultVerifyMaxRssBytes))
    assertPublicCorpusRunNotStopped({
      commandName: 'public-workbook-corpus verify',
      stopMarkerPath: corpusRunStopMarkerPath,
    })
    const scorecard = await withPublicWorkbookCorpusCacheLock(cacheDir, 'verify', async () => {
      const manifest = readManifest(manifestPath)
      const checkpointInterval = readNumberArg('--verify-checkpoint-interval', 10)
      const reusableCases = readFlagArg('--no-verify-resume')
        ? []
        : readReusablePublicWorkbookCorpusCases([scorecardPath, verifyCheckpointPath])
      const checkpointCasesById = readFlagArg('--no-verify-resume')
        ? new Map<string, PublicWorkbookCorpusCase>()
        : new Map(readReusablePublicWorkbookCorpusCases([verifyCheckpointPath]).map((entry) => [entry.id, entry]))
      const progressIntervalMs = Math.max(1_000, readNumberArg('--verify-progress-interval-ms', 15_000))
      let completedCount = 0
      let latestArtifactId = 'none'
      const startedAt = Date.now()
      const progressTimer = setInterval(() => {
        const elapsedSeconds = Math.max(1, Math.round((Date.now() - startedAt) / 1000))
        console.error(
          `Public workbook corpus verify progress: ${String(completedCount)}/${String(manifest.artifacts.length)} completed; latest=${latestArtifactId}; elapsed=${String(elapsedSeconds)}s`,
        )
      }, progressIntervalMs)
      progressTimer.unref()
      try {
        const verifiedScorecard = await buildPublicWorkbookCorpusScorecard({
          manifest,
          cacheDir,
          manifestPath,
          isolatedVerification: !inProcessVerification,
          structuralSmokeSampleLimit: readNumberArg('--structural-smoke-sample-limit', 50),
          verifyConcurrency,
          verifyTimeoutMs: readNumberArg('--verify-timeout-ms', defaultVerifyTimeoutMs),
          verifyMaxRssBytes,
          verifyMaxCellCount: readNumberArg('--verify-max-cells', defaultVerifyMaxCellCount),
          reusableCases,
          onCaseVerified: (progress) => {
            completedCount = progress.completedCount
            latestArtifactId = progress.latestCase.id
            checkpointCasesById.set(progress.latestCase.id, progress.latestCase)
            if (progress.completedCount % checkpointInterval === 0 || progress.completedCount === progress.totalCount) {
              writePublicWorkbookCorpusVerificationCheckpoint({
                path: verifyCheckpointPath,
                manifest,
                casesById: checkpointCasesById,
              })
            }
            if (
              progress.completedCount === progress.totalCount ||
              progress.completedCount % 50 === 0 ||
              progress.latestCase.status === 'failed' ||
              progress.latestCase.status === 'error'
            ) {
              console.error(
                `Public workbook corpus verified ${String(progress.completedCount)}/${String(progress.totalCount)}; latest=${progress.latestCase.id}; status=${progress.latestCase.status}`,
              )
            }
          },
        })
        writeJson(scorecardPath, verifiedScorecard, 'public-workbook-corpus-scorecard')
        return verifiedScorecard
      } finally {
        clearInterval(progressTimer)
      }
    })
    console.log(
      `Verified ${String(scorecard.summary.cachedWorkbookCount)} cached workbooks; ${String(scorecard.summary.remainingToTarget)} remaining`,
    )
    return
  }
  if (command === 'verify-missing') {
    await runPublicWorkbookCorpusVerifyMissingCommand({
      cacheDir,
      corpusRunStopMarkerPath,
      manifestPath,
      scorecardPath,
      verifyCheckpointPath,
    })
    return
  }
  if (command === 'verify-stale') {
    await runPublicWorkbookCorpusVerifyStaleCommand({
      cacheDir,
      corpusRunStopMarkerPath,
      manifestPath,
      scorecardPath,
      verifyCheckpointPath,
    })
    return
  }
  if (command === 'refresh-scorecard-from-checkpoint') {
    const manifest = readManifest(manifestPath)
    const recordedCases = readReusablePublicWorkbookCorpusCases([scorecardPath, verifyCheckpointPath])
    const refreshed = buildPublicWorkbookCorpusScorecardFromCases({
      manifest: selectManifestArtifactsWithRecordedCases(manifest, recordedCases),
      cases: selectRecordedCasesInManifestOrder(manifest.artifacts, recordedCases),
      generatedAt: existingScorecardGeneratedAt(scorecardPath),
    })
    validatePublicWorkbookCorpusScorecard(refreshed)
    const serialized = serializeJsonForRepo(refreshed, 'public-workbook-corpus-scorecard')
    const summary = {
      mode: readFlagArg('--check') ? 'check' : readFlagArg('--dry-run') ? 'dry-run' : 'write',
      outputPath: scorecardPath,
      targetWorkbookCount: refreshed.summary.targetWorkbookCount,
      sourceCount: refreshed.summary.sourceCount,
      cachedWorkbookCount: refreshed.summary.cachedWorkbookCount,
      passedWorkbookCount: refreshed.summary.passedWorkbookCount,
      formulaOracleComparisonCount: refreshed.summary.formulaOracleComparisonCount,
      remainingToTarget: refreshed.summary.remainingToTarget,
    }
    if (refreshed.cases.length === 0) {
      throw new Error('No reusable checkpoint cases match the current public workbook manifest')
    }
    if (readFlagArg('--check')) {
      const existing = existsSync(scorecardPath) ? readFileSync(scorecardPath, 'utf8') : ''
      if (existing !== serialized) {
        throw new Error(
          `Public workbook corpus scorecard is stale: ${String(
            refreshed.summary.cachedWorkbookCount,
          )} checkpoint-backed cases are available for ${scorecardPath}`,
        )
      }
      process.stdout.write(`${JSON.stringify(summary, null, 2)}\n`)
      return
    }
    if (!readFlagArg('--dry-run')) {
      mkdirSync(dirname(scorecardPath), { recursive: true })
      writeFileSync(scorecardPath, serialized)
    }
    process.stdout.write(`${JSON.stringify(summary, null, 2)}\n`)
    return
  }
  if (command === 'check') {
    writePublicWorkbookCorpusCheck({
      manifestPath,
      scorecardPath,
      cacheDir,
      verifyCheckpointPath,
      skipManifestCheck: readFlagArg('--skip-manifest-check'),
      requireTarget: readFlagArg('--require-target'),
      corpusRunStopMarkerPath,
    })
    return
  }
  if (command === 'status') {
    writePublicWorkbookCorpusStatus({
      manifestPath,
      scorecardPath,
      cacheDir,
      verifyCheckpointPath,
      requireTarget: readFlagArg('--require-target'),
      corpusRunStopMarkerPath,
    })
    return
  }
  throw new Error(`Unknown public workbook corpus command: ${command}`)
}

function formatFetchCheckpointProgress(progress: PublicWorkbookCorpusFetchCheckpointProgress): string {
  const parts = [
    `Cached ${String(progress.artifactCount)} public workbook artifacts`,
    `exhausted ${String(progress.exhaustedSourceCount)} sources`,
    `+${String(progress.exhaustedSourceDelta)} exhausted this batch`,
    `${String(progress.committedArtifactCount)} committed`,
    `${String(progress.failedSourceCount)} failed`,
    `${String(progress.duplicateHashSourceCount)} duplicate hashes`,
    `${String(progress.duplicateFingerprintSourceCount)} duplicate fingerprints`,
  ]
  if (progress.failedSourceSamples.length > 0) {
    parts.push(`failure samples: ${progress.failedSourceSamples.map(formatFetchFailureSample).join(' | ')}`)
  }
  return parts.join('; ')
}

function formatFetchFailureSample(sample: PublicWorkbookCorpusFetchCheckpointProgress['failedSourceSamples'][number]): string {
  return `${sample.sourceId} ${sample.fileName}: ${sample.error.replace(/\s+/gu, ' ').slice(0, 240)}`
}

function selectManifestArtifactsWithRecordedCases(
  manifest: PublicWorkbookManifest,
  recordedCases: readonly PublicWorkbookCorpusCase[],
): PublicWorkbookManifest {
  return {
    ...manifest,
    artifacts: selectRecordedArtifactsInManifestOrder(manifest.artifacts, recordedCases),
  }
}

function existingScorecardGeneratedAt(scorecardPath: string): string | undefined {
  if (!existsSync(scorecardPath)) {
    return undefined
  }
  const parsed: unknown = JSON.parse(readFileSync(scorecardPath, 'utf8'))
  if (typeof parsed !== 'object' || parsed === null) {
    return undefined
  }
  const generatedAt = Reflect.get(parsed, 'generatedAt')
  return typeof generatedAt === 'string' && generatedAt.trim().length > 0 ? generatedAt : undefined
}

function selectRecordedArtifactsInManifestOrder(
  artifacts: readonly PublicWorkbookArtifact[],
  recordedCases: readonly PublicWorkbookCorpusCase[],
): PublicWorkbookArtifact[] {
  const casesById = indexPublicWorkbookCorpusCases(recordedCases)
  return artifacts.filter((artifact) => {
    const recordedCase = casesById.get(artifact.id)
    return recordedCase?.passed === true && publicWorkbookCorpusCaseMatchesArtifact(recordedCase, artifact)
  })
}

function selectRecordedCasesInManifestOrder(
  artifacts: readonly PublicWorkbookArtifact[],
  recordedCases: readonly PublicWorkbookCorpusCase[],
): PublicWorkbookCorpusCase[] {
  const casesById = indexPublicWorkbookCorpusCases(recordedCases)
  return artifacts.flatMap((artifact) => {
    const recordedCase = casesById.get(artifact.id)
    return recordedCase?.passed === true && publicWorkbookCorpusCaseMatchesArtifact(recordedCase, artifact) ? [recordedCase] : []
  })
}

function readPublicWorkbookLinkInput(commandName: string): PublicWorkbookLinkInput {
  const sourceUrl = readStringArg('--source-url', readStringArg('--url', ''))
  if (!sourceUrl) {
    throw new Error(`Expected --source-url for ${commandName}`)
  }
  return {
    sourceUrl,
    downloadUrl: readStringArg('--download-url', ''),
    fileName: readStringArg('--file-name', ''),
    licenseTitle: readStringArg('--license-title', ''),
    licenseUrl: readStringArg('--license-url', ''),
    licenseSpdxId: readStringArg('--license-spdx', '') || null,
  }
}

function addPublicWorkbookLinkSourceFromInput(manifest: PublicWorkbookManifest, input: PublicWorkbookLinkInput) {
  return addPublicWorkbookLinkSource({
    manifest,
    sourceUrl: input.sourceUrl,
    downloadUrl: input.downloadUrl,
    fileName: input.fileName,
    licenseTitle: input.licenseTitle,
    licenseUrl: input.licenseUrl,
    licenseSpdxId: input.licenseSpdxId,
  })
}

if (process.argv[1] && pathToFileURL(resolve(process.argv[1])).href === import.meta.url) {
  await main()
}
