import { existsSync, readFileSync } from 'node:fs'
import { isAbsolute, relative, resolve } from 'node:path'

import {
  assertPublicCorpusRunNotStopped,
  publicCorpusStopMarkerOverrideEnvVar,
  publicCorpusStopMarkerOverrideFlag,
  readDebugOnlyFlagArg,
  readFlagArg,
  readMegabytesArg,
  readNumberArg,
  readPublicCorpusVerificationBatchLimitArg,
  readVerifyConcurrencyArg,
  readVerifyMissingLimitArg,
} from './public-workbook-corpus-cli.ts'
import { parsePublicWorkbookManifestJson } from './public-workbook-corpus-json.ts'
import { withPublicWorkbookCorpusCacheLock } from './public-workbook-corpus-lock.ts'
import {
  indexPublicWorkbookCorpusCases,
  listMissingPublicWorkbookArtifacts,
  listStalePublicWorkbookArtifacts,
  selectMissingPublicWorkbookArtifacts,
  selectStalePublicWorkbookArtifacts,
} from './public-workbook-corpus-missing.ts'
import { publicWorkbookCorpusCaseEvidenceRefreshReasons } from './public-workbook-corpus-evidence.ts'
import {
  readReusablePublicWorkbookCorpusCases,
  writePublicWorkbookCorpusVerificationCheckpoint,
} from './public-workbook-corpus-verify-checkpoint.ts'
import {
  buildPublicWorkbookCorpusScorecard,
  capVerifyMaxRssBytes,
  defaultVerifyConcurrency,
  defaultVerifyMaxCellCount,
  defaultVerifyMaxRssBytes,
  defaultVerifyTimeoutMs,
} from './public-workbook-corpus-verify.ts'
import type { PublicWorkbookCorpusCase, PublicWorkbookManifest } from './public-workbook-corpus-types.ts'

export interface PublicWorkbookCorpusVerifySliceCommandArgs {
  readonly cacheDir: string
  readonly corpusRunStopMarkerPath: string
  readonly manifestPath: string
  readonly scorecardPath: string
  readonly verifyCheckpointPath: string
}

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const largeVerifyStaleLimitEnvVar = 'BILIG_ALLOW_LARGE_PUBLIC_CORPUS_VERIFY_STALE'

export async function runPublicWorkbookCorpusVerifyMissingCommand(args: PublicWorkbookCorpusVerifySliceCommandArgs): Promise<void> {
  const dryRun = readFlagArg('--dry-run') || readFlagArg('--list')
  const limit = readVerifyMissingLimitArg(1, dryRun)
  if (dryRun) {
    const manifest = readManifest(args.manifestPath)
    const recordedCases = readReusablePublicWorkbookCorpusCases([args.scorecardPath, args.verifyCheckpointPath])
    const missingArtifacts = listMissingPublicWorkbookArtifacts({ manifest, cases: recordedCases })
    const selectedArtifactCount = Math.min(limit, missingArtifacts.length)
    process.stdout.write(
      `${JSON.stringify(
        {
          totalMissingArtifactCount: missingArtifacts.length,
          selectedArtifactCount,
          artifacts: missingArtifacts.slice(0, limit).map((artifact) => ({
            id: artifact.id,
            fileName: artifact.fileName,
            byteSize: artifact.byteSize,
            sourceUrl: artifact.sourceUrl,
            cachePath: artifact.cachePath,
          })),
          nextVerificationCommand:
            selectedArtifactCount > 0 ? formatVerifySliceCommand({ ...args, limit: selectedArtifactCount, slice: 'verify-missing' }) : null,
        },
        null,
        2,
      )}\n`,
    )
    return
  }
  const verifiedCount = await verifySelectedPublicWorkbookArtifacts({
    ...args,
    cacheLockName: 'verify-missing',
    commandName: 'public-workbook-corpus verify-missing',
    limit,
    progressLabel: 'verified missing',
    selectArtifacts: (manifest, recordedCases) => selectMissingPublicWorkbookArtifacts({ manifest, cases: recordedCases, limit }),
    structuralSmokeSampleLimit: 0,
  })
  console.log(`Verified ${String(verifiedCount)} missing public workbook cases into ${formatCommandPath(args.verifyCheckpointPath)}`)
}

export async function runPublicWorkbookCorpusVerifyStaleCommand(args: PublicWorkbookCorpusVerifySliceCommandArgs): Promise<void> {
  const dryRun = readFlagArg('--dry-run') || readFlagArg('--list')
  const limit = readPublicCorpusVerificationBatchLimitArg(1, dryRun, {
    commandName: 'verify-stale',
    envVar: largeVerifyStaleLimitEnvVar,
  })
  if (dryRun) {
    const manifest = readManifest(args.manifestPath)
    const recordedCases = readReusablePublicWorkbookCorpusCases([args.scorecardPath, args.verifyCheckpointPath])
    const staleArtifacts = listStalePublicWorkbookArtifacts({ manifest, cases: recordedCases })
    const recordedCasesById = indexPublicWorkbookCorpusCases(recordedCases)
    const selectedArtifactCount = Math.min(limit, staleArtifacts.length)
    process.stdout.write(
      `${JSON.stringify(
        {
          totalStaleArtifactCount: staleArtifacts.length,
          selectedArtifactCount,
          artifacts: staleArtifacts.slice(0, limit).map((artifact) => ({
            id: artifact.id,
            fileName: artifact.fileName,
            byteSize: artifact.byteSize,
            sourceUrl: artifact.sourceUrl,
            cachePath: artifact.cachePath,
            reason: staleArtifactReason(artifact.id, recordedCasesById),
          })),
          nextVerificationCommand:
            selectedArtifactCount > 0 ? formatVerifySliceCommand({ ...args, limit: selectedArtifactCount, slice: 'verify-stale' }) : null,
        },
        null,
        2,
      )}\n`,
    )
    return
  }
  const verifiedCount = await verifySelectedPublicWorkbookArtifacts({
    ...args,
    cacheLockName: 'verify-stale',
    commandName: 'public-workbook-corpus verify-stale',
    limit,
    progressLabel: 'refreshed stale',
    selectArtifacts: (manifest, recordedCases) => selectStalePublicWorkbookArtifacts({ manifest, cases: recordedCases, limit }),
    structuralSmokeSampleLimit: readNumberArg('--structural-smoke-sample-limit', 50),
  })
  console.log(`Refreshed ${String(verifiedCount)} stale public workbook cases into ${formatCommandPath(args.verifyCheckpointPath)}`)
}

async function verifySelectedPublicWorkbookArtifacts(
  args: PublicWorkbookCorpusVerifySliceCommandArgs & {
    readonly cacheLockName: string
    readonly commandName: string
    readonly limit: number
    readonly progressLabel: string
    readonly selectArtifacts: (
      manifest: PublicWorkbookManifest,
      recordedCases: ReturnType<typeof readReusablePublicWorkbookCorpusCases>,
    ) => PublicWorkbookManifest['artifacts']
    readonly structuralSmokeSampleLimit: number
  },
): Promise<number> {
  const inProcessVerification = readDebugOnlyFlagArg(
    '--in-process',
    'BILIG_ALLOW_IN_PROCESS_PUBLIC_CORPUS_VERIFY',
    'it can retain workbook verification memory across large corpus runs',
  )
  const verifyConcurrency = readVerifyConcurrencyArg(defaultVerifyConcurrency)
  const verifyMaxRssBytes = capVerifyMaxRssBytes(readMegabytesArg('--verify-max-rss-mb', defaultVerifyMaxRssBytes))
  assertPublicCorpusRunNotStopped({
    commandName: args.commandName,
    stopMarkerPath: args.corpusRunStopMarkerPath,
  })
  return await withPublicWorkbookCorpusCacheLock(args.cacheDir, args.cacheLockName, async () => {
    const manifest = readManifest(args.manifestPath)
    const recordedCases = readReusablePublicWorkbookCorpusCases([args.scorecardPath, args.verifyCheckpointPath])
    const selectedArtifacts = args.selectArtifacts(manifest, recordedCases)
    if (selectedArtifacts.length === 0) {
      return 0
    }
    const checkpointCasesById = indexPublicWorkbookCorpusCases(readReusablePublicWorkbookCorpusCases([args.verifyCheckpointPath]))
    const scorecard = await buildPublicWorkbookCorpusScorecard({
      manifest: { ...manifest, artifacts: selectedArtifacts },
      cacheDir: args.cacheDir,
      manifestPath: args.manifestPath,
      isolatedVerification: !inProcessVerification,
      structuralSmokeSampleLimit: args.structuralSmokeSampleLimit,
      verifyConcurrency,
      verifyTimeoutMs: readNumberArg('--verify-timeout-ms', defaultVerifyTimeoutMs),
      verifyMaxRssBytes,
      verifyMaxCellCount: readNumberArg('--verify-max-cells', defaultVerifyMaxCellCount),
      reusableCases: [],
      onCaseVerified: (progress) => {
        checkpointCasesById.set(progress.latestCase.id, progress.latestCase)
        writePublicWorkbookCorpusVerificationCheckpoint({
          path: args.verifyCheckpointPath,
          manifest,
          casesById: checkpointCasesById,
        })
        console.error(
          `Public workbook corpus ${args.progressLabel} ${String(progress.completedCount)}/${String(
            progress.totalCount,
          )}; latest=${progress.latestCase.id}; status=${progress.latestCase.status}`,
        )
      },
    })
    return scorecard.cases.length
  })
}

function staleArtifactReason(
  artifactId: string,
  recordedCasesById: ReadonlyMap<string, PublicWorkbookCorpusCase>,
): ReturnType<typeof publicWorkbookCorpusCaseEvidenceRefreshReasons>[number] {
  const recordedCase = recordedCasesById.get(artifactId)
  const reasons = recordedCase ? publicWorkbookCorpusCaseEvidenceRefreshReasons(recordedCase) : []
  return reasons[0] ?? 'missing-used-range-evidence'
}

function readManifest(path: string): PublicWorkbookManifest {
  const parsed: unknown = JSON.parse(readFileSync(path, 'utf8'))
  return parsePublicWorkbookManifestJson(parsed)
}

function formatCommandPath(path: string): string {
  const absolutePath = resolve(path)
  const relativePath = relative(rootDir, absolutePath)
  return relativePath && !relativePath.startsWith('..') && !isAbsolute(relativePath) ? relativePath : path
}

function formatVerifySliceCommand(
  args: PublicWorkbookCorpusVerifySliceCommandArgs & {
    readonly limit: number
    readonly slice: 'verify-missing' | 'verify-stale'
  },
): string {
  const command = [
    'pnpm',
    `public-workbook-corpus:${args.slice}`,
    '--',
    '--manifest',
    formatCommandPath(args.manifestPath),
    '--scorecard',
    formatCommandPath(args.scorecardPath),
    '--verify-checkpoint',
    formatCommandPath(args.verifyCheckpointPath),
    '--cache-dir',
    formatCommandPath(args.cacheDir),
    '--limit',
    String(args.limit),
  ]
  if (!existsSync(args.corpusRunStopMarkerPath)) {
    return command.map(shellQuote).join(' ')
  }
  return `${publicCorpusStopMarkerOverrideEnvVar}=1 ${[...command, publicCorpusStopMarkerOverrideFlag].map(shellQuote).join(' ')}`
}

function shellQuote(value: string): string {
  return /^[A-Za-z0-9_./:=@+-]+$/u.test(value) ? value : `'${value.replaceAll("'", "'\\''")}'`
}
