#!/usr/bin/env bun

import { existsSync, readFileSync } from 'node:fs'
import { isAbsolute, join, relative, resolve } from 'node:path'
import { pathToFileURL } from 'node:url'

import { publicWorkbookCorpusCaseNeedsEvidenceRefresh } from './public-workbook-corpus-evidence.ts'
import { parsePublicWorkbookManifestJson } from './public-workbook-corpus-json.ts'
import { publicWorkbookCorpusCaseMatchesArtifact } from './public-workbook-corpus-missing.ts'
import { readReusablePublicWorkbookCorpusCases } from './public-workbook-corpus-verify-checkpoint.ts'
import {
  publicCorpusStopMarkerOverrideEnvVar,
  publicCorpusStopMarkerOverrideFlag,
  readNumberArg,
  readStringArg,
} from './public-workbook-corpus-cli.ts'
import {
  buildUnsupportedClassificationCounts,
  type PublicWorkbookCorpusUnsupportedClassificationCount,
} from './public-workbook-corpus-completion-audit-helpers.ts'
import type { PublicWorkbookArtifact, PublicWorkbookCorpusCase, PublicWorkbookManifest } from './public-workbook-corpus-types.ts'

export interface PublicWorkbookCorpusResourceLimitPlan {
  readonly schemaVersion: 1
  readonly mode: 'resource-limit-plan'
  readonly generatedAt: string
  readonly stopMarker: {
    readonly active: boolean
    readonly path: string
    readonly overrideFlag: string
    readonly overrideEnvVar: string
  }
  readonly currentState: {
    readonly manifestArtifactCount: number
    readonly recordedCaseCount: number
    readonly resourceLimitCaseCount: number
    readonly currentResourceLimitCaseCount: number
    readonly staleResourceLimitCaseCount: number
    readonly currentClassifications: readonly PublicWorkbookCorpusUnsupportedClassificationCount[]
    readonly staleClassifications: readonly PublicWorkbookCorpusUnsupportedClassificationCount[]
  }
  readonly currentSamples: readonly PublicWorkbookCorpusResourceLimitPlanEntry[]
  readonly staleSamples: readonly PublicWorkbookCorpusResourceLimitPlanEntry[]
}

export interface PublicWorkbookCorpusResourceLimitPlanEntry {
  readonly id: string
  readonly fileName: string
  readonly byteSize: number
  readonly cachePath: string
  readonly sourceUrl: string
  readonly classifications: readonly string[]
  readonly rssEvidence: readonly string[]
  readonly probeCommand: string
  readonly checkpointRefreshCommand: string
}

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const defaultCacheDir = join(rootDir, '.cache', 'public-workbook-corpus')
const defaultManifestPath = join(defaultCacheDir, 'manifest.json')
const defaultScorecardPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'public-workbook-corpus-scorecard.json')
const defaultVerifyCheckpointPath = join(defaultCacheDir, 'verification-checkpoint.json')
const defaultCorpusRunStopMarkerPath = join(rootDir, '.agent-coordination', '20260507T074946Z-codex-stop-interactive-corpus-runs.md')
const resourceLimitClassificationPrefix = 'xlsx.publicCorpus.resourceLimit:'
const defaultVerifyMaxRssMiB = 1536

function main(): void {
  const plan = buildPublicWorkbookCorpusResourceLimitPlanFromArgs()
  process.stdout.write(`${JSON.stringify(plan, null, 2)}\n`)
}

export function buildPublicWorkbookCorpusResourceLimitPlanFromArgs(): PublicWorkbookCorpusResourceLimitPlan {
  const cacheDir = resolve(readStringArg('--cache-dir', defaultCacheDir))
  const manifestPath = resolve(readStringArg('--manifest', defaultManifestPath))
  const scorecardPath = resolve(readStringArg('--scorecard', defaultScorecardPath))
  const verifyCheckpointPath = resolve(readStringArg('--verify-checkpoint', defaultVerifyCheckpointPath))
  const stopMarkerPath = resolve(readStringArg('--corpus-run-stop-marker', defaultCorpusRunStopMarkerPath))
  const displayRootDir = resolve(readStringArg('--display-root', rootDir))
  const sampleLimit = readNumberArg('--sample-limit', 20)
  const verifyMaxRssMiB = readNumberArg('--verify-max-rss-mb', defaultVerifyMaxRssMiB)
  const generatedAt = readStringArg('--generated-at', new Date().toISOString())
  const manifest = parsePublicWorkbookManifestJson(JSON.parse(readFileSync(manifestPath, 'utf8')))
  const recordedCases = readReusablePublicWorkbookCorpusCases([scorecardPath, verifyCheckpointPath])
  return buildPublicWorkbookCorpusResourceLimitPlan({
    cacheDir,
    displayRootDir,
    generatedAt,
    manifest,
    manifestPath,
    recordedCases,
    sampleLimit,
    scorecardPath,
    stopMarkerActive: existsSync(stopMarkerPath),
    stopMarkerPath,
    verifyCheckpointPath,
    verifyMaxRssMiB,
  })
}

export function buildPublicWorkbookCorpusResourceLimitPlan(args: {
  readonly cacheDir: string
  readonly displayRootDir?: string
  readonly generatedAt: string
  readonly manifest: PublicWorkbookManifest
  readonly manifestPath: string
  readonly recordedCases: readonly PublicWorkbookCorpusCase[]
  readonly sampleLimit: number
  readonly scorecardPath: string
  readonly stopMarkerActive: boolean
  readonly stopMarkerPath: string
  readonly verifyCheckpointPath: string
  readonly verifyMaxRssMiB: number
}): PublicWorkbookCorpusResourceLimitPlan {
  const casesById = new Map(args.recordedCases.map((entry) => [entry.id, entry]))
  const matchedEntries = args.manifest.artifacts.flatMap((artifact) => {
    const recordedCase = casesById.get(artifact.id)
    return recordedCase && publicWorkbookCorpusCaseMatchesArtifact(recordedCase, artifact) ? [{ artifact, recordedCase }] : []
  })
  const resourceLimitEntries = matchedEntries.filter(({ recordedCase }) =>
    recordedCase.unsupportedFeatureClassifications.some((entry) => entry.startsWith(resourceLimitClassificationPrefix)),
  )
  const currentEntries = resourceLimitEntries.filter(({ recordedCase }) => !publicWorkbookCorpusCaseNeedsEvidenceRefresh(recordedCase))
  const staleEntries = resourceLimitEntries.filter(({ recordedCase }) => publicWorkbookCorpusCaseNeedsEvidenceRefresh(recordedCase))
  return {
    schemaVersion: 1,
    mode: 'resource-limit-plan',
    generatedAt: args.generatedAt,
    stopMarker: {
      active: args.stopMarkerActive,
      path: commandPath(args.stopMarkerPath, args.displayRootDir),
      overrideFlag: publicCorpusStopMarkerOverrideFlag,
      overrideEnvVar: publicCorpusStopMarkerOverrideEnvVar,
    },
    currentState: {
      manifestArtifactCount: args.manifest.artifacts.length,
      recordedCaseCount: matchedEntries.length,
      resourceLimitCaseCount: resourceLimitEntries.length,
      currentResourceLimitCaseCount: currentEntries.length,
      staleResourceLimitCaseCount: staleEntries.length,
      currentClassifications: buildUnsupportedClassificationCounts(currentEntries.map((entry) => entry.recordedCase)),
      staleClassifications: buildUnsupportedClassificationCounts(staleEntries.map((entry) => entry.recordedCase)),
    },
    currentSamples: samplePlanEntries(currentEntries, args),
    staleSamples: samplePlanEntries(staleEntries, args),
  }
}

function samplePlanEntries(
  entries: readonly { readonly artifact: PublicWorkbookArtifact; readonly recordedCase: PublicWorkbookCorpusCase }[],
  args: Parameters<typeof buildPublicWorkbookCorpusResourceLimitPlan>[0],
): PublicWorkbookCorpusResourceLimitPlanEntry[] {
  return entries
    .toSorted((left, right) => right.artifact.byteSize - left.artifact.byteSize || left.artifact.id.localeCompare(right.artifact.id))
    .slice(0, Math.max(0, Math.trunc(args.sampleLimit)))
    .map(({ artifact, recordedCase }) => ({
      id: artifact.id,
      fileName: artifact.fileName,
      byteSize: artifact.byteSize,
      cachePath: artifact.cachePath,
      sourceUrl: artifact.sourceUrl,
      classifications: recordedCase.unsupportedFeatureClassifications.filter((entry) =>
        entry.startsWith(resourceLimitClassificationPrefix),
      ),
      rssEvidence: recordedCase.evidence.filter((entry) => entry.startsWith('Public corpus verification RSS limit exceeded:')),
      probeCommand: formatVerifyArtifactCommand({ ...args, artifactId: artifact.id, updateCheckpoint: false }),
      checkpointRefreshCommand: formatVerifyArtifactCommand({ ...args, artifactId: artifact.id, updateCheckpoint: true }),
    }))
}

function formatVerifyArtifactCommand(
  args: Parameters<typeof buildPublicWorkbookCorpusResourceLimitPlan>[0] & {
    readonly artifactId: string
    readonly updateCheckpoint: boolean
  },
): string {
  const command = [
    'pnpm',
    'public-workbook-corpus:verify-artifact',
    '--',
    '--manifest',
    commandPath(args.manifestPath, args.displayRootDir),
    '--cache-dir',
    commandPath(args.cacheDir, args.displayRootDir),
    '--verify-checkpoint',
    commandPath(args.verifyCheckpointPath, args.displayRootDir),
    '--artifact-id',
    args.artifactId,
    '--verify-max-rss-mb',
    String(args.verifyMaxRssMiB),
    ...(args.updateCheckpoint ? ['--update-verify-checkpoint'] : []),
  ]
  const formatted = command.map(shellQuote).join(' ')
  return args.updateCheckpoint && args.stopMarkerActive
    ? `${publicCorpusStopMarkerOverrideEnvVar}=1 ${formatted} ${publicCorpusStopMarkerOverrideFlag}`
    : formatted
}

function commandPath(path: string, displayRootDir: string | undefined): string {
  if (!displayRootDir) {
    return path
  }
  const relativePath = relative(displayRootDir, resolve(path))
  return relativePath && !relativePath.startsWith('..') && !isAbsolute(relativePath) ? relativePath : path
}

function shellQuote(value: string): string {
  return /^[A-Za-z0-9_./:=@+-]+$/u.test(value) ? value : `'${value.replaceAll("'", "'\\''")}'`
}

if (process.argv[1] && pathToFileURL(resolve(process.argv[1])).href === import.meta.url) {
  main()
}
