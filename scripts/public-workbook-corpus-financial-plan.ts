#!/usr/bin/env bun

import { existsSync, readFileSync } from 'node:fs'
import { isAbsolute, join, relative, resolve } from 'node:path'

import { createEmptyPublicWorkbookManifest, parsePublicWorkbookManifestJson } from './public-workbook-corpus-json.ts'
import { planPublicWorkbookCorpusFetch } from './public-workbook-corpus-fetch.ts'
import {
  publicCorpusStopMarkerOverrideEnvVar,
  publicCorpusStopMarkerOverrideFlag,
  readNumberArg,
  readStringArg,
} from './public-workbook-corpus-cli.ts'
import type { PublicWorkbookManifest } from './public-workbook-corpus-types.ts'

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const defaultFinancialCacheDir = join(rootDir, '.cache', 'public-workbook-corpus-financial')
const defaultFinancialManifestPath = join(defaultFinancialCacheDir, 'manifest.json')
const defaultFinancialScorecardPath = join(defaultFinancialCacheDir, 'scorecard.json')
const defaultFinancialVerifyCheckpointPath = join(defaultFinancialCacheDir, 'verification-checkpoint.json')
const defaultCorpusRunStopMarkerPath = join(rootDir, '.agent-coordination', '20260507T074946Z-codex-stop-interactive-corpus-runs.md')

async function main(): Promise<void> {
  const cacheDir = resolve(readStringArg('--cache-dir', defaultFinancialCacheDir))
  const manifestPath = resolve(readStringArg('--manifest', defaultFinancialManifestPath))
  const scorecardPath = resolve(readStringArg('--scorecard', defaultFinancialScorecardPath))
  const verifyCheckpointPath = resolve(readStringArg('--verify-checkpoint', defaultFinancialVerifyCheckpointPath))
  const corpusRunStopMarkerPath = resolve(readStringArg('--corpus-run-stop-marker', defaultCorpusRunStopMarkerPath))
  const targetWorkbookCount = readNumberArg('--target-workbook-count', 5_000)
  const limit = readNumberArg('--limit', targetWorkbookCount)
  const sampleLimit = readNumberArg('--sample-limit', 20)
  const stopMarkerActive = existsSync(corpusRunStopMarkerPath)
  const manifest = readOrCreateFinancialManifest(manifestPath, targetWorkbookCount)
  const plan = planPublicWorkbookCorpusFetch({
    manifest,
    limit,
    sampleLimit,
  })
  const needsAdditionalDiscovery = !plan.targetReachableFromKnownCandidates

  process.stdout.write(
    `${JSON.stringify(
      {
        mode: 'plan',
        corpus: 'financial-accounting-workpapers',
        manifestExists: existsSync(manifestPath),
        targetWorkbookCount,
        manifestPath: formatCommandPath(manifestPath),
        cacheDir: formatCommandPath(cacheDir),
        scorecardPath: formatCommandPath(scorecardPath),
        verifyCheckpointPath: formatCommandPath(verifyCheckpointPath),
        stopMarker: {
          active: stopMarkerActive,
          path: formatCommandPath(corpusRunStopMarkerPath),
          overrideFlag: publicCorpusStopMarkerOverrideFlag,
          overrideEnvVar: publicCorpusStopMarkerOverrideEnvVar,
        },
        sourceCount: plan.sourceCount,
        targetArtifactCount: plan.targetArtifactCount,
        cachedArtifactCount: plan.cachedArtifactCount,
        remainingArtifactSlots: plan.remainingArtifactSlots,
        candidateSourceCount: plan.candidateSourceCount,
        candidateSourceDeficitCount: plan.candidateSourceDeficitCount,
        minimumAdditionalSourceCount: plan.minimumAdditionalSourceCount,
        recommendedDiscoveryLimit: Math.max(plan.recommendedDiscoveryLimit, targetWorkbookCount),
        needsAdditionalDiscovery,
        targetReachableFromKnownCandidates: plan.targetReachableFromKnownCandidates,
        commands: {
          discoverPlan: needsAdditionalDiscovery
            ? formatFinancialDiscoveryPlanCommand({ cacheDir, limit, manifestPath, targetWorkbookCount })
            : null,
          discover: needsAdditionalDiscovery
            ? formatFinancialDiscoveryCommand({
                cacheDir,
                limit: Math.max(plan.recommendedDiscoveryLimit, targetWorkbookCount),
                manifestPath,
                stopMarkerActive,
                targetWorkbookCount,
              })
            : null,
          fetchPlan: formatFinancialFetchPlanCommand({ cacheDir, limit, manifestPath }),
          fetch: formatFinancialFetchCommand({ cacheDir, limit, manifestPath, stopMarkerActive }),
          verify: formatFinancialVerifyCommand({ cacheDir, manifestPath, scorecardPath, stopMarkerActive, verifyCheckpointPath }),
          check: formatFinancialCheckCommand({ cacheDir, manifestPath, scorecardPath, verifyCheckpointPath }),
        },
        sampledCandidateSources: plan.sampledCandidateSources.map((source) => ({
          id: source.id,
          kind: source.kind,
          fileName: source.fileName,
          sourceUrl: source.sourceUrl,
          downloadUrl: source.downloadUrl,
          topicEvidence: source.topicEvidence ?? [],
        })),
      },
      null,
      2,
    )}\n`,
  )
}

function readOrCreateFinancialManifest(path: string, targetWorkbookCount: number): PublicWorkbookManifest {
  if (!existsSync(path)) {
    return createEmptyPublicWorkbookManifest(undefined, targetWorkbookCount)
  }
  return parsePublicWorkbookManifestJson(JSON.parse(readFileSync(path, 'utf8')))
}

function formatFinancialDiscoveryPlanCommand(args: {
  readonly cacheDir: string
  readonly limit: number
  readonly manifestPath: string
  readonly targetWorkbookCount: number
}): string {
  return formatCommand([
    'pnpm',
    'public-workbook-corpus:discover-financial:plan',
    '--',
    '--manifest',
    formatCommandPath(args.manifestPath),
    '--cache-dir',
    formatCommandPath(args.cacheDir),
    '--target-workbook-count',
    String(args.targetWorkbookCount),
    '--limit',
    String(args.limit),
  ])
}

function formatFinancialDiscoveryCommand(args: {
  readonly cacheDir: string
  readonly limit: number
  readonly manifestPath: string
  readonly stopMarkerActive: boolean
  readonly targetWorkbookCount: number
}): string {
  return formatMutatingCommand(
    [
      'pnpm',
      'public-workbook-corpus:discover-financial',
      '--',
      '--manifest',
      formatCommandPath(args.manifestPath),
      '--cache-dir',
      formatCommandPath(args.cacheDir),
      '--target-workbook-count',
      String(args.targetWorkbookCount),
      '--limit',
      String(args.limit),
    ],
    args.stopMarkerActive,
  )
}

function formatFinancialFetchPlanCommand(args: {
  readonly cacheDir: string
  readonly limit: number
  readonly manifestPath: string
}): string {
  return formatCommand([
    'pnpm',
    'public-workbook-corpus:fetch-financial:plan',
    '--',
    '--manifest',
    formatCommandPath(args.manifestPath),
    '--cache-dir',
    formatCommandPath(args.cacheDir),
    '--limit',
    String(args.limit),
  ])
}

function formatFinancialFetchCommand(args: {
  readonly cacheDir: string
  readonly limit: number
  readonly manifestPath: string
  readonly stopMarkerActive: boolean
}): string {
  return formatMutatingCommand(
    [
      'pnpm',
      'public-workbook-corpus:fetch-financial',
      '--',
      '--manifest',
      formatCommandPath(args.manifestPath),
      '--cache-dir',
      formatCommandPath(args.cacheDir),
      '--limit',
      String(args.limit),
    ],
    args.stopMarkerActive,
  )
}

function formatFinancialVerifyCommand(args: {
  readonly cacheDir: string
  readonly manifestPath: string
  readonly scorecardPath: string
  readonly stopMarkerActive: boolean
  readonly verifyCheckpointPath: string
}): string {
  return formatMutatingCommand(
    [
      'pnpm',
      'public-workbook-corpus:verify-financial',
      '--',
      '--manifest',
      formatCommandPath(args.manifestPath),
      '--cache-dir',
      formatCommandPath(args.cacheDir),
      '--scorecard',
      formatCommandPath(args.scorecardPath),
      '--verify-checkpoint',
      formatCommandPath(args.verifyCheckpointPath),
    ],
    args.stopMarkerActive,
  )
}

function formatFinancialCheckCommand(args: {
  readonly cacheDir: string
  readonly manifestPath: string
  readonly scorecardPath: string
  readonly verifyCheckpointPath: string
}): string {
  return formatCommand([
    'pnpm',
    'public-workbook-corpus:check-financial',
    '--',
    '--manifest',
    formatCommandPath(args.manifestPath),
    '--cache-dir',
    formatCommandPath(args.cacheDir),
    '--scorecard',
    formatCommandPath(args.scorecardPath),
    '--verify-checkpoint',
    formatCommandPath(args.verifyCheckpointPath),
  ])
}

function formatCommand(parts: readonly string[]): string {
  return parts.map(shellQuote).join(' ')
}

function formatMutatingCommand(parts: readonly string[], stopMarkerActive: boolean): string {
  if (!stopMarkerActive) {
    return formatCommand(parts)
  }
  return `${publicCorpusStopMarkerOverrideEnvVar}=1 ${formatCommand([...parts, publicCorpusStopMarkerOverrideFlag])}`
}

function shellQuote(value: string): string {
  return /^[A-Za-z0-9_./:=@+-]+$/u.test(value) ? value : `'${value.replaceAll("'", "'\\''")}'`
}

function formatCommandPath(path: string): string {
  const absolutePath = resolve(path)
  const relativePath = relative(rootDir, absolutePath)
  return relativePath && !relativePath.startsWith('..') && !isAbsolute(relativePath) ? relativePath : path
}

await main()
