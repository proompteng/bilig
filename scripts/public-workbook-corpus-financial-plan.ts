#!/usr/bin/env bun

import { existsSync, readFileSync } from 'node:fs'
import { isAbsolute, join, relative, resolve } from 'node:path'
import { pathToFileURL } from 'node:url'

import { createEmptyPublicWorkbookManifest, parsePublicWorkbookManifestJson } from './public-workbook-corpus-json.ts'
import { planPublicWorkbookCorpusFetch } from './public-workbook-corpus-fetch.ts'
import {
  publicCorpusStopMarkerOverrideEnvVar,
  publicCorpusStopMarkerOverrideFlag,
  readFlagArg,
  readNumberArg,
  readStringArg,
} from './public-workbook-corpus-cli.ts'
import type { PublicWorkbookManifest } from './public-workbook-corpus-types.ts'

export interface PublicWorkbookCorpusFinancialPlan {
  readonly schemaVersion: 1
  readonly mode: 'plan'
  readonly corpus: 'financial-accounting-workpapers'
  readonly generatedAt: string
  readonly manifestExists: boolean
  readonly targetWorkbookCount: number
  readonly manifestPath: string
  readonly cacheDir: string
  readonly scorecardPath: string
  readonly verifyCheckpointPath: string
  readonly stopMarker: {
    readonly active: boolean
    readonly path: string
    readonly overrideFlag: string
    readonly overrideEnvVar: string
  }
  readonly sourceCount: number
  readonly targetArtifactCount: number
  readonly cachedArtifactCount: number
  readonly remainingArtifactSlots: number
  readonly candidateSourceCount: number
  readonly candidateSourceDeficitCount: number
  readonly minimumAdditionalSourceCount: number
  readonly recommendedDiscoveryLimit: number
  readonly recommendedFetchTrancheSize: number
  readonly recommendedFetchLimit: number | null
  readonly needsAdditionalDiscovery: boolean
  readonly targetReachableFromKnownCandidates: boolean
  readonly commands: {
    readonly discoverPlan: string | null
    readonly discover: string | null
    readonly fetchPlan: string
    readonly fetch: string | null
    readonly fetchAll: string
    readonly resumePlan: string
    readonly resumeCheck: string
    readonly verify: string
    readonly check: string
  }
  readonly sampledCandidateSources: readonly {
    readonly id: string
    readonly kind: string
    readonly fileName: string
    readonly sourceUrl: string
    readonly downloadUrl: string
    readonly topicEvidence: readonly string[]
  }[]
}

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const defaultFinancialCacheDir = join(rootDir, '.cache', 'public-workbook-corpus-financial')
const defaultFinancialManifestPath = join(defaultFinancialCacheDir, 'manifest.json')
const defaultFinancialScorecardPath = join(defaultFinancialCacheDir, 'scorecard.json')
const defaultFinancialVerifyCheckpointPath = join(defaultFinancialCacheDir, 'verification-checkpoint.json')
const defaultCorpusRunStopMarkerPath = join(rootDir, '.agent-coordination', '20260507T074946Z-codex-stop-interactive-corpus-runs.md')

async function main(): Promise<void> {
  const plan = buildPublicWorkbookCorpusFinancialPlanFromArgs()
  if (readFlagArg('--check')) {
    const findings = validatePublicWorkbookCorpusFinancialPlan(plan)
    if (findings.length > 0) {
      throw new Error(`Public workbook corpus financial plan is invalid: ${findings.join('; ')}`)
    }
    process.stdout.write(
      `${JSON.stringify(
        {
          mode: 'check',
          schemaVersion: plan.schemaVersion,
          corpus: plan.corpus,
          generatedAt: plan.generatedAt,
          targetWorkbookCount: plan.targetWorkbookCount,
          sourceCount: plan.sourceCount,
          cachedArtifactCount: plan.cachedArtifactCount,
          remainingArtifactSlots: plan.remainingArtifactSlots,
          candidateSourceCount: plan.candidateSourceCount,
          candidateSourceDeficitCount: plan.candidateSourceDeficitCount,
          recommendedFetchLimit: plan.recommendedFetchLimit,
          needsAdditionalDiscovery: plan.needsAdditionalDiscovery,
          nextCommands: {
            ...(plan.commands.discover ? { discover: plan.commands.discover } : {}),
            fetch: plan.commands.fetch,
            fetchPlan: plan.commands.fetchPlan,
            resumeCheck: plan.commands.resumeCheck,
            verify: plan.commands.verify,
            check: plan.commands.check,
          },
        },
        null,
        2,
      )}\n`,
    )
    return
  }
  process.stdout.write(`${JSON.stringify(plan, null, 2)}\n`)
}

export function buildPublicWorkbookCorpusFinancialPlanFromArgs(): PublicWorkbookCorpusFinancialPlan {
  const cacheDir = resolve(readStringArg('--cache-dir', defaultFinancialCacheDir))
  const manifestPath = resolve(readStringArg('--manifest', defaultFinancialManifestPath))
  const scorecardPath = resolve(readStringArg('--scorecard', defaultFinancialScorecardPath))
  const verifyCheckpointPath = resolve(readStringArg('--verify-checkpoint', defaultFinancialVerifyCheckpointPath))
  const corpusRunStopMarkerPath = resolve(readStringArg('--corpus-run-stop-marker', defaultCorpusRunStopMarkerPath))
  const generatedAt = readStringArg('--generated-at', new Date().toISOString())
  const targetWorkbookCount = readNumberArg('--target-workbook-count', 5_000)
  const limit = readNumberArg('--limit', targetWorkbookCount)
  const sampleLimit = readNumberArg('--sample-limit', 20)
  const fetchTrancheSize = readNumberArg('--fetch-tranche-size', 20)
  const stopMarkerActive = existsSync(corpusRunStopMarkerPath)
  const manifest = readOrCreateFinancialManifest(manifestPath, targetWorkbookCount)
  const plan = planPublicWorkbookCorpusFetch({
    manifest,
    limit,
    sampleLimit,
  })
  return buildPublicWorkbookCorpusFinancialPlan({
    cacheDir,
    fetchPlan: plan,
    fetchTrancheSize,
    generatedAt,
    limit,
    manifestExists: existsSync(manifestPath),
    manifestPath,
    scorecardPath,
    stopMarkerActive,
    stopMarkerPath: corpusRunStopMarkerPath,
    targetWorkbookCount,
    verifyCheckpointPath,
  })
}

export function buildPublicWorkbookCorpusFinancialPlan(args: {
  readonly cacheDir: string
  readonly fetchPlan: ReturnType<typeof planPublicWorkbookCorpusFetch>
  readonly fetchTrancheSize: number
  readonly generatedAt: string
  readonly limit: number
  readonly manifestExists: boolean
  readonly manifestPath: string
  readonly scorecardPath: string
  readonly stopMarkerActive: boolean
  readonly stopMarkerPath: string
  readonly targetWorkbookCount: number
  readonly verifyCheckpointPath: string
}): PublicWorkbookCorpusFinancialPlan {
  const plan = args.fetchPlan
  const needsAdditionalDiscovery = !plan.targetReachableFromKnownCandidates
  const nextFetchLimit =
    plan.remainingArtifactSlots === 0 ? null : Math.min(plan.targetArtifactCount, plan.cachedArtifactCount + args.fetchTrancheSize)
  return {
    schemaVersion: 1,
    mode: 'plan',
    corpus: 'financial-accounting-workpapers',
    generatedAt: args.generatedAt,
    manifestExists: args.manifestExists,
    targetWorkbookCount: args.targetWorkbookCount,
    manifestPath: formatCommandPath(args.manifestPath),
    cacheDir: formatCommandPath(args.cacheDir),
    scorecardPath: formatCommandPath(args.scorecardPath),
    verifyCheckpointPath: formatCommandPath(args.verifyCheckpointPath),
    stopMarker: {
      active: args.stopMarkerActive,
      path: formatCommandPath(args.stopMarkerPath),
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
    recommendedDiscoveryLimit: Math.max(plan.recommendedDiscoveryLimit, args.targetWorkbookCount),
    recommendedFetchTrancheSize: args.fetchTrancheSize,
    recommendedFetchLimit: nextFetchLimit,
    needsAdditionalDiscovery,
    targetReachableFromKnownCandidates: plan.targetReachableFromKnownCandidates,
    commands: {
      discoverPlan: needsAdditionalDiscovery
        ? formatFinancialDiscoveryPlanCommand({
            cacheDir: args.cacheDir,
            limit: args.limit,
            manifestPath: args.manifestPath,
            targetWorkbookCount: args.targetWorkbookCount,
          })
        : null,
      discover: needsAdditionalDiscovery
        ? formatFinancialDiscoveryCommand({
            cacheDir: args.cacheDir,
            limit: Math.max(plan.recommendedDiscoveryLimit, args.targetWorkbookCount),
            manifestPath: args.manifestPath,
            stopMarkerActive: args.stopMarkerActive,
            targetWorkbookCount: args.targetWorkbookCount,
          })
        : null,
      fetchPlan: formatFinancialFetchPlanCommand({ cacheDir: args.cacheDir, limit: args.limit, manifestPath: args.manifestPath }),
      fetch: nextFetchLimit
        ? formatFinancialFetchCommand({
            cacheDir: args.cacheDir,
            limit: nextFetchLimit,
            manifestPath: args.manifestPath,
            stopMarkerActive: args.stopMarkerActive,
          })
        : null,
      fetchAll: formatFinancialFetchCommand({
        cacheDir: args.cacheDir,
        limit: args.limit,
        manifestPath: args.manifestPath,
        stopMarkerActive: args.stopMarkerActive,
      }),
      resumePlan: formatFinancialResumePlanCommand({
        cacheDir: args.cacheDir,
        fetchLimit: args.limit,
        manifestPath: args.manifestPath,
        scorecardPath: args.scorecardPath,
        verifyCheckpointPath: args.verifyCheckpointPath,
      }),
      resumeCheck: formatFinancialResumePlanCommand({
        cacheDir: args.cacheDir,
        check: true,
        fetchLimit: args.limit,
        manifestPath: args.manifestPath,
        scorecardPath: args.scorecardPath,
        verifyCheckpointPath: args.verifyCheckpointPath,
      }),
      verify: formatFinancialVerifyCommand({
        cacheDir: args.cacheDir,
        manifestPath: args.manifestPath,
        scorecardPath: args.scorecardPath,
        stopMarkerActive: args.stopMarkerActive,
        verifyCheckpointPath: args.verifyCheckpointPath,
      }),
      check: formatFinancialCheckCommand({
        cacheDir: args.cacheDir,
        manifestPath: args.manifestPath,
        scorecardPath: args.scorecardPath,
        verifyCheckpointPath: args.verifyCheckpointPath,
      }),
    },
    sampledCandidateSources: plan.sampledCandidateSources.map((source) => ({
      id: source.id,
      kind: source.kind,
      fileName: source.fileName,
      sourceUrl: source.sourceUrl,
      downloadUrl: source.downloadUrl,
      topicEvidence: source.topicEvidence ?? [],
    })),
  }
}

export function validatePublicWorkbookCorpusFinancialPlan(plan: PublicWorkbookCorpusFinancialPlan): string[] {
  const findings: string[] = []
  if (plan.schemaVersion !== 1) {
    findings.push(`unexpected schema version: ${String(plan.schemaVersion)}`)
  }
  if (plan.mode !== 'plan') {
    findings.push(`unexpected mode: ${String(plan.mode)}`)
  }
  if (plan.corpus !== 'financial-accounting-workpapers') {
    findings.push(`unexpected corpus: ${String(plan.corpus)}`)
  }
  if (!plan.generatedAt.trim()) {
    findings.push('generatedAt is empty')
  }
  for (const [label, value] of [
    ['target workbook count', plan.targetWorkbookCount],
    ['source count', plan.sourceCount],
    ['target artifact count', plan.targetArtifactCount],
    ['cached artifact count', plan.cachedArtifactCount],
    ['remaining artifact slots', plan.remainingArtifactSlots],
    ['candidate source count', plan.candidateSourceCount],
    ['candidate source deficit count', plan.candidateSourceDeficitCount],
    ['minimum additional source count', plan.minimumAdditionalSourceCount],
    ['recommended discovery limit', plan.recommendedDiscoveryLimit],
    ['recommended fetch tranche size', plan.recommendedFetchTrancheSize],
  ] as const) {
    if (!Number.isFinite(value) || value < 0) {
      findings.push(`${label} must be non-negative`)
    }
  }
  if (plan.remainingArtifactSlots !== Math.max(0, plan.targetArtifactCount - plan.cachedArtifactCount)) {
    findings.push('remaining artifact slots do not match target and cached artifact counts')
  }
  if (plan.needsAdditionalDiscovery === plan.targetReachableFromKnownCandidates) {
    findings.push('needsAdditionalDiscovery contradicts targetReachableFromKnownCandidates')
  }
  if (plan.recommendedFetchLimit !== null) {
    if (plan.recommendedFetchLimit <= plan.cachedArtifactCount) {
      findings.push('recommended fetch limit does not advance cached artifact count')
    }
    if (plan.recommendedFetchLimit > plan.targetArtifactCount) {
      findings.push('recommended fetch limit exceeds target artifact count')
    }
  } else if (plan.remainingArtifactSlots > 0) {
    findings.push('recommended fetch limit is missing while artifacts remain')
  }
  if (plan.stopMarker.active && plan.stopMarker.overrideFlag !== publicCorpusStopMarkerOverrideFlag) {
    findings.push('stop-marker override flag does not match corpus CLI guard')
  }
  if (plan.stopMarker.active && plan.stopMarker.overrideEnvVar !== publicCorpusStopMarkerOverrideEnvVar) {
    findings.push('stop-marker override environment variable does not match corpus CLI guard')
  }
  validateFinancialCommands(plan, findings)
  return findings
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

function formatFinancialResumePlanCommand(args: {
  readonly cacheDir: string
  readonly check?: boolean
  readonly fetchLimit: number
  readonly manifestPath: string
  readonly scorecardPath: string
  readonly verifyCheckpointPath: string
}): string {
  return formatCommand([
    'pnpm',
    args.check ? 'public-workbook-corpus:resume-financial:check' : 'public-workbook-corpus:resume-financial:plan',
    '--',
    '--manifest',
    formatCommandPath(args.manifestPath),
    '--cache-dir',
    formatCommandPath(args.cacheDir),
    '--scorecard',
    formatCommandPath(args.scorecardPath),
    '--verify-checkpoint',
    formatCommandPath(args.verifyCheckpointPath),
    '--fetch-limit',
    String(args.fetchLimit),
  ])
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

function validateFinancialCommands(plan: PublicWorkbookCorpusFinancialPlan, findings: string[]): void {
  if (plan.needsAdditionalDiscovery) {
    if (!plan.commands.discoverPlan?.includes('public-workbook-corpus:discover-financial:plan')) {
      findings.push('discover plan command is missing while discovery is needed')
    }
    if (!plan.commands.discover?.includes('public-workbook-corpus:discover-financial')) {
      findings.push('discover command is missing while discovery is needed')
    }
  } else {
    if (plan.commands.discoverPlan !== null) {
      findings.push('discover plan command is present when discovery is not needed')
    }
    if (plan.commands.discover !== null) {
      findings.push('discover command is present when discovery is not needed')
    }
  }
  if (!plan.commands.fetchPlan.includes('public-workbook-corpus:fetch-financial:plan')) {
    findings.push('fetch plan command is missing')
  }
  if (plan.remainingArtifactSlots > 0 && !plan.commands.fetch?.includes('public-workbook-corpus:fetch-financial')) {
    findings.push('bounded fetch command is missing while artifacts remain')
  }
  if (!plan.commands.fetchAll.includes('public-workbook-corpus:fetch-financial')) {
    findings.push('fetch-all command is missing')
  }
  if (!plan.commands.resumePlan.includes('public-workbook-corpus:resume-financial:plan')) {
    findings.push('resume plan command is missing')
  }
  if (!plan.commands.resumeCheck.includes('public-workbook-corpus:resume-financial:check')) {
    findings.push('resume check command is missing')
  }
  if (!plan.commands.verify.includes('public-workbook-corpus:verify-financial')) {
    findings.push('verify command is missing')
  }
  if (!plan.commands.check.includes('public-workbook-corpus:check-financial')) {
    findings.push('check command is missing')
  }
  const mutatingCommands = [plan.commands.discover, plan.commands.fetch, plan.commands.fetchAll, plan.commands.verify].filter(
    (command): command is string => command !== null,
  )
  const nonMutatingCommands = [
    plan.commands.discoverPlan,
    plan.commands.fetchPlan,
    plan.commands.resumePlan,
    plan.commands.resumeCheck,
    plan.commands.check,
  ].filter((command): command is string => command !== null)
  if (plan.stopMarker.active) {
    for (const command of mutatingCommands) {
      if (!command.includes(publicCorpusStopMarkerOverrideFlag) || !command.includes(`${publicCorpusStopMarkerOverrideEnvVar}=1`)) {
        findings.push(`mutating command is missing stop-marker override: ${command}`)
      }
    }
  }
  for (const command of nonMutatingCommands) {
    if (command.includes(publicCorpusStopMarkerOverrideFlag) || command.includes(`${publicCorpusStopMarkerOverrideEnvVar}=1`)) {
      findings.push(`non-mutating command unexpectedly bypasses stop marker: ${command}`)
    }
  }
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

if (process.argv[1] && pathToFileURL(resolve(process.argv[1])).href === import.meta.url) {
  await main()
}
