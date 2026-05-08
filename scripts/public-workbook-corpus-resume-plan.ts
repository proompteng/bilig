#!/usr/bin/env bun

import { existsSync, readFileSync } from 'node:fs'
import { join, resolve } from 'node:path'
import { pathToFileURL } from 'node:url'

import {
  publicCorpusStopMarkerOverrideEnvVar,
  publicCorpusStopMarkerOverrideFlag,
  readNumberArg,
  readStringArg,
} from './public-workbook-corpus-cli.ts'
import { planPublicWorkbookCorpusFetch } from './public-workbook-corpus-fetch.ts'
import { createEmptyPublicWorkbookManifest, parsePublicWorkbookManifestJson } from './public-workbook-corpus-json.ts'
import { readPublicWorkbookCorpusStatus } from './public-workbook-corpus-status.ts'

export interface PublicWorkbookCorpusResumePlan {
  readonly schemaVersion: 1
  readonly generatedAt: string
  readonly stopMarker: {
    readonly active: boolean
    readonly path: string
    readonly requiresExplicitResume: boolean
    readonly overrideFlag: string
    readonly overrideEnvVar: string
  }
  readonly currentState: {
    readonly targetWorkbookCount: number
    readonly cachedArtifactCount: number
    readonly recordedManifestArtifactCount: number
    readonly missingCachedArtifactCount: number
    readonly missingVerificationCount: number
    readonly recordedAllCasesPassed: boolean
  }
  readonly phases: {
    readonly verifyMissingCachedArtifacts: ResumePlanPhase
    readonly discoverAdditionalSources: ResumePlanPhase
    readonly fetchAdditionalArtifacts: ResumePlanPhase
    readonly finalEvidenceRefresh: ResumePlanPhase
  }
}

export interface ResumePlanPhase {
  readonly status: 'blocked-by-stop-marker' | 'ready' | 'not-needed'
  readonly reason: string
  readonly totalWorkItems: number
  readonly batchSize: number
  readonly batchCount: number
  readonly commands: readonly string[]
}

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const defaultCacheDir = join(rootDir, '.cache', 'public-workbook-corpus')
const defaultManifestPath = join(defaultCacheDir, 'manifest.json')
const defaultScorecardPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'public-workbook-corpus-scorecard.json')
const defaultVerifyCheckpointPath = join(defaultCacheDir, 'verification-checkpoint.json')
const defaultCorpusRunStopMarkerPath = join(rootDir, '.agent-coordination', '20260507T074946Z-codex-stop-interactive-corpus-runs.md')

function main(): void {
  const plan = buildPublicWorkbookCorpusResumePlanFromArgs()
  if (process.argv.includes('--check')) {
    const findings = validatePublicWorkbookCorpusResumePlan(plan)
    if (findings.length > 0) {
      throw new Error(`Public workbook corpus resume plan is invalid: ${findings.join('; ')}`)
    }
    process.stdout.write(
      `${JSON.stringify(
        {
          mode: 'check',
          schemaVersion: plan.schemaVersion,
          generatedAt: plan.generatedAt,
          stopMarkerActive: plan.stopMarker.active,
          currentState: plan.currentState,
          phases: {
            verifyMissingCachedArtifacts: phaseCheckSummary(plan.phases.verifyMissingCachedArtifacts),
            discoverAdditionalSources: phaseCheckSummary(plan.phases.discoverAdditionalSources),
            fetchAdditionalArtifacts: phaseCheckSummary(plan.phases.fetchAdditionalArtifacts),
            finalEvidenceRefresh: phaseCheckSummary(plan.phases.finalEvidenceRefresh),
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

export function buildPublicWorkbookCorpusResumePlanFromArgs(): PublicWorkbookCorpusResumePlan {
  const cacheDir = resolve(readStringArg('--cache-dir', defaultCacheDir))
  const manifestPath = resolve(readStringArg('--manifest', defaultManifestPath))
  const scorecardPath = resolve(readStringArg('--scorecard', defaultScorecardPath))
  const verifyCheckpointPath = resolve(readStringArg('--verify-checkpoint', defaultVerifyCheckpointPath))
  const stopMarkerPath = resolve(readStringArg('--corpus-run-stop-marker', defaultCorpusRunStopMarkerPath))
  const verifyBatchSize = readNumberArg('--verify-batch-size', 20)
  const fetchLimit = readNumberArg('--fetch-limit', 10_000)
  const fetchBatchSize = readNumberArg('--fetch-batch-size', 6)
  const generatedAt = readStringArg('--generated-at', new Date().toISOString())
  const status = readPublicWorkbookCorpusStatus({
    manifestPath,
    scorecardPath,
    cacheDir,
    verifyCheckpointPath,
  })
  const fetchPlan = existsSync(manifestPath)
    ? planPublicWorkbookCorpusFetch({
        manifest: parsePublicWorkbookManifestJson(JSON.parse(readFileSync(manifestPath, 'utf8')) as unknown),
        limit: fetchLimit,
        sampleLimit: 0,
      })
    : fallbackPublicWorkbookCorpusFetchPlan(status)
  return buildPublicWorkbookCorpusResumePlan({
    cacheDir,
    fetchBatchSize,
    fetchLimit,
    fetchPlan,
    generatedAt,
    manifestPath,
    scorecardPath,
    status,
    stopMarkerActive: existsSync(stopMarkerPath),
    stopMarkerPath,
    verifyBatchSize,
    verifyCheckpointPath,
  })
}

function fallbackPublicWorkbookCorpusFetchPlan(status: {
  readonly cachedArtifactCount: number
  readonly targetWorkbookCount: number
}): Parameters<typeof buildPublicWorkbookCorpusResumePlan>[0]['fetchPlan'] {
  const targetManifest = createEmptyPublicWorkbookManifest(new Date(0).toISOString(), status.targetWorkbookCount)
  const remainingArtifactSlots = Math.max(0, status.targetWorkbookCount - status.cachedArtifactCount)
  if (remainingArtifactSlots === 0) {
    return {
      candidateSourceCount: 0,
      candidateSourceDeficitCount: 0,
      recommendedDiscoveryLimit: targetManifest.targetWorkbookCount,
      remainingArtifactSlots: 0,
      targetReachableFromKnownCandidates: true,
    }
  }
  return {
    candidateSourceCount: 0,
    candidateSourceDeficitCount: remainingArtifactSlots,
    recommendedDiscoveryLimit: targetManifest.targetWorkbookCount,
    remainingArtifactSlots,
    targetReachableFromKnownCandidates: false,
  }
}

export function buildPublicWorkbookCorpusResumePlan(args: {
  readonly cacheDir: string
  readonly fetchBatchSize: number
  readonly fetchLimit: number
  readonly fetchPlan: {
    readonly candidateSourceCount: number
    readonly candidateSourceDeficitCount: number
    readonly recommendedDiscoveryLimit: number
    readonly remainingArtifactSlots: number
    readonly targetReachableFromKnownCandidates: boolean
  }
  readonly generatedAt: string
  readonly manifestPath: string
  readonly scorecardPath: string
  readonly status: {
    readonly cachedArtifactCount: number
    readonly missingManifestArtifactCount: number
    readonly recordedAllCasesPassed: boolean
    readonly recordedManifestArtifactCount: number
    readonly targetWorkbookCount: number
  }
  readonly stopMarkerActive: boolean
  readonly stopMarkerPath: string
  readonly verifyBatchSize: number
  readonly verifyCheckpointPath: string
}): PublicWorkbookCorpusResumePlan {
  const missingCachedArtifactCount = Math.max(0, args.status.targetWorkbookCount - args.status.cachedArtifactCount)
  return {
    schemaVersion: 1,
    generatedAt: args.generatedAt,
    stopMarker: {
      active: args.stopMarkerActive,
      path: args.stopMarkerPath,
      requiresExplicitResume: args.stopMarkerActive,
      overrideFlag: publicCorpusStopMarkerOverrideFlag,
      overrideEnvVar: publicCorpusStopMarkerOverrideEnvVar,
    },
    currentState: {
      targetWorkbookCount: args.status.targetWorkbookCount,
      cachedArtifactCount: args.status.cachedArtifactCount,
      recordedManifestArtifactCount: args.status.recordedManifestArtifactCount,
      missingCachedArtifactCount,
      missingVerificationCount: args.status.missingManifestArtifactCount,
      recordedAllCasesPassed: args.status.recordedAllCasesPassed,
    },
    phases: {
      verifyMissingCachedArtifacts: buildVerifyMissingPhase(args),
      discoverAdditionalSources: buildDiscoverPhase(args),
      fetchAdditionalArtifacts: buildFetchPhase(args),
      finalEvidenceRefresh: buildFinalEvidenceRefreshPhase(args),
    },
  }
}

export function validatePublicWorkbookCorpusResumePlan(plan: PublicWorkbookCorpusResumePlan): string[] {
  const findings: string[] = []
  if (plan.schemaVersion !== 1) {
    findings.push(`unexpected schema version: ${String(plan.schemaVersion)}`)
  }
  if (!plan.generatedAt.trim()) {
    findings.push('generatedAt is empty')
  }
  if (plan.stopMarker.active !== plan.stopMarker.requiresExplicitResume) {
    findings.push('stop-marker active state does not match explicit-resume requirement')
  }
  if (!Number.isFinite(plan.currentState.targetWorkbookCount) || plan.currentState.targetWorkbookCount < 0) {
    findings.push('target workbook count must be non-negative')
  }
  if (!Number.isFinite(plan.currentState.cachedArtifactCount) || plan.currentState.cachedArtifactCount < 0) {
    findings.push('cached artifact count must be non-negative')
  }
  if (!Number.isFinite(plan.currentState.recordedManifestArtifactCount) || plan.currentState.recordedManifestArtifactCount < 0) {
    findings.push('recorded verification count must be non-negative')
  }
  if (!Number.isFinite(plan.currentState.missingCachedArtifactCount) || plan.currentState.missingCachedArtifactCount < 0) {
    findings.push('missing cached artifact count must be non-negative')
  }
  if (!Number.isFinite(plan.currentState.missingVerificationCount) || plan.currentState.missingVerificationCount < 0) {
    findings.push('missing verification count must be non-negative')
  }
  if (plan.stopMarker.active) {
    if (plan.stopMarker.overrideFlag !== publicCorpusStopMarkerOverrideFlag) {
      findings.push('stop-marker override flag does not match corpus CLI guard')
    }
    if (plan.stopMarker.overrideEnvVar !== publicCorpusStopMarkerOverrideEnvVar) {
      findings.push('stop-marker override environment variable does not match corpus CLI guard')
    }
  }
  const expectedMissingCachedArtifactCount = Math.max(0, plan.currentState.targetWorkbookCount - plan.currentState.cachedArtifactCount)
  if (plan.currentState.missingCachedArtifactCount !== expectedMissingCachedArtifactCount) {
    findings.push(
      `missing cached artifact count is ${String(plan.currentState.missingCachedArtifactCount)}, expected ${String(
        expectedMissingCachedArtifactCount,
      )}`,
    )
  }
  if (plan.currentState.recordedManifestArtifactCount > plan.currentState.cachedArtifactCount) {
    findings.push('recorded verification count exceeds cached artifact count')
  }
  validatePhase('verifyMissingCachedArtifacts', plan.phases.verifyMissingCachedArtifacts, plan.stopMarker.active, findings)
  validatePhase('discoverAdditionalSources', plan.phases.discoverAdditionalSources, plan.stopMarker.active, findings)
  validatePhase('fetchAdditionalArtifacts', plan.phases.fetchAdditionalArtifacts, plan.stopMarker.active, findings)
  validatePhase('finalEvidenceRefresh', plan.phases.finalEvidenceRefresh, plan.stopMarker.active, findings)
  if (plan.phases.verifyMissingCachedArtifacts.totalWorkItems !== plan.currentState.missingVerificationCount) {
    findings.push('verify-missing phase does not match missing verification count')
  }
  if (plan.phases.fetchAdditionalArtifacts.totalWorkItems !== plan.currentState.missingCachedArtifactCount) {
    findings.push('fetch phase does not match missing cached artifact count')
  }
  const finalCommands = plan.phases.finalEvidenceRefresh.commands
  for (const requiredCommand of ['pnpm dominance:generate', 'pnpm dominance:audit:check', 'pnpm dominance:check']) {
    if (!finalCommands.includes(requiredCommand)) {
      findings.push(`final evidence refresh is missing command: ${requiredCommand}`)
    }
  }
  return findings
}

function buildVerifyMissingPhase(args: Parameters<typeof buildPublicWorkbookCorpusResumePlan>[0]): ResumePlanPhase {
  const totalWorkItems = args.status.missingManifestArtifactCount
  if (totalWorkItems === 0) {
    return notNeededPhase('all cached artifacts already have recorded verification evidence')
  }
  const batchSize = normalizedBatchSize(args.verifyBatchSize)
  return {
    status: phaseStatus(args.stopMarkerActive),
    reason: 'cached workbook artifacts exist locally but do not all have recorded verification cases',
    totalWorkItems,
    batchSize,
    batchCount: batchCount(totalWorkItems, batchSize),
    commands: [
      command([
        'pnpm',
        'public-workbook-corpus:verify-missing:plan',
        '--',
        '--manifest',
        args.manifestPath,
        '--scorecard',
        args.scorecardPath,
        '--verify-checkpoint',
        args.verifyCheckpointPath,
        '--cache-dir',
        args.cacheDir,
      ]),
      guardedCommand(args.stopMarkerActive, [
        'pnpm',
        'public-workbook-corpus:verify-missing',
        '--',
        '--manifest',
        args.manifestPath,
        '--scorecard',
        args.scorecardPath,
        '--verify-checkpoint',
        args.verifyCheckpointPath,
        '--cache-dir',
        args.cacheDir,
        '--limit',
        String(batchSize),
      ]),
    ],
  }
}

function buildDiscoverPhase(args: Parameters<typeof buildPublicWorkbookCorpusResumePlan>[0]): ResumePlanPhase {
  const totalWorkItems = args.fetchPlan.candidateSourceDeficitCount
  if (totalWorkItems === 0) {
    return notNeededPhase('known candidate sources can reach the target artifact count')
  }
  return {
    status: phaseStatus(args.stopMarkerActive),
    reason: 'known candidate sources cannot fill the remaining artifact target',
    totalWorkItems,
    batchSize: totalWorkItems,
    batchCount: 1,
    commands: [
      command(['pnpm', 'public-workbook-corpus:discover:plan', '--', '--limit', String(args.fetchPlan.recommendedDiscoveryLimit)]),
      guardedCommand(args.stopMarkerActive, [
        'pnpm',
        'public-workbook-corpus:discover',
        '--',
        '--manifest',
        args.manifestPath,
        '--cache-dir',
        args.cacheDir,
        '--limit',
        String(args.fetchPlan.recommendedDiscoveryLimit),
      ]),
    ],
  }
}

function buildFetchPhase(args: Parameters<typeof buildPublicWorkbookCorpusResumePlan>[0]): ResumePlanPhase {
  const totalWorkItems = Math.max(0, args.status.targetWorkbookCount - args.status.cachedArtifactCount)
  if (totalWorkItems === 0) {
    return notNeededPhase('cached artifacts already reached the target workbook count')
  }
  const batchSize = normalizedBatchSize(args.fetchBatchSize)
  const reason = args.fetchPlan.targetReachableFromKnownCandidates
    ? 'known candidate sources can be fetched to fill the remaining artifact target'
    : 'candidate source discovery must run before fetch can fill the remaining artifact target'
  return {
    status: phaseStatus(args.stopMarkerActive),
    reason,
    totalWorkItems,
    batchSize,
    batchCount: batchCount(totalWorkItems, batchSize),
    commands: [
      command([
        'pnpm',
        'public-workbook-corpus:fetch:plan',
        '--',
        '--manifest',
        args.manifestPath,
        '--cache-dir',
        args.cacheDir,
        '--limit',
        String(args.fetchLimit),
      ]),
      guardedCommand(args.stopMarkerActive, [
        'pnpm',
        'public-workbook-corpus:fetch',
        '--',
        '--manifest',
        args.manifestPath,
        '--cache-dir',
        args.cacheDir,
        '--limit',
        String(args.fetchLimit),
        '--fetch-batch-size',
        String(batchSize),
      ]),
    ],
  }
}

function buildFinalEvidenceRefreshPhase(args: Parameters<typeof buildPublicWorkbookCorpusResumePlan>[0]): ResumePlanPhase {
  return {
    status: phaseStatus(args.stopMarkerActive),
    reason: 'refresh checked-in evidence after local corpus target and verification coverage are complete',
    totalWorkItems: 1,
    batchSize: 1,
    batchCount: 1,
    commands: [
      guardedCommand(args.stopMarkerActive, [
        'pnpm',
        'public-workbook-corpus:verify',
        '--',
        '--manifest',
        args.manifestPath,
        '--scorecard',
        args.scorecardPath,
        '--verify-checkpoint',
        args.verifyCheckpointPath,
        '--cache-dir',
        args.cacheDir,
      ]),
      command(['pnpm', 'dominance:generate']),
      command(['pnpm', 'dominance:audit:check']),
      command(['pnpm', 'dominance:check']),
    ],
  }
}

function notNeededPhase(reason: string): ResumePlanPhase {
  return {
    status: 'not-needed',
    reason,
    totalWorkItems: 0,
    batchSize: 0,
    batchCount: 0,
    commands: [],
  }
}

function phaseStatus(stopMarkerActive: boolean): ResumePlanPhase['status'] {
  return stopMarkerActive ? 'blocked-by-stop-marker' : 'ready'
}

function guardedCommand(stopMarkerActive: boolean, parts: readonly string[]): string {
  if (!stopMarkerActive) {
    return command(parts)
  }
  return `${publicCorpusStopMarkerOverrideEnvVar}=1 ${command([...parts, publicCorpusStopMarkerOverrideFlag])}`
}

function command(parts: readonly string[]): string {
  return parts.map(shellQuote).join(' ')
}

function batchCount(totalWorkItems: number, batchSize: number): number {
  return Math.ceil(totalWorkItems / batchSize)
}

function normalizedBatchSize(value: number): number {
  return Math.max(1, Math.trunc(value))
}

function phaseCheckSummary(phase: ResumePlanPhase): Pick<ResumePlanPhase, 'status' | 'totalWorkItems' | 'batchSize' | 'batchCount'> {
  return {
    status: phase.status,
    totalWorkItems: phase.totalWorkItems,
    batchSize: phase.batchSize,
    batchCount: phase.batchCount,
  }
}

function shellQuote(value: string): string {
  return /^[A-Za-z0-9_./:=@+-]+$/u.test(value) ? value : `'${value.replaceAll("'", "'\\''")}'`
}

function validatePhase(name: string, phase: ResumePlanPhase, stopMarkerActive: boolean, findings: string[]): void {
  if (!phase.reason.trim()) {
    findings.push(`${name} reason is empty`)
  }
  if (!Number.isFinite(phase.totalWorkItems) || phase.totalWorkItems < 0) {
    findings.push(`${name} total work items must be non-negative`)
  }
  if (phase.totalWorkItems === 0) {
    if (phase.status !== 'not-needed') {
      findings.push(`${name} has no work but is not marked not-needed`)
    }
    if (phase.batchSize !== 0 || phase.batchCount !== 0 || phase.commands.length !== 0) {
      findings.push(`${name} has no work but still has batch data or commands`)
    }
    return
  }
  if (phase.status === 'not-needed') {
    findings.push(`${name} has work items but is marked not-needed`)
  }
  if (stopMarkerActive && phase.status !== 'blocked-by-stop-marker') {
    findings.push(`${name} is not blocked while stop marker is active`)
  }
  if (!stopMarkerActive && phase.status !== 'ready') {
    findings.push(`${name} is not ready while stop marker is inactive`)
  }
  if (!Number.isFinite(phase.batchSize) || phase.batchSize <= 0) {
    findings.push(`${name} batch size must be positive`)
  } else if (!Number.isFinite(phase.batchCount) || phase.batchCount < 0) {
    findings.push(`${name} batch count must be non-negative`)
  } else if (phase.batchCount !== batchCount(phase.totalWorkItems, phase.batchSize)) {
    findings.push(`${name} batch count does not match total work items and batch size`)
  }
  if (phase.commands.length === 0) {
    findings.push(`${name} has work items but no commands`)
  }
  if (stopMarkerActive) {
    for (const mutatingCommand of phase.commands.filter(isCorpusMutatingCommand)) {
      if (
        !mutatingCommand.includes(`${publicCorpusStopMarkerOverrideEnvVar}=1`) ||
        !mutatingCommand.includes(publicCorpusStopMarkerOverrideFlag)
      ) {
        findings.push(`${name} mutating command is missing the explicit stop-marker override: ${mutatingCommand}`)
      }
    }
  }
}

function isCorpusMutatingCommand(value: string): boolean {
  return [
    'public-workbook-corpus:verify-missing --',
    'public-workbook-corpus:discover --',
    'public-workbook-corpus:fetch --',
    'public-workbook-corpus:verify --',
  ].some((needle) => value.includes(needle))
}

if (process.argv[1] && pathToFileURL(resolve(process.argv[1])).href === import.meta.url) {
  main()
}
