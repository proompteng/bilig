#!/usr/bin/env bun

import { existsSync, readFileSync } from 'node:fs'
import { isAbsolute, join, relative, resolve } from 'node:path'
import { pathToFileURL } from 'node:url'

import type { BuildScorecardInput } from './bilig-dominance-scorecard-types.ts'
import { loadBiligDominanceScorecardInput, rootDir } from './bilig-dominance-scorecard-input.ts'
import {
  localCiResourceGuardOverrideEnv,
  readLocalCiResourceGuardStatus,
  type LocalCiResourceGuardStatus,
} from './ci-local-resource-guard.ts'
import { buildBiligDominanceScorecard } from './gen-bilig-dominance-scorecard.ts'
import type { UiResponsivenessSameCorpusWorkload } from './gen-ui-responsiveness-live-browser-scorecard.ts'
import {
  parseSameCorpusPublicAccessCheckJson,
  type SameCorpusPublicAccessCheck,
} from './ui-responsiveness-same-corpus-public-access-check.ts'
import { buildWorkbookBenchmarkCorpus, type WorkbookBenchmarkCorpusId } from '../packages/benchmarks/src/workbook-corpus.js'
import { publicCorpusStopMarkerOverrideEnvVar, publicCorpusStopMarkerOverrideFlag, readStringArg } from './public-workbook-corpus-cli.ts'
import { planPublicWorkbookCorpusFetch, type PublicWorkbookCorpusFetchPlan } from './public-workbook-corpus-fetch.ts'
import { createEmptyPublicWorkbookManifest, parsePublicWorkbookManifestJson } from './public-workbook-corpus-json.ts'
import { publicWorkbookCorpusCaseMatchesArtifact } from './public-workbook-corpus-missing.ts'
import type { PublicWorkbookCorpusStatus } from './public-workbook-corpus-status.ts'
import { readPublicWorkbookCorpusStatus } from './public-workbook-corpus-status.ts'
import { readReusablePublicWorkbookCorpusCases } from './public-workbook-corpus-verify-checkpoint.ts'
import { financialWorkbookTargetCount } from './public-workbook-corpus-completion-audit-helpers.ts'
import type { PublicWorkbookManifest } from './public-workbook-corpus-types.ts'
import { buildPublicWorkbookCorpusFinancialPlan, type PublicWorkbookCorpusFinancialPlan } from './public-workbook-corpus-financial-plan.ts'
import {
  buildPublicWorkbookCorpusFeatureWitnessPlan,
  readPublicWorkbookCorpusFeatureWitnessCases,
  type PublicWorkbookCorpusFeatureWitnessPlan,
} from './public-workbook-corpus-feature-witness-plan.ts'

export interface BiligDominanceStatus {
  readonly goalStatus: 'achieved' | 'active-not-achieved'
  readonly blanketTenXClaimAllowed: boolean
  readonly unmetRequirements: readonly string[]
  readonly importExportBlockers: readonly string[]
  readonly localCiResourceGuard: {
    readonly active: boolean
    readonly activeMarkerPaths: readonly string[]
    readonly overrideEnvVar: string
    readonly overridePrefix: string | null
  }
  readonly publicWorkbookCorpus: {
    readonly targetWorkbookCount: number
    readonly cachedArtifactCount: number
    readonly missingCachedArtifactCount: number
    readonly recordedManifestArtifactCount: number
    readonly missingManifestArtifactCount: number
    readonly financialWorkbookTargetCount: number | null
    readonly financialSourceCount: number | null
    readonly financialCachedArtifactCount: number | null
    readonly recordedFinancialManifestArtifactCount: number | null
    readonly recordedFinancialNonPassingCaseCount: number | null
    readonly featureWitnessPlan: {
      readonly recordedCaseCount: number
      readonly missingWitnessCount: number
      readonly nextPlanCommand: string
      readonly nextCheckCommand: string
      readonly coverage: readonly {
        readonly id: string
        readonly label: string
        readonly totalCount: number
        readonly witnessCaseCount: number
        readonly needsWitness: boolean
        readonly discoveryQuery: string
        readonly nextDiscoverCommand: string | null
        readonly blockedDiscoverCommand: string | null
      }[]
    } | null
    readonly financialPlan: {
      readonly sourceCount: number
      readonly targetArtifactCount: number
      readonly cachedArtifactCount: number
      readonly remainingArtifactSlots: number
      readonly candidateSourceCount: number
      readonly candidateSourceDeficitCount: number
      readonly recommendedFetchLimit: number | null
      readonly needsAdditionalDiscovery: boolean
      readonly targetReachableFromKnownCandidates: boolean
      readonly nextPlanCommand: string
      readonly nextCheckCommand: string
      readonly nextFetchPlanCommand: string
      readonly nextFetchCommand: string | null
      readonly nextVerifyCommand: string | null
      readonly blockedCommands: readonly string[]
    } | null
    readonly fetchCandidateSourceCount: number | null
    readonly fetchCandidateSourceDeficitCount: number | null
    readonly minimumAdditionalSourceCount: number | null
    readonly recommendedDiscoveryLimit: number | null
    readonly targetReachableFromKnownCandidates: boolean | null
    readonly scorecardCaseCount: number
    readonly checkpointCaseCount: number
    readonly recordedAllCasesPassed: boolean
    readonly targetComplete: boolean
    readonly nextFetchPlanCommand: string | null
    readonly nextDiscoveryPlanCommand: string | null
    readonly nextDiscoveryCommand: string | null
    readonly nextMissingVerificationPlanCommand: string | null
    readonly nextMissingVerificationCommand: string | null
    readonly nextStaleVerificationPlanCommand: string | null
    readonly nextStaleVerificationCommand: string | null
    readonly blockedCommands: readonly string[]
    readonly corpusRunStopMarkerActive: boolean
    readonly corpusRunStopMarkerPath: string
    readonly nextCorpusRunRequiresExplicitResume: boolean
    readonly corpusRunStopMarkerOverrideFlag: string
    readonly corpusRunStopMarkerOverrideEnvVar: string
    readonly gaps: readonly string[]
  }
  readonly uiSameCorpus: {
    readonly captured: boolean
    readonly evidenceKind: 'same-corpus-browser-capture' | 'not-captured'
    readonly requiredProductCount: number
    readonly requiredCaseCount: number
    readonly tenXMeanAndP95CaseCount: number
    readonly tenXRequirementSatisfied: boolean
    readonly requiredWorkloads: readonly UiResponsivenessSameCorpusWorkload[]
    readonly missingRequiredWorkloads: readonly UiResponsivenessSameCorpusWorkload[]
    readonly coveredCorpusCaseIds: readonly string[]
    readonly limitations: readonly string[]
    readonly fixture: {
      readonly corpusCaseId: WorkbookBenchmarkCorpusId
      readonly materializedCells: number
      readonly localXlsxPath: string
      readonly publicGithubRawUrl: string
      readonly publicForgejoRawUrl: string
      readonly microsoftExcelWebUrl: string
    }
    readonly googleSheetsUrl: string | null
    readonly googleSheetsUrlSource: UiSameCorpusGoogleSheetsUrlSource
    readonly googleSheetsUrlEnvVar: string
    readonly publicAccessCheckPath: string
    readonly missingInputs: readonly string[]
    readonly nextFixtureCheckCommand: string
    readonly nextPublicAccessCheckCommand: string
    readonly nextGoogleSheetsUploadInstruction: string
    readonly nextPreflightCommand: string | null
    readonly nextCaptureCommand: string | null
    readonly blockedCommands: readonly string[]
    readonly browserCaptureGuard: {
      readonly active: boolean
      readonly activeMarkerPaths: readonly string[]
      readonly overrideEnvVar: string
      readonly overridePrefix: string | null
      readonly nextPreflightRequiresOverride: boolean
      readonly nextCaptureRequiresOverride: boolean
    }
    readonly nextScorecardGenerateCommand: string
    readonly nextDominanceCheckCommand: string
  }
}

type UiSameCorpusGoogleSheetsUrlSource = 'argument-or-environment' | 'public-access-check' | 'missing'

const defaultCacheDir = join(rootDir, '.cache', 'public-workbook-corpus')
const defaultManifestPath = join(defaultCacheDir, 'manifest.json')
const defaultScorecardPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'public-workbook-corpus-scorecard.json')
const defaultVerifyCheckpointPath = join(defaultCacheDir, 'verification-checkpoint.json')
const defaultFinancialCacheDir = join(rootDir, '.cache', 'public-workbook-corpus-financial')
const defaultFinancialManifestPath = join(defaultFinancialCacheDir, 'manifest.json')
const defaultFinancialScorecardPath = join(defaultFinancialCacheDir, 'scorecard.json')
const defaultFinancialVerifyCheckpointPath = join(defaultFinancialCacheDir, 'verification-checkpoint.json')
const defaultCorpusRunStopMarkerPath = join(rootDir, '.agent-coordination', '20260507T074946Z-codex-stop-interactive-corpus-runs.md')
const defaultUiSameCorpusId: WorkbookBenchmarkCorpusId = 'wide-mixed-250k'
const defaultUiSameCorpusPublicAccessCheckPath = join(rootDir, '.cache', 'ui-responsiveness', 'same-corpus-public-access-check.json')
const uiSameCorpusGoogleSheetsUrlEnvVar = 'BILIG_UI_SAME_CORPUS_GOOGLE_SHEETS_URL'
const requiredUiSameCorpusWorkloads = ['visible-scroll-response'] as const satisfies readonly UiResponsivenessSameCorpusWorkload[]

function main(): void {
  process.stdout.write(`${JSON.stringify(buildBiligDominanceStatusFromArgs(), null, 2)}\n`)
}

export function buildBiligDominanceStatusFromArgs(): BiligDominanceStatus {
  const cacheDir = resolve(readStringArg('--cache-dir', defaultCacheDir))
  const scorecardPath = resolve(readStringArg('--scorecard', defaultScorecardPath))
  const manifestPath = resolve(readStringArg('--manifest', defaultManifestPath))
  const verifyCheckpointPath = resolve(readStringArg('--verify-checkpoint', defaultVerifyCheckpointPath))
  const financialCacheDir = resolve(readStringArg('--financial-cache-dir', defaultFinancialCacheDir))
  const financialManifestPath = resolve(readStringArg('--financial-manifest', defaultFinancialManifestPath))
  const financialScorecardPath = resolve(readStringArg('--financial-scorecard', defaultFinancialScorecardPath))
  const financialVerifyCheckpointPath = resolve(readStringArg('--financial-verify-checkpoint', defaultFinancialVerifyCheckpointPath))
  const stopMarkerPath = resolve(readStringArg('--corpus-run-stop-marker', defaultCorpusRunStopMarkerPath))
  const uiSameCorpusPublicAccessCheckPath = resolve(
    readStringArg('--ui-same-corpus-public-access-check', defaultUiSameCorpusPublicAccessCheckPath),
  )
  const explicitUiSameCorpusGoogleSheetsUrl =
    readStringArg('--ui-same-corpus-google-sheets-url', process.env[uiSameCorpusGoogleSheetsUrlEnvVar] ?? '') || null
  const uiSameCorpusPublicAccessCheck = readSameCorpusPublicAccessCheckOrNull(uiSameCorpusPublicAccessCheckPath)
  const input = loadBiligDominanceScorecardInput()
  const publicWorkbookCorpusStatus = readPublicWorkbookCorpusStatus({
    manifestPath,
    scorecardPath,
    cacheDir,
    verifyCheckpointPath,
    corpusRunStopMarkerPath: stopMarkerPath,
  })
  const stopMarkerActive = existsSync(stopMarkerPath)
  const nextFetchPlanCommand =
    publicWorkbookCorpusStatus.cachedArtifactCount < publicWorkbookCorpusStatus.targetWorkbookCount
      ? formatPublicWorkbookCorpusFetchPlanCommand({
          cacheDir,
          displayRootDir: rootDir,
          limit: publicWorkbookCorpusStatus.targetWorkbookCount,
          manifestPath,
        })
      : null
  const fetchPlan = existsSync(manifestPath)
    ? planPublicWorkbookCorpusFetch({
        manifest: parsePublicWorkbookManifestJson(JSON.parse(readFileSync(manifestPath, 'utf8')) as unknown),
        limit: publicWorkbookCorpusStatus.targetWorkbookCount,
        sampleLimit: 0,
      })
    : null
  const financialTargetWorkbookCount = financialWorkbookTargetCount(publicWorkbookCorpusStatus.targetWorkbookCount)
  const financialManifest = readPublicWorkbookManifestOrNull(financialManifestPath)
  const financialCorpusStatus = buildFinancialWorkbookCorpusStatus({
    manifest: financialManifest,
    recordedCases: readRecordedFinancialCases({
      manifest: financialManifest,
      scorecardPath: financialScorecardPath,
      verifyCheckpointPath: financialVerifyCheckpointPath,
    }),
    targetWorkbookCount: financialTargetWorkbookCount,
  })
  const financialCorpusPlan = buildPublicWorkbookCorpusFinancialPlan({
    cacheDir: financialCacheDir,
    fetchPlan: planPublicWorkbookCorpusFetch({
      manifest: financialManifest ?? createEmptyPublicWorkbookManifest(undefined, financialTargetWorkbookCount),
      limit: financialTargetWorkbookCount,
      sampleLimit: 20,
    }),
    fetchTrancheSize: 20,
    generatedAt: new Date().toISOString(),
    limit: financialTargetWorkbookCount,
    manifestExists: financialManifest !== null,
    manifestPath: financialManifestPath,
    scorecardPath: financialScorecardPath,
    stopMarkerActive,
    stopMarkerPath,
    targetWorkbookCount: financialTargetWorkbookCount,
    verifyCheckpointPath: financialVerifyCheckpointPath,
  })
  const featureWitnessPlan = buildPublicWorkbookCorpusFeatureWitnessPlan({
    cacheDir,
    cases: readPublicWorkbookCorpusFeatureWitnessCases({ manifestPath, scorecardPath, verifyCheckpointPath }),
    discoveryLimit: publicWorkbookCorpusStatus.targetWorkbookCount,
    displayRootDir: rootDir,
    generatedAt: new Date().toISOString(),
    manifestPath,
    stopMarkerActive,
    stopMarkerPath,
  })
  return buildBiligDominanceStatus({
    featureWitnessPlan,
    fetchPlan,
    financialCorpusPlan,
    financialCorpusStatus,
    input,
    nextFetchPlanCommand,
    publicWorkbookCorpusStatus,
    stopMarkerActive,
    stopMarkerPath: formatBiligDominanceStatusPathForMessage(stopMarkerPath, rootDir),
    uiSameCorpusGoogleSheetsUrl: explicitUiSameCorpusGoogleSheetsUrl,
    uiSameCorpusLocalCiResourceGuardStatus: readLocalCiResourceGuardStatus(rootDir),
    uiSameCorpusPublicAccessCheck,
    uiSameCorpusPublicAccessCheckPath: formatBiligDominanceStatusPathForMessage(uiSameCorpusPublicAccessCheckPath, rootDir),
  })
}

export function buildBiligDominanceStatus(args: {
  readonly featureWitnessPlan?: PublicWorkbookCorpusFeatureWitnessPlan | null
  readonly fetchPlan?: PublicWorkbookCorpusFetchPlan | null
  readonly financialCorpusPlan?: PublicWorkbookCorpusFinancialPlan | null
  readonly financialCorpusStatus?: FinancialWorkbookCorpusStatus | null
  readonly input: BuildScorecardInput
  readonly nextFetchPlanCommand?: string | null
  readonly publicWorkbookCorpusStatus: PublicWorkbookCorpusStatus
  readonly stopMarkerActive: boolean
  readonly stopMarkerPath: string
  readonly uiSameCorpusGoogleSheetsUrl?: string | null
  readonly uiSameCorpusLocalCiResourceGuardStatus?: LocalCiResourceGuardStatus
  readonly uiSameCorpusPublicAccessCheck?: SameCorpusPublicAccessCheck | null
  readonly uiSameCorpusPublicAccessCheckPath?: string
}): BiligDominanceStatus {
  const scorecard = buildBiligDominanceScorecard({
    ...args.input,
    publicWorkbookCorpusStatus: args.publicWorkbookCorpusStatus,
  })
  const importExportCategory = scorecard.categories.find((category) => category.id === 'import-export-compatibility')
  const publicWorkbookCorpusBlockers = publicWorkbookCorpusDominanceBlockers(args.publicWorkbookCorpusStatus)
  const featureWitnessBlockers = publicWorkbookCorpusFeatureWitnessBlockers(args.featureWitnessPlan ?? null)
  const financialCorpusBlockers = financialWorkbookCorpusDominanceBlockers(args.financialCorpusStatus ?? null)
  const localCiResourceGuardStatus = args.uiSameCorpusLocalCiResourceGuardStatus ?? { activeMarkerPaths: [] }
  const localCiResourceGuardBlockers = localCiResourceGuardDominanceBlockers(localCiResourceGuardStatus)
  const corpusBlockers = [...publicWorkbookCorpusBlockers, ...featureWitnessBlockers, ...financialCorpusBlockers]
  const liveStatusBlockers = [...corpusBlockers, ...localCiResourceGuardBlockers]
  const unmetRequirements = [...scorecard.claimPolicy.unmetRequirements, ...liveStatusBlockers]
  const blanketTenXClaimAllowed = scorecard.claimPolicy.blanketTenXClaimAllowed && liveStatusBlockers.length === 0
  const nextDiscoveryCommand =
    args.fetchPlan && args.fetchPlan.candidateSourceDeficitCount > 0
      ? formatPublicWorkbookCorpusDiscoveryCommand(args.fetchPlan.recommendedDiscoveryLimit)
      : null
  const nextMissingVerificationCommand = args.publicWorkbookCorpusStatus.nextMissingVerificationCommand
  const nextStaleVerificationCommand = args.publicWorkbookCorpusStatus.nextStaleVerificationCommand
  const blockedCorpusCommands = args.stopMarkerActive
    ? nonEmptyCommands([nextDiscoveryCommand, nextMissingVerificationCommand, nextStaleVerificationCommand]).map(
        corpusStopMarkerOverrideCommand,
      )
    : []
  return {
    goalStatus: scorecard.goalStatus === 'achieved' && liveStatusBlockers.length === 0 ? 'achieved' : 'active-not-achieved',
    blanketTenXClaimAllowed,
    unmetRequirements,
    importExportBlockers: [...(importExportCategory?.blockers ?? []), ...corpusBlockers],
    localCiResourceGuard: buildLocalCiResourceGuardStatus(localCiResourceGuardStatus),
    publicWorkbookCorpus: {
      targetWorkbookCount: args.publicWorkbookCorpusStatus.targetWorkbookCount,
      cachedArtifactCount: args.publicWorkbookCorpusStatus.cachedArtifactCount,
      missingCachedArtifactCount: Math.max(
        0,
        args.publicWorkbookCorpusStatus.targetWorkbookCount - args.publicWorkbookCorpusStatus.cachedArtifactCount,
      ),
      recordedManifestArtifactCount: args.publicWorkbookCorpusStatus.recordedManifestArtifactCount,
      missingManifestArtifactCount: args.publicWorkbookCorpusStatus.missingManifestArtifactCount,
      financialWorkbookTargetCount: args.financialCorpusStatus?.targetWorkbookCount ?? null,
      financialSourceCount: args.financialCorpusStatus?.sourceCount ?? null,
      financialCachedArtifactCount: args.financialCorpusStatus?.cachedArtifactCount ?? null,
      recordedFinancialManifestArtifactCount: args.financialCorpusStatus?.recordedManifestArtifactCount ?? null,
      recordedFinancialNonPassingCaseCount: args.financialCorpusStatus?.recordedNonPassingCaseCount ?? null,
      featureWitnessPlan: args.featureWitnessPlan ? buildFeatureWitnessPlanStatus(args.featureWitnessPlan) : null,
      financialPlan: args.financialCorpusPlan ? buildFinancialPlanStatus(args.financialCorpusPlan) : null,
      fetchCandidateSourceCount: args.fetchPlan?.candidateSourceCount ?? null,
      fetchCandidateSourceDeficitCount: args.fetchPlan?.candidateSourceDeficitCount ?? null,
      targetReachableFromKnownCandidates: args.fetchPlan?.targetReachableFromKnownCandidates ?? null,
      minimumAdditionalSourceCount: args.fetchPlan?.minimumAdditionalSourceCount ?? null,
      recommendedDiscoveryLimit: args.fetchPlan?.recommendedDiscoveryLimit ?? null,
      scorecardCaseCount: args.publicWorkbookCorpusStatus.scorecardCaseCount,
      checkpointCaseCount: args.publicWorkbookCorpusStatus.checkpointCaseCount,
      recordedAllCasesPassed: args.publicWorkbookCorpusStatus.recordedAllCasesPassed,
      targetComplete: args.publicWorkbookCorpusStatus.targetComplete,
      nextFetchPlanCommand: args.nextFetchPlanCommand ?? null,
      nextDiscoveryPlanCommand:
        args.fetchPlan && args.fetchPlan.candidateSourceDeficitCount > 0
          ? formatPublicWorkbookCorpusDiscoveryPlanCommand(args.fetchPlan.recommendedDiscoveryLimit)
          : null,
      nextDiscoveryCommand: args.stopMarkerActive ? null : nextDiscoveryCommand,
      nextMissingVerificationPlanCommand: args.publicWorkbookCorpusStatus.nextMissingVerificationPlanCommand,
      nextMissingVerificationCommand: args.stopMarkerActive ? null : nextMissingVerificationCommand,
      nextStaleVerificationPlanCommand: args.publicWorkbookCorpusStatus.nextStaleVerificationPlanCommand,
      nextStaleVerificationCommand: args.stopMarkerActive ? null : nextStaleVerificationCommand,
      blockedCommands: blockedCorpusCommands,
      corpusRunStopMarkerActive: args.stopMarkerActive,
      corpusRunStopMarkerPath: args.stopMarkerPath,
      nextCorpusRunRequiresExplicitResume: args.stopMarkerActive,
      corpusRunStopMarkerOverrideFlag: publicCorpusStopMarkerOverrideFlag,
      corpusRunStopMarkerOverrideEnvVar: publicCorpusStopMarkerOverrideEnvVar,
      gaps: args.publicWorkbookCorpusStatus.gaps,
    },
    uiSameCorpus: buildUiSameCorpusStatus(args.input, {
      localCiResourceGuardStatus,
      publicAccessCheckPath: args.uiSameCorpusPublicAccessCheckPath ?? '.cache/ui-responsiveness/same-corpus-public-access-check.json',
      ...resolveUiSameCorpusGoogleSheetsUrl({
        corpusCaseId: defaultUiSameCorpusId,
        explicitGoogleSheetsUrl: args.uiSameCorpusGoogleSheetsUrl ?? null,
        publicAccessCheck: args.uiSameCorpusPublicAccessCheck ?? null,
      }),
    }),
  }
}

interface FinancialWorkbookCorpusStatus {
  readonly targetWorkbookCount: number
  readonly sourceCount: number
  readonly cachedArtifactCount: number
  readonly recordedManifestArtifactCount: number
  readonly recordedNonPassingCaseCount: number
}

function publicWorkbookCorpusDominanceBlockers(status: PublicWorkbookCorpusStatus): string[] {
  return status.gaps.map((gap) => {
    const scorecardCoverage = /^scorecard cases do not cover manifest artifacts: (.+)$/u.exec(gap)
    if (scorecardCoverage) {
      return `public workbook corpus scorecard cases below cached artifacts: ${scorecardCoverage[1]}`
    }
    return `public workbook corpus ${gap}`
  })
}

function publicWorkbookCorpusFeatureWitnessBlockers(plan: PublicWorkbookCorpusFeatureWitnessPlan | null): string[] {
  const missingWitnesses = plan?.coverage.filter((entry) => entry.needsWitness).map((entry) => entry.label) ?? []
  if (missingWitnesses.length === 0) {
    return []
  }
  return [`public workbook corpus missing feature witness coverage: ${missingWitnesses.join(', ')}`]
}

function financialWorkbookCorpusDominanceBlockers(status: FinancialWorkbookCorpusStatus | null): string[] {
  if (!status) {
    return []
  }
  return [
    ...countDominanceGap(
      status.cachedArtifactCount,
      status.targetWorkbookCount,
      'financial/accounting corpus cached artifacts below target',
    ),
    ...countDominanceGap(
      status.recordedManifestArtifactCount,
      status.targetWorkbookCount,
      'financial/accounting corpus recorded verification cases below target',
    ),
    ...(status.recordedManifestArtifactCount >= status.cachedArtifactCount
      ? []
      : [
          `financial/accounting corpus cached artifacts missing verification evidence: ${String(
            status.cachedArtifactCount - status.recordedManifestArtifactCount,
          )}`,
        ]),
    ...(status.recordedNonPassingCaseCount === 0
      ? []
      : [`financial/accounting corpus non-passing recorded cases: ${String(status.recordedNonPassingCaseCount)}`]),
  ]
}

function localCiResourceGuardDominanceBlockers(status: LocalCiResourceGuardStatus): string[] {
  return status.activeMarkerPaths.length === 0
    ? []
    : [`operator/developer workflow local CI resource guard active: ${status.activeMarkerPaths.join(', ')}`]
}

function countDominanceGap(actual: number, required: number, label: string): string[] {
  return actual >= required ? [] : [`${label}: ${String(actual)}/${String(required)}`]
}

function readPublicWorkbookManifestOrNull(path: string): PublicWorkbookManifest | null {
  return existsSync(path) ? parsePublicWorkbookManifestJson(JSON.parse(readFileSync(path, 'utf8')) as unknown) : null
}

function buildFinancialPlanStatus(
  plan: PublicWorkbookCorpusFinancialPlan,
): NonNullable<BiligDominanceStatus['publicWorkbookCorpus']['financialPlan']> {
  return {
    sourceCount: plan.sourceCount,
    targetArtifactCount: plan.targetArtifactCount,
    cachedArtifactCount: plan.cachedArtifactCount,
    remainingArtifactSlots: plan.remainingArtifactSlots,
    candidateSourceCount: plan.candidateSourceCount,
    candidateSourceDeficitCount: plan.candidateSourceDeficitCount,
    recommendedFetchLimit: plan.recommendedFetchLimit,
    needsAdditionalDiscovery: plan.needsAdditionalDiscovery,
    targetReachableFromKnownCandidates: plan.targetReachableFromKnownCandidates,
    nextPlanCommand: 'pnpm public-workbook-corpus:discover-financial:plan',
    nextCheckCommand: 'pnpm public-workbook-corpus:discover-financial:check',
    nextFetchPlanCommand: plan.commands.fetchPlan,
    nextFetchCommand: plan.stopMarker.active ? null : plan.commands.fetch,
    nextVerifyCommand: plan.stopMarker.active ? null : plan.commands.verify,
    blockedCommands: plan.stopMarker.active
      ? nonEmptyCommands([plan.commands.discover, plan.commands.fetch, plan.commands.fetchAll, plan.commands.verify])
      : [],
  }
}

function buildFeatureWitnessPlanStatus(
  plan: PublicWorkbookCorpusFeatureWitnessPlan,
): NonNullable<BiligDominanceStatus['publicWorkbookCorpus']['featureWitnessPlan']> {
  return {
    recordedCaseCount: plan.recordedCaseCount,
    missingWitnessCount: plan.missingWitnessCount,
    nextPlanCommand: 'pnpm public-workbook-corpus:feature-witness:plan',
    nextCheckCommand: 'pnpm public-workbook-corpus:feature-witness:check',
    coverage: plan.coverage.map((entry) => ({
      id: entry.id,
      label: entry.label,
      totalCount: entry.totalCount,
      witnessCaseCount: entry.witnessCaseCount,
      needsWitness: entry.needsWitness,
      discoveryQuery: entry.discoveryQuery,
      nextDiscoverCommand: entry.commands.discover,
      blockedDiscoverCommand: entry.blockedCommands.discover,
    })),
  }
}

function buildFinancialWorkbookCorpusStatus(args: {
  readonly manifest: PublicWorkbookManifest | null
  readonly recordedCases: ReturnType<typeof readReusablePublicWorkbookCorpusCases>
  readonly targetWorkbookCount: number
}): FinancialWorkbookCorpusStatus {
  const recordedCasesById = new Map(args.recordedCases.map((entry) => [entry.id, entry]))
  const recordedManifestCases = (args.manifest?.artifacts ?? []).flatMap((artifact) => {
    const candidate = recordedCasesById.get(artifact.id)
    return candidate && publicWorkbookCorpusCaseMatchesArtifact(candidate, artifact) ? [candidate] : []
  })
  return {
    targetWorkbookCount: args.manifest?.targetWorkbookCount ?? args.targetWorkbookCount,
    sourceCount: args.manifest?.sources.length ?? 0,
    cachedArtifactCount: args.manifest?.artifacts.length ?? 0,
    recordedManifestArtifactCount: recordedManifestCases.length,
    recordedNonPassingCaseCount: recordedManifestCases.filter((entry) => !entry.passed).length,
  }
}

function readRecordedFinancialCases(args: {
  readonly manifest: PublicWorkbookManifest | null
  readonly scorecardPath: string
  readonly verifyCheckpointPath: string
}): ReturnType<typeof readReusablePublicWorkbookCorpusCases> {
  const reusableCases = readReusablePublicWorkbookCorpusCases([args.scorecardPath, args.verifyCheckpointPath])
  if (!args.manifest) {
    return reusableCases
  }
  const casesById = new Map(reusableCases.map((entry) => [entry.id, entry]))
  return args.manifest.artifacts.flatMap((artifact) => {
    const candidate = casesById.get(artifact.id)
    return candidate && publicWorkbookCorpusCaseMatchesArtifact(candidate, artifact) ? [candidate] : []
  })
}

function buildUiSameCorpusStatus(
  input: BuildScorecardInput,
  args: {
    readonly googleSheetsUrl: string | null
    readonly googleSheetsUrlSource: UiSameCorpusGoogleSheetsUrlSource
    readonly localCiResourceGuardStatus: LocalCiResourceGuardStatus
    readonly publicAccessCheckPath: string
  },
): BiligDominanceStatus['uiSameCorpus'] {
  const proof = input.uiResponsivenessLiveBrowserScorecard.sameCorpusProof
  const fixture = uiSameCorpusFixtureStatus(defaultUiSameCorpusId)
  const coveredWorkloads = new Set(proof.cases.map((entry) => entry.workload))
  const missingRequiredWorkloads = requiredUiSameCorpusWorkloads.filter((workload) => !coveredWorkloads.has(workload))
  const tenXRequirementSatisfied = uiSameCorpusTenXRequirementSatisfied(proof, missingRequiredWorkloads)
  const googleSheetsUrlArgument = args.googleSheetsUrl ?? '<google-sheets-url>'
  const browserCaptureGuard = buildBrowserCaptureGuardStatus(args.localCiResourceGuardStatus)
  const nextPreflightCommand = [
    'pnpm',
    'ui:same-corpus:capture',
    '--',
    '--preflight',
    '--google-sheets-url',
    googleSheetsUrlArgument,
    '--microsoft-excel-web-url',
    fixture.microsoftExcelWebUrl,
  ]
    .map(shellQuote)
    .join(' ')
  const nextCaptureCommand = [
    'pnpm',
    'ui:same-corpus:capture',
    '--',
    '--output',
    '.cache/ui-responsiveness/same-corpus-capture.json',
    '--google-sheets-url',
    googleSheetsUrlArgument,
    '--microsoft-excel-web-url',
    fixture.microsoftExcelWebUrl,
  ]
    .map(shellQuote)
    .join(' ')
  return {
    captured: proof.captured,
    evidenceKind: proof.evidenceKind,
    requiredProductCount: proof.requiredProductCount,
    requiredCaseCount: proof.requiredCaseCount,
    tenXMeanAndP95CaseCount: proof.tenXMeanAndP95CaseCount,
    tenXRequirementSatisfied,
    requiredWorkloads: requiredUiSameCorpusWorkloads,
    missingRequiredWorkloads,
    coveredCorpusCaseIds: proof.coveredCorpusCaseIds,
    limitations: proof.limitations,
    fixture,
    googleSheetsUrl: args.googleSheetsUrl,
    googleSheetsUrlSource: args.googleSheetsUrlSource,
    googleSheetsUrlEnvVar: uiSameCorpusGoogleSheetsUrlEnvVar,
    publicAccessCheckPath: args.publicAccessCheckPath,
    missingInputs: args.googleSheetsUrl || tenXRequirementSatisfied ? [] : ['googleSheetsUrlForUploadedSameCorpusWorkbook'],
    nextFixtureCheckCommand: 'pnpm ui:same-corpus:fixture:check',
    nextPublicAccessCheckCommand: [
      'pnpm',
      'ui:same-corpus:public-check',
      '--',
      '--output',
      args.publicAccessCheckPath,
      '--google-sheets-url',
      googleSheetsUrlArgument,
      '--microsoft-excel-web-url',
      fixture.microsoftExcelWebUrl,
    ]
      .map(shellQuote)
      .join(' '),
    nextGoogleSheetsUploadInstruction: `Upload ${fixture.localXlsxPath} to Google Sheets as a native Google Sheet, share it to anyone with the link, then pass its edit URL as --google-sheets-url.`,
    nextPreflightCommand: browserCaptureGuard.active ? null : nextPreflightCommand,
    nextCaptureCommand: browserCaptureGuard.active ? null : nextCaptureCommand,
    blockedCommands: browserCaptureGuard.active ? [nextPreflightCommand, nextCaptureCommand].map(localCiResourceGuardOverrideCommand) : [],
    browserCaptureGuard,
    nextScorecardGenerateCommand: 'pnpm ui:browser-live:generate -- --capture .cache/ui-responsiveness/same-corpus-capture.json',
    nextDominanceCheckCommand: 'pnpm dominance:generate && pnpm dominance:check && pnpm dominance:audit:check',
  }
}

function buildBrowserCaptureGuardStatus(status: LocalCiResourceGuardStatus): BiligDominanceStatus['uiSameCorpus']['browserCaptureGuard'] {
  const active = status.activeMarkerPaths.length > 0
  return {
    active,
    activeMarkerPaths: status.activeMarkerPaths,
    overrideEnvVar: localCiResourceGuardOverrideEnv,
    overridePrefix: active ? `${localCiResourceGuardOverrideEnv}=1` : null,
    nextPreflightRequiresOverride: active,
    nextCaptureRequiresOverride: active,
  }
}

function buildLocalCiResourceGuardStatus(status: LocalCiResourceGuardStatus): BiligDominanceStatus['localCiResourceGuard'] {
  const active = status.activeMarkerPaths.length > 0
  return {
    active,
    activeMarkerPaths: status.activeMarkerPaths,
    overrideEnvVar: localCiResourceGuardOverrideEnv,
    overridePrefix: active ? `${localCiResourceGuardOverrideEnv}=1` : null,
  }
}

function readSameCorpusPublicAccessCheckOrNull(path: string): SameCorpusPublicAccessCheck | null {
  if (!existsSync(path)) {
    return null
  }
  try {
    const parsed: unknown = JSON.parse(readFileSync(path, 'utf8'))
    return parseSameCorpusPublicAccessCheckJson(parsed)
  } catch {
    return null
  }
}

function resolveUiSameCorpusGoogleSheetsUrl(args: {
  readonly corpusCaseId: WorkbookBenchmarkCorpusId
  readonly explicitGoogleSheetsUrl: string | null
  readonly publicAccessCheck: SameCorpusPublicAccessCheck | null
}): {
  readonly googleSheetsUrl: string | null
  readonly googleSheetsUrlSource: UiSameCorpusGoogleSheetsUrlSource
} {
  if (args.explicitGoogleSheetsUrl) {
    return {
      googleSheetsUrl: args.explicitGoogleSheetsUrl,
      googleSheetsUrlSource: 'argument-or-environment',
    }
  }
  const verifiedPublicAccessUrl = verifiedGoogleSheetsUrlFromPublicAccessCheck(args.publicAccessCheck, args.corpusCaseId)
  if (verifiedPublicAccessUrl) {
    return {
      googleSheetsUrl: verifiedPublicAccessUrl,
      googleSheetsUrlSource: 'public-access-check',
    }
  }
  return {
    googleSheetsUrl: null,
    googleSheetsUrlSource: 'missing',
  }
}

function verifiedGoogleSheetsUrlFromPublicAccessCheck(
  check: SameCorpusPublicAccessCheck | null,
  corpusCaseId: WorkbookBenchmarkCorpusId,
): string | null {
  if (!check || check.corpusCaseId !== corpusCaseId) {
    return null
  }
  const corpus = buildWorkbookBenchmarkCorpus(corpusCaseId)
  if (check.materializedCells !== corpus.materializedCellCount) {
    return null
  }
  const product = check.products.find((entry) => entry.product === 'google-sheets')
  return product?.corpusVerification.verified ? product.source : null
}

function uiSameCorpusTenXRequirementSatisfied(
  proof: BuildScorecardInput['uiResponsivenessLiveBrowserScorecard']['sameCorpusProof'],
  missingRequiredWorkloads: readonly UiResponsivenessSameCorpusWorkload[],
): boolean {
  return (
    proof.captured &&
    proof.evidenceKind === 'same-corpus-browser-capture' &&
    proof.requiredProductCount === 3 &&
    proof.requiredCaseCount > 0 &&
    proof.cases.length === proof.requiredCaseCount &&
    proof.tenXMeanAndP95CaseCount === proof.requiredCaseCount &&
    missingRequiredWorkloads.length === 0 &&
    proof.cases.every((entry) => entry.passed)
  )
}

function uiSameCorpusFixtureStatus(corpusCaseId: WorkbookBenchmarkCorpusId): BiligDominanceStatus['uiSameCorpus']['fixture'] {
  const corpus = buildWorkbookBenchmarkCorpus(corpusCaseId)
  const localXlsxPath = `packages/benchmarks/baselines/ui-same-corpus/${corpus.id}.xlsx`
  const publicGithubRawUrl = `https://raw.githubusercontent.com/proompteng/bilig/main/${localXlsxPath}`
  return {
    corpusCaseId,
    materializedCells: corpus.materializedCellCount,
    localXlsxPath,
    publicGithubRawUrl,
    publicForgejoRawUrl: `https://code.proompteng.ai/kalmyk/bilig/raw/branch/main/${localXlsxPath}`,
    microsoftExcelWebUrl: `https://view.officeapps.live.com/op/view.aspx?src=${encodeURIComponent(publicGithubRawUrl)}`,
  }
}

function formatPublicWorkbookCorpusDiscoveryPlanCommand(limit: number): string {
  return ['pnpm', 'public-workbook-corpus:discover:plan', '--', '--limit', String(limit)].map(shellQuote).join(' ')
}

function formatPublicWorkbookCorpusDiscoveryCommand(limit: number): string {
  return ['pnpm', 'public-workbook-corpus:discover', '--', '--limit', String(limit)].map(shellQuote).join(' ')
}

function formatPublicWorkbookCorpusFetchPlanCommand(args: {
  readonly cacheDir: string
  readonly displayRootDir?: string
  readonly limit: number
  readonly manifestPath: string
}): string {
  return [
    'pnpm',
    'public-workbook-corpus:fetch:plan',
    '--',
    '--manifest',
    formatBiligDominanceStatusPathForMessage(args.manifestPath, args.displayRootDir),
    '--cache-dir',
    formatBiligDominanceStatusPathForMessage(args.cacheDir, args.displayRootDir),
    '--limit',
    String(args.limit),
  ]
    .map(shellQuote)
    .join(' ')
}

export function formatBiligDominanceStatusPathForMessage(path: string, displayRootDir: string | undefined): string {
  if (!displayRootDir) {
    return path
  }
  const relativePath = relative(displayRootDir, path)
  if (!relativePath || relativePath.startsWith('..') || isAbsolute(relativePath)) {
    return path
  }
  return relativePath
}

function shellQuote(value: string): string {
  return /^[A-Za-z0-9_./:=@+-]+$/u.test(value) ? value : `'${value.replaceAll("'", "'\\''")}'`
}

function nonEmptyCommands(commands: readonly (string | null)[]): string[] {
  return commands.filter((command): command is string => typeof command === 'string' && command.trim().length > 0)
}

function corpusStopMarkerOverrideCommand(command: string): string {
  if (command.includes(`${publicCorpusStopMarkerOverrideEnvVar}=1`)) {
    return command
  }
  return `${publicCorpusStopMarkerOverrideEnvVar}=1 ${command} ${publicCorpusStopMarkerOverrideFlag}`
}

function localCiResourceGuardOverrideCommand(command: string): string {
  if (command.includes(`${localCiResourceGuardOverrideEnv}=1`)) {
    return command
  }
  return `${localCiResourceGuardOverrideEnv}=1 ${command}`
}

if (process.argv[1] && pathToFileURL(resolve(process.argv[1])).href === import.meta.url) {
  main()
}
