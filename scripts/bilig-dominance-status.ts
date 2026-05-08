#!/usr/bin/env bun

import { existsSync, readFileSync } from 'node:fs'
import { isAbsolute, join, relative, resolve } from 'node:path'
import { pathToFileURL } from 'node:url'

import type { BuildScorecardInput } from './bilig-dominance-scorecard-types.ts'
import { loadBiligDominanceScorecardInput, rootDir } from './bilig-dominance-scorecard-input.ts'
import { buildBiligDominanceScorecard } from './gen-bilig-dominance-scorecard.ts'
import type { UiResponsivenessSameCorpusWorkload } from './gen-ui-responsiveness-live-browser-scorecard.ts'
import { buildWorkbookBenchmarkCorpus, type WorkbookBenchmarkCorpusId } from '../packages/benchmarks/src/workbook-corpus.js'
import { publicCorpusStopMarkerOverrideEnvVar, publicCorpusStopMarkerOverrideFlag, readStringArg } from './public-workbook-corpus-cli.ts'
import { planPublicWorkbookCorpusFetch, type PublicWorkbookCorpusFetchPlan } from './public-workbook-corpus-fetch.ts'
import { parsePublicWorkbookManifestJson } from './public-workbook-corpus-json.ts'
import { publicWorkbookCorpusCaseMatchesArtifact } from './public-workbook-corpus-missing.ts'
import type { PublicWorkbookCorpusStatus } from './public-workbook-corpus-status.ts'
import { readPublicWorkbookCorpusStatus } from './public-workbook-corpus-status.ts'
import { readReusablePublicWorkbookCorpusCases } from './public-workbook-corpus-verify-checkpoint.ts'
import { financialWorkbookTargetCount } from './public-workbook-corpus-completion-audit-helpers.ts'
import type { PublicWorkbookManifest } from './public-workbook-corpus-types.ts'

export interface BiligDominanceStatus {
  readonly goalStatus: 'achieved' | 'active-not-achieved'
  readonly blanketTenXClaimAllowed: boolean
  readonly unmetRequirements: readonly string[]
  readonly importExportBlockers: readonly string[]
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
    readonly fetchCandidateSourceCount: number | null
    readonly fetchCandidateSourceDeficitCount: number | null
    readonly minimumAdditionalSourceCount: number | null
    readonly recommendedDiscoveryLimit: number | null
    readonly targetReachableFromKnownCandidates: boolean | null
    readonly scorecardCaseCount: number
    readonly checkpointCaseCount: number
    readonly recordedAllCasesPassed: boolean
    readonly nextFetchPlanCommand: string | null
    readonly nextDiscoveryPlanCommand: string | null
    readonly nextDiscoveryCommand: string | null
    readonly nextMissingVerificationPlanCommand: string | null
    readonly nextMissingVerificationCommand: string | null
    readonly nextStaleVerificationPlanCommand: string | null
    readonly nextStaleVerificationCommand: string | null
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
    readonly missingInputs: readonly string[]
    readonly nextFixtureCheckCommand: string
    readonly nextGoogleSheetsUploadInstruction: string
    readonly nextPreflightCommand: string
    readonly nextCaptureCommand: string
    readonly nextScorecardGenerateCommand: string
    readonly nextDominanceCheckCommand: string
  }
}

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
const requiredUiSameCorpusWorkloads = ['visible-scroll-response'] as const satisfies readonly UiResponsivenessSameCorpusWorkload[]

function main(): void {
  process.stdout.write(`${JSON.stringify(buildBiligDominanceStatusFromArgs(), null, 2)}\n`)
}

export function buildBiligDominanceStatusFromArgs(): BiligDominanceStatus {
  const cacheDir = resolve(readStringArg('--cache-dir', defaultCacheDir))
  const scorecardPath = resolve(readStringArg('--scorecard', defaultScorecardPath))
  const manifestPath = resolve(readStringArg('--manifest', defaultManifestPath))
  const verifyCheckpointPath = resolve(readStringArg('--verify-checkpoint', defaultVerifyCheckpointPath))
  const financialManifestPath = resolve(readStringArg('--financial-manifest', defaultFinancialManifestPath))
  const financialScorecardPath = resolve(readStringArg('--financial-scorecard', defaultFinancialScorecardPath))
  const financialVerifyCheckpointPath = resolve(readStringArg('--financial-verify-checkpoint', defaultFinancialVerifyCheckpointPath))
  const stopMarkerPath = resolve(readStringArg('--corpus-run-stop-marker', defaultCorpusRunStopMarkerPath))
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
  return buildBiligDominanceStatus({
    fetchPlan,
    financialCorpusStatus: readFinancialWorkbookCorpusStatus({
      manifestPath: financialManifestPath,
      scorecardPath: financialScorecardPath,
      targetWorkbookCount: publicWorkbookCorpusStatus.targetWorkbookCount,
      verifyCheckpointPath: financialVerifyCheckpointPath,
    }),
    input,
    nextFetchPlanCommand,
    publicWorkbookCorpusStatus,
    stopMarkerActive,
    stopMarkerPath: formatBiligDominanceStatusPathForMessage(stopMarkerPath, rootDir),
  })
}

export function buildBiligDominanceStatus(args: {
  readonly fetchPlan?: PublicWorkbookCorpusFetchPlan | null
  readonly financialCorpusStatus?: FinancialWorkbookCorpusStatus | null
  readonly input: BuildScorecardInput
  readonly nextFetchPlanCommand?: string | null
  readonly publicWorkbookCorpusStatus: PublicWorkbookCorpusStatus
  readonly stopMarkerActive: boolean
  readonly stopMarkerPath: string
}): BiligDominanceStatus {
  const scorecard = buildBiligDominanceScorecard({
    ...args.input,
    publicWorkbookCorpusStatus: args.publicWorkbookCorpusStatus,
  })
  const importExportCategory = scorecard.categories.find((category) => category.id === 'import-export-compatibility')
  const publicWorkbookCorpusBlockers = publicWorkbookCorpusDominanceBlockers(args.publicWorkbookCorpusStatus)
  const financialCorpusBlockers = financialWorkbookCorpusDominanceBlockers(args.financialCorpusStatus ?? null)
  const corpusBlockers = [...publicWorkbookCorpusBlockers, ...financialCorpusBlockers]
  const unmetRequirements = [...scorecard.claimPolicy.unmetRequirements, ...corpusBlockers]
  const blanketTenXClaimAllowed = scorecard.claimPolicy.blanketTenXClaimAllowed && corpusBlockers.length === 0
  return {
    goalStatus: scorecard.goalStatus === 'achieved' && corpusBlockers.length === 0 ? 'achieved' : 'active-not-achieved',
    blanketTenXClaimAllowed,
    unmetRequirements,
    importExportBlockers: [...(importExportCategory?.blockers ?? []), ...corpusBlockers],
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
      fetchCandidateSourceCount: args.fetchPlan?.candidateSourceCount ?? null,
      fetchCandidateSourceDeficitCount: args.fetchPlan?.candidateSourceDeficitCount ?? null,
      targetReachableFromKnownCandidates: args.fetchPlan?.targetReachableFromKnownCandidates ?? null,
      minimumAdditionalSourceCount: args.fetchPlan?.minimumAdditionalSourceCount ?? null,
      recommendedDiscoveryLimit: args.fetchPlan?.recommendedDiscoveryLimit ?? null,
      scorecardCaseCount: args.publicWorkbookCorpusStatus.scorecardCaseCount,
      checkpointCaseCount: args.publicWorkbookCorpusStatus.checkpointCaseCount,
      recordedAllCasesPassed: args.publicWorkbookCorpusStatus.recordedAllCasesPassed,
      nextFetchPlanCommand: args.nextFetchPlanCommand ?? null,
      nextDiscoveryPlanCommand:
        args.fetchPlan && args.fetchPlan.candidateSourceDeficitCount > 0
          ? formatPublicWorkbookCorpusDiscoveryPlanCommand(args.fetchPlan.recommendedDiscoveryLimit)
          : null,
      nextDiscoveryCommand:
        args.fetchPlan && args.fetchPlan.candidateSourceDeficitCount > 0
          ? formatPublicWorkbookCorpusDiscoveryCommand(args.fetchPlan.recommendedDiscoveryLimit, args.stopMarkerActive)
          : null,
      nextMissingVerificationPlanCommand: args.publicWorkbookCorpusStatus.nextMissingVerificationPlanCommand,
      nextMissingVerificationCommand: args.publicWorkbookCorpusStatus.nextMissingVerificationCommand
        ? stopMarkerGuardedCorpusCommand(args.publicWorkbookCorpusStatus.nextMissingVerificationCommand, args.stopMarkerActive)
        : null,
      nextStaleVerificationPlanCommand: args.publicWorkbookCorpusStatus.nextStaleVerificationPlanCommand,
      nextStaleVerificationCommand: args.publicWorkbookCorpusStatus.nextStaleVerificationCommand
        ? stopMarkerGuardedCorpusCommand(args.publicWorkbookCorpusStatus.nextStaleVerificationCommand, args.stopMarkerActive)
        : null,
      corpusRunStopMarkerActive: args.stopMarkerActive,
      corpusRunStopMarkerPath: args.stopMarkerPath,
      nextCorpusRunRequiresExplicitResume: args.stopMarkerActive,
      corpusRunStopMarkerOverrideFlag: publicCorpusStopMarkerOverrideFlag,
      corpusRunStopMarkerOverrideEnvVar: publicCorpusStopMarkerOverrideEnvVar,
      gaps: args.publicWorkbookCorpusStatus.gaps,
    },
    uiSameCorpus: buildUiSameCorpusStatus(args.input),
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

function countDominanceGap(actual: number, required: number, label: string): string[] {
  return actual >= required ? [] : [`${label}: ${String(actual)}/${String(required)}`]
}

function readFinancialWorkbookCorpusStatus(args: {
  readonly manifestPath: string
  readonly scorecardPath: string
  readonly targetWorkbookCount: number
  readonly verifyCheckpointPath: string
}): FinancialWorkbookCorpusStatus {
  const manifest = existsSync(args.manifestPath)
    ? parsePublicWorkbookManifestJson(JSON.parse(readFileSync(args.manifestPath, 'utf8')) as unknown)
    : null
  return buildFinancialWorkbookCorpusStatus({
    manifest,
    recordedCases: readRecordedFinancialCases({
      manifest,
      scorecardPath: args.scorecardPath,
      verifyCheckpointPath: args.verifyCheckpointPath,
    }),
    targetWorkbookCount: financialWorkbookTargetCount(args.targetWorkbookCount),
  })
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

function buildUiSameCorpusStatus(input: BuildScorecardInput): BiligDominanceStatus['uiSameCorpus'] {
  const proof = input.uiResponsivenessLiveBrowserScorecard.sameCorpusProof
  const fixture = uiSameCorpusFixtureStatus(defaultUiSameCorpusId)
  const coveredWorkloads = new Set(proof.cases.map((entry) => entry.workload))
  return {
    captured: proof.captured,
    evidenceKind: proof.evidenceKind,
    requiredProductCount: proof.requiredProductCount,
    requiredCaseCount: proof.requiredCaseCount,
    tenXMeanAndP95CaseCount: proof.tenXMeanAndP95CaseCount,
    requiredWorkloads: requiredUiSameCorpusWorkloads,
    missingRequiredWorkloads: requiredUiSameCorpusWorkloads.filter((workload) => !coveredWorkloads.has(workload)),
    coveredCorpusCaseIds: proof.coveredCorpusCaseIds,
    limitations: proof.limitations,
    fixture,
    missingInputs: proof.captured ? [] : ['googleSheetsUrlForUploadedSameCorpusWorkbook'],
    nextFixtureCheckCommand: 'pnpm ui:same-corpus:fixture:check',
    nextGoogleSheetsUploadInstruction: `Upload ${fixture.localXlsxPath} to Google Sheets as a native Google Sheet, share it to anyone with the link, then pass its edit URL as --google-sheets-url.`,
    nextPreflightCommand: [
      'pnpm',
      'ui:same-corpus:capture',
      '--',
      '--preflight',
      '--google-sheets-url',
      '<google-sheets-url>',
      '--microsoft-excel-web-url',
      fixture.microsoftExcelWebUrl,
    ]
      .map(shellQuote)
      .join(' '),
    nextCaptureCommand: [
      'pnpm',
      'ui:same-corpus:capture',
      '--',
      '--output',
      '.cache/ui-responsiveness/same-corpus-capture.json',
      '--google-sheets-url',
      '<google-sheets-url>',
      '--microsoft-excel-web-url',
      fixture.microsoftExcelWebUrl,
    ]
      .map(shellQuote)
      .join(' '),
    nextScorecardGenerateCommand: 'pnpm ui:browser-live:generate -- --capture .cache/ui-responsiveness/same-corpus-capture.json',
    nextDominanceCheckCommand: 'pnpm dominance:generate && pnpm dominance:check && pnpm dominance:audit:check',
  }
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

function formatPublicWorkbookCorpusDiscoveryCommand(limit: number, stopMarkerActive: boolean): string {
  return stopMarkerGuardedCorpusCommand(
    ['pnpm', 'public-workbook-corpus:discover', '--', '--limit', String(limit)].map(shellQuote).join(' '),
    stopMarkerActive,
  )
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

function stopMarkerGuardedCorpusCommand(command: string, stopMarkerActive: boolean): string {
  if (!stopMarkerActive || command.includes(`${publicCorpusStopMarkerOverrideEnvVar}=1`)) {
    return command
  }
  return `${publicCorpusStopMarkerOverrideEnvVar}=1 ${command} ${publicCorpusStopMarkerOverrideFlag}`
}

if (process.argv[1] && pathToFileURL(resolve(process.argv[1])).href === import.meta.url) {
  main()
}
