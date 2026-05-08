import { existsSync, readFileSync } from 'node:fs'
import { isAbsolute, relative, resolve } from 'node:path'

import { parsePublicWorkbookCorpusScorecardJson, parsePublicWorkbookManifestJson } from './public-workbook-corpus-json.ts'
import { publicCorpusStopMarkerOverrideEnvVar, publicCorpusStopMarkerOverrideFlag } from './public-workbook-corpus-cli.ts'
import { listStalePublicWorkbookArtifacts } from './public-workbook-corpus-missing.ts'
import { readReusablePublicWorkbookCorpusCases } from './public-workbook-corpus-verify-checkpoint.ts'
import {
  publicWorkbookCorpusCaseEvidenceRefreshReasons,
  publicWorkbookCorpusCaseNeedsEvidenceRefresh,
  type PublicWorkbookCorpusEvidenceRefreshReason,
} from './public-workbook-corpus-evidence.ts'
import type {
  PublicWorkbookArtifact,
  PublicWorkbookCorpusCase,
  PublicWorkbookCorpusScorecard,
  PublicWorkbookManifest,
} from './public-workbook-corpus-types.ts'

export interface PublicWorkbookCorpusStatus {
  readonly targetWorkbookCount: number
  readonly sourceCount: number
  readonly cachedArtifactCount: number
  readonly scorecardCaseCount: number
  readonly checkpointCaseCount: number
  readonly recordedManifestArtifactCount: number
  readonly missingManifestArtifactCount: number
  readonly staleRecordedVerificationCount: number
  readonly recordedPassedCaseCount: number
  readonly recordedUnsupportedCaseCount: number
  readonly currentRecordedUnsupportedCaseCount: number
  readonly staleRecordedUnsupportedCaseCount: number
  readonly currentUnsupportedClassifications: readonly PublicWorkbookCorpusUnsupportedClassificationCount[]
  readonly staleUnsupportedClassifications: readonly PublicWorkbookCorpusUnsupportedClassificationCount[]
  readonly recordedFailedCaseCount: number
  readonly recordedErrorCaseCount: number
  readonly recordedCoversManifest: boolean
  readonly recordedAllCasesPassed: boolean
  readonly missingManifestArtifactSample: readonly MissingManifestArtifactSummary[]
  readonly staleRecordedVerificationSample: readonly StaleRecordedVerificationSummary[]
  readonly nextMissingVerificationCommand: string | null
  readonly nextMissingVerificationPlanCommand: string | null
  readonly blockedMissingVerificationCommand: string | null
  readonly nextStaleVerificationCommand: string | null
  readonly nextStaleVerificationPlanCommand: string | null
  readonly blockedStaleVerificationCommand: string | null
  readonly scorecardCoversManifest: boolean
  readonly targetComplete: boolean
  readonly gaps: readonly string[]
}

export interface MissingManifestArtifactSummary {
  readonly id: string
  readonly fileName: string
  readonly byteSize: number
  readonly sourceUrl: string
}

export interface StaleRecordedVerificationSummary extends MissingManifestArtifactSummary {
  readonly reason: PublicWorkbookCorpusEvidenceRefreshReason
  readonly reasons: readonly PublicWorkbookCorpusEvidenceRefreshReason[]
}

export interface PublicWorkbookCorpusUnsupportedClassificationCount {
  readonly classification: string
  readonly count: number
}

const missingManifestArtifactSampleLimit = 20
const rootDir = resolve(new URL('..', import.meta.url).pathname)

export function writePublicWorkbookCorpusCheck(args: {
  readonly manifestPath: string
  readonly scorecardPath: string
  readonly cacheDir: string
  readonly verifyCheckpointPath: string
  readonly skipManifestCheck: boolean
  readonly requireTarget: boolean
  readonly corpusRunStopMarkerPath?: string
}): void {
  if (!args.skipManifestCheck && existsSync(args.manifestPath)) {
    const status = readPublicWorkbookCorpusStatus(args)
    const blockingGaps = publicWorkbookCorpusVerificationBlockingGaps(status)
    if (blockingGaps.length > 0) {
      const nextCommand =
        status.nextMissingVerificationPlanCommand ??
        status.nextStaleVerificationPlanCommand ??
        status.nextMissingVerificationCommand ??
        status.nextStaleVerificationCommand
      const nextCommandMessage = nextCommand ? `; next command: ${nextCommand}` : ''
      throw new Error(`Public workbook corpus verification incomplete: ${blockingGaps.join('; ')}${nextCommandMessage}`)
    }
    if (args.requireTarget && status.cachedArtifactCount < status.targetWorkbookCount) {
      throw new Error(
        `Public workbook corpus target incomplete: ${status.gaps.join('; ')}; next command: ${formatPublicWorkbookCorpusResumePlanCheckCommand(
          {
            cacheDir: args.cacheDir,
            displayRootDir: rootDir,
            manifestPath: args.manifestPath,
            scorecardPath: args.scorecardPath,
            verifyCheckpointPath: args.verifyCheckpointPath,
          },
        )}`,
      )
    }
    console.log(
      `Checked public workbook corpus with ${String(status.recordedManifestArtifactCount)}/${String(
        status.cachedArtifactCount,
      )} recorded cached workbooks`,
    )
    return
  }
  const scorecard = parsePublicWorkbookCorpusScorecardJson(JSON.parse(readFileSync(args.scorecardPath, 'utf8')))
  if (args.requireTarget && scorecard.summary.remainingToTarget > 0) {
    throw new Error(`Public workbook corpus target incomplete: ${String(scorecard.summary.remainingToTarget)} remaining`)
  }
  console.log(`Checked public workbook corpus scorecard with ${String(scorecard.summary.cachedWorkbookCount)} cached workbooks`)
}

export function writePublicWorkbookCorpusStatus(args: {
  readonly manifestPath: string
  readonly scorecardPath: string
  readonly cacheDir: string
  readonly verifyCheckpointPath: string
  readonly requireTarget: boolean
  readonly corpusRunStopMarkerPath?: string
}): void {
  const status = readPublicWorkbookCorpusStatus(args)
  process.stdout.write(`${JSON.stringify(status, null, 2)}\n`)
  if (args.requireTarget && !status.targetComplete) {
    throw new Error(`Public workbook corpus target incomplete: ${status.gaps.join('; ')}`)
  }
}

export function readPublicWorkbookCorpusStatus(args: {
  readonly manifestPath: string
  readonly scorecardPath: string
  readonly cacheDir: string
  readonly verifyCheckpointPath: string
  readonly corpusRunStopMarkerPath?: string
}): PublicWorkbookCorpusStatus {
  const manifest = existsSync(args.manifestPath)
    ? parsePublicWorkbookManifestJson(JSON.parse(readFileSync(args.manifestPath, 'utf8')))
    : null
  const scorecard = existsSync(args.scorecardPath)
    ? parsePublicWorkbookCorpusScorecardJson(JSON.parse(readFileSync(args.scorecardPath, 'utf8')))
    : null
  const checkpointCases = readReusablePublicWorkbookCorpusCases([args.verifyCheckpointPath])
  return buildPublicWorkbookCorpusStatus({
    manifest,
    scorecard,
    checkpointCases,
    commandPaths: {
      manifestPath: args.manifestPath,
      scorecardPath: args.scorecardPath,
      cacheDir: args.cacheDir,
      verifyCheckpointPath: args.verifyCheckpointPath,
      displayRootDir: rootDir,
      stopMarkerActive: args.corpusRunStopMarkerPath ? existsSync(args.corpusRunStopMarkerPath) : false,
    },
  })
}

export function buildPublicWorkbookCorpusStatus(args: {
  readonly manifest: PublicWorkbookManifest | null
  readonly scorecard: PublicWorkbookCorpusScorecard | null
  readonly checkpointCases: readonly PublicWorkbookCorpusCase[]
  readonly commandPaths?: PublicWorkbookCorpusCommandPaths
}): PublicWorkbookCorpusStatus {
  const targetWorkbookCount = args.manifest?.targetWorkbookCount ?? args.scorecard?.summary.targetWorkbookCount ?? 10_000
  const sourceCount = args.manifest?.sources.length ?? args.scorecard?.summary.sourceCount ?? 0
  const cachedArtifactCount = args.manifest?.artifacts.length ?? args.scorecard?.summary.cachedWorkbookCount ?? 0
  const scorecardCaseCount = args.scorecard?.cases.length ?? 0
  const checkpointCaseCount = args.checkpointCases.length
  const candidateCases = [...(args.scorecard?.cases ?? []), ...args.checkpointCases]
  const recordedCases = args.manifest ? manifestRecordedCases(args.manifest, candidateCases) : candidateCases
  const recordedManifestArtifactCount = args.manifest ? recordedCases.length : Math.max(scorecardCaseCount, checkpointCaseCount)
  const missingManifestArtifactCount = Math.max(0, cachedArtifactCount - recordedManifestArtifactCount)
  const staleRecordedVerificationArtifacts = args.manifest
    ? listStalePublicWorkbookArtifacts({ manifest: args.manifest, cases: recordedCases })
    : []
  const staleRecordedVerificationCount = args.manifest
    ? staleRecordedVerificationArtifacts.length
    : recordedCases.filter(publicWorkbookCorpusCaseNeedsEvidenceRefresh).length
  const recordedPassedCaseCount = recordedCases.filter((entry) => entry.status === 'passed').length
  const recordedUnsupportedCases = recordedCases.filter((entry) => entry.status === 'unsupported')
  const staleRecordedUnsupportedCases = recordedUnsupportedCases.filter(publicWorkbookCorpusCaseNeedsEvidenceRefresh)
  const currentRecordedUnsupportedCases = recordedUnsupportedCases.filter((entry) => !publicWorkbookCorpusCaseNeedsEvidenceRefresh(entry))
  const staleRecordedUnsupportedCaseCount = staleRecordedUnsupportedCases.length
  const currentRecordedUnsupportedCaseCount = currentRecordedUnsupportedCases.length
  const recordedUnsupportedCaseCount = recordedUnsupportedCases.length
  const recordedFailedCaseCount = recordedCases.filter((entry) => entry.status === 'failed').length
  const recordedErrorCaseCount = recordedCases.filter((entry) => entry.status === 'error').length
  const recordedCasesById = new Map(recordedCases.map((entry) => [entry.id, entry]))
  const recordedCoversManifest = recordedManifestArtifactCount >= cachedArtifactCount
  const recordedAllCasesPassed = recordedCases.every((entry) => entry.passed)
  const missingManifestArtifactSample = args.manifest ? manifestMissingArtifactSample(args.manifest, candidateCases) : []
  const staleRecordedVerificationSample = staleRecordedVerificationArtifacts
    .slice(0, missingManifestArtifactSampleLimit)
    .map((artifact) => {
      const recordedCase = recordedCasesById.get(artifact.id)
      const reasons = recordedCase ? publicWorkbookCorpusCaseEvidenceRefreshReasons(recordedCase) : []
      return {
        id: artifact.id,
        fileName: artifact.fileName,
        byteSize: artifact.byteSize,
        sourceUrl: artifact.sourceUrl,
        reason: reasons[0] ?? 'missing-used-range-evidence',
        reasons,
      }
    })
  const nextMissingVerificationRun =
    missingManifestArtifactCount > 0 && args.commandPaths
      ? splitPublicWorkbookCorpusVerifySliceCommand(args.commandPaths, 'verify-missing')
      : { command: null, blockedCommand: null }
  const nextMissingVerificationPlanCommand =
    missingManifestArtifactCount > 0 && args.commandPaths
      ? formatPublicWorkbookCorpusVerifySliceCommand(args.commandPaths, 'verify-missing', 'plan')
      : null
  const nextStaleVerificationRun =
    staleRecordedVerificationCount > 0 && args.commandPaths
      ? splitPublicWorkbookCorpusVerifySliceCommand(args.commandPaths, 'verify-stale')
      : { command: null, blockedCommand: null }
  const nextStaleVerificationPlanCommand =
    staleRecordedVerificationCount > 0 && args.commandPaths
      ? formatPublicWorkbookCorpusVerifySliceCommand(args.commandPaths, 'verify-stale', 'plan')
      : null
  const scorecardCoversManifest =
    args.manifest && args.scorecard ? scorecardMatchesManifest(args.scorecard, args.manifest) : scorecardCaseCount >= cachedArtifactCount
  const targetComplete =
    cachedArtifactCount >= targetWorkbookCount &&
    recordedCoversManifest &&
    recordedAllCasesPassed &&
    staleRecordedVerificationCount === 0 &&
    recordedFailedCaseCount === 0 &&
    recordedErrorCaseCount === 0
  const gaps = [
    ...(sourceCount >= targetWorkbookCount
      ? []
      : [`discovered sources below target: ${String(sourceCount)}/${String(targetWorkbookCount)}`]),
    ...(cachedArtifactCount >= targetWorkbookCount
      ? []
      : [`cached artifacts below target: ${String(cachedArtifactCount)}/${String(targetWorkbookCount)}`]),
    ...(scorecardCoversManifest
      ? []
      : [`scorecard cases do not cover manifest artifacts: ${String(scorecardCaseCount)}/${String(cachedArtifactCount)}`]),
    ...(recordedManifestArtifactCount >= cachedArtifactCount
      ? []
      : [`recorded verification cases below cached artifacts: ${String(recordedManifestArtifactCount)}/${String(cachedArtifactCount)}`]),
    ...(staleRecordedVerificationCount === 0
      ? []
      : [`recorded verification cases need evidence refresh: ${String(staleRecordedVerificationCount)}`]),
    ...(recordedFailedCaseCount === 0 ? [] : [`recorded failed cases: ${String(recordedFailedCaseCount)}`]),
    ...(recordedErrorCaseCount === 0 ? [] : [`recorded error cases: ${String(recordedErrorCaseCount)}`]),
    ...scorecardHealthGaps(args.scorecard, cachedArtifactCount),
  ]
  return {
    targetWorkbookCount,
    sourceCount,
    cachedArtifactCount,
    scorecardCaseCount,
    checkpointCaseCount,
    recordedManifestArtifactCount,
    missingManifestArtifactCount,
    staleRecordedVerificationCount,
    recordedPassedCaseCount,
    recordedUnsupportedCaseCount,
    currentRecordedUnsupportedCaseCount,
    staleRecordedUnsupportedCaseCount,
    currentUnsupportedClassifications: buildUnsupportedClassificationCounts(currentRecordedUnsupportedCases),
    staleUnsupportedClassifications: buildUnsupportedClassificationCounts(staleRecordedUnsupportedCases),
    recordedFailedCaseCount,
    recordedErrorCaseCount,
    recordedCoversManifest,
    recordedAllCasesPassed,
    missingManifestArtifactSample,
    staleRecordedVerificationSample,
    nextMissingVerificationCommand: nextMissingVerificationRun.command,
    nextMissingVerificationPlanCommand,
    blockedMissingVerificationCommand: nextMissingVerificationRun.blockedCommand,
    nextStaleVerificationCommand: nextStaleVerificationRun.command,
    nextStaleVerificationPlanCommand,
    blockedStaleVerificationCommand: nextStaleVerificationRun.blockedCommand,
    scorecardCoversManifest,
    targetComplete,
    gaps,
  }
}

function buildUnsupportedClassificationCounts(
  cases: readonly PublicWorkbookCorpusCase[],
): PublicWorkbookCorpusUnsupportedClassificationCount[] {
  const counts = new Map<string, number>()
  for (const entry of cases) {
    for (const classification of entry.unsupportedFeatureClassifications) {
      counts.set(classification, (counts.get(classification) ?? 0) + 1)
    }
  }
  return [...counts.entries()]
    .map(([classification, count]) => ({ classification, count }))
    .toSorted((left, right) => right.count - left.count || left.classification.localeCompare(right.classification))
}

function scorecardHealthGaps(scorecard: PublicWorkbookCorpusScorecard | null, cachedArtifactCount: number): string[] {
  if (cachedArtifactCount === 0) {
    return []
  }
  if (!scorecard) {
    return ['scorecard is missing for cached workbooks']
  }
  return scorecard.summary.allCachedWorkbooksPassed ? [] : ['scorecard has non-passing cached workbooks']
}

function publicWorkbookCorpusVerificationBlockingGaps(status: PublicWorkbookCorpusStatus): string[] {
  return [
    ...(status.recordedManifestArtifactCount >= status.cachedArtifactCount
      ? []
      : [
          `recorded verification cases below cached artifacts: ${String(status.recordedManifestArtifactCount)}/${String(
            status.cachedArtifactCount,
          )}`,
        ]),
    ...(status.staleRecordedVerificationCount === 0
      ? []
      : [`recorded verification cases need evidence refresh: ${String(status.staleRecordedVerificationCount)}`]),
    ...(status.recordedFailedCaseCount === 0 ? [] : [`recorded failed cases: ${String(status.recordedFailedCaseCount)}`]),
    ...(status.recordedErrorCaseCount === 0 ? [] : [`recorded error cases: ${String(status.recordedErrorCaseCount)}`]),
    ...(status.recordedAllCasesPassed ? [] : ['recorded verification contains non-passing cases']),
  ]
}

interface PublicWorkbookCorpusCommandPaths {
  readonly manifestPath: string
  readonly scorecardPath: string
  readonly cacheDir: string
  readonly verifyCheckpointPath: string
  readonly displayRootDir?: string
  readonly stopMarkerActive?: boolean
}

function formatPublicWorkbookCorpusVerifySliceCommand(
  paths: PublicWorkbookCorpusCommandPaths,
  slice: 'verify-missing' | 'verify-stale',
  mode: 'plan' | 'verify',
): string {
  const script = mode === 'plan' ? `public-workbook-corpus:${slice}:plan` : `public-workbook-corpus:${slice}`
  const args =
    mode === 'plan'
      ? [
          '--manifest',
          commandPath(paths.manifestPath, paths.displayRootDir),
          '--scorecard',
          commandPath(paths.scorecardPath, paths.displayRootDir),
          '--verify-checkpoint',
          commandPath(paths.verifyCheckpointPath, paths.displayRootDir),
          '--cache-dir',
          commandPath(paths.cacheDir, paths.displayRootDir),
        ]
      : [
          '--manifest',
          commandPath(paths.manifestPath, paths.displayRootDir),
          '--scorecard',
          commandPath(paths.scorecardPath, paths.displayRootDir),
          '--verify-checkpoint',
          commandPath(paths.verifyCheckpointPath, paths.displayRootDir),
          '--cache-dir',
          commandPath(paths.cacheDir, paths.displayRootDir),
          '--limit',
          '1',
        ]
  return ['pnpm', script, '--', ...args].map(shellQuote).join(' ')
}

function splitPublicWorkbookCorpusVerifySliceCommand(
  paths: PublicWorkbookCorpusCommandPaths,
  slice: 'verify-missing' | 'verify-stale',
): { readonly command: string | null; readonly blockedCommand: string | null } {
  const command = formatPublicWorkbookCorpusVerifySliceCommand(paths, slice, 'verify')
  if (paths.stopMarkerActive === true) {
    return {
      command: null,
      blockedCommand: `${publicCorpusStopMarkerOverrideEnvVar}=1 ${command} ${publicCorpusStopMarkerOverrideFlag}`,
    }
  }
  return {
    command,
    blockedCommand: null,
  }
}

function formatPublicWorkbookCorpusResumePlanCheckCommand(paths: PublicWorkbookCorpusCommandPaths): string {
  return [
    'pnpm',
    'public-workbook-corpus:resume-plan:check',
    '--',
    '--manifest',
    commandPath(paths.manifestPath, paths.displayRootDir),
    '--cache-dir',
    commandPath(paths.cacheDir, paths.displayRootDir),
    '--scorecard',
    commandPath(paths.scorecardPath, paths.displayRootDir),
    '--verify-checkpoint',
    commandPath(paths.verifyCheckpointPath, paths.displayRootDir),
  ]
    .map(shellQuote)
    .join(' ')
}

function commandPath(path: string, displayRootDir: string | undefined): string {
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

function manifestRecordedCases(manifest: PublicWorkbookManifest, cases: readonly PublicWorkbookCorpusCase[]): PublicWorkbookCorpusCase[] {
  const casesById = new Map(cases.map((entry) => [entry.id, entry]))
  return manifest.artifacts.flatMap((artifact) => {
    const entry = casesById.get(artifact.id)
    return entry && caseMatchesArtifact(entry, artifact) ? [entry] : []
  })
}

function manifestMissingArtifactSample(
  manifest: PublicWorkbookManifest,
  cases: readonly PublicWorkbookCorpusCase[],
): MissingManifestArtifactSummary[] {
  const casesById = new Map(cases.map((entry) => [entry.id, entry]))
  return manifest.artifacts
    .flatMap((artifact) => {
      const entry = casesById.get(artifact.id)
      if (entry && caseMatchesArtifact(entry, artifact)) {
        return []
      }
      return [
        {
          id: artifact.id,
          fileName: artifact.fileName,
          byteSize: artifact.byteSize,
          sourceUrl: artifact.sourceUrl,
        },
      ]
    })
    .slice(0, missingManifestArtifactSampleLimit)
}

function scorecardMatchesManifest(scorecard: PublicWorkbookCorpusScorecard, manifest: PublicWorkbookManifest): boolean {
  return (
    scorecard.summary.targetWorkbookCount === manifest.targetWorkbookCount &&
    scorecard.summary.sourceCount === manifest.sources.length &&
    scorecard.summary.cachedWorkbookCount === manifest.artifacts.length &&
    scorecard.cases.length === manifest.artifacts.length &&
    manifest.artifacts.every((artifact, index) => {
      const entry = scorecard.cases[index]
      return entry ? caseMatchesArtifact(entry, artifact) : false
    })
  )
}

function caseMatchesArtifact(entry: PublicWorkbookCorpusCase, artifact: PublicWorkbookArtifact): boolean {
  return (
    entry.id === artifact.id &&
    entry.sourceId === artifact.sourceId &&
    entry.sourceUrl === artifact.sourceUrl &&
    entry.fileName === artifact.fileName &&
    entry.sha256 === artifact.sha256 &&
    entry.byteSize === artifact.byteSize
  )
}
