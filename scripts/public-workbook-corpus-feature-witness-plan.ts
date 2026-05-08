#!/usr/bin/env bun

import { existsSync, readFileSync } from 'node:fs'
import { isAbsolute, join, relative, resolve } from 'node:path'
import { pathToFileURL } from 'node:url'

import {
  publicCorpusStopMarkerOverrideEnvVar,
  publicCorpusStopMarkerOverrideFlag,
  readNumberArg,
  readStringArg,
} from './public-workbook-corpus-cli.ts'
import { buildFeatureWitnessCoverage } from './public-workbook-corpus-completion-audit-helpers.ts'
import { parsePublicWorkbookManifestJson } from './public-workbook-corpus-json.ts'
import { publicWorkbookCorpusCaseMatchesArtifact } from './public-workbook-corpus-missing.ts'
import { readReusablePublicWorkbookCorpusCases } from './public-workbook-corpus-verify-checkpoint.ts'
import type { PublicWorkbookArtifact, PublicWorkbookCorpusCase, PublicWorkbookManifest } from './public-workbook-corpus-types.ts'

export interface PublicWorkbookCorpusFeatureWitnessPlan {
  readonly schemaVersion: 1
  readonly mode: 'feature-witness-plan'
  readonly generatedAt: string
  readonly stopMarker: {
    readonly active: boolean
    readonly path: string
    readonly overrideFlag: string
    readonly overrideEnvVar: string
  }
  readonly recordedCaseCount: number
  readonly coverage: readonly {
    readonly id: string
    readonly label: string
    readonly totalCount: number
    readonly witnessCaseCount: number
    readonly needsWitness: boolean
    readonly discoveryQuery: string
    readonly cachedCandidateCount: number
    readonly cachedCandidates: readonly PublicWorkbookCorpusFeatureWitnessCandidate[]
    readonly commands: {
      readonly discover: string | null
    }
    readonly blockedCommands: {
      readonly discover: string | null
    }
  }[]
  readonly missingWitnessCount: number
  readonly missingWitnesses: readonly {
    readonly id: string
    readonly label: string
    readonly discoveryQuery: string
    readonly discoverCommand: string | null
    readonly blockedDiscoverCommand: string | null
    readonly cachedCandidateCount: number
    readonly cachedCandidates: readonly PublicWorkbookCorpusFeatureWitnessCandidate[]
  }[]
}

export interface PublicWorkbookCorpusFeatureWitnessCandidate {
  readonly artifactId: string
  readonly fileName: string
  readonly byteSize: number
  readonly sourceUrl: string
  readonly verifyArtifactCommand: string | null
  readonly blockedVerifyArtifactCommand: string | null
}

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const defaultCacheDir = join(rootDir, '.cache', 'public-workbook-corpus')
const defaultManifestPath = join(defaultCacheDir, 'manifest.json')
const defaultScorecardPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'public-workbook-corpus-scorecard.json')
const defaultVerifyCheckpointPath = join(defaultCacheDir, 'verification-checkpoint.json')
const defaultCorpusRunStopMarkerPath = join(rootDir, '.agent-coordination', '20260507T074946Z-codex-stop-interactive-corpus-runs.md')

const featureDiscoveryQueries: Readonly<Record<string, string>> = {
  'conditional formats': 'conditional formatting xlsx',
  charts: 'chart xlsx',
  formulas: 'formula xlsx',
  'merged ranges': 'merged cells xlsx',
  names: 'defined names xlsx',
  pivots: 'pivot table xlsx',
  styles: 'cell styles xlsx',
  tables: 'table xlsx',
  values: 'xlsx',
}

const featureCandidateNeedles: Readonly<Record<string, readonly string[]>> = {
  'conditional formats': ['conditional format', 'conditional-format', 'conditional_format'],
  charts: ['chart'],
  formulas: ['formula'],
  'merged ranges': ['merged cell', 'merged-cell', 'merged_cell', 'merge'],
  names: ['defined name', 'defined-name', 'defined_name', 'named range', 'named-range', 'named_range'],
  pivots: ['pivot'],
  styles: ['style'],
  tables: ['table'],
  values: ['xlsx'],
}

function main(): void {
  const cacheDir = resolve(readStringArg('--cache-dir', defaultCacheDir))
  const manifestPath = resolve(readStringArg('--manifest', defaultManifestPath))
  const scorecardPath = resolve(readStringArg('--scorecard', defaultScorecardPath))
  const verifyCheckpointPath = resolve(readStringArg('--verify-checkpoint', defaultVerifyCheckpointPath))
  const stopMarkerPath = resolve(readStringArg('--corpus-run-stop-marker', defaultCorpusRunStopMarkerPath))
  const discoveryLimit = readNumberArg('--discovery-limit', 10_000)
  const generatedAt = readStringArg('--generated-at', new Date().toISOString())
  const plan = buildPublicWorkbookCorpusFeatureWitnessPlan({
    artifacts: readPublicWorkbookCorpusFeatureWitnessArtifacts(manifestPath),
    cacheDir,
    cases: readPublicWorkbookCorpusFeatureWitnessCases({ manifestPath, scorecardPath, verifyCheckpointPath }),
    discoveryLimit,
    displayRootDir: rootDir,
    generatedAt,
    manifestPath,
    stopMarkerActive: existsSync(stopMarkerPath),
    stopMarkerPath,
    verifyCheckpointPath,
  })
  if (process.argv.includes('--check')) {
    const findings = validatePublicWorkbookCorpusFeatureWitnessPlan(plan)
    if (findings.length > 0) {
      throw new Error(`Public workbook corpus feature witness plan is invalid: ${findings.join('; ')}`)
    }
    process.stdout.write(
      `${JSON.stringify(
        {
          mode: 'check',
          schemaVersion: plan.schemaVersion,
          generatedAt: plan.generatedAt,
          missingWitnessCount: plan.missingWitnessCount,
          missingWitnesses: plan.missingWitnesses,
          recordedCaseCount: plan.recordedCaseCount,
        },
        null,
        2,
      )}\n`,
    )
    return
  }
  process.stdout.write(`${JSON.stringify(plan, null, 2)}\n`)
}

export function readPublicWorkbookCorpusFeatureWitnessCases(args: {
  readonly manifestPath: string
  readonly scorecardPath: string
  readonly verifyCheckpointPath: string
}): PublicWorkbookCorpusCase[] {
  const reusableCases = readReusablePublicWorkbookCorpusCases([args.scorecardPath, args.verifyCheckpointPath])
  const reusableCasesById = new Map(reusableCases.map((entry) => [entry.id, entry]))
  if (!existsSync(args.manifestPath)) {
    return [...reusableCasesById.values()]
  }
  const manifest = parsePublicWorkbookManifestJson(JSON.parse(readFileSync(args.manifestPath, 'utf8')) as unknown)
  return featureWitnessCasesForManifest(manifest, reusableCasesById)
}

export function readPublicWorkbookCorpusFeatureWitnessArtifacts(manifestPath: string): PublicWorkbookArtifact[] {
  if (!existsSync(manifestPath)) {
    return []
  }
  const manifest = parsePublicWorkbookManifestJson(JSON.parse(readFileSync(manifestPath, 'utf8')) as unknown)
  return [...manifest.artifacts]
}

function featureWitnessCasesForManifest(
  manifest: PublicWorkbookManifest,
  casesById: ReadonlyMap<string, PublicWorkbookCorpusCase>,
): PublicWorkbookCorpusCase[] {
  return manifest.artifacts.flatMap((artifact) => {
    const candidate = casesById.get(artifact.id)
    return candidate && publicWorkbookCorpusCaseMatchesArtifact(candidate, artifact) ? [candidate] : []
  })
}

export function buildPublicWorkbookCorpusFeatureWitnessPlan(args: {
  readonly artifacts?: readonly PublicWorkbookArtifact[]
  readonly cacheDir: string
  readonly cases: readonly PublicWorkbookCorpusCase[]
  readonly discoveryLimit: number
  readonly displayRootDir?: string
  readonly generatedAt: string
  readonly manifestPath: string
  readonly stopMarkerActive: boolean
  readonly stopMarkerPath: string
  readonly verifyCheckpointPath?: string
}): PublicWorkbookCorpusFeatureWitnessPlan {
  const coverage = buildFeatureWitnessCoverage(args.cases).map((entry) => {
    const discoveryQuery = featureDiscoveryQueries[entry.id] ?? `${entry.label} xlsx`
    const candidateArtifacts = featureCandidateArtifacts(entry.id, args.artifacts ?? [])
    const discoverCommand = splitGuardedCommand(args.stopMarkerActive, [
      'pnpm',
      'public-workbook-corpus:discover',
      '--',
      '--manifest',
      commandPath(args.manifestPath, args.displayRootDir),
      '--cache-dir',
      commandPath(args.cacheDir, args.displayRootDir),
      '--query',
      discoveryQuery,
      '--limit',
      String(Math.max(0, Math.trunc(args.discoveryLimit))),
    ])
    const cachedCandidates = candidateArtifacts.slice(0, 5).map((artifact) =>
      featureWitnessCandidate({
        artifact,
        cacheDir: args.cacheDir,
        displayRootDir: args.displayRootDir,
        manifestPath: args.manifestPath,
        stopMarkerActive: args.stopMarkerActive,
        verifyCheckpointPath: args.verifyCheckpointPath ?? defaultVerifyCheckpointPath,
      }),
    )
    return {
      id: entry.id,
      label: entry.label,
      totalCount: entry.totalCount,
      witnessCaseCount: entry.witnessCaseCount,
      needsWitness: entry.witnessCaseCount === 0,
      discoveryQuery,
      cachedCandidateCount: candidateArtifacts.length,
      cachedCandidates,
      commands: {
        discover: discoverCommand.command,
      },
      blockedCommands: {
        discover: discoverCommand.blockedCommand,
      },
    }
  })
  const missingWitnesses = coverage
    .filter((entry) => entry.needsWitness)
    .map((entry) => ({
      id: entry.id,
      label: entry.label,
      discoveryQuery: entry.discoveryQuery,
      discoverCommand: entry.commands.discover,
      blockedDiscoverCommand: entry.blockedCommands.discover,
      cachedCandidateCount: entry.cachedCandidateCount,
      cachedCandidates: entry.cachedCandidates,
    }))
  return {
    schemaVersion: 1,
    mode: 'feature-witness-plan',
    generatedAt: args.generatedAt,
    stopMarker: {
      active: args.stopMarkerActive,
      path: commandPath(args.stopMarkerPath, args.displayRootDir),
      overrideFlag: publicCorpusStopMarkerOverrideFlag,
      overrideEnvVar: publicCorpusStopMarkerOverrideEnvVar,
    },
    recordedCaseCount: args.cases.length,
    coverage,
    missingWitnessCount: missingWitnesses.length,
    missingWitnesses,
  }
}

export function validatePublicWorkbookCorpusFeatureWitnessPlan(plan: PublicWorkbookCorpusFeatureWitnessPlan): string[] {
  const findings: string[] = []
  if (plan.schemaVersion !== 1) {
    findings.push(`unexpected schema version: ${String(plan.schemaVersion)}`)
  }
  if (!plan.generatedAt.trim()) {
    findings.push('generatedAt is empty')
  }
  if (plan.stopMarker.active && plan.stopMarker.overrideFlag !== publicCorpusStopMarkerOverrideFlag) {
    findings.push('stop-marker override flag does not match corpus CLI guard')
  }
  if (plan.stopMarker.active && plan.stopMarker.overrideEnvVar !== publicCorpusStopMarkerOverrideEnvVar) {
    findings.push('stop-marker override environment variable does not match corpus CLI guard')
  }
  if (plan.missingWitnessCount !== plan.coverage.filter((entry) => entry.needsWitness).length) {
    findings.push('missing witness count does not match coverage rows')
  }
  if (plan.missingWitnessCount !== plan.missingWitnesses.length) {
    findings.push('missing witness count does not match missing witness summaries')
  }
  const missingCoverageIds = plan.coverage.filter((entry) => entry.needsWitness).map((entry) => entry.id)
  const missingSummaryIds = plan.missingWitnesses.map((entry) => entry.id)
  if (JSON.stringify(missingCoverageIds) !== JSON.stringify(missingSummaryIds)) {
    findings.push('missing witness summaries do not match missing coverage rows')
  }
  for (const entry of plan.coverage) {
    if (!entry.id.trim() || !entry.label.trim() || !entry.discoveryQuery.trim()) {
      findings.push(`feature witness row has empty identity or query: ${entry.id}`)
    }
    if (entry.needsWitness !== (entry.witnessCaseCount === 0)) {
      findings.push(`feature witness row has inconsistent needsWitness flag: ${entry.id}`)
    }
    if (entry.cachedCandidateCount < entry.cachedCandidates.length) {
      findings.push(`feature witness candidate count is below listed candidates: ${entry.id}`)
    }
    if (entry.commands.discover !== null && !entry.commands.discover.includes('public-workbook-corpus:discover')) {
      findings.push(`feature witness discover command is missing discover script: ${entry.id}`)
    }
    if (plan.stopMarker.active && entry.commands.discover !== null) {
      findings.push(`feature witness discover command is runnable while stop marker is active: ${entry.id}`)
    }
    if (!plan.stopMarker.active && entry.commands.discover === null) {
      findings.push(`feature witness discover command is missing while stop marker is inactive: ${entry.id}`)
    }
    if (plan.stopMarker.active) {
      if (!entry.blockedCommands.discover?.includes(publicCorpusStopMarkerOverrideFlag)) {
        findings.push(`feature witness blocked discover command is missing stop-marker override: ${entry.id}`)
      }
    } else if (entry.blockedCommands.discover !== null) {
      findings.push(`feature witness blocked discover command is present while stop marker is inactive: ${entry.id}`)
    }
    for (const candidate of entry.cachedCandidates) {
      if (!candidate.artifactId.trim() || !candidate.fileName.trim() || !candidate.sourceUrl.trim()) {
        findings.push(`feature witness cached candidate has empty identity: ${entry.id}`)
      }
      if (candidate.verifyArtifactCommand !== null && !candidate.verifyArtifactCommand.includes('public-workbook-corpus:verify-artifact')) {
        findings.push(`feature witness cached candidate verify command is missing verify-artifact script: ${entry.id}`)
      }
      if (candidate.verifyArtifactCommand !== null && !candidate.verifyArtifactCommand.includes('--update-verify-checkpoint')) {
        findings.push(`feature witness cached candidate verify command is missing checkpoint update flag: ${entry.id}`)
      }
      if (plan.stopMarker.active && candidate.verifyArtifactCommand !== null) {
        findings.push(`feature witness cached candidate verify command is runnable while stop marker is active: ${entry.id}`)
      }
      if (!plan.stopMarker.active && candidate.verifyArtifactCommand === null) {
        findings.push(`feature witness cached candidate verify command is missing while stop marker is inactive: ${entry.id}`)
      }
      if (plan.stopMarker.active) {
        if (!candidate.blockedVerifyArtifactCommand?.includes(publicCorpusStopMarkerOverrideFlag)) {
          findings.push(`feature witness cached candidate blocked verify command is missing stop-marker override: ${entry.id}`)
        }
      } else if (candidate.blockedVerifyArtifactCommand !== null) {
        findings.push(`feature witness cached candidate blocked verify command is present while stop marker is inactive: ${entry.id}`)
      }
    }
  }
  for (const entry of plan.missingWitnesses) {
    if (!entry.id.trim() || !entry.label.trim() || !entry.discoveryQuery.trim()) {
      findings.push(`missing feature witness summary has empty fields: ${entry.id}`)
    }
    if (plan.stopMarker.active && entry.discoverCommand !== null) {
      findings.push(`missing feature witness summary discover command is runnable while stop marker is active: ${entry.id}`)
    }
    if (!plan.stopMarker.active && entry.discoverCommand === null) {
      findings.push(`missing feature witness summary discover command is missing while stop marker is inactive: ${entry.id}`)
    }
    if (plan.stopMarker.active) {
      if (!entry.blockedDiscoverCommand?.includes(publicCorpusStopMarkerOverrideFlag)) {
        findings.push(`missing feature witness summary blocked discover command is missing stop-marker override: ${entry.id}`)
      }
    } else if (entry.blockedDiscoverCommand !== null) {
      findings.push(`missing feature witness summary blocked discover command is present while stop marker is inactive: ${entry.id}`)
    }
  }
  return findings
}

function featureCandidateArtifacts(featureId: string, artifacts: readonly PublicWorkbookArtifact[]): PublicWorkbookArtifact[] {
  const needles = featureCandidateNeedles[featureId] ?? []
  if (needles.length === 0) {
    return []
  }
  return artifacts
    .filter((artifact) => {
      const haystack = [artifact.id, artifact.fileName, artifact.sourceUrl, artifact.downloadUrl, ...(artifact.topicEvidence ?? [])]
        .join(' ')
        .toLocaleLowerCase('en-US')
      return needles.some((needle) => haystack.includes(needle))
    })
    .toSorted(
      (left, right) => left.byteSize - right.byteSize || left.fileName.localeCompare(right.fileName) || left.id.localeCompare(right.id),
    )
}

function featureWitnessCandidate(args: {
  readonly artifact: PublicWorkbookArtifact
  readonly cacheDir: string
  readonly displayRootDir?: string
  readonly manifestPath: string
  readonly stopMarkerActive: boolean
  readonly verifyCheckpointPath: string
}): PublicWorkbookCorpusFeatureWitnessCandidate {
  const verifyArtifactCommand = splitGuardedCommand(args.stopMarkerActive, [
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
    args.artifact.id,
    '--update-verify-checkpoint',
  ])
  return {
    artifactId: args.artifact.id,
    fileName: args.artifact.fileName,
    byteSize: args.artifact.byteSize,
    sourceUrl: args.artifact.sourceUrl,
    verifyArtifactCommand: verifyArtifactCommand.command,
    blockedVerifyArtifactCommand: verifyArtifactCommand.blockedCommand,
  }
}

function splitGuardedCommand(
  stopMarkerActive: boolean,
  parts: readonly string[],
): { readonly command: string | null; readonly blockedCommand: string | null } {
  if (!stopMarkerActive) {
    return {
      command: command(parts),
      blockedCommand: null,
    }
  }
  return {
    command: null,
    blockedCommand: guardedCommand(true, parts),
  }
}

function guardedCommand(stopMarkerActive: boolean, parts: readonly string[]): string {
  if (!stopMarkerActive) {
    return command(parts)
  }
  return `${publicCorpusStopMarkerOverrideEnvVar}=1 ${command([...parts, publicCorpusStopMarkerOverrideFlag])}`
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

function command(parts: readonly string[]): string {
  return parts.map(shellQuote).join(' ')
}

function shellQuote(value: string): string {
  return /^[A-Za-z0-9_./:=@+-]+$/u.test(value) ? value : `'${value.replaceAll("'", "'\\''")}'`
}

if (process.argv[1] && pathToFileURL(resolve(process.argv[1])).href === import.meta.url) {
  main()
}
