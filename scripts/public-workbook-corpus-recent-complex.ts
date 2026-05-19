#!/usr/bin/env bun

import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { pathToFileURL } from 'node:url'

import { runWorkPaperXlsxCorpusInChildProcesses } from './check-workpaper-xlsx-corpus.ts'
import type { WorkPaperXlsxCorpusResult } from './check-workpaper-xlsx-corpus-types.ts'
import { assertPublicCorpusRunNotStopped, readFlagArg, readNumberArg, readStringArg } from './public-workbook-corpus-cli.ts'
import { formatCommandPath } from './public-workbook-corpus-command-format.ts'
import { parsePublicWorkbookManifestJson } from './public-workbook-corpus-json.ts'
import type { PublicWorkbookArtifact, PublicWorkbookCorpusCase, PublicWorkbookCorpusScorecard } from './public-workbook-corpus-types.ts'
import { formatJsonForRepo } from './scorecard-format.ts'

export interface RecentComplexWorkbookCandidate {
  readonly artifact: PublicWorkbookArtifact
  readonly corpusCase: PublicWorkbookCorpusCase
  readonly cachePath: string
  readonly complexityScore: number
}

export interface PublicWorkbookCorpusRecentComplexSummary {
  readonly schemaVersion: 1
  readonly suite: 'public-workbook-corpus-recent-complex-headless'
  readonly generatedAt: string
  readonly targetWorkbookCount: number
  readonly manifestTargetWorkbookCount: number | null
  readonly manifestSourceCount: number
  readonly manifestArtifactCount: number
  readonly recommendedManifestTargetWorkbookCount: number
  readonly recommendedDiscoverySourceLimit: number
  readonly recommendedFetchArtifactLimit: number
  readonly publicScorecardCaseCount: number
  readonly recentArtifactCount: number
  readonly publicPassingRecentComplexCount: number
  readonly headlessFileCount: number
  readonly headlessOkFileCount: number
  readonly headlessComparableFormulaFileCount: number
  readonly endToEndPassingWorkbookCount: number
  readonly remainingToTarget: number
  readonly allSelectedHeadlessWorkbooksPassed: boolean
  readonly minFormulaCells: number
  readonly minComplexityScore: number
  readonly samplePassingArtifactIds: readonly string[]
  readonly sampleMissingHeadlessArtifactIds: readonly string[]
  readonly commands: {
    readonly retarget: string
    readonly discover: string
    readonly discoverHdx: string
    readonly discoverGithub: string
    readonly discoverZenodo: string
    readonly discoverFigshare: string
    readonly fetch: string
    readonly publicVerify: string
    readonly headlessVerify: string
    readonly check: string
  }
}

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const defaultCacheDir = join(rootDir, '.cache', 'public-workbook-corpus-recent-complex')
const defaultManifestPath = join(defaultCacheDir, 'manifest.json')
const defaultScorecardPath = join(defaultCacheDir, 'scorecard.json')
const defaultHeadlessScorecardPath = join(defaultCacheDir, 'headless-scorecard.json')
const defaultCorpusRunStopMarkerPath = join(rootDir, '.agent-coordination', '20260507T074946Z-codex-stop-interactive-corpus-runs.md')
const defaultTargetWorkbookCount = 500
const defaultMinFormulaCells = 1
const defaultMinComplexityScore = 5

async function main(): Promise<void> {
  const command = process.argv[2] ?? 'plan'
  const args = readRecentComplexArgs()
  if (command === 'headless') {
    assertPublicCorpusRunNotStopped({
      commandName: 'public-workbook-corpus recent-complex headless verification',
      stopMarkerPath: args.corpusRunStopMarkerPath,
    })
    const candidates = readRecentComplexCandidates(args)
    const selected = candidates.slice(0, args.targetWorkbookCount)
    if (selected.length === 0) {
      throw new Error('No recent complex public workbook candidates are available for headless verification')
    }
    const result = runWorkPaperXlsxCorpusInChildProcesses(
      selected.map((candidate) => candidate.cachePath),
      {
        childProcessTimeoutMs: args.childTimeoutMs,
        evaluationTimeoutMs: args.timeoutMs,
        maxFileBytes: args.maxFileBytes,
      },
    )
    mkdirSync(dirname(args.headlessScorecardPath), { recursive: true })
    writeFileSync(args.headlessScorecardPath, serializeJson(result, 'public-workbook-corpus-recent-complex-headless'))
    process.stdout.write(`${JSON.stringify(result.summary, null, 2)}\n`)
    if (result.summary.failedErrors > 0 || result.summary.failedTimeouts > 0 || result.summary.mismatchedFormulaCells > 0) {
      process.exitCode = 1
    }
    return
  }
  const summary = buildPublicWorkbookCorpusRecentComplexSummary(args)
  if (command === 'check') {
    const findings = validatePublicWorkbookCorpusRecentComplexSummary(summary, { requireTarget: readFlagArg('--require-target') })
    process.stdout.write(`${JSON.stringify(summary, null, 2)}\n`)
    if (findings.length > 0) {
      process.stderr.write(`Recent complex headless corpus check failed: ${findings.join('; ')}\n`)
      process.exitCode = 1
    }
    return
  }
  if (command === 'plan') {
    process.stdout.write(`${JSON.stringify(summary, null, 2)}\n`)
    return
  }
  throw new Error(`Unknown recent complex corpus command: ${command}`)
}

function readRecentComplexArgs(): RecentComplexArgs {
  const cacheDir = resolve(readStringArg('--cache-dir', defaultCacheDir))
  return {
    cacheDir,
    childTimeoutMs: readNumberArg('--child-timeout-ms', 31_000),
    corpusRunStopMarkerPath: resolve(readStringArg('--corpus-run-stop-marker', defaultCorpusRunStopMarkerPath)),
    generatedAt: readStringArg('--generated-at', new Date().toISOString()),
    headlessScorecardPath: resolve(readStringArg('--headless-scorecard', defaultHeadlessScorecardPath)),
    manifestPath: resolve(readStringArg('--manifest', defaultManifestPath)),
    maxFileBytes: readNumberArg('--max-file-bytes', 50 * 1024 * 1024),
    minComplexityScore: readNumberArg('--min-complexity-score', defaultMinComplexityScore),
    minFormulaCells: readNonNegativeIntegerArg('--min-formula-cells', defaultMinFormulaCells),
    scorecardPath: resolve(readStringArg('--scorecard', defaultScorecardPath)),
    targetWorkbookCount: readNumberArg('--target-workbook-count', defaultTargetWorkbookCount),
    timeoutMs: readNumberArg('--timeout-ms', 30_000),
  }
}

function readNonNegativeIntegerArg(name: string, fallback: number): number {
  const raw = readStringArg(name, String(fallback))
  const parsed = Number(raw)
  if (!/^\d+$/u.test(raw) || !Number.isSafeInteger(parsed)) {
    throw new Error(`Expected ${name} to be a non-negative integer`)
  }
  return parsed
}

interface RecentComplexArgs {
  readonly cacheDir: string
  readonly childTimeoutMs: number
  readonly corpusRunStopMarkerPath: string
  readonly generatedAt: string
  readonly headlessScorecardPath: string
  readonly manifestPath: string
  readonly maxFileBytes: number
  readonly minComplexityScore: number
  readonly minFormulaCells: number
  readonly scorecardPath: string
  readonly targetWorkbookCount: number
  readonly timeoutMs: number
}

export function selectRecentComplexWorkbookCandidates(args: {
  readonly cacheDir: string
  readonly manifestArtifacts: readonly PublicWorkbookArtifact[]
  readonly scorecard: PublicWorkbookCorpusScorecard
  readonly minFormulaCells?: number
  readonly minComplexityScore?: number
}): RecentComplexWorkbookCandidate[] {
  const minFormulaCells = args.minFormulaCells ?? defaultMinFormulaCells
  const minComplexityScore = args.minComplexityScore ?? defaultMinComplexityScore
  const casesByArtifactId = new Map(args.scorecard.cases.map((corpusCase) => [corpusCase.id, corpusCase]))
  return args.manifestArtifacts.flatMap((artifact) => {
    const corpusCase = casesByArtifactId.get(artifact.id)
    if (!corpusCase || corpusCase.status !== 'passed' || !corpusCase.passed || !hasRecentWorkbookEvidence(artifact)) {
      return []
    }
    const complexityScore = recentComplexityScore(corpusCase)
    if (
      corpusCase.featureCounts.formulaCellCount < minFormulaCells ||
      corpusCase.validation.formulaOracleComparisons < minFormulaCells ||
      complexityScore < minComplexityScore
    ) {
      return []
    }
    return [
      {
        artifact,
        corpusCase,
        cachePath: join(args.cacheDir, artifact.cachePath),
        complexityScore,
      },
    ]
  })
}

export function recentComplexityScore(corpusCase: PublicWorkbookCorpusCase): number {
  const featureCounts = corpusCase.featureCounts
  const formulaScore = Math.min(8, Math.floor(featureCounts.formulaCellCount / 10))
  const cellScore = Math.min(5, Math.floor(featureCounts.cellCount / 1_000))
  const sheetScore = Math.min(4, Math.max(0, featureCounts.sheetCount - 1))
  const modelFeatureScore =
    featureCounts.definedNameCount +
    featureCounts.tableCount * 2 +
    featureCounts.chartCount * 2 +
    featureCounts.pivotCount * 3 +
    featureCounts.conditionalFormatCount +
    featureCounts.dataValidationCount
  return formulaScore + cellScore + sheetScore + Math.min(10, modelFeatureScore)
}

export function hasRecentWorkbookEvidence(entry: { readonly topicEvidence?: readonly string[] }): boolean {
  return (entry.topicEvidence ?? []).some((evidence) => /^recent-202[56]:/u.test(evidence))
}

export function buildPublicWorkbookCorpusRecentComplexSummary(args: RecentComplexArgs): PublicWorkbookCorpusRecentComplexSummary {
  const manifest = readManifestIfExists(args.manifestPath)
  const scorecard = readScorecardIfExists(args.scorecardPath)
  const headless = readHeadlessScorecardIfExists(args.headlessScorecardPath)
  const artifacts = manifest?.artifacts ?? []
  const candidates =
    scorecard && manifest
      ? selectRecentComplexWorkbookCandidates({
          cacheDir: args.cacheDir,
          manifestArtifacts: artifacts,
          scorecard,
          minFormulaCells: args.minFormulaCells,
          minComplexityScore: args.minComplexityScore,
        })
      : []
  const selected = candidates.slice(0, args.targetWorkbookCount)
  const headlessByPath = new Map((headless?.files ?? []).map((file) => [resolve(file.path), file]))
  const selectedWithHeadless = selected.map((candidate) => ({ candidate, headless: headlessByPath.get(resolve(candidate.cachePath)) }))
  const passingHeadless = selectedWithHeadless.filter((entry) => isPassingHeadlessResult(entry.headless))
  return {
    schemaVersion: 1,
    suite: 'public-workbook-corpus-recent-complex-headless',
    generatedAt: args.generatedAt,
    targetWorkbookCount: args.targetWorkbookCount,
    manifestTargetWorkbookCount: manifest?.targetWorkbookCount ?? null,
    manifestSourceCount: manifest?.sources.length ?? 0,
    manifestArtifactCount: artifacts.length,
    publicScorecardCaseCount: scorecard?.cases.length ?? 0,
    recentArtifactCount: artifacts.filter(hasRecentWorkbookEvidence).length,
    publicPassingRecentComplexCount: candidates.length,
    headlessFileCount: headless?.summary.totalFiles ?? 0,
    headlessOkFileCount: headless?.summary.ok ?? 0,
    headlessComparableFormulaFileCount: (headless?.files ?? []).filter((file) => file.comparableFormulaCells > 0).length,
    endToEndPassingWorkbookCount: passingHeadless.length,
    remainingToTarget: Math.max(0, args.targetWorkbookCount - passingHeadless.length),
    allSelectedHeadlessWorkbooksPassed:
      selected.length > 0 &&
      selected.length === passingHeadless.length &&
      (headless?.summary.failedErrors ?? 0) === 0 &&
      (headless?.summary.failedTimeouts ?? 0) === 0 &&
      (headless?.summary.mismatchedFormulaCells ?? 0) === 0,
    minFormulaCells: args.minFormulaCells,
    minComplexityScore: args.minComplexityScore,
    samplePassingArtifactIds: passingHeadless.slice(0, 20).map((entry) => entry.candidate.artifact.id),
    sampleMissingHeadlessArtifactIds: selectedWithHeadless
      .filter((entry) => !isPassingHeadlessResult(entry.headless))
      .slice(0, 20)
      .map((entry) => entry.candidate.artifact.id),
    recommendedManifestTargetWorkbookCount: recommendedManifestTargetWorkbookCount(args, manifest, artifacts.length),
    recommendedDiscoverySourceLimit: recommendedDiscoverySourceLimit(args, manifest),
    recommendedFetchArtifactLimit: recommendedFetchArtifactLimit(args, manifest, artifacts.length),
    commands: recentComplexCommands(args, {
      minimumManifestTargetWorkbookCount: recommendedManifestTargetWorkbookCount(args, manifest, artifacts.length),
      minimumDiscoverySourceLimit: recommendedDiscoverySourceLimit(args, manifest),
    }),
  }
}

function isPassingHeadlessResult(file: WorkPaperXlsxCorpusResult['files'][number] | undefined): boolean {
  return file?.status === 'ok' && file.comparableFormulaCells > 0 && file.mismatchedFormulaCells === 0 && file.matchRate === 1
}

export function validatePublicWorkbookCorpusRecentComplexSummary(
  summary: PublicWorkbookCorpusRecentComplexSummary,
  options: { readonly requireTarget?: boolean } = {},
): string[] {
  const findings: string[] = []
  if (summary.schemaVersion !== 1 || summary.suite !== 'public-workbook-corpus-recent-complex-headless') {
    findings.push('unexpected recent complex headless summary header')
  }
  if (summary.endToEndPassingWorkbookCount > summary.publicPassingRecentComplexCount) {
    findings.push('end-to-end passing workbook count exceeds public passing recent complex count')
  }
  if (summary.manifestTargetWorkbookCount !== null && summary.manifestTargetWorkbookCount < summary.targetWorkbookCount) {
    findings.push('manifest target workbook count is below the recent complex target')
  }
  if (summary.remainingToTarget !== Math.max(0, summary.targetWorkbookCount - summary.endToEndPassingWorkbookCount)) {
    findings.push('remaining target count is stale')
  }
  if (
    summary.headlessFileCount > 0 &&
    summary.headlessFileCount < Math.min(summary.targetWorkbookCount, summary.publicPassingRecentComplexCount)
  ) {
    findings.push('headless verifier did not cover every selected recent complex candidate')
  }
  if (summary.headlessFileCount > 0 && !summary.allSelectedHeadlessWorkbooksPassed) {
    findings.push('one or more selected recent complex workbooks did not pass headless verification')
  }
  if (options.requireTarget === true && summary.endToEndPassingWorkbookCount < summary.targetWorkbookCount) {
    findings.push(
      `end-to-end recent complex headless target not met: ${String(summary.endToEndPassingWorkbookCount)}/${String(
        summary.targetWorkbookCount,
      )}`,
    )
  }
  return findings
}

function readRecentComplexCandidates(args: RecentComplexArgs): RecentComplexWorkbookCandidate[] {
  const manifest = parsePublicWorkbookManifestJson(JSON.parse(readFileSync(args.manifestPath, 'utf8')))
  const scorecard = readScorecard(args.scorecardPath)
  return selectRecentComplexWorkbookCandidates({
    cacheDir: args.cacheDir,
    manifestArtifacts: manifest.artifacts,
    scorecard,
    minFormulaCells: args.minFormulaCells,
    minComplexityScore: args.minComplexityScore,
  })
}

function readManifestIfExists(path: string): ReturnType<typeof parsePublicWorkbookManifestJson> | null {
  return existsSync(path) ? parsePublicWorkbookManifestJson(JSON.parse(readFileSync(path, 'utf8'))) : null
}

function readScorecardIfExists(path: string): PublicWorkbookCorpusScorecard | null {
  return existsSync(path) ? readScorecard(path) : null
}

function readScorecard(path: string): PublicWorkbookCorpusScorecard {
  const parsed: unknown = JSON.parse(readFileSync(path, 'utf8'))
  if (!isPublicWorkbookCorpusScorecard(parsed)) {
    throw new Error(`Public workbook corpus scorecard has an invalid shape: ${path}`)
  }
  return parsed
}

function readHeadlessScorecardIfExists(path: string): WorkPaperXlsxCorpusResult | null {
  if (!existsSync(path)) {
    return null
  }
  const parsed: unknown = JSON.parse(readFileSync(path, 'utf8'))
  if (!isWorkPaperXlsxCorpusResult(parsed)) {
    throw new Error(`Recent complex headless scorecard has an invalid shape: ${path}`)
  }
  return parsed
}

function isWorkPaperXlsxCorpusResult(value: unknown): value is WorkPaperXlsxCorpusResult {
  if (!isRecord(value) || !isRecord(value['summary']) || !Array.isArray(value['files']) || !Array.isArray(value['mismatches'])) {
    return false
  }
  const summary = value['summary']
  return (
    typeof summary['totalFiles'] === 'number' &&
    typeof summary['ok'] === 'number' &&
    typeof summary['failedErrors'] === 'number' &&
    typeof summary['failedTimeouts'] === 'number' &&
    typeof summary['mismatchedFormulaCells'] === 'number' &&
    value['files'].every(isWorkPaperXlsxCorpusFileResult)
  )
}

function isWorkPaperXlsxCorpusFileResult(value: unknown): value is WorkPaperXlsxCorpusResult['files'][number] {
  return (
    isRecord(value) &&
    typeof value['path'] === 'string' &&
    typeof value['status'] === 'string' &&
    typeof value['comparableFormulaCells'] === 'number' &&
    typeof value['mismatchedFormulaCells'] === 'number' &&
    typeof value['matchRate'] === 'number'
  )
}

function isPublicWorkbookCorpusScorecard(value: unknown): value is PublicWorkbookCorpusScorecard {
  return (
    isRecord(value) &&
    value['schemaVersion'] === 1 &&
    value['suite'] === 'public-workbook-corpus' &&
    isRecord(value['summary']) &&
    Array.isArray(value['cases']) &&
    value['cases'].every(isPublicWorkbookCorpusCase)
  )
}

function isPublicWorkbookCorpusCase(value: unknown): value is PublicWorkbookCorpusCase {
  return (
    isRecord(value) &&
    typeof value['id'] === 'string' &&
    typeof value['sourceId'] === 'string' &&
    typeof value['sourceUrl'] === 'string' &&
    typeof value['fileName'] === 'string' &&
    typeof value['sha256'] === 'string' &&
    typeof value['byteSize'] === 'number' &&
    (value['status'] === 'passed' || value['status'] === 'failed' || value['status'] === 'error' || value['status'] === 'unsupported') &&
    typeof value['passed'] === 'boolean' &&
    isPublicWorkbookFeatureCounts(value['featureCounts']) &&
    isRecord(value['validation']) &&
    Array.isArray(value['unsupportedFeatureClassifications']) &&
    Array.isArray(value['evidence'])
  )
}

function isPublicWorkbookFeatureCounts(value: unknown): value is PublicWorkbookCorpusCase['featureCounts'] {
  const keys: readonly (keyof PublicWorkbookCorpusCase['featureCounts'])[] = [
    'sheetCount',
    'cellCount',
    'formulaCellCount',
    'valueCellCount',
    'definedNameCount',
    'tableCount',
    'chartCount',
    'pivotCount',
    'mergeCount',
    'styleRangeCount',
    'conditionalFormatCount',
    'dataValidationCount',
    'macroPayloadCount',
    'warningCount',
  ]
  return isRecord(value) && keys.every((key) => typeof value[key] === 'number')
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function recentComplexCommands(
  args: RecentComplexArgs,
  options: { readonly minimumDiscoverySourceLimit?: number; readonly minimumManifestTargetWorkbookCount?: number } = {},
): PublicWorkbookCorpusRecentComplexSummary['commands'] {
  const retargetWorkbookCount = Math.max(args.targetWorkbookCount, options.minimumManifestTargetWorkbookCount ?? 0)
  const fetchArtifactLimit = Math.max(args.targetWorkbookCount, options.minimumManifestTargetWorkbookCount ?? 0)
  const discoverySourceLimit = Math.max(args.targetWorkbookCount * 5, options.minimumDiscoverySourceLimit ?? 0)
  const sharedArgs = [
    '--manifest',
    formatCommandPath(args.manifestPath),
    '--cache-dir',
    formatCommandPath(args.cacheDir),
    '--target-workbook-count',
    String(args.targetWorkbookCount),
  ]
  const checkArgs = [
    ...sharedArgs,
    '--scorecard',
    formatCommandPath(args.scorecardPath),
    '--headless-scorecard',
    formatCommandPath(args.headlessScorecardPath),
    '--min-formula-cells',
    String(args.minFormulaCells),
    '--min-complexity-score',
    String(args.minComplexityScore),
  ]
  return {
    retarget: formatShellCommand([
      'bun',
      'scripts/public-workbook-corpus.ts',
      'retarget',
      '--manifest',
      formatCommandPath(args.manifestPath),
      '--cache-dir',
      formatCommandPath(args.cacheDir),
      '--target-workbook-count',
      String(retargetWorkbookCount),
    ]),
    discover: formatShellCommand([
      'bun',
      'scripts/public-workbook-corpus.ts',
      'discover-recent-complex-ckan',
      ...sharedArgs,
      '--limit',
      String(discoverySourceLimit),
    ]),
    discoverHdx: formatShellCommand([
      'bun',
      'scripts/public-workbook-corpus.ts',
      'discover-recent-complex-ckan',
      ...sharedArgs,
      '--limit',
      String(discoverySourceLimit),
      '--portal',
      'https://data.humdata.org/api/3/action',
      '--query',
      '2025',
      '--query',
      '2026',
    ]),
    discoverGithub: formatShellCommand([
      'bun',
      'scripts/public-workbook-corpus.ts',
      'discover-recent-complex-github',
      ...sharedArgs,
      '--limit',
      String(discoverySourceLimit),
      '--skip-code-search',
      '--max-repository-pages-per-query',
      '3',
      '--max-repositories-per-query',
      '20',
    ]),
    discoverZenodo: formatShellCommand([
      'bun',
      'scripts/public-workbook-corpus.ts',
      'discover-recent-complex-zenodo',
      ...sharedArgs,
      '--limit',
      String(discoverySourceLimit),
      '--per-page',
      '100',
      '--max-pages-per-query',
      '20',
    ]),
    discoverFigshare: formatShellCommand([
      'bun',
      'scripts/public-workbook-corpus.ts',
      'discover-recent-complex-figshare',
      ...sharedArgs,
      '--limit',
      String(discoverySourceLimit),
      '--page-size',
      '100',
      '--max-pages-per-query',
      '20',
    ]),
    fetch: formatShellCommand([
      'bun',
      'scripts/public-workbook-corpus.ts',
      'fetch',
      ...sharedArgs,
      '--limit',
      String(fetchArtifactLimit),
      '--fetch-batch-size',
      '2',
      '--max-bytes',
      String(50 * 1024 * 1024),
    ]),
    publicVerify: formatShellCommand([
      'bun',
      'scripts/public-workbook-corpus.ts',
      'verify',
      ...sharedArgs,
      '--scorecard',
      formatCommandPath(args.scorecardPath),
      '--verify-checkpoint',
      formatCommandPath(join(args.cacheDir, 'verification-checkpoint.json')),
    ]),
    headlessVerify: formatShellCommand(['bun', 'scripts/public-workbook-corpus-recent-complex.ts', 'headless', ...checkArgs]),
    check: formatShellCommand(['bun', 'scripts/public-workbook-corpus-recent-complex.ts', 'check', ...checkArgs, '--require-target']),
  }
}

function recommendedManifestTargetWorkbookCount(
  args: RecentComplexArgs,
  manifest: ReturnType<typeof readManifestIfExists>,
  artifactCount: number,
): number {
  return Math.max(args.targetWorkbookCount, manifest?.targetWorkbookCount ?? 0, artifactCount)
}

function recommendedDiscoverySourceLimit(args: RecentComplexArgs, manifest: ReturnType<typeof readManifestIfExists>): number {
  const currentSourceCount = manifest?.sources.length ?? 0
  return Math.max(args.targetWorkbookCount * 5, currentSourceCount + args.targetWorkbookCount * 5)
}

function recommendedFetchArtifactLimit(
  args: RecentComplexArgs,
  manifest: ReturnType<typeof readManifestIfExists>,
  artifactCount: number,
): number {
  return recommendedManifestTargetWorkbookCount(args, manifest, artifactCount)
}

function formatShellCommand(parts: readonly string[]): string {
  return parts.map(shellQuote).join(' ')
}

function shellQuote(value: string): string {
  return /^[A-Za-z0-9_./:=@+-]+$/u.test(value) ? value : `'${value.replaceAll("'", "'\\''")}'`
}

function serializeJson(value: unknown, tempPrefix: string): string {
  return formatJsonForRepo({
    rootDir,
    serializedJson: `${JSON.stringify(value, null, 2)}\n`,
    tempPrefix,
  })
}

if (process.argv[1] && pathToFileURL(resolve(process.argv[1])).href === import.meta.url) {
  try {
    await main()
  } catch (error) {
    process.stderr.write(`${error instanceof Error ? error.message : String(error)}\n`)
    process.exit(1)
  }
}
