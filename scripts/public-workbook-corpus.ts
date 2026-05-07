#!/usr/bin/env bun

import { spawn } from 'node:child_process'
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'

import { SpreadsheetEngine } from '../packages/core/src/engine.js'
import { exportXlsx, importXlsx } from '../packages/excel-import/src/index.js'
import type { WorkbookSnapshot } from '../packages/protocol/src/types.js'
import {
  createEmptyPublicWorkbookManifest,
  parsePublicWorkbookCorpusCase,
  parsePublicWorkbookCorpusScorecardJson,
  parsePublicWorkbookManifestJson,
  validatePublicWorkbookManifest,
} from './public-workbook-corpus-json.ts'
import { defaultCkanPortalBases, discoverCkanWorkbookSources, discoverFinancialCkanQueries } from './public-workbook-corpus-discovery.ts'
import {
  defaultDownloadTimeoutMs,
  defaultFingerprintMaxRssBytes,
  defaultFingerprintTimeoutMs,
  fetchPublicWorkbookArtifacts,
  fingerprintWorkbookFileIsolated,
} from './public-workbook-corpus-fetch.ts'
import { withPublicWorkbookCorpusCacheLock } from './public-workbook-corpus-lock.ts'
import { formatByteSize, startChildRssWatchdog, terminateChildProcess } from './public-workbook-corpus-process.ts'
import { roundTripSemanticsDigest } from './public-workbook-corpus-roundtrip.ts'
import { buildPublicWorkbookCorpusScorecardFromCases, validatePublicWorkbookCorpusScorecard } from './public-workbook-corpus-scorecard.ts'
import { defaultFinancialWorkbookQueries } from './public-workbook-corpus-topics.ts'
import {
  indexReusablePublicWorkbookCorpusCases,
  readReusablePublicWorkbookCorpusCases,
  writePublicWorkbookCorpusVerificationCheckpoint,
} from './public-workbook-corpus-verify-checkpoint.ts'
import {
  cellValuesMatchOracle,
  countWorkbookFeatures,
  emptyFeatureCounts,
  extractFormulaOracles,
  fingerprintWorkbookBytes,
  formatCellValue,
  inspectWorkbookFootprint,
  sha256HexSync,
  workbookMetadata,
} from './public-workbook-corpus-workbook.ts'
import type {
  BuildScorecardArgs,
  FormulaOracle,
  FormulaOracleValidationResult,
  PublicWorkbookArtifact,
  PublicWorkbookCaseStatus,
  PublicWorkbookCorpusCase,
  PublicWorkbookCorpusScorecard,
  PublicWorkbookFeatureCounts,
  PublicWorkbookManifest,
  PublicWorkbookValidationSummary,
} from './public-workbook-corpus-types.ts'
import { formatJsonForRepo } from './scorecard-format.ts'

declare const Bun:
  | {
      gc(force?: boolean): void
    }
  | undefined

export {
  createEmptyPublicWorkbookManifest,
  discoverCkanWorkbookSources,
  parsePublicWorkbookCorpusScorecardJson,
  parsePublicWorkbookManifestJson,
  validatePublicWorkbookCorpusScorecard,
  validatePublicWorkbookManifest,
  fetchPublicWorkbookArtifacts,
}
export type {
  PublicWorkbookArtifact,
  PublicWorkbookCaseStatus,
  PublicWorkbookCorpusCase,
  PublicWorkbookCorpusScorecard,
  PublicWorkbookFeatureCounts,
  PublicWorkbookLicenseEvidence,
  PublicWorkbookManifest,
  PublicWorkbookSource,
  PublicWorkbookSourceKind,
  PublicWorkbookValidationSummary,
} from './public-workbook-corpus-types.ts'

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const publicWorkbookCorpusScriptPath = fileURLToPath(import.meta.url)
const defaultCacheDir = join(rootDir, '.cache', 'public-workbook-corpus')
const defaultManifestPath = join(defaultCacheDir, 'manifest.json')
const defaultScorecardPath = join(rootDir, 'packages', 'benchmarks', 'baselines', 'public-workbook-corpus-scorecard.json')
const defaultVerifyTimeoutMs = 180_000
const defaultVerifyMaxRssBytes = 4096 * 1024 * 1024
const maxVerifyMaxRssBytes = 4096 * 1024 * 1024
const defaultVerifyMaxCellCount = 1_500_000
const defaultSelfRssCheckIntervalMs = 500
const noop = (): void => undefined

export async function sha256Hex(bytes: Uint8Array): Promise<string> {
  return sha256HexSync(bytes)
}

export async function buildPublicWorkbookCorpusScorecard(args: BuildScorecardArgs): Promise<PublicWorkbookCorpusScorecard> {
  validatePublicWorkbookManifest(args.manifest)
  const structuralSmokeSampleLimit = args.structuralSmokeSampleLimit ?? 50
  const verifyConcurrency = Math.max(1, Math.trunc(args.verifyConcurrency ?? 2))
  const verificationManifestPath = args.manifestPath
  const isolatedVerification = args.isolatedVerification === true && verificationManifestPath !== undefined
  const verifyTimeoutMs = Math.max(1, Math.trunc(args.verifyTimeoutMs ?? defaultVerifyTimeoutMs))
  const verifyMaxRssBytes = capVerifyMaxRssBytes(Math.max(1, Math.trunc(args.verifyMaxRssBytes ?? defaultVerifyMaxRssBytes)))
  const verifyMaxCellCount = Math.max(1, Math.trunc(args.verifyMaxCellCount ?? defaultVerifyMaxCellCount))
  const verifyRssCheckIntervalMs = Math.max(100, Math.trunc(args.verifyRssCheckIntervalMs ?? 250))
  const reusableCasesById = indexReusablePublicWorkbookCorpusCases({
    manifest: args.manifest,
    cases: args.reusableCases ?? [],
    structuralSmokeSampleLimit,
  })
  let completedCount = 0
  const reportVerifiedCase = (verifiedCase: PublicWorkbookCorpusCase): PublicWorkbookCorpusCase => {
    completedCount += 1
    args.onCaseVerified?.({
      completedCount,
      totalCount: args.manifest.artifacts.length,
      latestCase: verifiedCase,
    })
    return verifiedCase
  }
  const cases = await mapWithConcurrency(args.manifest.artifacts, verifyConcurrency, (artifact, index) => {
    const reusableCase = reusableCasesById.get(artifact.id)
    if (reusableCase) {
      return Promise.resolve(reportVerifiedCase(reusableCase))
    }
    const runStructuralSmoke = index < structuralSmokeSampleLimit
    const verifiedCasePromise =
      isolatedVerification && verificationManifestPath
        ? verifyCachedWorkbookArtifactIsolated({
            artifact,
            cacheDir: args.cacheDir,
            manifestPath: verificationManifestPath,
            runStructuralSmoke,
            timeoutMs: verifyTimeoutMs,
            maxRssBytes: verifyMaxRssBytes,
            maxCellCount: verifyMaxCellCount,
            rssCheckIntervalMs: verifyRssCheckIntervalMs,
          })
        : verifyCachedWorkbookArtifact(artifact, args.cacheDir, runStructuralSmoke, verifyMaxCellCount)
    return verifiedCasePromise.then(reportVerifiedCase)
  })
  return buildPublicWorkbookCorpusScorecardFromCases({
    manifest: args.manifest,
    generatedAt: args.generatedAt,
    cases,
  })
}

function verifyCachedWorkbookArtifactIsolated(args: {
  readonly artifact: PublicWorkbookArtifact
  readonly cacheDir: string
  readonly manifestPath: string
  readonly runStructuralSmoke: boolean
  readonly timeoutMs: number
  readonly maxRssBytes: number
  readonly maxCellCount: number
  readonly rssCheckIntervalMs?: number
}): Promise<PublicWorkbookCorpusCase> {
  const baseEvidence = artifactBaseEvidence(args.artifact)
  return new Promise<PublicWorkbookCorpusCase>((resolvePromise) => {
    const childArgs = [
      publicWorkbookCorpusScriptPath,
      'verify-artifact-worker',
      '--manifest',
      args.manifestPath,
      '--cache-dir',
      args.cacheDir,
      '--artifact-id',
      args.artifact.id,
      '--verify-max-rss-mb',
      String(Math.ceil(args.maxRssBytes / 1024 / 1024)),
      '--verify-max-cells',
      String(args.maxCellCount),
      ...(args.runStructuralSmoke ? ['--structural-smoke'] : []),
    ]
    const child = spawn(process.execPath, childArgs, {
      stdio: ['ignore', 'pipe', 'pipe'],
    })
    let stdout = ''
    let stderr = ''
    let timer: ReturnType<typeof setTimeout>
    let stopRssWatchdog = noop
    const finish = createOneShotResolver(resolvePromise, () => {
      clearTimeout(timer)
      stopRssWatchdog()
    })
    const terminateChild = (signal: 'SIGTERM' | 'SIGKILL'): void => {
      terminateChildProcess(child, signal)
    }
    stopRssWatchdog = startChildRssWatchdog(child, {
      maxRssBytes: args.maxRssBytes,
      intervalMs: args.rssCheckIntervalMs,
      onLimitExceeded: (rssBytes) => {
        terminateChild('SIGTERM')
        const forceKillTimer = setTimeout(() => terminateChild('SIGKILL'), 5_000)
        forceKillTimer.unref()
        finish(
          failedCase(args.artifact, 'error', baseEvidence, [
            `Verification subprocess exceeded RSS limit: ${formatByteSize(rssBytes)} > ${formatByteSize(args.maxRssBytes)}`,
            'The workbook was isolated in a subprocess so the corpus verification run could continue.',
          ]),
        )
      },
    })
    timer = setTimeout(() => {
      terminateChild('SIGTERM')
      const forceKillTimer = setTimeout(() => terminateChild('SIGKILL'), 5_000)
      forceKillTimer.unref()
      finish(
        failedCase(args.artifact, 'error', baseEvidence, [
          `Verification timed out after ${String(args.timeoutMs)}ms`,
          'The workbook was isolated in a subprocess so the corpus verification run could continue.',
        ]),
      )
    }, args.timeoutMs)
    child.stdout.setEncoding('utf8')
    child.stderr.setEncoding('utf8')
    child.stdout.on('data', (chunk: string) => {
      stdout += chunk
    })
    child.stderr.on('data', (chunk: string) => {
      stderr += chunk
    })
    child.on('error', (error) => {
      finish(failedCase(args.artifact, 'error', baseEvidence, [`Verification subprocess failed to start: ${error.message}`]))
    })
    child.on('close', (code, signal) => {
      if (code !== 0) {
        const failureDetails = compactProcessOutput(stderr || stdout)
        finish(
          failedCase(args.artifact, 'error', baseEvidence, [
            `Verification subprocess exited with ${signal ? `signal ${signal}` : `code ${String(code)}`}`,
            ...(failureDetails ? [failureDetails] : []),
          ]),
        )
        return
      }
      try {
        const parsed: unknown = JSON.parse(stdout)
        finish(parsePublicWorkbookCorpusCase(parsed))
      } catch (error) {
        const details = compactProcessOutput(stderr || stdout)
        finish(
          failedCase(args.artifact, 'error', baseEvidence, [
            `Verification subprocess returned invalid JSON: ${error instanceof Error ? error.message : String(error)}`,
            ...(details ? [details] : []),
          ]),
        )
      }
    })
  })
}

function createOneShotResolver<T>(resolveValue: (value: T) => void, cleanup: () => void): (value: T) => void {
  let settled = false
  return (value) => {
    if (settled) {
      return
    }
    settled = true
    cleanup()
    resolveValue(value)
  }
}

async function verifyCachedWorkbookArtifact(
  artifact: PublicWorkbookArtifact,
  cacheDir: string,
  runStructuralSmoke: boolean,
  maxCellCount: number,
): Promise<PublicWorkbookCorpusCase> {
  const cachePath = join(cacheDir, artifact.cachePath)
  const baseEvidence = artifactBaseEvidence(artifact)
  if (!existsSync(cachePath)) {
    return failedCase(artifact, 'error', baseEvidence, [`Missing cached workbook file: ${artifact.cachePath}`])
  }
  try {
    const bytes = readFileSync(cachePath)
    const actualHash = sha256HexSync(bytes)
    if (actualHash !== artifact.sha256) {
      return failedCase(artifact, 'failed', baseEvidence, [
        `Cached workbook hash mismatch: expected ${artifact.sha256}, received ${actualHash}`,
      ])
    }
    const footprint = inspectWorkbookFootprint(bytes, artifact.fileName)
    if (footprint.featureCounts.cellCount > maxCellCount) {
      return unsupportedResourceLimitCase(artifact, baseEvidence, footprint, maxCellCount)
    }
    collectGarbage()
    const imported = importXlsx(bytes, artifact.fileName)
    const featureCounts = countWorkbookFeatures(imported.snapshot, imported.warnings)
    const metadata = workbookMetadata(imported.snapshot)
    const formulaOracleValidation =
      footprint.featureCounts.formulaCellCount === 0
        ? { comparisons: 0, mismatches: [] }
        : await validateFormulaOracles(imported.snapshot, bytes)
    collectGarbage()
    const structuralSmokePassed = runStructuralSmoke ? runStructuralSmokeOps(imported.snapshot) : null
    const unsupportedFeatureClassifications = classifyUnsupportedFeatures(imported.snapshot, imported.warnings)
    const roundTripPassed = roundTripsSupportedSemantics(detachImportedWorkbookSnapshot(imported))
    const validation: PublicWorkbookValidationSummary = {
      importPassed: true,
      formulaOraclePassed: formulaOracleValidation.mismatches.length === 0,
      formulaOracleComparisons: formulaOracleValidation.comparisons,
      formulaOracleMismatches: formulaOracleValidation.mismatches,
      roundTripPassed,
      structuralSmokePassed,
    }
    const passed =
      validation.importPassed &&
      validation.formulaOraclePassed &&
      validation.roundTripPassed &&
      (validation.structuralSmokePassed === null || validation.structuralSmokePassed)
    const status: PublicWorkbookCaseStatus = passed ? (unsupportedFeatureClassifications.length > 0 ? 'unsupported' : 'passed') : 'failed'
    return {
      id: artifact.id,
      sourceId: artifact.sourceId,
      sourceUrl: artifact.sourceUrl,
      fileName: artifact.fileName,
      sha256: artifact.sha256,
      byteSize: artifact.byteSize,
      license: artifact.license,
      status,
      passed,
      featureCounts,
      workbookMetadata: metadata,
      validation,
      unsupportedFeatureClassifications,
      evidence: [
        ...baseEvidence,
        `sheets=${String(featureCounts.sheetCount)}`,
        `cells=${String(featureCounts.cellCount)}`,
        `formulas=${String(featureCounts.formulaCellCount)}`,
        ...validationEvidence(validation),
      ],
    }
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error)
    return failedCase(artifact, 'error', baseEvidence, [message])
  }
}

function artifactBaseEvidence(artifact: PublicWorkbookArtifact): string[] {
  return [
    `source=${artifact.sourceUrl}`,
    `license=${artifact.license.title}`,
    `sha256=${artifact.sha256}`,
    ...(artifact.topicEvidence ?? []).map((entry) => `topic=${entry}`),
  ]
}

async function validateFormulaOracles(snapshot: WorkbookSnapshot, bytes: Uint8Array): Promise<FormulaOracleValidationResult> {
  try {
    const formulaOracles = extractFormulaOracles(bytes)
    return {
      comparisons: formulaOracles.length,
      mismatches: await compareFormulaOracles(snapshot, formulaOracles),
    }
  } catch (error) {
    return {
      comparisons: 0,
      mismatches: [`Formula oracle check failed: ${error instanceof Error ? error.message : String(error)}`],
    }
  }
}

function validationEvidence(validation: PublicWorkbookValidationSummary): string[] {
  const evidence: string[] = []
  if (!validation.formulaOraclePassed) {
    evidence.push(...validation.formulaOracleMismatches.slice(0, 25))
  }
  if (!validation.roundTripPassed) {
    evidence.push('Round-trip projection failed')
  }
  if (validation.structuralSmokePassed === false) {
    evidence.push('Structural smoke operations failed')
  }
  return evidence
}

function failedCase(
  artifact: PublicWorkbookArtifact,
  status: 'failed' | 'error',
  evidence: readonly string[],
  errors: readonly string[],
): PublicWorkbookCorpusCase {
  return {
    id: artifact.id,
    sourceId: artifact.sourceId,
    sourceUrl: artifact.sourceUrl,
    fileName: artifact.fileName,
    sha256: artifact.sha256,
    byteSize: artifact.byteSize,
    license: artifact.license,
    status,
    passed: false,
    featureCounts: emptyFeatureCounts(),
    workbookMetadata: { workbookName: artifact.fileName, sheetNames: [], dimensions: [] },
    validation: {
      importPassed: false,
      formulaOraclePassed: false,
      formulaOracleComparisons: 0,
      formulaOracleMismatches: [],
      roundTripPassed: false,
      structuralSmokePassed: null,
    },
    unsupportedFeatureClassifications: [],
    evidence: [...evidence, ...errors],
  }
}

function unsupportedResourceLimitCase(
  artifact: PublicWorkbookArtifact,
  evidence: readonly string[],
  footprint: ReturnType<typeof inspectWorkbookFootprint>,
  maxCellCount: number,
): PublicWorkbookCorpusCase {
  return {
    id: artifact.id,
    sourceId: artifact.sourceId,
    sourceUrl: artifact.sourceUrl,
    fileName: artifact.fileName,
    sha256: artifact.sha256,
    byteSize: artifact.byteSize,
    license: artifact.license,
    status: 'unsupported',
    passed: true,
    featureCounts: footprint.featureCounts,
    workbookMetadata: footprint.workbookMetadata,
    validation: {
      importPassed: false,
      formulaOraclePassed: true,
      formulaOracleComparisons: 0,
      formulaOracleMismatches: [],
      roundTripPassed: true,
      structuralSmokePassed: null,
    },
    unsupportedFeatureClassifications: [`xlsx.publicCorpus.resourceLimit:cellCount>${String(maxCellCount)}`],
    evidence: [
      ...evidence,
      `cells=${String(footprint.featureCounts.cellCount)}`,
      `Public corpus verification cell-count limit exceeded: ${String(footprint.featureCounts.cellCount)} > ${String(maxCellCount)}`,
    ],
  }
}

async function compareFormulaOracles(snapshot: WorkbookSnapshot, oracles: readonly FormulaOracle[]): Promise<string[]> {
  if (oracles.length === 0) {
    return []
  }
  const engine = new SpreadsheetEngine({
    workbookName: snapshot.workbook.name,
    replicaId: `public-corpus-${stableId(snapshot.workbook.name)}`,
  })
  await engine.ready()
  engine.importSnapshot(snapshot)
  engine.recalculateNow()
  const mismatches: string[] = []
  for (const oracle of oracles) {
    const actual = engine.getCellValue(oracle.sheetName, oracle.address)
    if (!cellValuesMatchOracle(actual, oracle.expected)) {
      mismatches.push(`${oracle.sheetName}!${oracle.address} expected ${formatCellValue(oracle.expected)} got ${formatCellValue(actual)}`)
    }
  }
  return mismatches
}

function roundTripsSupportedSemantics(snapshot: WorkbookSnapshot): boolean {
  try {
    const workbookName = snapshot.workbook.name
    const expectedDigest = roundTripSemanticsDigest(snapshot)
    collectGarbage()
    const exported = exportXlsx(snapshot)
    snapshot = createDetachedWorkbookSnapshot(workbookName)
    collectGarbage()
    const actualDigest = roundTripSemanticsDigest(importXlsx(exported, `${workbookName}.xlsx`).snapshot)
    return actualDigest === expectedDigest
  } catch {
    return false
  }
}

function detachImportedWorkbookSnapshot(imported: ReturnType<typeof importXlsx>): WorkbookSnapshot {
  const snapshot = imported.snapshot
  imported.snapshot = createDetachedWorkbookSnapshot(snapshot.workbook.name)
  collectGarbage()
  return snapshot
}

function createDetachedWorkbookSnapshot(workbookName: string): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: workbookName },
    sheets: [],
  }
}

function collectGarbage(): void {
  if (typeof Bun !== 'undefined' && typeof Bun.gc === 'function') {
    Bun.gc(true)
    return
  }
  const gc = Reflect.get(globalThis, 'gc')
  if (typeof gc === 'function') {
    gc()
  }
}

function runStructuralSmokeOps(snapshot: WorkbookSnapshot): boolean | null {
  try {
    const sheetName = findStructuralSmokeSheetName(snapshot)
    if (!sheetName) {
      return null
    }
    const engine = new SpreadsheetEngine({ workbookName: `${snapshot.workbook.name}-structural-smoke` })
    engine.importSnapshot(structuredClone(snapshot))
    engine.insertRows(sheetName, 0, 1)
    engine.deleteRows(sheetName, 0, 1)
    engine.recalculateNow()
    return true
  } catch {
    return false
  }
}

function findStructuralSmokeSheetName(snapshot: WorkbookSnapshot): string | null {
  const sheet = snapshot.sheets.find((entry) => !entry.metadata?.sheetProtection && (entry.metadata?.protectedRanges?.length ?? 0) === 0)
  return sheet?.name ?? null
}

function classifyUnsupportedFeatures(snapshot: WorkbookSnapshot, warnings: readonly string[]): string[] {
  const classifications = new Set<string>()
  if ((snapshot.workbook.metadata?.macroPayloads?.length ?? 0) > 0) {
    classifications.add('xlsx.macros.execution.declined')
  }
  for (const warning of warnings) {
    classifications.add(`xlsx.import.warning:${warning}`)
  }
  return [...classifications].toSorted()
}

function stableId(value: string): string {
  return sha256HexSync(Buffer.from(value)).slice(0, 16)
}

async function mapWithConcurrency<T, R>(
  items: readonly T[],
  concurrency: number,
  mapper: (item: T, index: number) => Promise<R>,
): Promise<R[]> {
  const results: R[] = []
  let nextIndex = 0
  const workerCount = Math.min(Math.max(1, concurrency), items.length)
  const runNext = async (): Promise<void> => {
    const index = nextIndex
    nextIndex += 1
    if (index >= items.length) {
      return
    }
    results[index] = await mapper(items[index], index)
    await runNext()
  }
  await Promise.all(Array.from({ length: workerCount }, () => runNext()))
  return results
}

function compactProcessOutput(value: string): string | null {
  const compacted = value.replaceAll(rootDir, '<repo>').replace(/\s+/gu, ' ').trim()
  return compacted.length > 0 ? compacted.slice(0, 1_000) : null
}

function readManifest(path: string): PublicWorkbookManifest {
  const parsed: unknown = JSON.parse(readFileSync(path, 'utf8'))
  return parsePublicWorkbookManifestJson(parsed)
}

function writeJson(path: string, value: unknown, tempPrefix: string): void {
  mkdirSync(dirname(path), { recursive: true })
  writeFileSync(
    path,
    formatJsonForRepo({
      rootDir,
      serializedJson: `${JSON.stringify(value, null, 2)}\n`,
      tempPrefix,
    }),
  )
}

function readStringArg(name: string, fallback: string): string {
  const index = process.argv.indexOf(name)
  return index >= 0 ? (process.argv[index + 1] ?? fallback) : fallback
}

function readNumberArg(name: string, fallback: number): number {
  const raw = readStringArg(name, String(fallback))
  const parsed = Number(raw)
  if (!Number.isFinite(parsed) || parsed <= 0) {
    throw new Error(`Expected ${name} to be a positive number`)
  }
  return Math.trunc(parsed)
}

function readMegabytesArg(name: string, fallbackBytes: number): number {
  const raw = readStringArg(name, String(Math.ceil(fallbackBytes / 1024 / 1024)))
  const parsed = Number(raw)
  if (!Number.isFinite(parsed) || parsed <= 0) {
    throw new Error(`Expected ${name} to be a positive number of MiB`)
  }
  return Math.trunc(parsed * 1024 * 1024)
}

function capVerifyMaxRssBytes(value: number): number {
  return Math.min(Math.max(1, Math.trunc(value)), maxVerifyMaxRssBytes)
}

function readFlagArg(name: string): boolean {
  return process.argv.includes(name)
}

function startSelfRssGuard(maxRssBytes: number, label: string): () => void {
  const normalizedMaxRssBytes = Math.max(1, Math.trunc(maxRssBytes))
  const timer = setInterval(() => {
    const rssBytes = process.memoryUsage().rss
    if (rssBytes <= normalizedMaxRssBytes) {
      return
    }
    console.error(`${label} exceeded RSS limit: ${formatByteSize(rssBytes)} > ${formatByteSize(normalizedMaxRssBytes)}`)
    process.exit(70)
  }, defaultSelfRssCheckIntervalMs)
  timer.unref()
  return () => clearInterval(timer)
}

function readRepeatedStringArg(name: string): string[] {
  const values: string[] = []
  process.argv.forEach((arg, index) => {
    if (arg === name && process.argv[index + 1]) {
      values.push(process.argv[index + 1])
    }
  })
  return values
}

function readOrCreateManifest(path: string, targetWorkbookCount = 10_000): PublicWorkbookManifest {
  return existsSync(path) ? readManifest(path) : createEmptyPublicWorkbookManifest(undefined, targetWorkbookCount)
}

async function main(): Promise<void> {
  const command = process.argv[2] ?? 'verify'
  const manifestPath = resolve(readStringArg('--manifest', defaultManifestPath))
  const cacheDir = resolve(readStringArg('--cache-dir', defaultCacheDir))
  const scorecardPath = resolve(readStringArg('--scorecard', defaultScorecardPath))
  const verifyCheckpointPath = resolve(readStringArg('--verify-checkpoint', join(cacheDir, 'verification-checkpoint.json')))
  const targetWorkbookCount = readNumberArg('--target-workbook-count', 10_000)
  if (command === 'init') {
    await withPublicWorkbookCorpusCacheLock(cacheDir, 'init', async () => {
      writeJson(manifestPath, createEmptyPublicWorkbookManifest(undefined, targetWorkbookCount), 'public-workbook-corpus-manifest')
    })
    return
  }
  if (command === 'discover-ckan') {
    const portalBases = readRepeatedStringArg('--ckan-base')
    const manifest = await withPublicWorkbookCorpusCacheLock(cacheDir, 'discover-ckan', async () => {
      const discoveredManifest = await discoverCkanWorkbookSources({
        manifest: readOrCreateManifest(manifestPath, targetWorkbookCount),
        portalBases: portalBases.length > 0 ? portalBases : defaultCkanPortalBases,
        query: readStringArg('--query', 'xlsx'),
        limit: readNumberArg('--limit', 10_000),
        rowsPerRequest: readNumberArg('--rows', 100),
        ...(readStringArg('--required-topic', '') === 'financial-workpapers' ? { requiredTopic: 'financial-workpapers' as const } : {}),
      })
      writeJson(manifestPath, discoveredManifest, 'public-workbook-corpus-manifest')
      return discoveredManifest
    })
    console.log(`Discovered ${String(manifest.sources.length)} public workbook sources`)
    return
  }
  if (command === 'discover-financial-ckan') {
    const portalBases = readRepeatedStringArg('--ckan-base')
    const queries = readRepeatedStringArg('--query')
    const limit = readNumberArg('--limit', 5_000)
    const rowsPerRequest = readNumberArg('--rows', 100)
    const manifest = await withPublicWorkbookCorpusCacheLock(cacheDir, 'discover-financial-ckan', async () => {
      const discoveredManifest = await discoverFinancialCkanQueries({
        manifest: readOrCreateManifest(manifestPath, targetWorkbookCount),
        portalBases: portalBases.length > 0 ? portalBases : defaultCkanPortalBases,
        queries: queries.length > 0 ? queries : defaultFinancialWorkbookQueries,
        limit,
        rowsPerRequest,
        onQueryDiscovered: (checkpointManifest) => {
          writeJson(manifestPath, checkpointManifest, 'public-workbook-corpus-manifest')
        },
      })
      writeJson(manifestPath, discoveredManifest, 'public-workbook-corpus-manifest')
      return discoveredManifest
    })
    console.log(`Discovered ${String(manifest.sources.length)} financial workbook sources`)
    return
  }
  if (command === 'fetch') {
    const manifest = await withPublicWorkbookCorpusCacheLock(cacheDir, 'fetch', async () => {
      const fetchedManifest = await fetchPublicWorkbookArtifacts({
        manifest: readManifest(manifestPath),
        cacheDir,
        limit: readNumberArg('--limit', 10_000),
        downloadTimeoutMs: readNumberArg('--download-timeout-ms', defaultDownloadTimeoutMs),
        fingerprintTimeoutMs: readNumberArg('--fingerprint-timeout-ms', defaultFingerprintTimeoutMs),
        fingerprintMaxRssBytes: readMegabytesArg('--fingerprint-max-rss-mb', defaultFingerprintMaxRssBytes),
        isolatedFingerprinting: !readFlagArg('--in-process-fingerprint'),
        maxBytes: readNumberArg('--max-bytes', 50 * 1024 * 1024),
        onArtifactsCommitted: (checkpointManifest) => {
          writeJson(manifestPath, checkpointManifest, 'public-workbook-corpus-manifest')
          console.error(`Cached ${String(checkpointManifest.artifacts.length)} public workbook artifacts`)
        },
      })
      writeJson(manifestPath, fetchedManifest, 'public-workbook-corpus-manifest')
      return fetchedManifest
    })
    console.log(`Cached ${String(manifest.artifacts.length)} public workbook artifacts`)
    return
  }
  if (command === 'fingerprint-artifact') {
    const rawFilePath = readStringArg('--file', '')
    if (!rawFilePath) {
      throw new Error('Expected --file for fingerprint-artifact')
    }
    const workbookFingerprint = await fingerprintWorkbookFileIsolated(
      resolve(rawFilePath),
      readStringArg('--file-name', 'workbook.xlsx'),
      readNumberArg('--fingerprint-timeout-ms', defaultFingerprintTimeoutMs),
      {
        maxRssBytes: readMegabytesArg('--fingerprint-max-rss-mb', defaultFingerprintMaxRssBytes),
        rssCheckIntervalMs: 250,
      },
    )
    process.stdout.write(`${JSON.stringify({ workbookFingerprint })}\n`)
    return
  }
  if (command === 'fingerprint-artifact-worker') {
    const stopSelfRssGuard = startSelfRssGuard(
      readMegabytesArg('--fingerprint-max-rss-mb', defaultFingerprintMaxRssBytes),
      'Workbook fingerprinting worker',
    )
    const rawFilePath = readStringArg('--file', '')
    try {
      if (!rawFilePath) {
        throw new Error('Expected --file for fingerprint-artifact-worker')
      }
      const filePath = resolve(rawFilePath)
      const fileName = readStringArg('--file-name', 'workbook.xlsx')
      const workbookFingerprint = fingerprintWorkbookBytes(readFileSync(filePath), fileName)
      process.stdout.write(`${JSON.stringify({ workbookFingerprint })}\n`)
    } finally {
      stopSelfRssGuard()
    }
    return
  }
  if (command === 'verify-artifact') {
    const artifactId = readStringArg('--artifact-id', '')
    if (!artifactId) {
      throw new Error('Expected --artifact-id for verify-artifact')
    }
    const manifest = readManifest(manifestPath)
    const artifact = manifest.artifacts.find((entry) => entry.id === artifactId)
    if (!artifact) {
      throw new Error(`Manifest does not contain public workbook artifact ${artifactId}`)
    }
    const result = await verifyCachedWorkbookArtifactIsolated({
      artifact,
      cacheDir,
      manifestPath,
      runStructuralSmoke: readFlagArg('--structural-smoke'),
      timeoutMs: readNumberArg('--verify-timeout-ms', defaultVerifyTimeoutMs),
      maxRssBytes: capVerifyMaxRssBytes(readMegabytesArg('--verify-max-rss-mb', defaultVerifyMaxRssBytes)),
      maxCellCount: readNumberArg('--verify-max-cells', defaultVerifyMaxCellCount),
      rssCheckIntervalMs: 250,
    })
    process.stdout.write(`${JSON.stringify(result)}\n`)
    return
  }
  if (command === 'verify-artifact-worker') {
    const stopSelfRssGuard = startSelfRssGuard(
      capVerifyMaxRssBytes(readMegabytesArg('--verify-max-rss-mb', defaultVerifyMaxRssBytes)),
      'Workbook verification worker',
    )
    const artifactId = readStringArg('--artifact-id', '')
    try {
      if (!artifactId) {
        throw new Error('Expected --artifact-id for verify-artifact-worker')
      }
      const manifest = readManifest(manifestPath)
      const artifact = manifest.artifacts.find((entry) => entry.id === artifactId)
      if (!artifact) {
        throw new Error(`Manifest does not contain public workbook artifact ${artifactId}`)
      }
      const result = await verifyCachedWorkbookArtifact(
        artifact,
        cacheDir,
        readFlagArg('--structural-smoke'),
        readNumberArg('--verify-max-cells', defaultVerifyMaxCellCount),
      )
      process.stdout.write(`${JSON.stringify(result)}\n`)
    } finally {
      stopSelfRssGuard()
    }
    return
  }
  if (command === 'verify') {
    const scorecard = await withPublicWorkbookCorpusCacheLock(cacheDir, 'verify', async () => {
      const manifest = readManifest(manifestPath)
      const checkpointInterval = readNumberArg('--verify-checkpoint-interval', 10)
      const reusableCases = readFlagArg('--no-verify-resume')
        ? []
        : readReusablePublicWorkbookCorpusCases([scorecardPath, verifyCheckpointPath])
      const checkpointCasesById = new Map(reusableCases.map((entry) => [entry.id, entry]))
      let completedCount = 0
      let latestArtifactId = 'none'
      const startedAt = Date.now()
      const progressTimer = setInterval(() => {
        const elapsedSeconds = Math.max(1, Math.round((Date.now() - startedAt) / 1000))
        console.error(
          `Public workbook corpus verify progress: ${String(completedCount)}/${String(manifest.artifacts.length)} completed; latest=${latestArtifactId}; elapsed=${String(elapsedSeconds)}s`,
        )
      }, 15_000)
      progressTimer.unref()
      try {
        const verifiedScorecard = await buildPublicWorkbookCorpusScorecard({
          manifest,
          cacheDir,
          manifestPath,
          isolatedVerification: !readFlagArg('--in-process'),
          structuralSmokeSampleLimit: readNumberArg('--structural-smoke-sample-limit', 50),
          verifyConcurrency: readNumberArg('--verify-concurrency', 2),
          verifyTimeoutMs: readNumberArg('--verify-timeout-ms', defaultVerifyTimeoutMs),
          verifyMaxRssBytes: capVerifyMaxRssBytes(readMegabytesArg('--verify-max-rss-mb', defaultVerifyMaxRssBytes)),
          verifyMaxCellCount: readNumberArg('--verify-max-cells', defaultVerifyMaxCellCount),
          reusableCases,
          onCaseVerified: (progress) => {
            completedCount = progress.completedCount
            latestArtifactId = progress.latestCase.id
            checkpointCasesById.set(progress.latestCase.id, progress.latestCase)
            if (progress.completedCount % checkpointInterval === 0 || progress.completedCount === progress.totalCount) {
              writePublicWorkbookCorpusVerificationCheckpoint({
                path: verifyCheckpointPath,
                manifest,
                casesById: checkpointCasesById,
              })
            }
            if (
              progress.completedCount === progress.totalCount ||
              progress.completedCount % 50 === 0 ||
              progress.latestCase.status === 'failed' ||
              progress.latestCase.status === 'error'
            ) {
              console.error(
                `Public workbook corpus verified ${String(progress.completedCount)}/${String(progress.totalCount)}; latest=${progress.latestCase.id}; status=${progress.latestCase.status}`,
              )
            }
          },
        })
        writeJson(scorecardPath, verifiedScorecard, 'public-workbook-corpus-scorecard')
        return verifiedScorecard
      } finally {
        clearInterval(progressTimer)
      }
    })
    console.log(
      `Verified ${String(scorecard.summary.cachedWorkbookCount)} cached workbooks; ${String(scorecard.summary.remainingToTarget)} remaining`,
    )
    return
  }
  if (command === 'check') {
    const parsed: unknown = JSON.parse(readFileSync(scorecardPath, 'utf8'))
    const scorecard = parsePublicWorkbookCorpusScorecardJson(parsed)
    if (process.argv.includes('--require-target') && scorecard.summary.remainingToTarget > 0) {
      throw new Error(`Public workbook corpus target incomplete: ${String(scorecard.summary.remainingToTarget)} remaining`)
    }
    console.log(`Checked public workbook corpus scorecard with ${String(scorecard.summary.cachedWorkbookCount)} cached workbooks`)
    return
  }
  throw new Error(`Unknown public workbook corpus command: ${command}`)
}

if (process.argv[1] && pathToFileURL(resolve(process.argv[1])).href === import.meta.url) {
  await main()
}
