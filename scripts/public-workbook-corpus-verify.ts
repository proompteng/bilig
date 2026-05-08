import { spawn } from 'node:child_process'
import { existsSync, readFileSync } from 'node:fs'
import { join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

import { SpreadsheetEngine } from '../packages/core/src/engine.js'
import {
  exportXlsx,
  externalWorkbookReferencesWarning,
  importXlsx,
  manualCalculationModeWarning,
  precisionAsDisplayedCalculationWarning,
  volatileFormulasWarning,
} from '../packages/excel-import/src/index.js'
import type { WorkbookSnapshot } from '../packages/protocol/src/types.js'
import { parsePublicWorkbookCorpusCase, validatePublicWorkbookManifest } from './public-workbook-corpus-json.ts'
import {
  hasImportWarningUnsupportedClassifications,
  publicWorkbookImportWarningClassifierEvidence,
} from './public-workbook-corpus-evidence.ts'
import { inspectWorkbookFootprintIsolated, type PublicWorkbookCorpusWorkerOptions } from './public-workbook-corpus-footprint.ts'
import { formatByteSize, startChildRssWatchdog, terminateChildProcess } from './public-workbook-corpus-process.ts'
import { roundTripSemanticsDigest } from './public-workbook-corpus-roundtrip.ts'
import { buildPublicWorkbookCorpusScorecardFromCases } from './public-workbook-corpus-scorecard.ts'
import { indexReusablePublicWorkbookCorpusCases } from './public-workbook-corpus-verify-checkpoint.ts'
import {
  cellValuesMatchOracle,
  countWorkbookFeatures,
  emptyFeatureCounts,
  extractFormulaOracles,
  formatCellValue,
  inspectWorkbookFootprint,
  isUnsupportedCycleOracleMismatch,
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
  PublicWorkbookValidationSummary,
} from './public-workbook-corpus-types.ts'

declare const Bun:
  | {
      gc(force?: boolean): void
    }
  | undefined

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const publicWorkbookCorpusScriptPath = fileURLToPath(new URL('./public-workbook-corpus.ts', import.meta.url))
const noop = (): void => undefined

export const defaultVerifyTimeoutMs = 180_000
export const defaultVerifyConcurrency = 1
export const defaultVerifyMaxRssBytes = 1536 * 1024 * 1024
export const defaultVerifyMaxCellCount = 1_500_000
export const isolatedFootprintByteThreshold = 1_000_000

const externalWorkbookRoundTripSkipEvidence =
  'Round-trip projection skipped because external workbook links are not recalculated during XLSX import.'

export async function buildPublicWorkbookCorpusScorecard(args: BuildScorecardArgs): Promise<PublicWorkbookCorpusScorecard> {
  validatePublicWorkbookManifest(args.manifest)
  const structuralSmokeSampleLimit = args.structuralSmokeSampleLimit ?? 50
  const verifyConcurrency = Math.max(1, Math.trunc(args.verifyConcurrency ?? defaultVerifyConcurrency))
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
        : verifyCachedWorkbookArtifact(artifact, args.cacheDir, runStructuralSmoke, verifyMaxCellCount, {
            timeoutMs: verifyTimeoutMs,
            maxRssBytes: verifyMaxRssBytes,
            rssCheckIntervalMs: verifyRssCheckIntervalMs,
          })
    return verifiedCasePromise.then(reportVerifiedCase)
  })
  return buildPublicWorkbookCorpusScorecardFromCases({
    manifest: args.manifest,
    generatedAt: args.generatedAt,
    cases,
  })
}

export function verifyCachedWorkbookArtifactIsolated(args: {
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
          unsupportedRssLimitCase(args.artifact, baseEvidence, rssBytes, args.maxRssBytes, [
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

export async function verifyCachedWorkbookArtifact(
  artifact: PublicWorkbookArtifact,
  cacheDir: string,
  runStructuralSmoke: boolean,
  maxCellCount: number,
  workerOptions: PublicWorkbookCorpusWorkerOptions,
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
    const footprint =
      bytes.byteLength >= isolatedFootprintByteThreshold
        ? await inspectWorkbookFootprintIsolated({
            bytes,
            fileName: artifact.fileName,
            scriptPath: publicWorkbookCorpusScriptPath,
            options: workerOptions,
          })
        : inspectWorkbookFootprint(bytes, artifact.fileName)
    if (!footprint) {
      return failedCase(artifact, 'error', baseEvidence, [
        'Workbook footprint subprocess did not return a valid footprint.',
        'The workbook was isolated in a subprocess so the corpus verification run could continue.',
      ])
    }
    if (footprint.featureCounts.cellCount > maxCellCount) {
      return unsupportedResourceLimitCase(artifact, baseEvidence, footprint, maxCellCount)
    }
    collectGarbage()
    const imported = importXlsx(bytes, artifact.fileName)
    const featureCounts = countWorkbookFeatures(imported.snapshot, imported.warnings)
    const metadata = workbookMetadata(imported.snapshot)
    const formulaOracleValidation =
      footprint.featureCounts.formulaCellCount === 0 || hasUnsupportedFormulaOracleWarning(imported.warnings)
        ? { comparisons: 0, mismatches: [] }
        : await validateFormulaOracles(imported.snapshot, bytes)
    collectGarbage()
    const structuralSmokePassed = runStructuralSmoke ? runStructuralSmokeOps(imported.snapshot) : null
    const unsupportedFeatureClassifications = classifyUnsupportedFeatures(imported.snapshot, imported.warnings)
    const roundTripSkipEvidence = roundTripValidationSkipEvidence(imported.warnings)
    const roundTripPassed = roundTripSkipEvidence ? true : roundTripsSupportedSemantics(detachImportedWorkbookSnapshot(imported))
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
        ...(hasImportWarningUnsupportedClassifications(unsupportedFeatureClassifications)
          ? [publicWorkbookImportWarningClassifierEvidence]
          : []),
        ...(roundTripSkipEvidence ? [roundTripSkipEvidence] : []),
        ...validationEvidence(validation),
      ],
    }
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error)
    return failedCase(artifact, 'error', baseEvidence, [message])
  }
}

export function capVerifyMaxRssBytes(value: number): number {
  const normalizedValue = Math.max(1, Math.trunc(value))
  if (normalizedValue > defaultVerifyMaxRssBytes) {
    throw new Error(
      `Public workbook corpus verification RSS limits above ${String(Math.ceil(defaultVerifyMaxRssBytes / 1024 / 1024))} MiB are disabled because workbook workers can hang interactive hosts.`,
    )
  }
  return normalizedValue
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

function hasUnsupportedFormulaOracleWarning(warnings: readonly string[]): boolean {
  return warnings.some(
    (warning) =>
      warning === precisionAsDisplayedCalculationWarning ||
      warning === externalWorkbookReferencesWarning ||
      warning === manualCalculationModeWarning ||
      warning === volatileFormulasWarning,
  )
}

function roundTripValidationSkipEvidence(warnings: readonly string[]): string | null {
  return warnings.includes(externalWorkbookReferencesWarning) ? externalWorkbookRoundTripSkipEvidence : null
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

function unsupportedRssLimitCase(
  artifact: PublicWorkbookArtifact,
  evidence: readonly string[],
  rssBytes: number,
  maxRssBytes: number,
  details: readonly string[],
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
    featureCounts: emptyFeatureCounts(),
    workbookMetadata: { workbookName: artifact.fileName, sheetNames: [], dimensions: [] },
    validation: {
      importPassed: false,
      formulaOraclePassed: true,
      formulaOracleComparisons: 0,
      formulaOracleMismatches: [],
      roundTripPassed: true,
      structuralSmokePassed: null,
    },
    unsupportedFeatureClassifications: [`xlsx.publicCorpus.resourceLimit:rss>${String(Math.ceil(maxRssBytes / 1024 / 1024))}MiB`],
    evidence: [
      ...evidence,
      `Public corpus verification RSS limit exceeded: ${formatByteSize(rssBytes)} > ${formatByteSize(maxRssBytes)}`,
      ...details,
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
      const explanation = engine.explainCell(oracle.sheetName, oracle.address)
      if (isUnsupportedCycleOracleMismatch(actual, oracle.expected, explanation.inCycle)) {
        continue
      }
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
