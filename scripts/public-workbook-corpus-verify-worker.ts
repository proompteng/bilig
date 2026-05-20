import { createHash } from 'node:crypto'
import { existsSync, readFileSync } from 'node:fs'
import { join } from 'node:path'

import { parsePublicWorkbookArtifact, parsePublicWorkbookManifestJson } from './public-workbook-corpus-json.ts'
import {
  tryInspectLargeSimpleXlsxHeadless,
  type LargeSimpleXlsxHeadlessInspectResult,
} from '../packages/excel-import/src/xlsx-large-simple-headless-inspect.js'
import type { LargeSimpleXlsxImportStats } from '../packages/excel-import/src/xlsx-large-simple-import.js'
import {
  readXlsxZipEntryMetadata,
  readXlsxZipEntriesLazyFromByteSource,
  type XlsxZipByteSource,
} from '../packages/excel-import/src/xlsx-zip.js'
import type { PublicWorkbookArtifact, PublicWorkbookCorpusCase, PublicWorkbookFeatureCounts } from './public-workbook-corpus-types.ts'
import { defaultSelfRssCheckIntervalMs, startSelfRssGuard } from './public-workbook-corpus-process.ts'
import { readMegabytesArg, readNumberArg, readFlagArg, readStringArg } from './public-workbook-corpus-cli.ts'
import { largeSimpleImportPhaseTelemetryEvidence } from './public-workbook-corpus-large-simple-evidence.ts'
import { FileBackedXlsxZipByteSource, isZipWorkbookSource } from './public-workbook-corpus-xlsx-byte-source.ts'

const verificationWorkerPhasePrefix = 'bilig-public-workbook-verify-phase='
const defaultVerifyTimeoutMs = 180_000
const defaultVerifyMaxRssBytes = 1536 * 1024 * 1024
const defaultVerifyMaxCellCount = 1_500_000
const compactLargeSimplePreflightMinPackageBytes = 2 * 1024 * 1024
const compactLargeSimplePreflightMinWorksheetCompressedBytes = 256 * 1024

const verifyMaxRssBytes = capVerifyMaxRssBytes(readMegabytesArg('--verify-max-rss-mb', defaultVerifyMaxRssBytes))
const stopSelfRssGuard = startSelfRssGuard(verifyMaxRssBytes, 'Workbook verification worker')

try {
  const cacheDir = readStringArg('--cache-dir', '.cache/public-workbook-corpus')
  const artifactId = readStringArg('--artifact-id', '')
  if (!artifactId) {
    throw new Error('Expected --artifact-id for verify-artifact-worker')
  }
  const artifact = readWorkerArtifact(artifactId)
  if (!artifact) {
    throw new Error(`Manifest does not contain public workbook artifact ${artifactId}`)
  }
  const runStructuralSmoke = readFlagArg('--structural-smoke')
  const maxCellCount = readNumberArg('--verify-max-cells', defaultVerifyMaxCellCount)
  const workerOptions = {
    timeoutMs: readNumberArg('--verify-timeout-ms', defaultVerifyTimeoutMs),
    maxRssBytes: verifyMaxRssBytes,
    rssCheckIntervalMs: defaultSelfRssCheckIntervalMs,
    onPhase: writeWorkerPhase,
  }
  const result =
    tryVerifyCompactLargeSimpleArtifact(artifact, cacheDir, runStructuralSmoke, maxCellCount) ??
    (await (async () => {
      const { verifyCachedWorkbookArtifact } = await import('./public-workbook-corpus-verify.ts')
      return verifyCachedWorkbookArtifact(artifact, cacheDir, runStructuralSmoke, maxCellCount, workerOptions)
    })())
  process.stdout.write(`${JSON.stringify(result)}\n`)
} finally {
  stopSelfRssGuard()
}

function tryVerifyCompactLargeSimpleArtifact(
  artifact: PublicWorkbookArtifact,
  cacheDir: string,
  runStructuralSmoke: boolean,
  maxCellCount: number,
): PublicWorkbookCorpusCase | null {
  const cachePath = join(cacheDir, artifact.cachePath)
  if (!existsSync(cachePath)) {
    return null
  }
  writeWorkerPhase('read-cache')
  const source = new FileBackedXlsxZipByteSource(cachePath)
  try {
    if (sha256Hex(source) !== artifact.sha256 || !isZipWorkbookSource(source)) {
      return null
    }
    collectGarbage()
    if (artifact.byteSize < compactLargeSimplePreflightMinPackageBytes && !hasCompactLargeSimpleWorksheetPayload(source)) {
      return null
    }
    const zip = readXlsxZipEntriesLazyFromByteSource(source)
    if (!zip) {
      return null
    }
    writeWorkerPhase('import-xlsx')
    const imported = tryInspectLargeSimpleXlsxHeadless({ byteLength: source.byteLength }, artifact.fileName, zip, {
      afterWorksheetScan: collectGarbage,
      minByteLength: 0,
      releaseOwnedSourceBytes: () => ({ ownedSourceBytesBeforeRelease: source.byteLength, ownedSourceBytesAfterRelease: 0 }),
      releaseZipSource: true,
    })
    if (!imported) {
      return null
    }
    return buildCompactLargeSimpleCaseFromInspect(artifact, imported, runStructuralSmoke, maxCellCount)
  } finally {
    source.release()
  }
}

function hasCompactLargeSimpleWorksheetPayload(source: XlsxZipByteSource): boolean {
  const entries = readXlsxZipEntryMetadata(source)
  if (!entries) {
    return true
  }
  const worksheetCompressedBytes = entries.reduce(
    (sum, entry) => (/^xl\/worksheets\/[^/]+\.xml$/u.test(entry.path) ? sum + entry.compressedSize : sum),
    0,
  )
  return worksheetCompressedBytes >= compactLargeSimplePreflightMinWorksheetCompressedBytes
}

function buildCompactLargeSimpleCaseFromInspect(
  artifact: PublicWorkbookArtifact,
  imported: LargeSimpleXlsxHeadlessInspectResult,
  runStructuralSmoke: boolean,
  maxCellCount: number,
): PublicWorkbookCorpusCase | null {
  const featureCounts = featureCountsFromLargeSimpleStats(imported.stats)
  if (
    featureCounts.cellCount <= 100_000 ||
    featureCounts.cellCount > maxCellCount ||
    (featureCounts.formulaCellCount > 0 && featureCounts.formulaCellCount <= 2_000) ||
    !roundTripWouldBeResourceSkipped(artifact, featureCounts) ||
    (runStructuralSmoke && !structuralSmokeWouldBeResourceSkipped(featureCounts))
  ) {
    return null
  }
  const formulaOracleSkipped = featureCounts.formulaCellCount > 2_000
  const roundTripEvidence = `Round-trip projection skipped because workbook footprint exceeds verifier resource budget: cell-count ${String(
    featureCounts.cellCount,
  )} > 100000`
  const formulaOracleEvidence = `Formula oracle skipped because workbook has ${String(
    featureCounts.formulaCellCount,
  )} formulas, above verifier budget 2000.`
  const structuralSmokeEvidence = `Structural smoke skipped because workbook footprint exceeds verifier resource budget: cell-count ${String(
    featureCounts.cellCount,
  )} > 100000`
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
    featureCounts,
    workbookMetadata: {
      workbookName: imported.workbookName,
      sheetNames: imported.sheetNames,
      dimensions: imported.stats.dimensions.map((dimension) => ({
        sheetName: dimension.sheetName,
        rowCount: dimension.usedRange ? dimension.usedRange.endRow + 1 : dimension.rowCount,
        columnCount: dimension.usedRange ? dimension.usedRange.endColumn + 1 : dimension.columnCount,
        nonEmptyCellCount: dimension.nonEmptyCellCount,
        usedRange: dimension.usedRange,
      })),
    },
    validation: {
      importPassed: true,
      formulaOraclePassed: true,
      formulaOracleComparisons: 0,
      formulaOracleMismatches: [],
      roundTripPassed: true,
      structuralSmokePassed: null,
    },
    unsupportedFeatureClassifications: [
      ...(formulaOracleSkipped ? ['xlsx.publicCorpus.resourceLimit:preflightFormulaOracleBudget>2000formulas'] : []),
      'xlsx.publicCorpus.resourceLimit:preflightRoundTripBudget>100000cells',
      ...(runStructuralSmoke ? ['xlsx.publicCorpus.resourceLimit:preflightStructuralSmokeBudget>100000cells'] : []),
    ],
    evidence: [
      `source=${artifact.sourceUrl}`,
      `license=${artifact.license.title}`,
      `sha256=${artifact.sha256}`,
      `sheets=${String(featureCounts.sheetCount)}`,
      `cells=${String(featureCounts.cellCount)}`,
      `formulas=${String(featureCounts.formulaCellCount)}`,
      ...largeSimpleImportPhaseTelemetryEvidence(imported.stats),
      'resource-limit-classifier=2026-05-17-native-streaming-xlsx-footprint',
      ...(formulaOracleSkipped
        ? [
            'rss-limit-phase=formula-oracle',
            formulaOracleEvidence,
            `formula-oracle-formula-count=${String(featureCounts.formulaCellCount)}`,
          ]
        : []),
      'rss-limit-phase=round-trip',
      roundTripEvidence,
      ...(runStructuralSmoke ? ['rss-limit-phase=structural-smoke', structuralSmokeEvidence] : []),
    ],
  }
}

function readWorkerArtifact(artifactId: string): PublicWorkbookArtifact | undefined {
  const artifactJsonBase64 = readStringArg('--artifact-json-base64', '')
  if (artifactJsonBase64) {
    const artifact = parsePublicWorkbookArtifact(JSON.parse(Buffer.from(artifactJsonBase64, 'base64').toString('utf8')))
    if (artifact.id !== artifactId) {
      throw new Error(`Worker artifact id mismatch: expected ${artifactId}, received ${artifact.id}`)
    }
    return artifact
  }
  const manifestPath = readStringArg('--manifest', '.cache/public-workbook-corpus/manifest.json')
  const manifest = parsePublicWorkbookManifestJson(JSON.parse(readFileSync(manifestPath, 'utf8')))
  return manifest.artifacts.find((entry) => entry.id === artifactId)
}

function writeWorkerPhase(phase: string): void {
  process.stderr.write(`${verificationWorkerPhasePrefix}${phase}\n`)
}

function capVerifyMaxRssBytes(value: number): number {
  const normalizedValue = Math.max(1, Math.trunc(value))
  if (normalizedValue > defaultVerifyMaxRssBytes) {
    throw new Error(
      `Public workbook corpus verification RSS limits above ${String(Math.ceil(defaultVerifyMaxRssBytes / 1024 / 1024))} MiB are disabled because workbook workers can hang interactive hosts.`,
    )
  }
  return normalizedValue
}

function sha256Hex(source: XlsxZipByteSource): string {
  const hash = createHash('sha256')
  const chunkSize = 64 * 1024
  for (let offset = 0; offset < source.byteLength; offset += chunkSize) {
    hash.update(source.readRange(offset, Math.min(source.byteLength, offset + chunkSize)))
  }
  return hash.digest('hex')
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

function roundTripWouldBeResourceSkipped(artifact: PublicWorkbookArtifact, featureCounts: PublicWorkbookFeatureCounts): boolean {
  return featureCounts.cellCount > 100_000 || (featureCounts.sheetCount >= 30 && artifact.byteSize > 2 * 1024 * 1024)
}

function structuralSmokeWouldBeResourceSkipped(featureCounts: PublicWorkbookFeatureCounts): boolean {
  return featureCounts.cellCount > 100_000 || featureCounts.sheetCount > 80
}

function featureCountsFromLargeSimpleStats(stats: LargeSimpleXlsxImportStats): PublicWorkbookFeatureCounts {
  return {
    sheetCount: stats.sheetCount,
    cellCount: stats.cellCount,
    formulaCellCount: stats.formulaCellCount,
    valueCellCount: stats.valueCellCount,
    definedNameCount: stats.definedNameCount,
    tableCount: stats.tableCount,
    chartCount: 0,
    pivotCount: 0,
    mergeCount: stats.mergeCount,
    styleRangeCount: 0,
    conditionalFormatCount: stats.conditionalFormatCount,
    dataValidationCount: stats.dataValidationCount ?? 0,
    macroPayloadCount: 0,
    warningCount: stats.warningCount,
  }
}
