import {
  tryInspectLargeSimpleXlsxHeadless,
  type LargeSimpleXlsxHeadlessInspectResult,
} from '../packages/excel-import/src/xlsx-large-simple-headless-inspect.js'
import type { LargeSimpleXlsxImportStats } from '../packages/excel-import/src/xlsx-large-simple-import.js'
import { readXlsxZipEntriesLazyFromByteSource, type XlsxZipByteSource, type XlsxZipEntries } from '../packages/excel-import/src/xlsx-zip.js'
import type { WorkbookSnapshot } from '../packages/protocol/src/types.js'
import {
  publicWorkbookResourceLimitClassifierEvidence,
  hasResourceLimitUnsupportedClassifications,
} from './public-workbook-corpus-evidence.ts'
import type { PublicWorkbookCorpusWorkerOptions } from './public-workbook-corpus-footprint.ts'
import { largeSimpleImportPhaseTelemetryEvidence } from './public-workbook-corpus-large-simple-evidence.ts'
import {
  formulaOracleFormulaCountResourceLimitPreflight,
  roundTripResourceLimitPreflight,
  structuralSmokeResourceLimitPreflight,
  unsupportedResourceLimitCase,
} from './public-workbook-corpus-resource-limits.ts'
import { timeVerificationPhase, type startVerificationRuntimeMetrics } from './public-workbook-corpus-verification-metrics.ts'
import { isZipWorkbookSource } from './public-workbook-corpus-xlsx-byte-source.ts'
import type {
  PublicWorkbookArtifact,
  PublicWorkbookCorpusCase,
  PublicWorkbookFeatureCounts,
  PublicWorkbookValidationSummary,
} from './public-workbook-corpus-types.ts'
import type { WorkbookFootprint } from './public-workbook-corpus-workbook.ts'

declare const Bun:
  | {
      gc(force?: boolean): void
    }
  | undefined

export type LargeSimpleUnsupportedFeatureClassifier = (
  snapshot: WorkbookSnapshot,
  warnings: readonly string[],
  featureCounts: PublicWorkbookFeatureCounts,
  options: { readonly extraClassifications?: readonly string[] },
) => string[]

export function shouldUseCompactLargeSimpleVerification(
  artifact: PublicWorkbookArtifact,
  footprint: WorkbookFootprint,
  runStructuralSmoke: boolean,
): boolean {
  return (
    footprint.largeSimpleXlsxImport?.eligible === true &&
    shouldUseCompactLargeSimpleFeatureCounts(artifact, footprint.featureCounts, runStructuralSmoke)
  )
}

export function verifyLargeSimpleWorkbookCompactPreflight(args: {
  readonly artifact: PublicWorkbookArtifact
  readonly source: XlsxZipByteSource
  readonly baseEvidence: readonly string[]
  readonly classifyUnsupportedFeatures: LargeSimpleUnsupportedFeatureClassifier
  readonly maxCellCount: number
  readonly minByteLength: number
  readonly runStructuralSmoke: boolean
  readonly runtimeMetrics: ReturnType<typeof startVerificationRuntimeMetrics>
  readonly workerOptions: PublicWorkbookCorpusWorkerOptions
}): PublicWorkbookCorpusCase | null {
  if (args.source.byteLength < args.minByteLength) {
    return null
  }
  const zip = readLargeSimpleVerifierZipEntries(args.source)
  if (!zip) {
    return null
  }
  args.workerOptions.onPhase?.('import-xlsx')
  const startedAt = performance.now()
  const inspected = tryInspectLargeSimpleHeadless({
    byteLength: args.source.byteLength,
    fileName: args.artifact.fileName,
    zip,
    options: {
      minByteLength: 0,
      releaseZipSource: true,
    },
  })
  if (!inspected) {
    return null
  }
  const featureCounts = featureCountsFromLargeSimpleStats(inspected.stats)
  const footprint = footprintFromLargeSimpleInspect(inspected, featureCounts)
  const recordImportTiming = (): void => {
    args.runtimeMetrics.phaseTimings.push({ phase: 'import-xlsx', elapsedMs: Math.max(0, Math.round(performance.now() - startedAt)) })
  }
  if (featureCounts.cellCount > args.maxCellCount) {
    recordImportTiming()
    return unsupportedResourceLimitCase(args.artifact, args.baseEvidence, footprint, args.maxCellCount)
  }
  if (!shouldUseCompactLargeSimpleFeatureCounts(args.artifact, featureCounts, args.runStructuralSmoke)) {
    return null
  }
  recordImportTiming()
  return buildLargeSimpleCompactCase(
    args.artifact,
    inspected,
    featureCounts,
    args.baseEvidence,
    args.runStructuralSmoke,
    args.classifyUnsupportedFeatures,
  )
}

export async function verifyLargeSimpleWorkbookCompact(args: {
  readonly artifact: PublicWorkbookArtifact
  readonly source: XlsxZipByteSource
  readonly footprint: WorkbookFootprint
  readonly baseEvidence: readonly string[]
  readonly classifyUnsupportedFeatures: LargeSimpleUnsupportedFeatureClassifier
  readonly runStructuralSmoke: boolean
  readonly runtimeMetrics: ReturnType<typeof startVerificationRuntimeMetrics>
  readonly workerOptions: PublicWorkbookCorpusWorkerOptions
}): Promise<PublicWorkbookCorpusCase | null> {
  const inspected = await timeVerificationPhase(args.runtimeMetrics, args.workerOptions, 'import-xlsx', () => {
    const zip = readLargeSimpleVerifierZipEntries(args.source)
    return zip
      ? tryInspectLargeSimpleHeadless({
          byteLength: args.source.byteLength,
          fileName: args.artifact.fileName,
          zip,
          options: {
            minByteLength: 0,
            releaseZipSource: true,
          },
        })
      : null
  })
  if (!inspected) {
    return null
  }
  const featureCounts = mergeFeatureCounts(featureCountsFromLargeSimpleStats(inspected.stats), args.footprint.featureCounts)
  return buildLargeSimpleCompactCase(
    args.artifact,
    inspected,
    featureCounts,
    args.baseEvidence,
    args.runStructuralSmoke,
    args.classifyUnsupportedFeatures,
  )
}

function tryInspectLargeSimpleHeadless(args: {
  readonly byteLength: number
  readonly fileName: string
  readonly zip: XlsxZipEntries
  readonly options: Parameters<typeof tryInspectLargeSimpleXlsxHeadless>[3]
}): LargeSimpleXlsxHeadlessInspectResult | null {
  try {
    return tryInspectLargeSimpleXlsxHeadless({ byteLength: args.byteLength }, args.fileName, args.zip, args.options)
  } catch {
    return null
  }
}

function readLargeSimpleVerifierZipEntries(source: XlsxZipByteSource): XlsxZipEntries | null {
  return isZipWorkbookSource(source) ? readXlsxZipEntriesLazyFromByteSource(borrowXlsxZipByteSource(source)) : null
}

function borrowXlsxZipByteSource(source: XlsxZipByteSource): XlsxZipByteSource {
  return {
    byteLength: source.byteLength,
    readRange: (start, end) => source.readRange(start, end),
  }
}

function shouldUseCompactLargeSimpleFeatureCounts(
  artifact: PublicWorkbookArtifact,
  counts: PublicWorkbookFeatureCounts,
  runStructuralSmoke: boolean,
): boolean {
  const formulaOracleResourceLimit = formulaOracleFormulaCountResourceLimitPreflight(counts)
  const roundTripResourceLimit = roundTripResourceLimitPreflight(artifact, counts)
  const structuralSmokeResourceLimit = runStructuralSmoke ? structuralSmokeResourceLimitPreflight(counts) : null
  return (
    counts.cellCount > 0 &&
    (counts.formulaCellCount === 0 || formulaOracleResourceLimit !== null) &&
    roundTripResourceLimit !== null &&
    (!runStructuralSmoke || structuralSmokeResourceLimit !== null)
  )
}

function buildLargeSimpleCompactCase(
  artifact: PublicWorkbookArtifact,
  inspected: LargeSimpleXlsxHeadlessInspectResult,
  featureCounts: PublicWorkbookFeatureCounts,
  baseEvidence: readonly string[],
  runStructuralSmoke: boolean,
  classifyUnsupportedFeatures: LargeSimpleUnsupportedFeatureClassifier,
): PublicWorkbookCorpusCase {
  const metadata: PublicWorkbookCorpusCase['workbookMetadata'] = {
    workbookName: inspected.workbookName,
    sheetNames: inspected.sheetNames,
    dimensions: inspected.stats.dimensions,
  }
  collectGarbage()
  const formulaOracleResourceLimit = formulaOracleFormulaCountResourceLimitPreflight(featureCounts)
  const roundTripResourceLimit = roundTripResourceLimitPreflight(artifact, featureCounts)
  const structuralSmokeResourceLimit = runStructuralSmoke ? structuralSmokeResourceLimitPreflight(featureCounts) : null
  if (
    (featureCounts.formulaCellCount > 0 && !formulaOracleResourceLimit) ||
    !roundTripResourceLimit ||
    (runStructuralSmoke && !structuralSmokeResourceLimit)
  ) {
    throw new Error('Large-simple compact verification requires resource-skipped round-trip and structural phases.')
  }
  const phaseResourceLimitClassifications = [
    ...(formulaOracleResourceLimit ? [formulaOracleResourceLimit.classification] : []),
    roundTripResourceLimit.classification,
    ...(structuralSmokeResourceLimit ? [structuralSmokeResourceLimit.classification] : []),
  ]
  const phaseResourceLimitEvidence = [
    ...(formulaOracleResourceLimit?.evidence ?? []),
    ...roundTripResourceLimit.evidence,
    ...(structuralSmokeResourceLimit?.evidence ?? []),
  ]
  const unsupportedFeatureClassifications = classifyUnsupportedFeatures(
    minimalSnapshotFromLargeSimpleInspect(inspected),
    inspected.warnings,
    featureCounts,
    { extraClassifications: phaseResourceLimitClassifications },
  )
  const validation: PublicWorkbookValidationSummary = {
    importPassed: true,
    formulaOraclePassed: true,
    formulaOracleComparisons: 0,
    formulaOracleMismatches: [],
    roundTripPassed: true,
    structuralSmokePassed: null,
  }
  return {
    id: artifact.id,
    sourceId: artifact.sourceId,
    sourceUrl: artifact.sourceUrl,
    fileName: artifact.fileName,
    sha256: artifact.sha256,
    byteSize: artifact.byteSize,
    license: artifact.license,
    status: unsupportedFeatureClassifications.length > 0 ? 'unsupported' : 'passed',
    passed: true,
    featureCounts,
    workbookMetadata: metadata,
    validation,
    unsupportedFeatureClassifications,
    evidence: [
      ...baseEvidence,
      `sheets=${String(featureCounts.sheetCount)}`,
      `cells=${String(featureCounts.cellCount)}`,
      `formulas=${String(featureCounts.formulaCellCount)}`,
      ...largeSimpleImportPhaseTelemetryEvidence(inspected.stats),
      ...(hasResourceLimitUnsupportedClassifications(unsupportedFeatureClassifications)
        ? [publicWorkbookResourceLimitClassifierEvidence, ...phaseResourceLimitEvidence]
        : []),
    ],
  }
}

function minimalSnapshotFromLargeSimpleInspect(inspected: LargeSimpleXlsxHeadlessInspectResult): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: inspected.workbookName },
    sheets: inspected.sheetNames.map((sheetName, order) => ({
      id: order + 1,
      name: sheetName,
      order,
      cells: [],
    })),
  }
}

function footprintFromLargeSimpleInspect(
  inspected: LargeSimpleXlsxHeadlessInspectResult,
  featureCounts: PublicWorkbookFeatureCounts,
): WorkbookFootprint {
  return {
    featureCounts,
    workbookMetadata: {
      workbookName: inspected.workbookName,
      sheetNames: inspected.sheetNames,
      dimensions: inspected.stats.dimensions,
    },
    externalWorkbookReferences: [],
    largeSimpleXlsxImport: { eligible: true, blockers: [] },
  }
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

function mergeFeatureCounts(
  importedFeatureCounts: PublicWorkbookFeatureCounts,
  footprintFeatureCounts: PublicWorkbookFeatureCounts,
): PublicWorkbookFeatureCounts {
  return {
    sheetCount: Math.max(importedFeatureCounts.sheetCount, footprintFeatureCounts.sheetCount),
    cellCount: Math.max(importedFeatureCounts.cellCount, footprintFeatureCounts.cellCount),
    formulaCellCount: Math.max(importedFeatureCounts.formulaCellCount, footprintFeatureCounts.formulaCellCount),
    valueCellCount: Math.max(importedFeatureCounts.valueCellCount, footprintFeatureCounts.valueCellCount),
    definedNameCount: Math.max(importedFeatureCounts.definedNameCount, footprintFeatureCounts.definedNameCount),
    tableCount: Math.max(importedFeatureCounts.tableCount, footprintFeatureCounts.tableCount),
    chartCount: Math.max(importedFeatureCounts.chartCount, footprintFeatureCounts.chartCount),
    pivotCount: Math.max(importedFeatureCounts.pivotCount, footprintFeatureCounts.pivotCount),
    mergeCount: Math.max(importedFeatureCounts.mergeCount, footprintFeatureCounts.mergeCount),
    styleRangeCount: Math.max(importedFeatureCounts.styleRangeCount, footprintFeatureCounts.styleRangeCount),
    conditionalFormatCount: Math.max(importedFeatureCounts.conditionalFormatCount, footprintFeatureCounts.conditionalFormatCount),
    dataValidationCount: Math.max(importedFeatureCounts.dataValidationCount, footprintFeatureCounts.dataValidationCount),
    macroPayloadCount: Math.max(importedFeatureCounts.macroPayloadCount, footprintFeatureCounts.macroPayloadCount),
    warningCount: Math.max(importedFeatureCounts.warningCount, footprintFeatureCounts.warningCount),
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
