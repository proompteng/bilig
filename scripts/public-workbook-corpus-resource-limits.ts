import { publicWorkbookResourceLimitClassifierEvidence } from './public-workbook-corpus-evidence.ts'
import { formatByteSize } from './public-workbook-corpus-process.ts'
import type { PublicWorkbookArtifact, PublicWorkbookCorpusCase, PublicWorkbookFeatureCounts } from './public-workbook-corpus-types.ts'
import { emptyFeatureCounts, type WorkbookFootprint } from './public-workbook-corpus-workbook.ts'

const preflightImportCellCountLimit = 200_000
const preflightImportPackageByteLimit = 8 * 1024 * 1024
const preflightRoundTripCellCountLimit = 100_000
const preflightRoundTripSheetCountLimit = 30
const preflightRoundTripPackageByteLimit = 2 * 1024 * 1024
const preflightStructuralSmokeCellCountLimit = 100_000
const preflightStructuralSmokeSheetCountLimit = 80

export interface ResourceLimitPreflight {
  readonly classification: string
  readonly evidence: readonly string[]
}

export function importResourceLimitPreflight(
  artifact: PublicWorkbookArtifact,
  footprint: WorkbookFootprint,
): ResourceLimitPreflight | null {
  const reasons: string[] = []
  if (footprint.featureCounts.cellCount > preflightImportCellCountLimit) {
    reasons.push(`cell-count ${String(footprint.featureCounts.cellCount)} > ${String(preflightImportCellCountLimit)}`)
  }
  if (artifact.byteSize > preflightImportPackageByteLimit) {
    reasons.push(`package-bytes ${String(artifact.byteSize)} > ${String(preflightImportPackageByteLimit)}`)
  }
  if (footprint.featureCounts.sheetCount >= preflightRoundTripSheetCountLimit && artifact.byteSize > preflightRoundTripPackageByteLimit) {
    reasons.push(
      `sheet/package budget ${String(footprint.featureCounts.sheetCount)} sheets and ${String(artifact.byteSize)} bytes exceeds ${String(
        preflightRoundTripSheetCountLimit,
      )} sheets / ${String(preflightRoundTripPackageByteLimit)} bytes`,
    )
  }
  if (footprint.featureCounts.cellCount > preflightStructuralSmokeCellCountLimit && footprint.featureCounts.formulaCellCount > 2_000) {
    reasons.push(
      `formula-oracle budget ${String(footprint.featureCounts.formulaCellCount)} formulas across ${String(
        footprint.featureCounts.cellCount,
      )} cells exceeds verifier preflight budget`,
    )
  }
  if (reasons.length === 0) {
    return null
  }
  return {
    classification: 'xlsx.publicCorpus.resourceLimit:preflightWorkbookBudget',
    evidence: [
      'rss-limit-phase=import-xlsx',
      `Public corpus verification import preflight limit exceeded: ${reasons.join('; ')}`,
      'The workbook was rejected before SheetJS import to avoid exceeding the worker RSS guard.',
    ],
  }
}

export function roundTripResourceLimitPreflight(
  artifact: PublicWorkbookArtifact,
  featureCounts: PublicWorkbookFeatureCounts,
): ResourceLimitPreflight | null {
  const reasons: string[] = []
  if (featureCounts.cellCount > preflightRoundTripCellCountLimit) {
    reasons.push(`cell-count ${String(featureCounts.cellCount)} > ${String(preflightRoundTripCellCountLimit)}`)
  }
  if (featureCounts.sheetCount >= preflightRoundTripSheetCountLimit && artifact.byteSize > preflightRoundTripPackageByteLimit) {
    reasons.push(
      `sheet/package budget ${String(featureCounts.sheetCount)} sheets and ${String(artifact.byteSize)} bytes exceeds ${String(
        preflightRoundTripSheetCountLimit,
      )} sheets / ${String(preflightRoundTripPackageByteLimit)} bytes`,
    )
  }
  if (reasons.length === 0) {
    return null
  }
  return {
    classification: `xlsx.publicCorpus.resourceLimit:preflightRoundTripBudget>${String(preflightRoundTripCellCountLimit)}cells`,
    evidence: [
      'rss-limit-phase=round-trip',
      `Round-trip projection skipped because workbook footprint exceeds verifier resource budget: ${reasons.join('; ')}`,
    ],
  }
}

export function structuralSmokeResourceLimitPreflight(featureCounts: PublicWorkbookFeatureCounts): ResourceLimitPreflight | null {
  const reasons: string[] = []
  if (featureCounts.cellCount > preflightStructuralSmokeCellCountLimit) {
    reasons.push(`cell-count ${String(featureCounts.cellCount)} > ${String(preflightStructuralSmokeCellCountLimit)}`)
  }
  if (featureCounts.sheetCount > preflightStructuralSmokeSheetCountLimit) {
    reasons.push(`sheet-count ${String(featureCounts.sheetCount)} > ${String(preflightStructuralSmokeSheetCountLimit)}`)
  }
  if (reasons.length === 0) {
    return null
  }
  return {
    classification: `xlsx.publicCorpus.resourceLimit:preflightStructuralSmokeBudget>${String(preflightStructuralSmokeCellCountLimit)}cells`,
    evidence: [
      'rss-limit-phase=structural-smoke',
      `Structural smoke skipped because workbook footprint exceeds verifier resource budget: ${reasons.join('; ')}`,
    ],
  }
}

export function unsupportedResourceLimitCase(
  artifact: PublicWorkbookArtifact,
  evidence: readonly string[],
  footprint: WorkbookFootprint,
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
      publicWorkbookResourceLimitClassifierEvidence,
      `cells=${String(footprint.featureCounts.cellCount)}`,
      `Public corpus verification cell-count limit exceeded: ${String(footprint.featureCounts.cellCount)} > ${String(maxCellCount)}`,
    ],
  }
}

export function unsupportedPreflightResourceLimitCase(
  artifact: PublicWorkbookArtifact,
  evidence: readonly string[],
  footprint: WorkbookFootprint,
  resourceLimit: ResourceLimitPreflight,
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
    unsupportedFeatureClassifications: [resourceLimit.classification],
    evidence: [
      ...evidence,
      publicWorkbookResourceLimitClassifierEvidence,
      `sheets=${String(footprint.featureCounts.sheetCount)}`,
      `cells=${String(footprint.featureCounts.cellCount)}`,
      `formulas=${String(footprint.featureCounts.formulaCellCount)}`,
      ...resourceLimit.evidence,
    ],
  }
}

export function unsupportedRssLimitCase(
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
      publicWorkbookResourceLimitClassifierEvidence,
      `Public corpus verification RSS limit exceeded: ${formatByteSize(rssBytes)} > ${formatByteSize(maxRssBytes)}`,
      ...details,
    ],
  }
}
