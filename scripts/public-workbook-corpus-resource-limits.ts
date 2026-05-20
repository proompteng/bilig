import { parseCellAddress, parseRangeAddress } from '../packages/formula/src/addressing.js'
import { collectFormulaDependencyMetadata } from '../packages/formula/src/binder-dependencies.js'
import { parseFormula } from '../packages/formula/src/parser.js'
import type { WorkbookSnapshot } from '../packages/protocol/src/types.js'
import { publicWorkbookResourceLimitClassifierEvidence } from './public-workbook-corpus-evidence.ts'
import { formatByteSize } from './public-workbook-corpus-process.ts'
import type {
  PublicWorkbookArtifact,
  PublicWorkbookCorpusCase,
  PublicWorkbookExternalReferenceSummary,
  PublicWorkbookFeatureCounts,
} from './public-workbook-corpus-types.ts'
import { emptyFeatureCounts, type WorkbookFootprint } from './public-workbook-corpus-workbook.ts'

const preflightImportCellCountLimit = 200_000
const preflightLargeSimpleImportCellCountLimit = 750_000
const preflightImportPackageByteLimit = 8 * 1024 * 1024
const preflightRoundTripCellCountLimit = 100_000
const preflightRoundTripSheetCountLimit = 30
const preflightRoundTripPackageByteLimit = 2 * 1024 * 1024
const preflightStructuralSmokeCellCountLimit = 100_000
const preflightStructuralSmokeSheetCountLimit = 80
const preflightFormulaOracleFormulaCellLimit = 2_000
const preflightFormulaOracleDependencyCellLimit = 2_000_000
const preflightFormulaOracleRangeCellLimit = 50_000
const excelMaxRows = 1_048_576
const excelMaxColumns = 16_384

export interface ResourceLimitPreflight {
  readonly classification: string
  readonly evidence: readonly string[]
}

export interface FormulaOracleDependencyFootprint {
  readonly formulaCellCount: number
  readonly dependencyReferenceCount: number
  readonly totalDependencyCellReferences: number
  readonly maxDependencyCellReferences: number
  readonly maxDependencyReference: string | null
  readonly unparseableDependencyReferenceCount: number
}

export function importResourceLimitPreflight(
  artifact: PublicWorkbookArtifact,
  footprint: WorkbookFootprint,
): ResourceLimitPreflight | null {
  const reasons: string[] = []
  const usesLargeSimpleBudget = usesLargeSimpleXlsxImportBudget(footprint)
  const importCellCountLimit = usesLargeSimpleBudget ? preflightLargeSimpleImportCellCountLimit : preflightImportCellCountLimit
  if (footprint.featureCounts.cellCount > importCellCountLimit) {
    reasons.push(`cell-count ${String(footprint.featureCounts.cellCount)} > ${String(importCellCountLimit)}`)
  }
  if (artifact.byteSize > preflightImportPackageByteLimit) {
    reasons.push(`package-bytes ${String(artifact.byteSize)} > ${String(preflightImportPackageByteLimit)}`)
  }
  if (
    !usesLargeSimpleBudget &&
    footprint.featureCounts.sheetCount >= preflightRoundTripSheetCountLimit &&
    artifact.byteSize > preflightRoundTripPackageByteLimit
  ) {
    reasons.push(
      `sheet/package budget ${String(footprint.featureCounts.sheetCount)} sheets and ${String(artifact.byteSize)} bytes exceeds ${String(
        preflightRoundTripSheetCountLimit,
      )} sheets / ${String(preflightRoundTripPackageByteLimit)} bytes`,
    )
  }
  if (
    !usesLargeSimpleBudget &&
    footprint.featureCounts.cellCount > preflightStructuralSmokeCellCountLimit &&
    footprint.featureCounts.formulaCellCount > 2_000
  ) {
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

function usesLargeSimpleXlsxImportBudget(footprint: WorkbookFootprint): boolean {
  const counts = footprint.featureCounts
  if (footprint.largeSimpleXlsxImport) {
    return (
      footprint.largeSimpleXlsxImport.eligible &&
      counts.chartCount === 0 &&
      counts.pivotCount === 0 &&
      counts.macroPayloadCount === 0 &&
      footprint.externalWorkbookReferences.length === 0
    )
  }
  return (
    counts.formulaCellCount === 0 &&
    counts.definedNameCount === 0 &&
    counts.tableCount === 0 &&
    counts.chartCount === 0 &&
    counts.pivotCount === 0 &&
    counts.conditionalFormatCount === 0 &&
    counts.dataValidationCount === 0 &&
    counts.macroPayloadCount === 0 &&
    footprint.externalWorkbookReferences.length === 0
  )
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

export function formulaOracleResourceLimitPreflight(snapshot: WorkbookSnapshot): ResourceLimitPreflight | null {
  const formulaCellCount = countFormulaCells(snapshot)
  const formulaCountLimit = formulaOracleFormulaCountResourceLimitPreflight({ formulaCellCount })
  if (formulaCountLimit) {
    return formulaCountLimit
  }
  const footprint = inspectFormulaOracleDependencyFootprint(snapshot)
  const reasons: string[] = []
  if (footprint.totalDependencyCellReferences > preflightFormulaOracleDependencyCellLimit) {
    reasons.push(
      `dependency-cell-references ${String(footprint.totalDependencyCellReferences)} > ${String(
        preflightFormulaOracleDependencyCellLimit,
      )}`,
    )
  }
  if (footprint.maxDependencyCellReferences > preflightFormulaOracleRangeCellLimit) {
    reasons.push(
      `largest-dependency-range ${String(footprint.maxDependencyCellReferences)} cells at ${
        footprint.maxDependencyReference ?? 'unknown'
      } > ${String(preflightFormulaOracleRangeCellLimit)}`,
    )
  }
  if (reasons.length === 0) {
    return null
  }
  return {
    classification: `xlsx.publicCorpus.resourceLimit:preflightFormulaOracleBudget>${String(
      preflightFormulaOracleDependencyCellLimit,
    )}dependencyCells`,
    evidence: [
      'rss-limit-phase=formula-oracle',
      `Formula oracle skipped because formula dependency footprint exceeds verifier resource budget: ${reasons.join('; ')}`,
      `formula-oracle-dependency-footprint=${String(footprint.totalDependencyCellReferences)}`,
      `formula-oracle-largest-dependency=${footprint.maxDependencyReference ?? 'unknown'}:${String(footprint.maxDependencyCellReferences)}`,
    ],
  }
}

export function formulaOracleFormulaCountResourceLimitPreflight(input: {
  readonly formulaCellCount: number
}): ResourceLimitPreflight | null {
  if (input.formulaCellCount > preflightFormulaOracleFormulaCellLimit) {
    return {
      classification: `xlsx.publicCorpus.resourceLimit:preflightFormulaOracleBudget>${String(
        preflightFormulaOracleFormulaCellLimit,
      )}formulas`,
      evidence: [
        'rss-limit-phase=formula-oracle',
        `Formula oracle skipped because workbook has ${String(input.formulaCellCount)} formulas, above verifier budget ${String(
          preflightFormulaOracleFormulaCellLimit,
        )}.`,
        `formula-oracle-formula-count=${String(input.formulaCellCount)}`,
      ],
    }
  }
  return null
}

function countFormulaCells(snapshot: WorkbookSnapshot): number {
  let formulaCellCount = 0
  for (const sheet of snapshot.sheets) {
    for (const cell of sheet.cells) {
      if (cell.formula !== undefined) {
        formulaCellCount += 1
      }
    }
  }
  return formulaCellCount
}

export function inspectFormulaOracleDependencyFootprint(snapshot: WorkbookSnapshot): FormulaOracleDependencyFootprint {
  let formulaCellCount = 0
  let dependencyReferenceCount = 0
  let totalDependencyCellReferences = 0
  let maxDependencyCellReferences = 0
  let maxDependencyReference: string | null = null
  let unparseableDependencyReferenceCount = 0
  for (const sheet of snapshot.sheets) {
    for (const cell of sheet.cells) {
      if (cell.formula === undefined) {
        continue
      }
      formulaCellCount += 1
      let dependencies: readonly string[]
      try {
        dependencies = collectFormulaDependencyMetadata(parseFormula(cell.formula)).deps
      } catch {
        continue
      }
      for (const dependency of dependencies) {
        dependencyReferenceCount += 1
        const dependencyCellReferences = countFormulaDependencyCells(dependency)
        if (dependencyCellReferences === null) {
          unparseableDependencyReferenceCount += 1
          continue
        }
        totalDependencyCellReferences += dependencyCellReferences
        if (dependencyCellReferences > maxDependencyCellReferences) {
          maxDependencyCellReferences = dependencyCellReferences
          maxDependencyReference = dependency
        }
      }
    }
  }
  return {
    formulaCellCount,
    dependencyReferenceCount,
    totalDependencyCellReferences,
    maxDependencyCellReferences,
    maxDependencyReference,
    unparseableDependencyReferenceCount,
  }
}

function countFormulaDependencyCells(dependency: string): number | null {
  try {
    if (!dependency.includes(':')) {
      parseCellAddress(dependency)
      return 1
    }
    const range = parseRangeAddress(dependency)
    if (range.kind === 'cells') {
      return (Math.abs(range.end.row - range.start.row) + 1) * (Math.abs(range.end.col - range.start.col) + 1)
    }
    if (range.kind === 'rows') {
      return (Math.abs(range.end.row - range.start.row) + 1) * excelMaxColumns
    }
    return (Math.abs(range.end.col - range.start.col) + 1) * excelMaxRows
  } catch {
    return null
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
    ...externalWorkbookReferenceSummaryFields(footprint),
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
    unsupportedFeatureClassifications: [
      `xlsx.publicCorpus.resourceLimit:cellCount>${String(maxCellCount)}`,
      ...externalWorkbookReferenceClassifications(footprint),
    ],
    evidence: [
      ...evidence,
      publicWorkbookResourceLimitClassifierEvidence,
      `cells=${String(footprint.featureCounts.cellCount)}`,
      ...externalWorkbookReferenceEvidence(footprint),
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
    ...externalWorkbookReferenceSummaryFields(footprint),
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
    unsupportedFeatureClassifications: [resourceLimit.classification, ...externalWorkbookReferenceClassifications(footprint)],
    evidence: [
      ...evidence,
      publicWorkbookResourceLimitClassifierEvidence,
      `sheets=${String(footprint.featureCounts.sheetCount)}`,
      `cells=${String(footprint.featureCounts.cellCount)}`,
      `formulas=${String(footprint.featureCounts.formulaCellCount)}`,
      ...externalWorkbookReferenceEvidence(footprint),
      ...resourceLimit.evidence,
    ],
  }
}

function externalWorkbookReferenceSummary(footprint: WorkbookFootprint): PublicWorkbookExternalReferenceSummary | undefined {
  if (footprint.externalWorkbookReferences.length === 0) {
    return undefined
  }
  return {
    linkedWorkbookCount: footprint.externalWorkbookReferences.length,
    formulaDependencyCount: 0,
    cachedValueDependencyCount: 0,
  }
}

function externalWorkbookReferenceSummaryFields(footprint: WorkbookFootprint): {
  readonly externalWorkbookReferences?: PublicWorkbookExternalReferenceSummary
} {
  const summary = externalWorkbookReferenceSummary(footprint)
  return summary ? { externalWorkbookReferences: summary } : {}
}

function externalWorkbookReferenceClassifications(footprint: WorkbookFootprint): readonly string[] {
  return footprint.externalWorkbookReferences.length > 0 ? ['xlsx.externalLinks.workbookReferencesPreserved'] : []
}

function externalWorkbookReferenceEvidence(footprint: WorkbookFootprint): readonly string[] {
  if (footprint.externalWorkbookReferences.length === 0) {
    return []
  }
  return [
    `external-workbook-links=${String(footprint.externalWorkbookReferences.length)}`,
    'external-workbook-formula-dependencies=0',
    'external-workbook-cached-value-dependencies=0',
    ...footprint.externalWorkbookReferences
      .slice(0, 10)
      .map((entry) => `external-workbook=${entry.workbookName ?? entry.target ?? entry.packagePath ?? `book#${String(entry.bookIndex)}`}`),
  ]
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
