import type { PublicWorkbookCorpusCase } from './public-workbook-corpus-types.ts'

export const publicWorkbookImportWarningClassifierEvidence = 'import-warning-classifier=2026-05-08-full-precision-formula-aware'
export const publicWorkbookPivotClassifierEvidence = 'pivot-classifier=2026-05-08-external-cache-warning'
export const publicWorkbookResourceLimitClassifierEvidence = 'resource-limit-classifier=2026-05-17-native-streaming-xlsx-footprint'
export const publicWorkbookFormulaOracleCacheClassifierEvidence =
  'formula-oracle-cache-classifier=2026-05-12-independent-recalculation-cross-check'

export type PublicWorkbookCorpusEvidenceRefreshReason =
  | 'missing-used-range-evidence'
  | 'missing-import-warning-classifier-evidence'
  | 'missing-pivot-classifier-evidence'
  | 'missing-resource-limit-classifier-evidence'
  | 'missing-formula-oracle-cache-classifier-evidence'

export function publicWorkbookCorpusCaseNeedsEvidenceRefresh(entry: PublicWorkbookCorpusCase): boolean {
  return publicWorkbookCorpusCaseEvidenceRefreshReasons(entry).length > 0
}

export function publicWorkbookCorpusCaseEvidenceRefreshReasons(
  entry: PublicWorkbookCorpusCase,
): readonly PublicWorkbookCorpusEvidenceRefreshReason[] {
  const reasons: PublicWorkbookCorpusEvidenceRefreshReason[] = []
  if (!hasPublicWorkbookCorpusUsedRangeEvidence(entry)) {
    reasons.push('missing-used-range-evidence')
  }
  if (hasImportWarningUnsupportedClassification(entry) && !hasCurrentImportWarningClassifierEvidence(entry)) {
    reasons.push('missing-import-warning-classifier-evidence')
  }
  if (hasPivotUnsupportedClassification(entry) && !hasCurrentPivotClassifierEvidence(entry)) {
    reasons.push('missing-pivot-classifier-evidence')
  }
  if (hasResourceLimitUnsupportedClassification(entry) && !hasCurrentResourceLimitClassifierEvidence(entry)) {
    reasons.push('missing-resource-limit-classifier-evidence')
  }
  if (hasFormulaOracleCacheUnsupportedClassification(entry) && !hasCurrentFormulaOracleCacheClassifierEvidence(entry)) {
    reasons.push('missing-formula-oracle-cache-classifier-evidence')
  }
  return reasons
}

export function hasPublicWorkbookCorpusUsedRangeEvidence(entry: PublicWorkbookCorpusCase): boolean {
  return entry.workbookMetadata.dimensions.every((dimension) => {
    if (!Object.hasOwn(dimension, 'usedRange')) {
      return false
    }
    const range = dimension.usedRange
    if (dimension.nonEmptyCellCount === 0) {
      return range === null
    }
    return (
      range !== null &&
      range !== undefined &&
      range.startRow >= 0 &&
      range.startColumn >= 0 &&
      range.endRow >= range.startRow &&
      range.endColumn >= range.startColumn &&
      dimension.rowCount === range.endRow + 1 &&
      dimension.columnCount === range.endColumn + 1
    )
  })
}

export function hasImportWarningUnsupportedClassification(entry: PublicWorkbookCorpusCase): boolean {
  return hasImportWarningUnsupportedClassifications(entry.unsupportedFeatureClassifications)
}

export function hasImportWarningUnsupportedClassifications(classifications: readonly string[]): boolean {
  return classifications.some((classification) => classification.startsWith('xlsx.import.warning:'))
}

export function hasPivotUnsupportedClassifications(classifications: readonly string[]): boolean {
  return classifications.some((classification) => classification.startsWith('xlsx.pivots.'))
}

export function hasResourceLimitUnsupportedClassifications(classifications: readonly string[]): boolean {
  return classifications.some((classification) => classification.startsWith('xlsx.publicCorpus.resourceLimit:'))
}

export function hasFormulaOracleCacheUnsupportedClassifications(classifications: readonly string[]): boolean {
  return classifications.some((classification) => classification.startsWith('xlsx.publicCorpus.formulaOracleCache:'))
}

function hasPivotUnsupportedClassification(entry: PublicWorkbookCorpusCase): boolean {
  return hasPivotUnsupportedClassifications(entry.unsupportedFeatureClassifications)
}

function hasResourceLimitUnsupportedClassification(entry: PublicWorkbookCorpusCase): boolean {
  return hasResourceLimitUnsupportedClassifications(entry.unsupportedFeatureClassifications)
}

function hasFormulaOracleCacheUnsupportedClassification(entry: PublicWorkbookCorpusCase): boolean {
  return hasFormulaOracleCacheUnsupportedClassifications(entry.unsupportedFeatureClassifications)
}

function hasCurrentImportWarningClassifierEvidence(entry: PublicWorkbookCorpusCase): boolean {
  return entry.evidence.includes(publicWorkbookImportWarningClassifierEvidence)
}

function hasCurrentPivotClassifierEvidence(entry: PublicWorkbookCorpusCase): boolean {
  return entry.evidence.includes(publicWorkbookPivotClassifierEvidence)
}

function hasCurrentResourceLimitClassifierEvidence(entry: PublicWorkbookCorpusCase): boolean {
  return entry.evidence.includes(publicWorkbookResourceLimitClassifierEvidence)
}

function hasCurrentFormulaOracleCacheClassifierEvidence(entry: PublicWorkbookCorpusCase): boolean {
  return entry.evidence.includes(publicWorkbookFormulaOracleCacheClassifierEvidence)
}
