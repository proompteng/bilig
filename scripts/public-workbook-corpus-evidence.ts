import type { PublicWorkbookCorpusCase } from './public-workbook-corpus-types.ts'

export const publicWorkbookImportWarningClassifierEvidence = 'import-warning-classifier=2026-05-08'

export type PublicWorkbookCorpusEvidenceRefreshReason = 'missing-used-range-evidence' | 'missing-import-warning-classifier-evidence'

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

function hasCurrentImportWarningClassifierEvidence(entry: PublicWorkbookCorpusCase): boolean {
  return entry.evidence.includes(publicWorkbookImportWarningClassifierEvidence)
}
