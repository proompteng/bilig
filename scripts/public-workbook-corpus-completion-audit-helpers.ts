import type { PublicWorkbookCorpusCase } from './public-workbook-corpus-types.ts'
import { publicWorkbookCorpusCaseNeedsEvidenceRefresh } from './public-workbook-corpus-evidence.ts'

const requiredFeatureWitnesses = [
  { id: 'formulas', label: 'formulas', count: (entry: PublicWorkbookCorpusCase) => entry.featureCounts.formulaCellCount },
  { id: 'values', label: 'values', count: (entry: PublicWorkbookCorpusCase) => entry.featureCounts.valueCellCount },
  { id: 'names', label: 'defined names', count: (entry: PublicWorkbookCorpusCase) => entry.featureCounts.definedNameCount },
  { id: 'tables', label: 'tables', count: (entry: PublicWorkbookCorpusCase) => entry.featureCounts.tableCount },
  { id: 'charts', label: 'charts', count: (entry: PublicWorkbookCorpusCase) => entry.featureCounts.chartCount },
  { id: 'pivots', label: 'pivots', count: (entry: PublicWorkbookCorpusCase) => entry.featureCounts.pivotCount },
  { id: 'styles', label: 'styles', count: (entry: PublicWorkbookCorpusCase) => entry.featureCounts.styleRangeCount },
  { id: 'merged ranges', label: 'merged ranges', count: (entry: PublicWorkbookCorpusCase) => entry.featureCounts.mergeCount },
  {
    id: 'conditional formats',
    label: 'conditional formats',
    count: (entry: PublicWorkbookCorpusCase) => entry.featureCounts.conditionalFormatCount,
  },
] as const

export function financialWorkbookTargetCount(targetWorkbookCount: number): number {
  return Math.min(5_000, targetWorkbookCount)
}

export function hasFinancialTopicEvidence(entry: { readonly topicEvidence?: readonly string[] }): boolean {
  return (entry.topicEvidence?.length ?? 0) > 0
}

export interface PublicWorkbookCorpusUnsupportedClassificationCount {
  readonly classification: string
  readonly count: number
}

export function buildUnsupportedClassificationCounts(
  cases: readonly PublicWorkbookCorpusCase[],
): PublicWorkbookCorpusUnsupportedClassificationCount[] {
  const counts = new Map<string, number>()
  for (const entry of cases) {
    for (const classification of entry.unsupportedFeatureClassifications) {
      counts.set(classification, (counts.get(classification) ?? 0) + 1)
    }
  }
  return [...counts.entries()]
    .map(([classification, count]) => ({ classification, count }))
    .toSorted((left, right) => right.count - left.count || left.classification.localeCompare(right.classification))
}

export function formatUnsupportedClassificationCounts(counts: readonly PublicWorkbookCorpusUnsupportedClassificationCount[]): string {
  return counts.length === 0 ? 'none' : counts.map((entry) => `${entry.classification}=${String(entry.count)}`).join('; ')
}

export interface PublicWorkbookCorpusUnsupportedCaseSummary {
  readonly currentRecordedUnsupportedCaseCount: number
  readonly staleRecordedUnsupportedCaseCount: number
  readonly currentUnsupportedClassifications: readonly PublicWorkbookCorpusUnsupportedClassificationCount[]
  readonly staleUnsupportedClassifications: readonly PublicWorkbookCorpusUnsupportedClassificationCount[]
}

export function buildUnsupportedCaseSummary(cases: readonly PublicWorkbookCorpusCase[]): PublicWorkbookCorpusUnsupportedCaseSummary {
  const staleCases = cases.filter((entry) => entry.status === 'unsupported' && publicWorkbookCorpusCaseNeedsEvidenceRefresh(entry))
  const currentCases = cases.filter((entry) => entry.status === 'unsupported' && !publicWorkbookCorpusCaseNeedsEvidenceRefresh(entry))
  return {
    currentRecordedUnsupportedCaseCount: currentCases.length,
    staleRecordedUnsupportedCaseCount: staleCases.length,
    currentUnsupportedClassifications: buildUnsupportedClassificationCounts(currentCases),
    staleUnsupportedClassifications: buildUnsupportedClassificationCounts(staleCases),
  }
}

export function countGap(actual: number, required: number, label: string): string[] {
  return actual >= required ? [] : [`${label}: ${String(actual)}/${String(required)}`]
}

export function duplicateGap(values: readonly string[], label: string): string[] {
  return new Set(values).size === values.length ? [] : [`${label}: ${String(values.length - new Set(values).size)}`]
}

export function readNonNegativeInteger(record: Record<string, unknown>, key: string, fallback: number): number {
  const value = record[key]
  return typeof value === 'number' && Number.isInteger(value) && value >= 0 ? value : fallback
}

export function isRepoEvidenceArtifact(artifact: string): boolean {
  return (
    artifact.length > 0 &&
    !artifact.startsWith('.cache/') &&
    !artifact.startsWith('http://') &&
    !artifact.startsWith('https://') &&
    !artifact.includes('<') &&
    !artifact.includes('>') &&
    !artifact.startsWith('$') &&
    !artifact.startsWith('/')
  )
}

export function pnpmScriptName(command: string): string | null {
  const parts = command
    .trim()
    .split(/\s+/u)
    .filter((part) => !/^[A-Za-z_][A-Za-z0-9_]*=.*/u.test(part))
  const pnpmIndex = parts.indexOf('pnpm')
  if (pnpmIndex < 0) {
    return null
  }
  const candidate = parts[pnpmIndex + 1]
  if (!candidate || candidate === '--' || candidate === 'exec' || candidate === 'dlx') {
    return null
  }
  if (candidate === 'run') {
    return parts[pnpmIndex + 2] ?? null
  }
  if (candidate.startsWith('-')) {
    return null
  }
  return candidate
}

export function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

export function hasCacheIntegrityFailureEvidence(entry: PublicWorkbookCorpusCase): boolean {
  return entry.evidence.some(
    (line) => line.startsWith('Missing cached workbook file:') || line.startsWith('Cached workbook hash mismatch:'),
  )
}

export function isHashAddressedCachePath(cachePath: string, sha256: string): boolean {
  const parts = cachePath.split(/[\\/]/u)
  return !cachePath.startsWith('/') && !parts.includes('..') && parts.some((part) => part.startsWith(sha256))
}

export function hasFeatureValidationEvidence(entry: PublicWorkbookCorpusCase): boolean {
  const dimensionCellCount = entry.workbookMetadata.dimensions.reduce((sum, dimension) => sum + dimension.nonEmptyCellCount, 0)
  return (
    entry.workbookMetadata.workbookName.trim().length > 0 &&
    entry.workbookMetadata.sheetNames.length === entry.featureCounts.sheetCount &&
    entry.workbookMetadata.dimensions.length === entry.featureCounts.sheetCount &&
    dimensionCellCount === entry.featureCounts.cellCount &&
    entry.workbookMetadata.dimensions.every(
      (dimension) =>
        dimension.sheetName.trim().length > 0 &&
        dimension.rowCount >= 0 &&
        dimension.columnCount >= 0 &&
        dimension.nonEmptyCellCount >= 0 &&
        (dimension.nonEmptyCellCount === 0 || (dimension.rowCount > 0 && dimension.columnCount > 0)),
    ) &&
    entry.featureCounts.cellCount >= entry.featureCounts.formulaCellCount + entry.featureCounts.valueCellCount &&
    entry.featureCounts.definedNameCount >= 0 &&
    entry.featureCounts.tableCount >= 0 &&
    entry.featureCounts.chartCount >= 0 &&
    entry.featureCounts.pivotCount >= 0 &&
    entry.featureCounts.mergeCount >= 0 &&
    entry.featureCounts.styleRangeCount >= 0 &&
    entry.featureCounts.conditionalFormatCount >= 0 &&
    entry.featureCounts.dataValidationCount >= 0 &&
    entry.featureCounts.macroPayloadCount >= 0 &&
    entry.featureCounts.warningCount >= 0
  )
}

export function hasRecordedProvenanceEvidence(entry: PublicWorkbookCorpusCase): boolean {
  return (
    entry.evidence.includes(`source=${entry.sourceUrl}`) &&
    entry.evidence.includes(`license=${entry.license.title}`) &&
    entry.evidence.includes(`sha256=${entry.sha256}`)
  )
}

export function hasWorkbookMetadata(entry: PublicWorkbookCorpusCase): boolean {
  return entry.workbookMetadata.sheetNames.length > 0 && entry.workbookMetadata.dimensions.length > 0
}

export function isResourceLimitedUnsupportedCase(entry: PublicWorkbookCorpusCase): boolean {
  return (
    entry.status === 'unsupported' &&
    entry.passed &&
    entry.unsupportedFeatureClassifications.some((classification) => classification.startsWith('xlsx.publicCorpus.resourceLimit:')) &&
    entry.evidence.some((line) => line.startsWith('Public corpus verification RSS limit exceeded:'))
  )
}

export function buildFeatureWitnessCoverage(cases: readonly PublicWorkbookCorpusCase[]): {
  readonly id: string
  readonly label: string
  readonly totalCount: number
  readonly witnessCaseCount: number
}[] {
  return requiredFeatureWitnesses.map((family) => ({
    id: family.id,
    label: family.label,
    totalCount: cases.reduce((sum, entry) => sum + family.count(entry), 0),
    witnessCaseCount: cases.filter((entry) => family.count(entry) > 0).length,
  }))
}
