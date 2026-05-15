import type { WorkbookPackageRelationshipSnapshot } from './package-artifacts.js'
import type { CellRangeRef } from './types.js'

export type PivotAggregation = 'sum' | 'count'

export interface WorkbookPivotValueSnapshot {
  sourceColumn: string
  summarizeBy: PivotAggregation
  outputLabel?: string
}

export interface WorkbookPivotSnapshot {
  name: string
  sheetName: string
  address: string
  source: CellRangeRef
  groupBy: string[]
  values: WorkbookPivotValueSnapshot[]
  rows: number
  cols: number
}

export interface WorkbookUnsupportedFormulaDependencySnapshot {
  kind: 'external-workbook-reference'
  sheetName: string
  address: string
  formula: string
  importedFormula: string
  resolvedExternalReferenceCount: number
  unresolvedExternalReferenceCount: number
  reason: string
}

export interface WorkbookUnsupportedPivotSnapshot {
  kind: 'external-cache' | 'raw-part'
  reason: string
  cacheId?: number
  sourceType?: string
  sheetName?: string
  address?: string
  name?: string
  packagePart?: string
}

export interface WorkbookPivotPackagePartSnapshot {
  path: string
  xml: string
}

export interface WorkbookPivotArtifactsSnapshot {
  parts: WorkbookPivotPackagePartSnapshot[]
  workbookPivotCachesXml?: string
  workbookRelationships?: WorkbookPackageRelationshipSnapshot[]
}

export interface WorkbookSheetPivotArtifactsSnapshot {
  relationships: WorkbookPackageRelationshipSnapshot[]
  pivotTableDefinitionsXml?: string
}
