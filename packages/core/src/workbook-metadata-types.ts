import type {
  WorkbookChartSnapshot,
  WorkbookImageSnapshot,
  WorkbookMacroPayloadSnapshot,
  WorkbookMergeRangeSnapshot,
  WorkbookShapeSnapshot,
  CellNumberFormatRecord,
  CellRangeRef,
  CellStylePatch,
  CellStyleRecord,
  LiteralInput,
  WorkbookAutoFilterSnapshot,
  WorkbookCalculationSettingsSnapshot,
  WorkbookCommentEntrySnapshot,
  WorkbookCommentThreadSnapshot,
  WorkbookConditionalFormatSnapshot,
  WorkbookDataValidationSnapshot,
  WorkbookDefinedNameValueSnapshot,
  WorkbookFreezePaneSnapshot,
  WorkbookNoteSnapshot,
  WorkbookRangeProtectionSnapshot,
  WorkbookSheetProtectionSnapshot,
  WorkbookPivotSnapshot,
  WorkbookPivotValueSnapshot,
  WorkbookTableSnapshot,
  WorkbookVolatileContextSnapshot,
} from '@bilig/protocol'
import { canonicalWorkbookAddress } from './workbook-range-records.js'

export interface WorkbookDefinedNameRecord {
  name: string
  scopeSheetName?: string
  value: WorkbookDefinedNameValueSnapshot
}

export interface WorkbookPropertyRecord {
  key: string
  value: LiteralInput
}

export interface WorkbookMacroPayloadRecord extends WorkbookMacroPayloadSnapshot {}

export interface WorkbookSpillRecord {
  sheetName: string
  address: string
  rows: number
  cols: number
}

export interface WorkbookPivotRecord extends WorkbookPivotSnapshot {
  values: WorkbookPivotValueSnapshot[]
}

export interface WorkbookChartRecord extends WorkbookChartSnapshot {}
export interface WorkbookImageRecord extends WorkbookImageSnapshot {}
export interface WorkbookShapeRecord extends WorkbookShapeSnapshot {}

export interface WorkbookTableRecord extends WorkbookTableSnapshot {}

export interface WorkbookAxisMetadataRecord {
  sheetName: string
  start: number
  count: number
  size: number | null
  hidden: boolean | null
}

export interface WorkbookAxisEntryRecord {
  id: string
  size: number | null
  hidden: boolean | null
}

export interface WorkbookCellStyleRecord extends CellStyleRecord {}

export interface WorkbookStyleRangeRecord {
  range: CellRangeRef
  styleId: string
}

export interface WorkbookCellNumberFormatRecord extends CellNumberFormatRecord {}

export interface WorkbookFormatRangeRecord {
  range: CellRangeRef
  formatId: string
}

export interface WorkbookCalculationSettingsRecord extends WorkbookCalculationSettingsSnapshot {}

export interface WorkbookVolatileContextRecord extends WorkbookVolatileContextSnapshot {}

export interface WorkbookFreezePaneRecord extends WorkbookFreezePaneSnapshot {
  sheetName: string
}

export interface WorkbookMergeRangeRecord extends WorkbookMergeRangeSnapshot {}

export interface WorkbookFilterRecord {
  sheetName: string
  range: WorkbookAutoFilterSnapshot
}

export interface WorkbookSortKeyRecord {
  keyAddress: string
  direction: 'asc' | 'desc'
}

export interface WorkbookSortRecord {
  sheetName: string
  range: CellRangeRef
  keys: WorkbookSortKeyRecord[]
}

export interface WorkbookDataValidationRecord extends WorkbookDataValidationSnapshot {}
export interface WorkbookConditionalFormatRecord extends WorkbookConditionalFormatSnapshot {
  style: CellStylePatch
}
export interface WorkbookSheetProtectionRecord extends WorkbookSheetProtectionSnapshot {}
export interface WorkbookRangeProtectionRecord extends WorkbookRangeProtectionSnapshot {}
export interface WorkbookCommentEntryRecord extends WorkbookCommentEntrySnapshot {}
export interface WorkbookCommentThreadRecord extends WorkbookCommentThreadSnapshot {
  comments: WorkbookCommentEntryRecord[]
}
export interface WorkbookNoteRecord extends WorkbookNoteSnapshot {}

export interface WorkbookMetadataRecord {
  properties: Map<string, WorkbookPropertyRecord>
  macroPayloads: Map<string, WorkbookMacroPayloadRecord>
  definedNames: Map<string, WorkbookDefinedNameRecord>
  tables: Map<string, WorkbookTableRecord>
  spills: Map<string, WorkbookSpillRecord>
  pivots: Map<string, WorkbookPivotRecord>
  charts: Map<string, WorkbookChartRecord>
  images: Map<string, WorkbookImageRecord>
  shapes: Map<string, WorkbookShapeRecord>
  rowMetadata: Map<string, WorkbookAxisMetadataRecord>
  columnMetadata: Map<string, WorkbookAxisMetadataRecord>
  calculationSettings: WorkbookCalculationSettingsRecord
  volatileContext: WorkbookVolatileContextRecord
  freezePanes: Map<string, WorkbookFreezePaneRecord>
  merges: Map<string, WorkbookMergeRangeRecord>
  sheetProtections: Map<string, WorkbookSheetProtectionRecord>
  filters: Map<string, WorkbookFilterRecord>
  sorts: Map<string, WorkbookSortRecord>
  dataValidations: Map<string, WorkbookDataValidationRecord>
  conditionalFormats: Map<string, WorkbookConditionalFormatRecord>
  rangeProtections: Map<string, WorkbookRangeProtectionRecord>
  commentThreads: Map<string, WorkbookCommentThreadRecord>
  notes: Map<string, WorkbookNoteRecord>
}

export function createWorkbookMetadataRecord(): WorkbookMetadataRecord {
  return {
    properties: new Map(),
    macroPayloads: new Map(),
    definedNames: new Map(),
    tables: new Map(),
    spills: new Map(),
    pivots: new Map(),
    charts: new Map(),
    images: new Map(),
    shapes: new Map(),
    rowMetadata: new Map(),
    columnMetadata: new Map(),
    calculationSettings: { mode: 'automatic', compatibilityMode: 'excel-modern' },
    volatileContext: { recalcEpoch: 0 },
    freezePanes: new Map(),
    merges: new Map(),
    sheetProtections: new Map(),
    filters: new Map(),
    sorts: new Map(),
    dataValidations: new Map(),
    conditionalFormats: new Map(),
    rangeProtections: new Map(),
    commentThreads: new Map(),
    notes: new Map(),
  }
}

export function normalizeDefinedName(name: string): string {
  return normalizeWorkbookObjectName(name, 'Defined names')
}

export function normalizeDefinedNameScope(scopeSheetName: string | undefined): string | undefined {
  const scope = scopeSheetName?.trim()
  return scope && scope.length > 0 ? scope : undefined
}

export function definedNameKey(name: string, scopeSheetName?: string): string {
  return `${normalizeDefinedNameScope(scopeSheetName) ?? '<workbook>'}\u0000${normalizeDefinedName(name)}`
}

export function compareDefinedNameRecords(left: WorkbookDefinedNameRecord, right: WorkbookDefinedNameRecord): number {
  return (
    normalizeDefinedName(left.name).localeCompare(normalizeDefinedName(right.name)) ||
    (left.scopeSheetName ?? '').localeCompare(right.scopeSheetName ?? '')
  )
}

export function normalizeWorkbookObjectName(name: string, label = 'Workbook object'): string {
  const normalized = name.trim().toUpperCase()
  if (normalized.length === 0) {
    throw new Error(`${label} must be non-empty`)
  }
  return normalized
}

export function macroPayloadKey(kind: WorkbookMacroPayloadSnapshot['kind']): string {
  return normalizeWorkbookObjectName(kind, 'Macro payloads')
}

export function pivotKey(sheetName: string, address: string): string {
  return `${sheetName}!${canonicalWorkbookAddress(sheetName, address)}`
}

export function chartKey(id: string): string {
  return normalizeWorkbookObjectName(id, 'Charts')
}

export function imageKey(id: string): string {
  return normalizeWorkbookObjectName(id, 'Images')
}

export function shapeKey(id: string): string {
  return normalizeWorkbookObjectName(id, 'Shapes')
}
