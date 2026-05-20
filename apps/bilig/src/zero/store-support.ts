import { isEngineReplicaSnapshot, type EngineReplicaSnapshot } from '@bilig/core'
import {
  isCellValue,
  isWorkbookSnapshot,
  sanitizeCellStyleRecord,
  ValueTag,
  type CellRangeRef,
  type CellStyleRecord,
  type CellValue,
  type WorkbookSnapshot,
} from '@bilig/protocol'
import {
  cellCoordinatesWithinBounds,
  createEmptyWorkbookSnapshot as createSharedEmptyWorkbookSnapshot,
  normalizeRangeBounds,
  type DirtyRegion,
  type WorkbookEventPayload,
} from '@bilig/zero-sync'
import type {
  AxisMetadataSourceRow,
  CellEvalRow,
  CellSourceRow,
  DefinedNameSourceRow,
  NumberFormatSourceRow,
  SheetSourceRow,
  StyleSourceRow,
  WorkbookMetadataSourceRow,
} from './projection.js'

export { normalizeRangeBounds } from '@bilig/zero-sync'

export type FocusedCellEventPayload = Extract<WorkbookEventPayload, { kind: 'setCellValue' | 'setCellFormula' | 'clearCell' }>

export type StyleRangeEventPayload = Extract<WorkbookEventPayload, { kind: 'setRangeStyle' | 'clearRangeStyle' }>

export type NumberFormatRangeEventPayload = Extract<WorkbookEventPayload, { kind: 'setRangeNumberFormat' | 'clearRangeNumberFormat' }>

export type ColumnMetadataEventPayload = Extract<WorkbookEventPayload, { kind: 'updateColumnWidth' | 'updateColumnMetadata' }>

export type RowMetadataEventPayload = Extract<WorkbookEventPayload, { kind: 'updateRowMetadata' }>

export const createEmptyWorkbookSnapshot = createSharedEmptyWorkbookSnapshot

export function nowIso(): string {
  return new Date().toISOString()
}

export function parseNullableInteger(value: unknown): number | null {
  if (typeof value === 'number' && Number.isSafeInteger(value)) {
    return value
  }
  if (typeof value === 'string') {
    const trimmed = value.trim()
    if (/^-?\d+$/u.test(trimmed)) {
      const parsed = Number(trimmed)
      return Number.isSafeInteger(parsed) ? parsed : null
    }
  }
  return null
}

export function parsePositiveInteger(value: unknown): number | null {
  const parsed = parseNullableInteger(value)
  return parsed !== null && parsed > 0 ? parsed : null
}

export function parseNonNegativeInteger(value: unknown): number | null {
  const parsed = parseNullableInteger(value)
  return parsed !== null && parsed >= 0 ? parsed : null
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function isSafeNonNegativeInteger(value: unknown): value is number {
  return typeof value === 'number' && Number.isSafeInteger(value) && value >= 0
}

export function isDirtyRegion(value: unknown): value is DirtyRegion {
  return (
    isRecord(value) &&
    typeof value['sheetName'] === 'string' &&
    isSafeNonNegativeInteger(value['rowStart']) &&
    isSafeNonNegativeInteger(value['rowEnd']) &&
    isSafeNonNegativeInteger(value['colStart']) &&
    isSafeNonNegativeInteger(value['colEnd']) &&
    value['rowEnd'] >= value['rowStart'] &&
    value['colEnd'] >= value['colStart']
  )
}

export function parseCheckpointPayload(value: unknown, documentId: string): WorkbookSnapshot {
  return isWorkbookSnapshot(value) ? value : createEmptyWorkbookSnapshot(documentId)
}

export function parseCheckpointReplicaState(value: unknown): EngineReplicaSnapshot | null {
  return isEngineReplicaSnapshot(value) ? value : null
}

export function isFocusedCellEventPayload(payload: WorkbookEventPayload): payload is FocusedCellEventPayload {
  return payload.kind === 'setCellValue' || payload.kind === 'setCellFormula' || payload.kind === 'clearCell'
}

export function isStyleRangeEventPayload(payload: WorkbookEventPayload): payload is StyleRangeEventPayload {
  return payload.kind === 'setRangeStyle' || payload.kind === 'clearRangeStyle'
}

export function isNumberFormatRangeEventPayload(payload: WorkbookEventPayload): payload is NumberFormatRangeEventPayload {
  return payload.kind === 'setRangeNumberFormat' || payload.kind === 'clearRangeNumberFormat'
}

export function isColumnMetadataEventPayload(payload: WorkbookEventPayload): payload is ColumnMetadataEventPayload {
  return payload.kind === 'updateColumnWidth' || payload.kind === 'updateColumnMetadata'
}

export function isRowMetadataEventPayload(payload: WorkbookEventPayload): payload is RowMetadataEventPayload {
  return payload.kind === 'updateRowMetadata'
}

export function eventRequiresRecalc(payload: WorkbookEventPayload): boolean {
  return !(
    payload.kind === 'setRangeStyle' ||
    payload.kind === 'clearRangeStyle' ||
    payload.kind === 'setRangeNumberFormat' ||
    payload.kind === 'clearRangeNumberFormat' ||
    payload.kind === 'updateRowMetadata' ||
    payload.kind === 'updateColumnMetadata' ||
    payload.kind === 'updateColumnWidth' ||
    payload.kind === 'setFreezePane' ||
    payload.kind === 'mergeCells' ||
    payload.kind === 'unmergeCells'
  )
}

function semanticSignature(value: unknown): string {
  return JSON.stringify(value)
}

export function sheetSignature(row: SheetSourceRow): string {
  return semanticSignature([row.name, row.sortOrder, row.freezeRows, row.freezeCols])
}

export function cellSignature(row: CellSourceRow): string {
  return semanticSignature([
    row.sheetName,
    row.address,
    row.rowNum,
    row.colNum,
    row.inputValue ?? null,
    row.formula ?? null,
    row.format ?? null,
    row.styleId ?? null,
    row.explicitFormatId ?? null,
  ])
}

export function axisSignature(row: AxisMetadataSourceRow): string {
  return semanticSignature([row.sheetName, row.startIndex, row.count, row.size ?? null, row.hidden ?? null])
}

export function definedNameSignature(row: DefinedNameSourceRow): string {
  return semanticSignature([row.name, row.value])
}

export function workbookMetadataSignature(row: WorkbookMetadataSourceRow): string {
  return semanticSignature([row.key, row.value])
}

export function styleSignature(row: StyleSourceRow): string {
  return semanticSignature([row.id, row.recordJSON, row.hash])
}

export function numberFormatSignature(row: NumberFormatSourceRow): string {
  return semanticSignature([row.id, row.code, row.kind])
}

export function cellEvalSignature(row: CellEvalRow): string {
  return semanticSignature([
    row.sheetName,
    row.address,
    row.rowNum,
    row.colNum,
    row.value,
    row.flags,
    row.version,
    row.styleId,
    row.styleJson,
    row.formatId,
    row.formatCode,
  ])
}

export function parseJsonKey(key: string): unknown[] {
  const parsed = JSON.parse(key) as unknown
  if (!Array.isArray(parsed)) {
    throw new Error(`Invalid projection key: ${key}`)
  }
  return parsed
}

export function parseCellEvalValue(value: unknown): CellValue {
  return isCellValue(value) ? value : { tag: ValueTag.Empty }
}

export function parseCellStyleRecord(value: unknown): CellStyleRecord | null {
  if (!isRecord(value) || typeof value['id'] !== 'string') {
    return null
  }
  return sanitizeCellStyleRecord(value['id'], value)
}

export function cellEvalRowInRange(row: Pick<CellEvalRow, 'sheetName' | 'rowNum' | 'colNum'>, range: CellRangeRef): boolean {
  const bounds = normalizeRangeBounds(range)
  return row.sheetName === bounds.sheetName && cellCoordinatesWithinBounds(row.rowNum, row.colNum, bounds)
}

export function cellSourceRowInRange(row: Pick<CellSourceRow, 'sheetName' | 'rowNum' | 'colNum'>, range: CellRangeRef): boolean {
  const bounds = normalizeRangeBounds(range)
  return row.sheetName === bounds.sheetName && cellCoordinatesWithinBounds(row.rowNum, row.colNum, bounds)
}
