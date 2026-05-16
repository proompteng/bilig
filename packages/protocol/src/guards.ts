import { ErrorCode, ValueTag } from './enums.js'
import type { CellRangeRef, CellSnapshot, LiteralInput, WorkbookSnapshot } from './types.js'

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function isSafeNonNegativeInteger(value: unknown): value is number {
  return typeof value === 'number' && Number.isSafeInteger(value) && value >= 0
}

function isCellValueTag(value: unknown): value is ValueTag {
  return (
    value === ValueTag.Empty ||
    value === ValueTag.Number ||
    value === ValueTag.Boolean ||
    value === ValueTag.String ||
    value === ValueTag.Error
  )
}

function isErrorCode(value: unknown): value is ErrorCode {
  return (
    value === ErrorCode.None ||
    value === ErrorCode.Div0 ||
    value === ErrorCode.Ref ||
    value === ErrorCode.Value ||
    value === ErrorCode.Name ||
    value === ErrorCode.NA ||
    value === ErrorCode.Cycle ||
    value === ErrorCode.Spill ||
    value === ErrorCode.Blocked
  )
}

function isCellValue(value: unknown): boolean {
  if (!isRecord(value) || !isCellValueTag(value['tag'])) {
    return false
  }
  switch (value['tag']) {
    case ValueTag.Empty:
      return true
    case ValueTag.Number:
      return typeof value['value'] === 'number' && Number.isFinite(value['value'])
    case ValueTag.Boolean:
      return typeof value['value'] === 'boolean'
    case ValueTag.String:
      return typeof value['value'] === 'string' && isSafeNonNegativeInteger(value['stringId'])
    case ValueTag.Error:
      return isErrorCode(value['code'])
  }
}

function isWorkbookSnapshotCell(value: unknown): boolean {
  return (
    isRecord(value) &&
    typeof value['address'] === 'string' &&
    (value['row'] === undefined || isSafeNonNegativeInteger(value['row'])) &&
    (value['col'] === undefined || isSafeNonNegativeInteger(value['col'])) &&
    (value['value'] === undefined || isLiteralInput(value['value'])) &&
    (value['formula'] === undefined || typeof value['formula'] === 'string') &&
    (value['format'] === undefined || typeof value['format'] === 'string')
  )
}

function isWorkbookSnapshotSheet(value: unknown): boolean {
  return (
    isRecord(value) &&
    (value['id'] === undefined || isSafeNonNegativeInteger(value['id'])) &&
    typeof value['name'] === 'string' &&
    isSafeNonNegativeInteger(value['order']) &&
    Array.isArray(value['cells']) &&
    value['cells'].every((cell) => isWorkbookSnapshotCell(cell))
  )
}

export function isLiteralInput(value: unknown): value is LiteralInput {
  return value === null || typeof value === 'boolean' || typeof value === 'string' || (typeof value === 'number' && Number.isFinite(value))
}

export function isCellRangeRef(value: unknown): value is CellRangeRef {
  return (
    isRecord(value) &&
    typeof value['sheetName'] === 'string' &&
    typeof value['startAddress'] === 'string' &&
    typeof value['endAddress'] === 'string'
  )
}

export function isWorkbookSnapshot(value: unknown): value is WorkbookSnapshot {
  return (
    isRecord(value) &&
    value['version'] === 1 &&
    isRecord(value['workbook']) &&
    typeof value['workbook']['name'] === 'string' &&
    Array.isArray(value['sheets']) &&
    value['sheets'].every((sheet) => isWorkbookSnapshotSheet(sheet))
  )
}

export function isCellSnapshot(value: unknown): value is CellSnapshot {
  return (
    isRecord(value) &&
    typeof value['sheetName'] === 'string' &&
    typeof value['address'] === 'string' &&
    isSafeNonNegativeInteger(value['flags']) &&
    isSafeNonNegativeInteger(value['version']) &&
    isCellValue(value['value'])
  )
}
