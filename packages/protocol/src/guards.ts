import { ValueTag } from './enums.js'
import type { CellRangeRef, CellSnapshot, LiteralInput, WorkbookSnapshot } from './types.js'

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
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

export function isLiteralInput(value: unknown): value is LiteralInput {
  return value === null || typeof value === 'boolean' || typeof value === 'number' || typeof value === 'string'
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
    Array.isArray(value['sheets'])
  )
}

export function isCellSnapshot(value: unknown): value is CellSnapshot {
  return (
    isRecord(value) &&
    typeof value['sheetName'] === 'string' &&
    typeof value['address'] === 'string' &&
    typeof value['flags'] === 'number' &&
    typeof value['version'] === 'number' &&
    isRecord(value['value']) &&
    isCellValueTag(value['value']['tag'])
  )
}
