import {
  getCellNumberFormatKind,
  type CellNumberFormatRecord,
  type CellRangeRef,
  type CellStyleRecord,
  type SheetFormatRangeSnapshot,
  type SheetStyleRangeSnapshot,
} from '@bilig/protocol'
import { formatAddress, parseCellAddress } from '@bilig/formula'
import type {
  WorkbookCellNumberFormatRecord,
  WorkbookCellStyleRecord,
  WorkbookFormatRangeRecord,
  WorkbookStyleRangeRecord,
} from './workbook-metadata-types.js'
import {
  cloneWorkbookRangeRecords,
  findWorkbookRangeRecord,
  overlayWorkbookRangeRecords,
  replaceWorkbookRangeRecords,
} from './workbook-range-records.js'
import {
  cellNumberFormatIdForCode,
  cellStyleIdForKey,
  cellStyleKey,
  normalizeCellNumberFormatRecord,
  normalizeCellStyleRecord,
} from './workbook-store-records.js'

interface WorkbookStyleFormatCatalog {
  cellStyles: Map<string, WorkbookCellStyleRecord>
  styleKeys: Map<string, string>
  cellNumberFormats: Map<string, WorkbookCellNumberFormatRecord>
  numberFormatKeys: Map<string, string>
}

interface WorkbookStyleFormatSheet {
  styleRanges: WorkbookStyleRangeRecord[]
  formatRanges: WorkbookFormatRangeRecord[]
}

export function upsertCellStyle(
  catalog: WorkbookStyleFormatCatalog,
  style: CellStyleRecord,
  bumpStyleId: (id: string) => void,
): WorkbookCellStyleRecord {
  const normalized = normalizeCellStyleRecord(style)
  const existing = catalog.cellStyles.get(normalized.id)
  if (existing) {
    catalog.styleKeys.delete(cellStyleKey(existing))
  }
  catalog.cellStyles.set(normalized.id, normalized)
  catalog.styleKeys.set(cellStyleKey(normalized), normalized.id)
  bumpStyleId(normalized.id)
  return normalized
}

export function internCellStyle(
  catalog: WorkbookStyleFormatCatalog,
  style: Omit<WorkbookCellStyleRecord, 'id'>,
  defaultStyleId: string,
): WorkbookCellStyleRecord {
  const normalized = normalizeCellStyleRecord({
    id: defaultStyleId,
    ...style,
  })
  const key = cellStyleKey(normalized)
  const existingId = catalog.styleKeys.get(key)
  if (existingId) {
    return catalog.cellStyles.get(existingId)!
  }
  return { ...normalized, id: cellStyleIdForKey(key) }
}

export function getCellStyle(
  catalog: WorkbookStyleFormatCatalog,
  id: string | undefined,
  defaultStyleId: string,
): WorkbookCellStyleRecord | undefined {
  if (!id) {
    return catalog.cellStyles.get(defaultStyleId)
  }
  return catalog.cellStyles.get(id) ?? catalog.cellStyles.get(defaultStyleId)
}

export function listCellStyles(catalog: WorkbookStyleFormatCatalog): WorkbookCellStyleRecord[] {
  return [...catalog.cellStyles.values()].toSorted((left, right) => left.id.localeCompare(right.id))
}

export function upsertCellNumberFormat(
  catalog: WorkbookStyleFormatCatalog,
  format: CellNumberFormatRecord,
  bumpFormatId: (id: string) => void,
): WorkbookCellNumberFormatRecord {
  const normalized = normalizeCellNumberFormatRecord(format)
  const existing = catalog.cellNumberFormats.get(normalized.id)
  if (existing) {
    catalog.numberFormatKeys.delete(existing.code)
  }
  catalog.cellNumberFormats.set(normalized.id, normalized)
  catalog.numberFormatKeys.set(normalized.code, normalized.id)
  bumpFormatId(normalized.id)
  return normalized
}

export function internCellNumberFormat(
  catalog: WorkbookStyleFormatCatalog,
  format: string | CellNumberFormatRecord,
  defaultFormatId: string,
): WorkbookCellNumberFormatRecord {
  const normalized =
    typeof format === 'string'
      ? normalizeCellNumberFormatRecord({
          id: defaultFormatId,
          code: format,
          kind: getCellNumberFormatKind(format),
        })
      : normalizeCellNumberFormatRecord(format)
  const existingId = catalog.numberFormatKeys.get(normalized.code)
  if (existingId) {
    return catalog.cellNumberFormats.get(existingId)!
  }
  return { ...normalized, id: cellNumberFormatIdForCode(normalized.code) }
}

export function getCellNumberFormat(
  catalog: WorkbookStyleFormatCatalog,
  id: string | undefined,
  defaultFormatId: string,
): WorkbookCellNumberFormatRecord | undefined {
  if (!id) {
    return catalog.cellNumberFormats.get(defaultFormatId)
  }
  return catalog.cellNumberFormats.get(id) ?? catalog.cellNumberFormats.get(defaultFormatId)
}

export function listCellNumberFormats(catalog: WorkbookStyleFormatCatalog): WorkbookCellNumberFormatRecord[] {
  return [...catalog.cellNumberFormats.values()].toSorted((left, right) => left.id.localeCompare(right.id))
}

export function setStyleRange(
  catalog: WorkbookStyleFormatCatalog,
  sheet: WorkbookStyleFormatSheet,
  range: CellRangeRef,
  styleId: string,
  defaultStyleId: string,
): WorkbookStyleRangeRecord {
  if (!catalog.cellStyles.has(styleId)) {
    throw new Error(`Unknown cell style: ${styleId}`)
  }
  const stored: WorkbookStyleRangeRecord = {
    range: { ...range },
    styleId,
  }
  sheet.styleRanges = overlayWorkbookRangeRecords(
    sheet.styleRanges,
    stored,
    (nextRange, record) => ({
      range: nextRange,
      styleId: record.styleId,
    }),
    (record) => record.styleId === defaultStyleId,
  )
  sheet.styleRanges = coalesceStyleRangeRecords(sheet.styleRanges)
  return stored
}

export function listStyleRanges(sheet: WorkbookStyleFormatSheet | undefined): WorkbookStyleRangeRecord[] {
  return cloneWorkbookRangeRecords(sheet?.styleRanges ?? [], (range, record) => ({
    range,
    styleId: record.styleId,
  }))
}

export function setStyleRanges(
  catalog: WorkbookStyleFormatCatalog,
  sheet: WorkbookStyleFormatSheet,
  ranges: readonly SheetStyleRangeSnapshot[],
): WorkbookStyleRangeRecord[] {
  const nextRanges = replaceWorkbookRangeRecords(
    ranges.map((entry) => ({
      range: { ...entry.range },
      styleId: entry.styleId,
    })),
    (range, record) => ({
      range,
      styleId: record.styleId,
    }),
    (entry) => {
      if (!catalog.cellStyles.has(entry.styleId)) {
        throw new Error(`Unknown cell style: ${entry.styleId}`)
      }
    },
  )
  sheet.styleRanges = coalesceStyleRangeRecords(nextRanges)
  return listStyleRanges(sheet)
}

export function getStyleId(sheet: WorkbookStyleFormatSheet | undefined, row: number, col: number, defaultStyleId: string): string {
  if (!sheet) {
    return defaultStyleId
  }
  return findWorkbookRangeRecord(sheet.styleRanges, row, col)?.styleId ?? defaultStyleId
}

export function setFormatRange(
  catalog: WorkbookStyleFormatCatalog,
  sheet: WorkbookStyleFormatSheet,
  range: CellRangeRef,
  formatId: string,
  defaultFormatId: string,
): WorkbookFormatRangeRecord {
  if (!catalog.cellNumberFormats.has(formatId)) {
    throw new Error(`Unknown cell number format: ${formatId}`)
  }
  const stored: WorkbookFormatRangeRecord = {
    range: { ...range },
    formatId,
  }
  sheet.formatRanges = overlayWorkbookRangeRecords(
    sheet.formatRanges,
    stored,
    (nextRange, record) => ({
      range: nextRange,
      formatId: record.formatId,
    }),
    (record) => record.formatId === defaultFormatId,
  )
  sheet.formatRanges = coalesceFormatRangeRecords(sheet.formatRanges)
  return stored
}

export function listFormatRanges(sheet: WorkbookStyleFormatSheet | undefined): WorkbookFormatRangeRecord[] {
  return cloneWorkbookRangeRecords(sheet?.formatRanges ?? [], (range, record) => ({
    range,
    formatId: record.formatId,
  }))
}

export function setFormatRanges(
  catalog: WorkbookStyleFormatCatalog,
  sheet: WorkbookStyleFormatSheet,
  ranges: readonly SheetFormatRangeSnapshot[],
): WorkbookFormatRangeRecord[] {
  const nextRanges = replaceWorkbookRangeRecords(
    ranges.map((entry) => ({
      range: { ...entry.range },
      formatId: entry.formatId,
    })),
    (range, record) => ({
      range,
      formatId: record.formatId,
    }),
    (entry) => {
      if (!catalog.cellNumberFormats.has(entry.formatId)) {
        throw new Error(`Unknown cell number format: ${entry.formatId}`)
      }
    },
  )
  sheet.formatRanges = coalesceFormatRangeRecords(nextRanges)
  return listFormatRanges(sheet)
}

export function getRangeFormatId(sheet: WorkbookStyleFormatSheet | undefined, row: number, col: number, defaultFormatId: string): string {
  if (!sheet) {
    return defaultFormatId
  }
  return findWorkbookRangeRecord(sheet.formatRanges, row, col)?.formatId ?? defaultFormatId
}

function coalesceStyleRangeRecords(records: readonly WorkbookStyleRangeRecord[]): WorkbookStyleRangeRecord[] {
  return coalesceWorkbookRangeRecords(
    records,
    (left, right) => left.styleId === right.styleId,
    (range, record) => ({
      range,
      styleId: record.styleId,
    }),
  )
}

function coalesceFormatRangeRecords(records: readonly WorkbookFormatRangeRecord[]): WorkbookFormatRangeRecord[] {
  return coalesceWorkbookRangeRecords(
    records,
    (left, right) => left.formatId === right.formatId,
    (range, record) => ({
      range,
      formatId: record.formatId,
    }),
  )
}

function coalesceWorkbookRangeRecords<RecordType extends { range: CellRangeRef }>(
  records: readonly RecordType[],
  canMerge: (left: RecordType, right: RecordType) => boolean,
  cloneRecord: (range: CellRangeRef, record: RecordType) => RecordType,
): RecordType[] {
  const sorted = records
    .map((record) => ({
      record,
      start: parseCellAddress(record.range.startAddress, record.range.sheetName),
      end: parseCellAddress(record.range.endAddress, record.range.sheetName),
    }))
    .toSorted((left, right) => {
      const sheetCompare = left.record.range.sheetName.localeCompare(right.record.range.sheetName)
      if (sheetCompare !== 0) {
        return sheetCompare
      }
      if (left.start.row !== right.start.row) {
        return left.start.row - right.start.row
      }
      if (left.start.col !== right.start.col) {
        return left.start.col - right.start.col
      }
      if (left.end.row !== right.end.row) {
        return left.end.row - right.end.row
      }
      return left.end.col - right.end.col
    })

  const merged: Array<{ record: RecordType; startRow: number; endRow: number; startCol: number; endCol: number }> = []
  sorted.forEach(({ record, start, end }) => {
    const startRow = Math.min(start.row, end.row)
    const endRow = Math.max(start.row, end.row)
    const startCol = Math.min(start.col, end.col)
    const endCol = Math.max(start.col, end.col)
    const previous = merged[merged.length - 1]
    if (
      previous &&
      canMerge(previous.record, record) &&
      previous.record.range.sheetName === record.range.sheetName &&
      ((previous.startCol === startCol && previous.endCol === endCol && startRow <= previous.endRow + 1) ||
        (previous.startRow === startRow && previous.endRow === endRow && startCol <= previous.endCol + 1))
    ) {
      previous.startRow = Math.min(previous.startRow, startRow)
      previous.endRow = Math.max(previous.endRow, endRow)
      previous.startCol = Math.min(previous.startCol, startCol)
      previous.endCol = Math.max(previous.endCol, endCol)
      previous.record = cloneRecord(
        {
          sheetName: record.range.sheetName,
          startAddress: formatAddress(previous.startRow, previous.startCol),
          endAddress: formatAddress(previous.endRow, previous.endCol),
        },
        previous.record,
      )
      return
    }
    merged.push({
      record: cloneRecord(
        {
          sheetName: record.range.sheetName,
          startAddress: formatAddress(startRow, startCol),
          endAddress: formatAddress(endRow, endCol),
        },
        record,
      ),
      startRow,
      endRow,
      startCol,
      endCol,
    })
  })

  return merged.map(({ record }) => record)
}
