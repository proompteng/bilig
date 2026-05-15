import type { CellRangeRef } from '@bilig/protocol'
import { formatAddress, parseCellAddress } from '@bilig/formula'

export type WorkbookRangeRecord = {
  range: CellRangeRef
}

export type NormalizedWorkbookRangeRef = CellRangeRef & {
  startRow: number
  endRow: number
  startCol: number
  endCol: number
}

type NormalizedWorkbookRangeRecord = NormalizedWorkbookRangeRef & {
  recordIndex: number
}

const normalizedRecordsByRangeArray = new WeakMap<readonly WorkbookRangeRecord[], readonly NormalizedWorkbookRangeRecord[]>()

export function cloneWorkbookRangeRecords<RecordType extends WorkbookRangeRecord>(
  records: readonly RecordType[],
  cloneRecord: (range: CellRangeRef, record: RecordType) => RecordType,
): RecordType[] {
  return records.map((record) => cloneRecord(canonicalWorkbookRangeRef(record.range), record))
}

export function overlayWorkbookRangeRecords<RecordType extends WorkbookRangeRecord>(
  records: readonly RecordType[],
  nextRecord: RecordType,
  cloneRecord: (range: CellRangeRef, record: RecordType) => RecordType,
  isDefaultRecord: (record: RecordType) => boolean,
): RecordType[] {
  const normalizedNext = cloneRecord(canonicalWorkbookRangeRef(nextRecord.range), nextRecord)
  const remainders = records.flatMap((record) =>
    subtractWorkbookRangeRecord(record.range, normalizedNext.range).map((remainder) => cloneRecord(remainder, record)),
  )
  return isDefaultRecord(normalizedNext) ? remainders : [...remainders, normalizedNext]
}

export function replaceWorkbookRangeRecords<RecordType extends WorkbookRangeRecord>(
  records: readonly RecordType[],
  cloneRecord: (range: CellRangeRef, record: RecordType) => RecordType,
  validateRecord: (record: RecordType) => void,
): RecordType[] {
  const nextRecords = records.map((record) => cloneRecord(canonicalWorkbookRangeRef(record.range), record))
  nextRecords.forEach((record) => {
    validateRecord(record)
  })
  return nextRecords
}

export function findWorkbookRangeRecord<RecordType extends WorkbookRangeRecord>(
  records: readonly RecordType[],
  row: number,
  col: number,
): RecordType | undefined {
  const normalizedRecords = getNormalizedWorkbookRangeRecords(records)
  for (let index = normalizedRecords.length - 1; index >= 0; index -= 1) {
    const normalized = normalizedRecords[index]!
    if (normalized.startRow <= row && normalized.endRow >= row && normalized.startCol <= col && normalized.endCol >= col) {
      return records[normalized.recordIndex]!
    }
  }
  return undefined
}

function getNormalizedWorkbookRangeRecords(records: readonly WorkbookRangeRecord[]): readonly NormalizedWorkbookRangeRecord[] {
  const cached = normalizedRecordsByRangeArray.get(records)
  if (cached && cached.length === records.length) {
    return cached
  }
  const normalizedRecords = records.map((record, recordIndex) => ({
    ...normalizeWorkbookRangeRef(record.range),
    recordIndex,
  }))
  normalizedRecordsByRangeArray.set(records, normalizedRecords)
  return normalizedRecords
}

function normalizeWorkbookRangeRef(range: CellRangeRef): NormalizedWorkbookRangeRef {
  const start = parseCellAddress(range.startAddress, range.sheetName)
  const end = parseCellAddress(range.endAddress, range.sheetName)
  const startRow = Math.min(start.row, end.row)
  const endRow = Math.max(start.row, end.row)
  const startCol = Math.min(start.col, end.col)
  const endCol = Math.max(start.col, end.col)
  return {
    sheetName: range.sheetName,
    startAddress: formatAddress(startRow, startCol),
    endAddress: formatAddress(endRow, endCol),
    startRow,
    endRow,
    startCol,
    endCol,
  }
}

export function canonicalWorkbookRangeRef(range: CellRangeRef): CellRangeRef {
  return toPlainWorkbookRangeRef(normalizeWorkbookRangeRef(range))
}

export function canonicalWorkbookAddress(sheetName: string, address: string): string {
  const parsed = parseCellAddress(address, sheetName)
  return formatAddress(parsed.row, parsed.col)
}

function subtractWorkbookRangeRecord(existing: CellRangeRef, cut: CellRangeRef): CellRangeRef[] {
  if (existing.sheetName !== cut.sheetName) {
    return [toPlainWorkbookRangeRef(existing)]
  }
  const source = normalizeWorkbookRangeRef(existing)
  const removal = normalizeWorkbookRangeRef(cut)
  const startRow = Math.max(source.startRow, removal.startRow)
  const endRow = Math.min(source.endRow, removal.endRow)
  const startCol = Math.max(source.startCol, removal.startCol)
  const endCol = Math.min(source.endCol, removal.endCol)
  if (startRow > endRow || startCol > endCol) {
    return [toPlainWorkbookRangeRef(source)]
  }

  const remainders: CellRangeRef[] = []
  const pushRemainder = (rowStart: number, rowEnd: number, colStart: number, colEnd: number) => {
    if (rowStart > rowEnd || colStart > colEnd) {
      return
    }
    remainders.push({
      sheetName: source.sheetName,
      startAddress: formatAddress(rowStart, colStart),
      endAddress: formatAddress(rowEnd, colEnd),
    })
  }

  pushRemainder(source.startRow, startRow - 1, source.startCol, source.endCol)
  pushRemainder(endRow + 1, source.endRow, source.startCol, source.endCol)
  pushRemainder(startRow, endRow, source.startCol, startCol - 1)
  pushRemainder(startRow, endRow, endCol + 1, source.endCol)

  return remainders
}

function toPlainWorkbookRangeRef(range: CellRangeRef): CellRangeRef {
  return {
    sheetName: range.sheetName,
    startAddress: range.startAddress,
    endAddress: range.endAddress,
  }
}
