import { ValueTag, type CellValue } from '@bilig/protocol'
import type { EngineRuntimeState } from '../runtime-state.js'

export interface RuntimeColumnSlice {
  readonly sheetName: string
  readonly rowStart: number
  readonly rowEnd: number
  readonly col: number
  readonly length: number
  columnVersion: number
  structureVersion: number
  sheetColumnVersions: Uint32Array
  tags: Uint8Array
  numbers: Float64Array
  stringIds: Uint32Array
  errors: Uint16Array
}

export interface EngineRuntimeColumnStoreService {
  readonly getColumnSlice: (request: { sheetName: string; rowStart: number; rowEnd: number; col: number }) => RuntimeColumnSlice
  readonly readCellValue: (sheetName: string, row: number, col: number) => CellValue
  readonly readRangeValues: (request: {
    sheetName: string
    rowStart: number
    rowEnd: number
    colStart: number
    colEnd: number
  }) => CellValue[]
  readonly normalizeStringId: (stringId: number) => string
  readonly normalizeLookupText: (value: Extract<CellValue, { tag: ValueTag.String }>) => string
}

function getColumnSliceCacheKey(sheetName: string, col: number, rowStart: number, rowEnd: number): string {
  return `${sheetName}\t${col}\t${rowStart}\t${rowEnd}`
}

function decodeValueTag(rawTag: number | undefined): ValueTag {
  if (rawTag === undefined) {
    return ValueTag.Empty
  }
  switch (rawTag) {
    case 1:
      return ValueTag.Number
    case 2:
      return ValueTag.Boolean
    case 3:
      return ValueTag.String
    case 4:
      return ValueTag.Error
    case 0:
    default:
      return ValueTag.Empty
  }
}

export function createEngineRuntimeColumnStoreService(args: {
  readonly state: Pick<EngineRuntimeState, 'workbook' | 'strings'>
}): EngineRuntimeColumnStoreService {
  const emptyColumnVersions = new Uint32Array(0)
  const normalizedStrings = new Map<number, string>()
  const columnSliceCache = new Map<string, RuntimeColumnSlice>()

  const normalizeStringId = (stringId: number): string => {
    let normalized = normalizedStrings.get(stringId)
    if (normalized === undefined) {
      normalized = args.state.strings.get(stringId).toUpperCase()
      normalizedStrings.set(stringId, normalized)
    }
    return normalized
  }

  const materializeCellValueFromSlice = (slice: RuntimeColumnSlice, offset: number): CellValue => {
    const tag = decodeValueTag(slice.tags[offset])
    switch (tag) {
      case ValueTag.Empty:
        return { tag: ValueTag.Empty }
      case ValueTag.Number:
        return { tag: ValueTag.Number, value: slice.numbers[offset] ?? 0 }
      case ValueTag.Boolean:
        return { tag: ValueTag.Boolean, value: (slice.numbers[offset] ?? 0) !== 0 }
      case ValueTag.String: {
        const stringId = slice.stringIds[offset] ?? 0
        return {
          tag: ValueTag.String,
          value: stringId === 0 ? '' : args.state.strings.get(stringId),
          stringId,
        }
      }
      case ValueTag.Error: {
        return { tag: ValueTag.Error, code: slice.errors[offset]! }
      }
      default:
        return { tag: ValueTag.Empty }
    }
  }

  const buildColumnSlice = (request: { sheetName: string; rowStart: number; rowEnd: number; col: number }): RuntimeColumnSlice => {
    const sheet = args.state.workbook.getSheet(request.sheetName)
    const sheetColumnVersions = sheet?.columnVersions ?? emptyColumnVersions
    const structureVersion = sheet?.structureVersion ?? 0
    const length = request.rowEnd - request.rowStart + 1
    const tags = new Uint8Array(length)
    const numbers = new Float64Array(length)
    const stringIds = new Uint32Array(length)
    const errors = new Uint16Array(length)

    if (sheet) {
      for (let offset = 0; offset < length; offset += 1) {
        const row = request.rowStart + offset
        const cellIndex = sheet.grid.get(row, request.col)
        if (cellIndex === -1) {
          tags[offset] = ValueTag.Empty
          continue
        }
        const tag = decodeValueTag(args.state.workbook.cellStore.tags[cellIndex])
        tags[offset] = tag
        if (tag === ValueTag.Number || tag === ValueTag.Boolean) {
          const numeric = args.state.workbook.cellStore.numbers[cellIndex] ?? 0
          numbers[offset] = Object.is(numeric, -0) ? 0 : numeric
          continue
        }
        if (tag === ValueTag.String) {
          stringIds[offset] = args.state.workbook.cellStore.stringIds[cellIndex] ?? 0
          continue
        }
        if (tag === ValueTag.Error) {
          errors[offset] = args.state.workbook.cellStore.errors[cellIndex] ?? 0
        }
      }
    }

    return {
      sheetName: request.sheetName,
      rowStart: request.rowStart,
      rowEnd: request.rowEnd,
      col: request.col,
      length,
      columnVersion: sheetColumnVersions[request.col] ?? 0,
      structureVersion,
      sheetColumnVersions,
      tags,
      numbers,
      stringIds,
      errors,
    }
  }

  const getColumnSlice = (request: { sheetName: string; rowStart: number; rowEnd: number; col: number }): RuntimeColumnSlice => {
    const cacheKey = getColumnSliceCacheKey(request.sheetName, request.col, request.rowStart, request.rowEnd)
    const currentSheet = args.state.workbook.getSheet(request.sheetName)
    const currentSheetColumnVersions = currentSheet?.columnVersions ?? emptyColumnVersions
    const currentColumnVersion = currentSheetColumnVersions[request.col] ?? 0
    const currentStructureVersion = currentSheet?.structureVersion ?? 0
    let slice = columnSliceCache.get(cacheKey)
    if (
      !slice ||
      slice.columnVersion !== currentColumnVersion ||
      slice.structureVersion !== currentStructureVersion ||
      slice.sheetColumnVersions !== currentSheetColumnVersions
    ) {
      slice = buildColumnSlice(request)
      columnSliceCache.set(cacheKey, slice)
    }
    return slice
  }

  return {
    getColumnSlice,
    readCellValue(sheetName, row, col) {
      const slice = getColumnSlice({ sheetName, rowStart: row, rowEnd: row, col })
      return materializeCellValueFromSlice(slice, 0)
    },
    readRangeValues({ sheetName, rowStart, rowEnd, colStart, colEnd }) {
      const columnSlices: RuntimeColumnSlice[] = []
      for (let col = colStart; col <= colEnd; col += 1) {
        columnSlices.push(getColumnSlice({ sheetName, rowStart, rowEnd, col }))
      }
      const values: CellValue[] = []
      for (let rowOffset = 0; rowOffset <= rowEnd - rowStart; rowOffset += 1) {
        for (let colOffset = 0; colOffset < columnSlices.length; colOffset += 1) {
          values.push(materializeCellValueFromSlice(columnSlices[colOffset]!, rowOffset))
        }
      }
      return values
    },
    normalizeStringId,
    normalizeLookupText(value) {
      return (value.stringId !== undefined && value.stringId !== 0 ? normalizeStringId(value.stringId) : value.value).toUpperCase()
    },
  }
}
