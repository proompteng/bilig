import { ValueTag, type CellValue } from '@bilig/protocol'
import type { EngineRuntimeState } from '../runtime-state.js'
import { BLOCK_ROWS } from '../../sheet-grid.js'
import { addEngineCounter, type EngineCounters } from '../../perf/engine-counters.js'

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

interface RuntimeColumnPage {
  readonly rowStart: number
  readonly tags: Uint8Array
  readonly numbers: Float64Array
  readonly stringIds: Uint32Array
  readonly errors: Uint16Array
}

export interface RuntimeColumnOwner {
  readonly sheetName: string
  readonly col: number
  columnVersion: number
  structureVersion: number
  sheetColumnVersions: Uint32Array
  readonly pages: ReadonlyMap<number, RuntimeColumnPage>
}

export interface RuntimeColumnView {
  readonly owner: RuntimeColumnOwner
  readonly sheetName: string
  readonly rowStart: number
  readonly rowEnd: number
  readonly col: number
  readonly length: number
  readonly columnVersion: number
  readonly structureVersion: number
  readonly sheetColumnVersions: Uint32Array
  readonly readTagAt: (offset: number) => number
  readonly readNumberAt: (offset: number) => number
  readonly readStringIdAt: (offset: number) => number
  readonly readErrorAt: (offset: number) => number
  readonly readCellValueAt: (offset: number) => CellValue
}

export interface EngineRuntimeColumnStoreService {
  readonly getColumnOwner: (request: { sheetName: string; col: number }) => RuntimeColumnOwner
  readonly getColumnView: (request: { sheetName: string; rowStart: number; rowEnd: number; col: number }) => RuntimeColumnView
  readonly getColumnSlice: (request: { sheetName: string; rowStart: number; rowEnd: number; col: number }) => RuntimeColumnSlice
  readonly readCellValue: (sheetName: string, row: number, col: number) => CellValue
  readonly readRangeValues: (request: {
    sheetName: string
    rowStart: number
    rowEnd: number
    colStart: number
    colEnd: number
  }) => CellValue[]
  readonly readRangeValueMatrix: (request: {
    sheetName: string
    rowStart: number
    rowEnd: number
    colStart: number
    colEnd: number
  }) => CellValue[][]
  readonly normalizeStringId: (stringId: number) => string
  readonly normalizeLookupText: (value: Extract<CellValue, { tag: ValueTag.String }>) => string
}

function getColumnSliceCacheKey(sheetName: string, col: number, rowStart: number, rowEnd: number): string {
  return `${sheetName}\t${col}\t${rowStart}\t${rowEnd}`
}

function getColumnOwnerCacheKey(sheetName: string, col: number): string {
  return `${sheetName}\t${col}`
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
  readonly state: Pick<EngineRuntimeState, 'workbook' | 'strings'> & { counters?: EngineCounters }
}): EngineRuntimeColumnStoreService {
  const emptyColumnVersions = new Uint32Array(0)
  const normalizedStrings = new Map<number, string>()
  const columnOwnerCache = new Map<string, RuntimeColumnOwner>()
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

  const readOwnerPageEntry = (
    owner: RuntimeColumnOwner,
    absoluteRow: number,
  ): {
    rawTag: number
    number: number
    stringId: number
    error: number
  } => {
    const pageRowStart = Math.floor(absoluteRow / BLOCK_ROWS) * BLOCK_ROWS
    const page = owner.pages.get(pageRowStart)
    if (!page) {
      return {
        rawTag: ValueTag.Empty,
        number: 0,
        stringId: 0,
        error: 0,
      }
    }
    const localRow = absoluteRow - pageRowStart
    return {
      rawTag: page.tags[localRow] ?? ValueTag.Empty,
      number: page.numbers[localRow] ?? 0,
      stringId: page.stringIds[localRow] ?? 0,
      error: page.errors[localRow] ?? 0,
    }
  }

  const materializeCellValueFromOwner = (owner: RuntimeColumnOwner, absoluteRow: number): CellValue => {
    const entry = readOwnerPageEntry(owner, absoluteRow)
    const tag = decodeValueTag(entry.rawTag)
    switch (tag) {
      case ValueTag.Empty:
        return { tag: ValueTag.Empty }
      case ValueTag.Number:
        return { tag: ValueTag.Number, value: entry.number }
      case ValueTag.Boolean:
        return { tag: ValueTag.Boolean, value: entry.number !== 0 }
      case ValueTag.String:
        return {
          tag: ValueTag.String,
          value: entry.stringId === 0 ? '' : args.state.strings.get(entry.stringId),
          stringId: entry.stringId,
        }
      case ValueTag.Error:
        return { tag: ValueTag.Error, code: entry.error }
      default:
        return { tag: ValueTag.Empty }
    }
  }

  const buildColumnOwner = (request: { sheetName: string; col: number }): RuntimeColumnOwner => {
    if (args.state.counters) {
      addEngineCounter(args.state.counters, 'columnOwnerBuilds')
    }
    const sheet = args.state.workbook.getSheet(request.sheetName)
    const sheetColumnVersions = sheet?.columnVersions ?? emptyColumnVersions
    const structureVersion = sheet?.structureVersion ?? 0
    const pages = new Map<number, RuntimeColumnPage>()

    if (sheet) {
      sheet.logical.forEachVisibleColumnCellEntry(request.col, (cellIndex, row) => {
        const pageRowStart = Math.floor(row / BLOCK_ROWS) * BLOCK_ROWS
        let page = pages.get(pageRowStart)
        if (!page) {
          page = {
            rowStart: pageRowStart,
            tags: new Uint8Array(BLOCK_ROWS),
            numbers: new Float64Array(BLOCK_ROWS),
            stringIds: new Uint32Array(BLOCK_ROWS),
            errors: new Uint16Array(BLOCK_ROWS),
          }
          pages.set(pageRowStart, page)
        }
        const localRow = row - pageRowStart
        const tag = decodeValueTag(args.state.workbook.cellStore.tags[cellIndex])
        page.tags[localRow] = tag
        if (tag === ValueTag.Number || tag === ValueTag.Boolean) {
          const numeric = args.state.workbook.cellStore.numbers[cellIndex] ?? 0
          page.numbers[localRow] = Object.is(numeric, -0) ? 0 : numeric
          return
        }
        if (tag === ValueTag.String) {
          page.stringIds[localRow] = args.state.workbook.cellStore.stringIds[cellIndex] ?? 0
          return
        }
        if (tag === ValueTag.Error) {
          page.errors[localRow] = args.state.workbook.cellStore.errors[cellIndex] ?? 0
        }
      })
    }

    return {
      sheetName: request.sheetName,
      col: request.col,
      columnVersion: sheetColumnVersions[request.col] ?? 0,
      structureVersion,
      sheetColumnVersions,
      pages,
    }
  }

  const getColumnOwner = (request: { sheetName: string; col: number }): RuntimeColumnOwner => {
    const cacheKey = getColumnOwnerCacheKey(request.sheetName, request.col)
    const currentSheet = args.state.workbook.getSheet(request.sheetName)
    const currentSheetColumnVersions = currentSheet?.columnVersions ?? emptyColumnVersions
    const currentColumnVersion = currentSheetColumnVersions[request.col] ?? 0
    const currentStructureVersion = currentSheet?.structureVersion ?? 0
    let owner = columnOwnerCache.get(cacheKey)
    if (
      !owner ||
      owner.columnVersion !== currentColumnVersion ||
      owner.structureVersion !== currentStructureVersion ||
      owner.sheetColumnVersions !== currentSheetColumnVersions
    ) {
      owner = buildColumnOwner(request)
      columnOwnerCache.set(cacheKey, owner)
    }
    return owner
  }

  const getColumnView = (request: { sheetName: string; rowStart: number; rowEnd: number; col: number }): RuntimeColumnView => {
    const owner = getColumnOwner(request)
    return {
      owner,
      sheetName: request.sheetName,
      rowStart: request.rowStart,
      rowEnd: request.rowEnd,
      col: request.col,
      length: request.rowEnd - request.rowStart + 1,
      columnVersion: owner.columnVersion,
      structureVersion: owner.structureVersion,
      sheetColumnVersions: owner.sheetColumnVersions,
      readTagAt(offset) {
        return readOwnerPageEntry(owner, request.rowStart + offset).rawTag
      },
      readNumberAt(offset) {
        return readOwnerPageEntry(owner, request.rowStart + offset).number
      },
      readStringIdAt(offset) {
        return readOwnerPageEntry(owner, request.rowStart + offset).stringId
      },
      readErrorAt(offset) {
        return readOwnerPageEntry(owner, request.rowStart + offset).error
      },
      readCellValueAt(offset) {
        return materializeCellValueFromOwner(owner, request.rowStart + offset)
      },
    }
  }

  const buildColumnSlice = (request: { sheetName: string; rowStart: number; rowEnd: number; col: number }): RuntimeColumnSlice => {
    if (args.state.counters) {
      addEngineCounter(args.state.counters, 'columnSliceBuilds')
    }
    const view = getColumnView(request)
    const tags = new Uint8Array(view.length)
    const numbers = new Float64Array(view.length)
    const stringIds = new Uint32Array(view.length)
    const errors = new Uint16Array(view.length)
    for (let offset = 0; offset < view.length; offset += 1) {
      tags[offset] = view.readTagAt(offset)
      numbers[offset] = view.readNumberAt(offset)
      stringIds[offset] = view.readStringIdAt(offset)
      errors[offset] = view.readErrorAt(offset)
    }
    return {
      sheetName: request.sheetName,
      rowStart: request.rowStart,
      rowEnd: request.rowEnd,
      col: request.col,
      length: view.length,
      columnVersion: view.columnVersion,
      structureVersion: view.structureVersion,
      sheetColumnVersions: view.sheetColumnVersions,
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
    getColumnOwner,
    getColumnView,
    getColumnSlice,
    readCellValue(sheetName, row, col) {
      const view = getColumnView({ sheetName, rowStart: row, rowEnd: row, col })
      return view.readCellValueAt(0)
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
    readRangeValueMatrix({ sheetName, rowStart, rowEnd, colStart, colEnd }) {
      const width = colEnd - colStart + 1
      const height = rowEnd - rowStart + 1
      const rows: CellValue[][] = Array.from({ length: height }, () => [])
      for (let colOffset = 0; colOffset < width; colOffset += 1) {
        const slice = getColumnSlice({ sheetName, rowStart, rowEnd, col: colStart + colOffset })
        for (let rowOffset = 0; rowOffset < height; rowOffset += 1) {
          rows[rowOffset]![colOffset] = materializeCellValueFromSlice(slice, rowOffset)
        }
      }
      return rows
    },
    normalizeStringId,
    normalizeLookupText(value) {
      return (value.stringId !== undefined && value.stringId !== 0 ? normalizeStringId(value.stringId) : value.value).toUpperCase()
    },
  }
}
