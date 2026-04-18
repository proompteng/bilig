import { ValueTag, type CellValue } from '@bilig/protocol'
import { parseRangeAddress } from '@bilig/formula'
import type { EngineRuntimeState, PreparedExactVectorLookup } from '../runtime-state.js'
import type { EngineRuntimeColumnStoreService, RuntimeColumnView } from './runtime-column-store-service.js'
import {
  applyLookupColumnOwnerLiteralWrite,
  buildLookupColumnOwner,
  findExactMatchInRange,
  isLookupColumnOwner,
  summarizeExactRange,
  type LookupColumnOwner,
} from './lookup-column-owner.js'

export interface ExactVectorMatchRequest {
  lookupValue: CellValue
  sheetName: string
  start: string
  end: string
  startRow?: number
  endRow?: number
  startCol?: number
  endCol?: number
  searchMode: 1 | -1
}

export type ExactVectorMatchResult = { handled: false } | { handled: true; position: number | undefined }

export interface ExactColumnIndexService {
  readonly primeColumnIndex: (request: { sheetName: string; rowStart: number; rowEnd: number; col: number }) => void
  readonly prepareVectorLookup: (request: { sheetName: string; rowStart: number; rowEnd: number; col: number }) => PreparedExactVectorLookup
  readonly findPreparedVectorMatch: (request: {
    lookupValue: CellValue
    prepared: PreparedExactVectorLookup
    searchMode: 1 | -1
  }) => ExactVectorMatchResult
  readonly findVectorMatch: (request: ExactVectorMatchRequest) => ExactVectorMatchResult
  readonly invalidateColumn: (request: { sheetName: string; col: number }) => void
  readonly recordLiteralWrite: (request: {
    sheetName: string
    row: number
    col: number
    oldValue: CellValue
    newValue: CellValue
    oldStringId?: number
    newStringId?: number
  }) => void
}

interface ExactColumnIndexEntry {
  sheetName: string
  rowStart: number
  rowEnd: number
  col: number
  columnVersion: number
  structureVersion: number
  comparableKind: 'numeric' | 'text' | 'mixed'
  uniformStart: number | undefined
  uniformStep: number | undefined
  rowLists: Map<string, number[]>
  firstPositions: Map<string, number>
  lastPositions: Map<string, number>
  firstNumericPositions: Map<number, number> | undefined
  lastNumericPositions: Map<number, number> | undefined
  firstTextPositions: Map<string, number> | undefined
  lastTextPositions: Map<string, number> | undefined
}

interface ExactColumnBounds {
  rowStart: number
  rowEnd: number
  col: number
}

interface VectorLookupBoundsRequest {
  sheetName: string
  start: string
  end: string
  startRow?: number
  endRow?: number
  startCol?: number
  endCol?: number
}

function getExactColumnCacheKey(sheetName: string, col: number, rowStart: number, rowEnd: number): string {
  return `${sheetName}\t${col}\t${rowStart}\t${rowEnd}`
}

function columnRegistryKey(sheetName: string, col: number): string {
  return `${sheetName}\t${col}`
}

function normalizeExactLookupKey(value: CellValue, lookupString: (id: number) => string, stringId = 0): string | undefined {
  switch (value.tag) {
    case ValueTag.Empty:
      return 'e:'
    case ValueTag.Number:
      return `n:${Object.is(value.value, -0) ? 0 : value.value}`
    case ValueTag.Boolean:
      return value.value ? 'b:1' : 'b:0'
    case ValueTag.String:
      return `s:${(stringId !== 0 ? lookupString(stringId) : value.value).toUpperCase()}`
    case ValueTag.Error:
      return undefined
  }
}

function detectUniformNumericStep(values: Float64Array): { start: number; step: number } | undefined {
  if (values.length < 2) {
    return undefined
  }
  const start = values[0]!
  const step = values[1]! - start
  if (!Number.isFinite(step) || step === 0) {
    return undefined
  }
  for (let index = 2; index < values.length; index += 1) {
    if (values[index]! - values[index - 1]! !== step) {
      return undefined
    }
  }
  return { start, step }
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

function resolveExactColumnBounds(request: VectorLookupBoundsRequest): ExactColumnBounds | undefined {
  if (request.startRow !== undefined && request.endRow !== undefined && request.startCol !== undefined && request.endCol !== undefined) {
    if (request.startCol !== request.endCol) {
      return undefined
    }
    return {
      rowStart: request.startRow,
      rowEnd: request.endRow,
      col: request.startCol,
    }
  }

  const parsedRange = parseRangeAddress(`${request.start}:${request.end}`, request.sheetName)
  if (parsedRange.kind !== 'cells' || parsedRange.start.col !== parsedRange.end.col) {
    return undefined
  }
  return {
    rowStart: parsedRange.start.row,
    rowEnd: parsedRange.end.row,
    col: parsedRange.start.col,
  }
}

function setPositionsForKey(entry: ExactColumnIndexEntry, key: string): void {
  const rows = entry.rowLists.get(key)
  if (!rows || rows.length === 0) {
    entry.rowLists.delete(key)
    entry.firstPositions.delete(key)
    entry.lastPositions.delete(key)
    if (key.startsWith('n:')) {
      const numericValue = Number(key.slice(2))
      entry.firstNumericPositions?.delete(numericValue)
      entry.lastNumericPositions?.delete(numericValue)
    } else if (key.startsWith('s:')) {
      const textValue = key.slice(2)
      entry.firstTextPositions?.delete(textValue)
      entry.lastTextPositions?.delete(textValue)
    }
    return
  }
  entry.firstPositions.set(key, rows[0]!)
  entry.lastPositions.set(key, rows[rows.length - 1]!)
  if (key.startsWith('n:')) {
    const numericValue = Number(key.slice(2))
    entry.firstNumericPositions?.set(numericValue, rows[0]!)
    entry.lastNumericPositions?.set(numericValue, rows[rows.length - 1]!)
  } else if (key.startsWith('s:')) {
    const textValue = key.slice(2)
    entry.firstTextPositions?.set(textValue, rows[0]!)
    entry.lastTextPositions?.set(textValue, rows[rows.length - 1]!)
  }
}

export function createExactColumnIndexService(args: {
  readonly state: Pick<EngineRuntimeState, 'workbook' | 'strings'>
  readonly runtimeColumnStore: EngineRuntimeColumnStoreService
}): ExactColumnIndexService {
  const emptyColumnVersions = new Uint32Array(0)
  const exactColumnIndices = new Map<string, ExactColumnIndexEntry>()
  const cacheKeysByColumn = new Map<string, Set<string>>()
  const ownerIndices = new Map<string, LookupColumnOwner>()

  const readPreparedOwner = (prepared: PreparedExactVectorLookup): LookupColumnOwner | undefined =>
    isLookupColumnOwner(prepared.internalOwner) ? prepared.internalOwner : undefined

  const getCurrentColumnVersions = (
    sheetName: string,
    col: number,
  ): {
    columnVersion: number
    structureVersion: number
    sheetColumnVersions: Uint32Array
  } => {
    const sheet = args.state.workbook.getSheet(sheetName)
    const sheetColumnVersions = sheet?.columnVersions ?? emptyColumnVersions
    return {
      columnVersion: sheetColumnVersions[col] ?? 0,
      structureVersion: sheet?.structureVersion ?? 0,
      sheetColumnVersions,
    }
  }

  const trackCacheKey = (sheetName: string, col: number, cacheKey: string): void => {
    const registryKey = columnRegistryKey(sheetName, col)
    const existing = cacheKeysByColumn.get(registryKey)
    if (existing) {
      existing.add(cacheKey)
      return
    }
    cacheKeysByColumn.set(registryKey, new Set([cacheKey]))
  }

  const ensureOwnerIndex = (sheetName: string, col: number): LookupColumnOwner | undefined => {
    const registryKey = columnRegistryKey(sheetName, col)
    const currentVersions = getCurrentColumnVersions(sheetName, col)
    let owner = ownerIndices.get(registryKey)
    if (
      !owner ||
      owner.columnVersion !== currentVersions.columnVersion ||
      owner.structureVersion !== currentVersions.structureVersion ||
      owner.sheetColumnVersions !== currentVersions.sheetColumnVersions
    ) {
      owner = buildLookupColumnOwner({
        owner: args.runtimeColumnStore.getColumnOwner({ sheetName, col }),
        normalizeStringId: args.runtimeColumnStore.normalizeStringId,
      })
      if (owner) {
        ownerIndices.set(registryKey, owner)
      } else {
        ownerIndices.delete(registryKey)
      }
    }
    return owner
  }

  const untrackCacheKey = (sheetName: string, col: number, cacheKey: string): void => {
    const registryKey = columnRegistryKey(sheetName, col)
    const existing = cacheKeysByColumn.get(registryKey)
    if (!existing) {
      return
    }
    existing.delete(cacheKey)
    if (existing.size === 0) {
      cacheKeysByColumn.delete(registryKey)
    }
  }

  const replaceExactColumnIndex = (cacheKey: string, entry: ExactColumnIndexEntry | undefined): void => {
    const existing = exactColumnIndices.get(cacheKey)
    if (existing) {
      untrackCacheKey(existing.sheetName, existing.col, cacheKey)
      exactColumnIndices.delete(cacheKey)
    }
    if (!entry) {
      return
    }
    exactColumnIndices.set(cacheKey, entry)
    trackCacheKey(entry.sheetName, entry.col, cacheKey)
  }

  const keyAtOffset = (view: RuntimeColumnView, offset: number): string | undefined => {
    const tag = decodeValueTag(view.readTagAt(offset))
    switch (tag) {
      case ValueTag.Empty:
        return 'e:'
      case ValueTag.Number:
        return `n:${view.readNumberAt(offset)}`
      case ValueTag.Boolean:
        return view.readNumberAt(offset) !== 0 ? 'b:1' : 'b:0'
      case ValueTag.String: {
        const stringId = view.readStringIdAt(offset)
        return `s:${args.runtimeColumnStore.normalizeStringId(stringId)}`
      }
      case ValueTag.Error:
      default:
        return undefined
    }
  }

  const buildExactColumnIndex = (sheetName: string, col: number, rowStart: number, rowEnd: number): ExactColumnIndexEntry => {
    const view = args.runtimeColumnStore.getColumnView({
      sheetName,
      rowStart,
      rowEnd,
      col,
    })
    const firstPositions = new Map<string, number>()
    const lastPositions = new Map<string, number>()
    const firstNumericPositions = new Map<number, number>()
    const lastNumericPositions = new Map<number, number>()
    const firstTextPositions = new Map<string, number>()
    const lastTextPositions = new Map<string, number>()
    const rowLists = new Map<string, number[]>()
    const numericSequence: number[] = []
    let sawNumeric = false
    let sawText = false
    let sawOther = false
    for (let offset = 0; offset < view.length; offset += 1) {
      const row = rowStart + offset
      const key = keyAtOffset(view, offset)
      if (key === undefined) {
        continue
      }
      if (!firstPositions.has(key)) {
        firstPositions.set(key, row)
      }
      lastPositions.set(key, row)
      const existingRows = rowLists.get(key)
      if (existingRows) {
        existingRows.push(row)
      } else {
        rowLists.set(key, [row])
      }
      if (key.startsWith('n:')) {
        const numericValue = Number(key.slice(2))
        if (!firstNumericPositions.has(numericValue)) {
          firstNumericPositions.set(numericValue, row)
        }
        lastNumericPositions.set(numericValue, row)
        numericSequence.push(numericValue)
        sawNumeric = true
        continue
      }
      if (key.startsWith('s:')) {
        const textValue = key.slice(2)
        if (!firstTextPositions.has(textValue)) {
          firstTextPositions.set(textValue, row)
        }
        lastTextPositions.set(textValue, row)
        sawText = true
        continue
      }
      sawOther = true
    }
    const comparableKind = sawOther || (sawNumeric && sawText) ? 'mixed' : sawNumeric ? 'numeric' : sawText ? 'text' : 'mixed'
    const uniformNumericStep = comparableKind === 'numeric' ? detectUniformNumericStep(Float64Array.from(numericSequence)) : undefined
    return {
      sheetName,
      rowStart,
      rowEnd,
      col,
      columnVersion: view.columnVersion,
      structureVersion: view.structureVersion,
      comparableKind,
      uniformStart: uniformNumericStep?.start,
      uniformStep: uniformNumericStep?.step,
      rowLists,
      firstPositions,
      lastPositions,
      firstNumericPositions: comparableKind === 'numeric' ? firstNumericPositions : undefined,
      lastNumericPositions: comparableKind === 'numeric' ? lastNumericPositions : undefined,
      firstTextPositions: comparableKind === 'text' ? firstTextPositions : undefined,
      lastTextPositions: comparableKind === 'text' ? lastTextPositions : undefined,
    }
  }

  const ensureExactColumnIndex = (sheetName: string, col: number, rowStart: number, rowEnd: number): ExactColumnIndexEntry => {
    const cacheKey = getExactColumnCacheKey(sheetName, col, rowStart, rowEnd)
    const currentVersions = getCurrentColumnVersions(sheetName, col)
    let entry = exactColumnIndices.get(cacheKey)
    if (!entry || entry.columnVersion !== currentVersions.columnVersion || entry.structureVersion !== currentVersions.structureVersion) {
      entry = buildExactColumnIndex(sheetName, col, rowStart, rowEnd)
      replaceExactColumnIndex(cacheKey, entry)
    }
    return entry
  }

  const updateEntryLiteralWrite = (
    entry: ExactColumnIndexEntry,
    row: number,
    oldKey: string | undefined,
    newKey: string | undefined,
    currentColumnVersion: number,
    currentStructureVersion: number,
  ): boolean => {
    entry.columnVersion = currentColumnVersion
    entry.structureVersion = currentStructureVersion
    if (row < entry.rowStart || row > entry.rowEnd) {
      return true
    }
    if (oldKey === newKey) {
      return true
    }
    if (entry.comparableKind === 'numeric') {
      const oldNumeric = oldKey?.startsWith('n:') ?? false
      const newNumeric = newKey?.startsWith('n:') ?? false
      if ((oldKey !== undefined && !oldNumeric) || (newKey !== undefined && !newNumeric)) {
        return false
      }
      entry.uniformStart = undefined
      entry.uniformStep = undefined
    } else if (entry.comparableKind === 'text') {
      const oldText = oldKey?.startsWith('s:') ?? false
      const newText = newKey?.startsWith('s:') ?? false
      if ((oldKey !== undefined && !oldText) || (newKey !== undefined && !newText)) {
        return false
      }
    }
    if (oldKey !== undefined) {
      const rows = entry.rowLists.get(oldKey)
      if (!rows) {
        return false
      }
      const rowIndex = rows.indexOf(row)
      if (rowIndex === -1) {
        return false
      }
      rows.splice(rowIndex, 1)
      setPositionsForKey(entry, oldKey)
    }
    if (newKey !== undefined) {
      const rows = entry.rowLists.get(newKey)
      if (rows) {
        let insertIndex = rows.length
        while (insertIndex > 0 && rows[insertIndex - 1]! > row) {
          insertIndex -= 1
        }
        rows.splice(insertIndex, 0, row)
      } else {
        entry.rowLists.set(newKey, [row])
      }
      setPositionsForKey(entry, newKey)
    }
    return true
  }

  const prepareVectorLookup = (request: {
    sheetName: string
    rowStart: number
    rowEnd: number
    col: number
  }): PreparedExactVectorLookup => {
    const owner = ensureOwnerIndex(request.sheetName, request.col)
    if (owner && request.rowStart >= owner.rowStart && request.rowEnd <= owner.rowEnd) {
      const summary = summarizeExactRange(owner, request.rowStart, request.rowEnd)
      return {
        sheetName: request.sheetName,
        rowStart: request.rowStart,
        rowEnd: request.rowEnd,
        col: request.col,
        length: request.rowEnd - request.rowStart + 1,
        columnVersion: owner.columnVersion,
        structureVersion: owner.structureVersion,
        sheetColumnVersions: owner.sheetColumnVersions,
        comparableKind: summary?.comparableKind ?? 'mixed',
        uniformStart: summary?.uniformStart,
        uniformStep: summary?.uniformStep,
        firstPositions: new Map(),
        lastPositions: new Map(),
        firstNumericPositions: undefined,
        lastNumericPositions: undefined,
        firstTextPositions: undefined,
        lastTextPositions: undefined,
        internalOwner: owner,
      }
    }
    const entry = ensureExactColumnIndex(request.sheetName, request.col, request.rowStart, request.rowEnd)
    return {
      sheetName: request.sheetName,
      rowStart: request.rowStart,
      rowEnd: request.rowEnd,
      col: request.col,
      length: request.rowEnd - request.rowStart + 1,
      columnVersion: entry.columnVersion,
      structureVersion: entry.structureVersion,
      sheetColumnVersions: getCurrentColumnVersions(request.sheetName, request.col).sheetColumnVersions,
      comparableKind: entry.comparableKind,
      uniformStart: entry.uniformStart,
      uniformStep: entry.uniformStep,
      firstPositions: entry.firstPositions,
      lastPositions: entry.lastPositions,
      firstNumericPositions: entry.firstNumericPositions,
      lastNumericPositions: entry.lastNumericPositions,
      firstTextPositions: entry.firstTextPositions,
      lastTextPositions: entry.lastTextPositions,
      internalOwner: undefined,
    }
  }

  const refreshPreparedVectorLookup = (prepared: PreparedExactVectorLookup): PreparedExactVectorLookup => {
    const owner = readPreparedOwner(prepared)
    if (owner) {
      const currentVersions = getCurrentColumnVersions(prepared.sheetName, prepared.col)
      if (currentVersions.columnVersion === prepared.columnVersion && currentVersions.structureVersion === prepared.structureVersion) {
        return prepared
      }
      const refreshedOwner = ensureOwnerIndex(prepared.sheetName, prepared.col)
      if (refreshedOwner && prepared.rowStart >= refreshedOwner.rowStart && prepared.rowEnd <= refreshedOwner.rowEnd) {
        const summary = summarizeExactRange(refreshedOwner, prepared.rowStart, prepared.rowEnd)
        prepared.length = prepared.rowEnd - prepared.rowStart + 1
        prepared.columnVersion = refreshedOwner.columnVersion
        prepared.structureVersion = refreshedOwner.structureVersion
        prepared.sheetColumnVersions = refreshedOwner.sheetColumnVersions
        prepared.comparableKind = summary?.comparableKind ?? 'mixed'
        prepared.uniformStart = summary?.uniformStart
        prepared.uniformStep = summary?.uniformStep
        prepared.internalOwner = refreshedOwner
        return prepared
      }
    }
    const currentVersions = getCurrentColumnVersions(prepared.sheetName, prepared.col)
    if (currentVersions.columnVersion === prepared.columnVersion && currentVersions.structureVersion === prepared.structureVersion) {
      return prepared
    }
    const refreshed = prepareVectorLookup(prepared)
    prepared.length = refreshed.length
    prepared.columnVersion = refreshed.columnVersion
    prepared.structureVersion = refreshed.structureVersion
    prepared.sheetColumnVersions = refreshed.sheetColumnVersions
    prepared.comparableKind = refreshed.comparableKind
    prepared.uniformStart = refreshed.uniformStart
    prepared.uniformStep = refreshed.uniformStep
    prepared.firstPositions = refreshed.firstPositions
    prepared.lastPositions = refreshed.lastPositions
    prepared.firstNumericPositions = refreshed.firstNumericPositions
    prepared.lastNumericPositions = refreshed.lastNumericPositions
    prepared.firstTextPositions = refreshed.firstTextPositions
    prepared.lastTextPositions = refreshed.lastTextPositions
    prepared.internalOwner = refreshed.internalOwner
    return prepared
  }

  const findPreparedVectorMatch = (request: {
    lookupValue: CellValue
    prepared: PreparedExactVectorLookup
    searchMode: 1 | -1
  }): ExactVectorMatchResult => {
    const prepared = refreshPreparedVectorLookup(request.prepared)
    const owner = readPreparedOwner(prepared)
    if (owner) {
      const normalizedLookupKey = normalizeExactLookupKey(
        request.lookupValue,
        (id) => args.state.strings.get(id),
        request.lookupValue.tag === ValueTag.String ? request.lookupValue.stringId : 0,
      )
      if (normalizedLookupKey === undefined) {
        return { handled: false }
      }
      const row = findExactMatchInRange(owner, normalizedLookupKey, prepared.rowStart, prepared.rowEnd, request.searchMode)
      return {
        handled: true,
        position: row === undefined ? undefined : row - prepared.rowStart + 1,
      }
    }
    if (prepared.comparableKind === 'numeric') {
      if (request.lookupValue.tag === ValueTag.Error) {
        return { handled: false }
      }
      if (request.lookupValue.tag !== ValueTag.Number) {
        return { handled: true, position: undefined }
      }
      const numericValue = Object.is(request.lookupValue.value, -0) ? 0 : request.lookupValue.value
      if (prepared.uniformStart !== undefined && prepared.uniformStep !== undefined) {
        const relative = (numericValue - prepared.uniformStart) / prepared.uniformStep
        const position = Number.isInteger(relative) ? relative + 1 : undefined
        return {
          handled: true,
          position: position !== undefined && position >= 1 && position <= prepared.length ? position : undefined,
        }
      }
      const numericMap = request.searchMode === -1 ? prepared.lastNumericPositions : prepared.firstNumericPositions
      const row = numericMap?.get(numericValue)
      return {
        handled: true,
        position: row === undefined ? undefined : row - prepared.rowStart + 1,
      }
    }
    if (prepared.comparableKind === 'text') {
      if (request.lookupValue.tag === ValueTag.Error) {
        return { handled: false }
      }
      if (request.lookupValue.tag !== ValueTag.String) {
        return { handled: true, position: undefined }
      }
      const textValue = args.runtimeColumnStore.normalizeLookupText(request.lookupValue)
      const textMap = request.searchMode === -1 ? prepared.lastTextPositions : prepared.firstTextPositions
      const row = textMap?.get(textValue)
      return {
        handled: true,
        position: row === undefined ? undefined : row - prepared.rowStart + 1,
      }
    }
    const normalizedLookupKey = normalizeExactLookupKey(request.lookupValue, (id) => args.state.strings.get(id))
    if (normalizedLookupKey === undefined) {
      return { handled: false }
    }
    const row =
      request.searchMode === -1 ? prepared.lastPositions.get(normalizedLookupKey) : prepared.firstPositions.get(normalizedLookupKey)
    return {
      handled: true,
      position: row === undefined ? undefined : row - prepared.rowStart + 1,
    }
  }

  return {
    primeColumnIndex(request) {
      ensureExactColumnIndex(request.sheetName, request.col, request.rowStart, request.rowEnd)
    },
    prepareVectorLookup(request) {
      return prepareVectorLookup(request)
    },
    findPreparedVectorMatch(request) {
      return findPreparedVectorMatch(request)
    },
    findVectorMatch(request) {
      const bounds = resolveExactColumnBounds(request)
      if (!bounds) {
        return { handled: false }
      }
      const prepared = prepareVectorLookup({
        sheetName: request.sheetName,
        rowStart: bounds.rowStart,
        rowEnd: bounds.rowEnd,
        col: bounds.col,
      })
      return findPreparedVectorMatch({
        lookupValue: request.lookupValue,
        prepared,
        searchMode: request.searchMode,
      })
    },
    /* c8 ignore start */
    invalidateColumn(request) {
      ownerIndices.delete(columnRegistryKey(request.sheetName, request.col))
      const cacheKeys = cacheKeysByColumn.get(columnRegistryKey(request.sheetName, request.col))
      if (!cacheKeys) {
        return
      }
      for (const cacheKey of cacheKeys) {
        replaceExactColumnIndex(cacheKey, undefined)
      }
    },
    recordLiteralWrite(request) {
      const registryKey = columnRegistryKey(request.sheetName, request.col)
      const owner = ownerIndices.get(registryKey)
      if (owner) {
        const currentVersions = getCurrentColumnVersions(request.sheetName, request.col)
        owner.columnVersion = currentVersions.columnVersion
        owner.structureVersion = currentVersions.structureVersion
        owner.sheetColumnVersions = currentVersions.sheetColumnVersions
        if (
          !applyLookupColumnOwnerLiteralWrite({
            owner,
            write: request,
            normalizeStringId: args.runtimeColumnStore.normalizeStringId,
          })
        ) {
          ownerIndices.delete(registryKey)
        }
      }
      const cacheKeys = cacheKeysByColumn.get(registryKey)
      if (!cacheKeys || cacheKeys.size === 0) {
        return
      }
      const sheet = args.state.workbook.getSheet(request.sheetName)
      if (!sheet) {
        return
      }
      const currentColumnVersion = sheet.columnVersions[request.col] ?? 0
      const currentStructureVersion = args.state.workbook.getSheetStructureVersion(request.sheetName)
      const oldKey = normalizeExactLookupKey(request.oldValue, (id) => args.state.strings.get(id), request.oldStringId)
      const newKey = normalizeExactLookupKey(request.newValue, (id) => args.state.strings.get(id), request.newStringId)
      for (const cacheKey of cacheKeys) {
        const entry = exactColumnIndices.get(cacheKey)
        if (!entry) {
          untrackCacheKey(request.sheetName, request.col, cacheKey)
          continue
        }
        if (entry.structureVersion !== currentStructureVersion) {
          replaceExactColumnIndex(cacheKey, undefined)
          continue
        }
        if (!updateEntryLiteralWrite(entry, request.row, oldKey, newKey, currentColumnVersion, currentStructureVersion)) {
          replaceExactColumnIndex(cacheKey, undefined)
        }
      }
    },
    /* c8 ignore stop */
  }
}
