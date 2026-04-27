import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import type { WorkbookStore } from '../workbook-store.js'
import type { EngineRuntimeColumnStoreService, RuntimeColumnView } from '../engine/services/runtime-column-store-service.js'
import type { RegionGraph } from './region-graph.js'
import type { RegionId } from './region-node-store.js'

const PREFIX_LITERAL_DELTA_MAX_SUFFIX_LENGTH = 16_384

export interface AggregateStateEntry {
  readonly regionId: RegionId
  readonly sheetName: string
  readonly col: number
  readonly rowStart: number
  rowEnd: number
  columnVersion: number
  structureVersion: number
  extremaValid: boolean
  prefixSums: Float64Array
  prefixCount: Uint32Array
  prefixAverageCount: Uint32Array
  prefixErrorCodes: Uint16Array
  prefixErrorCounts: Uint32Array
  prefixMinimums: Float64Array
  prefixMaximums: Float64Array
}

export interface AggregateStateStore {
  readonly getOrBuildPrefixForRegion: (
    regionId: RegionId,
    aggregateKind?: 'sum' | 'average' | 'count' | 'min' | 'max',
  ) => AggregateStateEntry
  readonly noteLiteralWrite: (request: {
    readonly sheetName: string
    readonly row: number
    readonly col: number
    readonly oldValue: CellValue
    readonly newValue: CellValue
  }) => void
  readonly invalidateColumn: (sheetName: string, col: number) => void
}

function columnKeyPrefix(sheetName: string, col: number): string {
  return `${sheetName}\t${col}`
}

function cacheKey(sheetName: string, col: number, rowStart: number): string {
  return `${columnKeyPrefix(sheetName, col)}\t${rowStart}`
}

function prefixAnchorForRegion(regionRowStart: number, aggregateKind?: 'sum' | 'average' | 'count' | 'min' | 'max'): number {
  return aggregateKind === 'count' ? 0 : regionRowStart
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

function ensureCapacity(entry: AggregateStateEntry, totalLength: number): void {
  if (entry.prefixSums.length >= totalLength) {
    return
  }
  const nextCapacity = Math.max(entry.prefixSums.length * 2, totalLength)
  const nextPrefixSums = new Float64Array(nextCapacity)
  const nextPrefixCount = new Uint32Array(nextCapacity)
  const nextPrefixAverageCount = new Uint32Array(nextCapacity)
  const nextPrefixErrorCodes = new Uint16Array(nextCapacity)
  const nextPrefixErrorCounts = new Uint32Array(nextCapacity)
  const nextPrefixMinimums = new Float64Array(nextCapacity)
  const nextPrefixMaximums = new Float64Array(nextCapacity)
  const currentLength = entry.rowEnd - entry.rowStart + 1
  nextPrefixSums.set(entry.prefixSums.subarray(0, currentLength), 0)
  nextPrefixCount.set(entry.prefixCount.subarray(0, currentLength), 0)
  nextPrefixAverageCount.set(entry.prefixAverageCount.subarray(0, currentLength), 0)
  nextPrefixErrorCodes.set(entry.prefixErrorCodes.subarray(0, currentLength), 0)
  nextPrefixErrorCounts.set(entry.prefixErrorCounts.subarray(0, currentLength), 0)
  nextPrefixMinimums.set(entry.prefixMinimums.subarray(0, currentLength), 0)
  nextPrefixMaximums.set(entry.prefixMaximums.subarray(0, currentLength), 0)
  entry.prefixSums = nextPrefixSums
  entry.prefixCount = nextPrefixCount
  entry.prefixAverageCount = nextPrefixAverageCount
  entry.prefixErrorCodes = nextPrefixErrorCodes
  entry.prefixErrorCounts = nextPrefixErrorCounts
  entry.prefixMinimums = nextPrefixMinimums
  entry.prefixMaximums = nextPrefixMaximums
}

function isErrorTag(value: CellValue): boolean {
  return value.tag === ValueTag.Error && value.code !== ErrorCode.None
}

function numericContribution(value: CellValue): number {
  switch (value.tag) {
    case ValueTag.Number:
      return value.value
    case ValueTag.Boolean:
      return value.value ? 1 : 0
    case ValueTag.Empty:
    case ValueTag.String:
    case ValueTag.Error:
    default:
      return 0
  }
}

function countContribution(value: CellValue): number {
  return value.tag === ValueTag.Number || value.tag === ValueTag.Boolean ? 1 : 0
}

function averageCountContribution(value: CellValue): number {
  return value.tag === ValueTag.Number || value.tag === ValueTag.Boolean || value.tag === ValueTag.Empty ? 1 : 0
}

export function createAggregateStateStore(args: {
  readonly workbook: Pick<WorkbookStore, 'getSheet'>
  readonly runtimeColumnStore: EngineRuntimeColumnStoreService
  readonly regionGraph: Pick<RegionGraph, 'getRegion'>
}): AggregateStateStore {
  const cache = new Map<string, AggregateStateEntry>()
  const entriesByColumn = new Map<string, AggregateStateEntry[]>()

  const registerEntry = (entry: AggregateStateEntry): void => {
    const key = columnKeyPrefix(entry.sheetName, entry.col)
    let entries = entriesByColumn.get(key)
    if (!entries) {
      entries = []
      entriesByColumn.set(key, entries)
    }
    const insertAt = entries.findIndex((candidate) => candidate.rowStart > entry.rowStart)
    if (insertAt === -1) {
      entries.push(entry)
    } else {
      entries.splice(insertAt, 0, entry)
    }
  }

  const deleteEntry = (entry: AggregateStateEntry): void => {
    cache.delete(cacheKey(entry.sheetName, entry.col, entry.rowStart))
    const entries = entriesByColumn.get(columnKeyPrefix(entry.sheetName, entry.col))
    if (!entries) {
      return
    }
    const index = entries.indexOf(entry)
    if (index !== -1) {
      entries.splice(index, 1)
    }
    if (entries.length === 0) {
      entriesByColumn.delete(columnKeyPrefix(entry.sheetName, entry.col))
    }
  }

  const getCurrentVersions = (sheetName: string, col: number) => {
    const sheet = args.workbook.getSheet(sheetName)
    return {
      columnVersion: sheet?.columnVersions[col] ?? 0,
      structureVersion: sheet?.structureVersion ?? 0,
    }
  }

  const getColumnView = (request: { sheetName: string; rowStart: number; rowEnd: number; col: number }): RuntimeColumnView => {
    const direct = Reflect.get(args.runtimeColumnStore, 'getColumnView')
    if (typeof direct === 'function') {
      return direct.call(args.runtimeColumnStore, request)
    }
    const slice = args.runtimeColumnStore.getColumnSlice(request)
    return {
      owner: {
        sheetName: slice.sheetName,
        col: slice.col,
        columnVersion: slice.columnVersion,
        structureVersion: slice.structureVersion,
        sheetColumnVersions: slice.sheetColumnVersions,
        pages: new Map(),
      },
      sheetName: slice.sheetName,
      rowStart: slice.rowStart,
      rowEnd: slice.rowEnd,
      col: slice.col,
      length: slice.length,
      columnVersion: slice.columnVersion,
      structureVersion: slice.structureVersion,
      sheetColumnVersions: slice.sheetColumnVersions,
      readTagAt(offset) {
        return slice.tags[offset] ?? ValueTag.Empty
      },
      readNumberAt(offset) {
        return slice.numbers[offset] ?? 0
      },
      readStringIdAt(offset) {
        return slice.stringIds[offset] ?? 0
      },
      readErrorAt(offset) {
        return slice.errors[offset] ?? ErrorCode.None
      },
      readCellValueAt(offset) {
        return args.runtimeColumnStore.readCellValue(request.sheetName, request.rowStart + offset, request.col)
      },
    }
  }

  const extendEntry = (entry: AggregateStateEntry, targetRowEnd: number): AggregateStateEntry => {
    const view = getColumnView({
      sheetName: entry.sheetName,
      rowStart: entry.rowEnd + 1,
      rowEnd: targetRowEnd,
      col: entry.col,
    })
    const currentLength = entry.rowEnd - entry.rowStart + 1
    const totalLength = targetRowEnd - entry.rowStart + 1
    ensureCapacity(entry, totalLength)
    let runningSum = currentLength > 0 ? (entry.prefixSums[currentLength - 1] ?? 0) : 0
    let runningCount = currentLength > 0 ? (entry.prefixCount[currentLength - 1] ?? 0) : 0
    let runningAverageCount = currentLength > 0 ? (entry.prefixAverageCount[currentLength - 1] ?? 0) : 0
    let runningErrorCode =
      currentLength > 0 ? ((entry.prefixErrorCodes[currentLength - 1] as ErrorCode | undefined) ?? ErrorCode.None) : ErrorCode.None
    let runningErrorCount = currentLength > 0 ? (entry.prefixErrorCounts[currentLength - 1] ?? 0) : 0
    let runningMinimum =
      currentLength > 0 ? (entry.prefixMinimums[currentLength - 1] ?? Number.POSITIVE_INFINITY) : Number.POSITIVE_INFINITY
    let runningMaximum =
      currentLength > 0 ? (entry.prefixMaximums[currentLength - 1] ?? Number.NEGATIVE_INFINITY) : Number.NEGATIVE_INFINITY

    for (let offset = 0; offset < view.length; offset += 1) {
      const tag = decodeValueTag(view.readTagAt(offset))
      switch (tag) {
        case ValueTag.Number: {
          const numeric = view.readNumberAt(offset)
          runningSum += numeric
          runningCount += 1
          runningAverageCount += 1
          runningMinimum = Math.min(runningMinimum, numeric)
          runningMaximum = Math.max(runningMaximum, numeric)
          break
        }
        case ValueTag.Boolean: {
          const booleanNumber = view.readNumberAt(offset) !== 0 ? 1 : 0
          runningSum += booleanNumber
          runningCount += 1
          runningAverageCount += 1
          runningMinimum = Math.min(runningMinimum, booleanNumber)
          runningMaximum = Math.max(runningMaximum, booleanNumber)
          break
        }
        case ValueTag.Empty:
          runningAverageCount += 1
          runningMinimum = Math.min(runningMinimum, 0)
          runningMaximum = Math.max(runningMaximum, 0)
          break
        case ValueTag.Error:
          runningErrorCode ||= view.readErrorAt(offset) ?? ErrorCode.None
          runningErrorCount += 1
          break
        case ValueTag.String:
        default:
          break
      }
      const targetOffset = currentLength + offset
      entry.prefixSums[targetOffset] = runningSum
      entry.prefixCount[targetOffset] = runningCount
      entry.prefixAverageCount[targetOffset] = runningAverageCount
      entry.prefixErrorCodes[targetOffset] = runningErrorCode
      entry.prefixErrorCounts[targetOffset] = runningErrorCount
      entry.prefixMinimums[targetOffset] = runningMinimum
      entry.prefixMaximums[targetOffset] = runningMaximum
    }
    entry.rowEnd = targetRowEnd
    entry.columnVersion = view.columnVersion
    entry.structureVersion = view.structureVersion
    return entry
  }

  return {
    getOrBuildPrefixForRegion(regionId, aggregateKind) {
      const region = args.regionGraph.getRegion(regionId)
      if (!region) {
        throw new Error(`Unknown region id: ${regionId}`)
      }
      const rowStart = prefixAnchorForRegion(region.rowStart, aggregateKind)
      const key = cacheKey(region.sheetName, region.col, rowStart)
      const currentVersions = getCurrentVersions(region.sheetName, region.col)
      const existing = cache.get(key)
      if (
        existing &&
        existing.columnVersion === currentVersions.columnVersion &&
        existing.structureVersion === currentVersions.structureVersion &&
        existing.rowStart <= region.rowStart &&
        (aggregateKind !== 'min' && aggregateKind !== 'max' ? true : existing.extremaValid)
      ) {
        return existing.rowEnd >= region.rowEnd ? existing : extendEntry(existing, region.rowEnd)
      }
      if (existing) {
        deleteEntry(existing)
      }
      const capacity = Math.max(region.rowEnd - rowStart + 1, 16)
      const next: AggregateStateEntry = {
        regionId,
        sheetName: region.sheetName,
        col: region.col,
        rowStart,
        rowEnd: rowStart - 1,
        columnVersion: currentVersions.columnVersion,
        structureVersion: currentVersions.structureVersion,
        extremaValid: true,
        prefixSums: new Float64Array(capacity),
        prefixCount: new Uint32Array(capacity),
        prefixAverageCount: new Uint32Array(capacity),
        prefixErrorCodes: new Uint16Array(capacity),
        prefixErrorCounts: new Uint32Array(capacity),
        prefixMinimums: new Float64Array(capacity),
        prefixMaximums: new Float64Array(capacity),
      }
      cache.set(key, next)
      registerEntry(next)
      return extendEntry(next, region.rowEnd)
    },
    noteLiteralWrite({ sheetName, row, col, oldValue, newValue }) {
      const entries = entriesByColumn.get(columnKeyPrefix(sheetName, col))
      if (!entries) {
        return
      }
      const currentVersions = getCurrentVersions(sheetName, col)
      for (let entryIndex = 0; entryIndex < entries.length; entryIndex += 1) {
        const entry = entries[entryIndex]!
        if (entry.rowStart > row) {
          break
        }
        if (entry.structureVersion !== currentVersions.structureVersion) {
          deleteEntry(entry)
          entryIndex -= 1
          continue
        }
        entry.columnVersion = currentVersions.columnVersion
        entry.structureVersion = currentVersions.structureVersion
        if (row < entry.rowStart || row > entry.rowEnd) {
          continue
        }
        if (isErrorTag(oldValue) || isErrorTag(newValue) || (entry.prefixErrorCounts[entry.rowEnd - entry.rowStart] ?? 0) > 0) {
          deleteEntry(entry)
          entryIndex -= 1
          continue
        }
        const sumDelta = numericContribution(newValue) - numericContribution(oldValue)
        const countDelta = countContribution(newValue) - countContribution(oldValue)
        const averageDelta = averageCountContribution(newValue) - averageCountContribution(oldValue)
        if (sumDelta === 0 && countDelta === 0 && averageDelta === 0) {
          continue
        }
        const offset = row - entry.rowStart
        const length = entry.rowEnd - entry.rowStart + 1
        if (length - offset > PREFIX_LITERAL_DELTA_MAX_SUFFIX_LENGTH) {
          deleteEntry(entry)
          entryIndex -= 1
          continue
        }
        for (let index = offset; index < length; index += 1) {
          entry.prefixSums[index] = (entry.prefixSums[index] ?? 0) + sumDelta
          entry.prefixCount[index] = (entry.prefixCount[index] ?? 0) + countDelta
          entry.prefixAverageCount[index] = (entry.prefixAverageCount[index] ?? 0) + averageDelta
        }
        entry.extremaValid = false
      }
    },
    invalidateColumn(sheetName, col) {
      const key = columnKeyPrefix(sheetName, col)
      const entries = entriesByColumn.get(key)
      if (!entries) {
        return
      }
      entries.forEach((entry) => {
        cache.delete(cacheKey(entry.sheetName, entry.col, entry.rowStart))
      })
      entriesByColumn.delete(key)
    },
  }
}
