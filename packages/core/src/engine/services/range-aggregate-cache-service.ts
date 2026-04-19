import { ErrorCode, ValueTag } from '@bilig/protocol'
import type { EngineRuntimeState } from '../runtime-state.js'
import type { EngineRuntimeColumnStoreService, RuntimeColumnView } from './runtime-column-store-service.js'

interface AggregatePrefixEntry {
  sheetName: string
  rowStart: number
  rowEnd: number
  col: number
  columnVersion: number
  structureVersion: number
  prefixSums: Float64Array
  prefixCount: Uint32Array
  prefixAverageCount: Uint32Array
  prefixErrorCodes: Uint16Array
  prefixErrorCounts: Uint32Array
  prefixMinimums: Float64Array
  prefixMaximums: Float64Array
}

export interface RangeAggregateCacheService {
  readonly getOrBuildPrefix: (request: { sheetName: string; rowStart: number; rowEnd: number; col: number }) => AggregatePrefixEntry
}

function cacheKey(sheetName: string, col: number): string {
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

function rawTagAt(view: RuntimeColumnView, offset: number): number {
  return view.readTagAt(offset)
}

function numericAt(view: RuntimeColumnView, offset: number): number {
  return view.readNumberAt(offset)
}

function errorAt(view: RuntimeColumnView, offset: number): number {
  return view.readErrorAt(offset)
}

function ensurePrefixCapacity(existing: AggregatePrefixEntry, totalLength: number): void {
  if (existing.prefixSums.length >= totalLength) {
    return
  }
  const nextCapacity = Math.max(existing.prefixSums.length * 2, totalLength)
  const nextPrefixSums = new Float64Array(nextCapacity)
  const nextPrefixCount = new Uint32Array(nextCapacity)
  const nextPrefixAverageCount = new Uint32Array(nextCapacity)
  const nextPrefixErrorCodes = new Uint16Array(nextCapacity)
  const nextPrefixErrorCounts = new Uint32Array(nextCapacity)
  const nextPrefixMinimums = new Float64Array(nextCapacity)
  const nextPrefixMaximums = new Float64Array(nextCapacity)
  const currentLength = existing.rowEnd - existing.rowStart + 1
  nextPrefixSums.set(existing.prefixSums.subarray(0, currentLength), 0)
  nextPrefixCount.set(existing.prefixCount.subarray(0, currentLength), 0)
  nextPrefixAverageCount.set(existing.prefixAverageCount.subarray(0, currentLength), 0)
  nextPrefixErrorCodes.set(existing.prefixErrorCodes.subarray(0, currentLength), 0)
  nextPrefixErrorCounts.set(existing.prefixErrorCounts.subarray(0, currentLength), 0)
  nextPrefixMinimums.set(existing.prefixMinimums.subarray(0, currentLength), 0)
  nextPrefixMaximums.set(existing.prefixMaximums.subarray(0, currentLength), 0)
  existing.prefixSums = nextPrefixSums
  existing.prefixCount = nextPrefixCount
  existing.prefixAverageCount = nextPrefixAverageCount
  existing.prefixErrorCodes = nextPrefixErrorCodes
  existing.prefixErrorCounts = nextPrefixErrorCounts
  existing.prefixMinimums = nextPrefixMinimums
  existing.prefixMaximums = nextPrefixMaximums
}

export function createRangeAggregateCacheService(args: {
  readonly state: Pick<EngineRuntimeState, 'workbook'>
  readonly runtimeColumnStore: EngineRuntimeColumnStoreService
}): RangeAggregateCacheService {
  const emptyColumnVersions = new Uint32Array(0)
  const cache = new Map<string, AggregatePrefixEntry[]>()
  const recentReusable = new Map<string, AggregatePrefixEntry>()

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
        const tag = decodeValueTag(slice.tags[offset])
        switch (tag) {
          case ValueTag.Empty:
            return { tag: ValueTag.Empty }
          case ValueTag.Number:
            return { tag: ValueTag.Number, value: slice.numbers[offset] ?? 0 }
          case ValueTag.Boolean:
            return { tag: ValueTag.Boolean, value: (slice.numbers[offset] ?? 0) !== 0 }
          case ValueTag.String:
            return { tag: ValueTag.String, value: '', stringId: slice.stringIds[offset] ?? 0 }
          case ValueTag.Error:
            return { tag: ValueTag.Error, code: slice.errors[offset] ?? ErrorCode.None }
          default:
            return { tag: ValueTag.Empty }
        }
      },
    }
  }

  const getCurrentVersions = (sheetName: string, col: number) => {
    const sheet = args.state.workbook.getSheet(sheetName)
    const columnVersions = sheet?.columnVersions ?? emptyColumnVersions
    return {
      columnVersion: columnVersions[col] ?? 0,
      structureVersion: sheet?.structureVersion ?? 0,
    }
  }

  const buildPrefix = (request: { sheetName: string; rowStart: number; rowEnd: number; col: number }): AggregatePrefixEntry => {
    const view = getColumnView(request)
    const prefixSums = new Float64Array(view.length)
    const prefixCount = new Uint32Array(view.length)
    const prefixAverageCount = new Uint32Array(view.length)
    const prefixErrorCodes = new Uint16Array(view.length)
    const prefixErrorCounts = new Uint32Array(view.length)
    const prefixMinimums = new Float64Array(view.length)
    const prefixMaximums = new Float64Array(view.length)
    let runningSum = 0
    let runningCount = 0
    let runningAverageCount = 0
    let runningErrorCode = ErrorCode.None
    let runningErrorCount = 0
    let runningMinimum = Number.POSITIVE_INFINITY
    let runningMaximum = Number.NEGATIVE_INFINITY
    for (let offset = 0; offset < view.length; offset += 1) {
      const tag = decodeValueTag(rawTagAt(view, offset))
      switch (tag) {
        case ValueTag.Number: {
          const numeric = numericAt(view, offset)
          runningSum += numeric
          runningCount += 1
          runningAverageCount += 1
          runningMinimum = Math.min(runningMinimum, numeric)
          runningMaximum = Math.max(runningMaximum, numeric)
          break
        }
        case ValueTag.Boolean: {
          const booleanNumber = numericAt(view, offset) !== 0 ? 1 : 0
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
          runningErrorCode ||= errorAt(view, offset) ?? ErrorCode.None
          runningErrorCount += 1
          break
        case ValueTag.String:
        default:
          break
      }
      prefixSums[offset] = runningSum
      prefixCount[offset] = runningCount
      prefixAverageCount[offset] = runningAverageCount
      prefixErrorCodes[offset] = runningErrorCode
      prefixErrorCounts[offset] = runningErrorCount
      prefixMinimums[offset] = runningMinimum
      prefixMaximums[offset] = runningMaximum
    }
    return {
      sheetName: request.sheetName,
      rowStart: request.rowStart,
      rowEnd: request.rowEnd,
      col: request.col,
      columnVersion: view.columnVersion,
      structureVersion: view.structureVersion,
      prefixSums,
      prefixCount,
      prefixAverageCount,
      prefixErrorCodes,
      prefixErrorCounts,
      prefixMinimums,
      prefixMaximums,
    }
  }

  const extendPrefix = (
    existing: AggregatePrefixEntry,
    request: {
      sheetName: string
      rowStart: number
      rowEnd: number
      col: number
    },
  ): AggregatePrefixEntry => {
    const deltaView = getColumnView({
      sheetName: request.sheetName,
      rowStart: existing.rowEnd + 1,
      rowEnd: request.rowEnd,
      col: request.col,
    })
    const totalLength = request.rowEnd - existing.rowStart + 1
    ensurePrefixCapacity(existing, totalLength)
    const currentLength = existing.rowEnd - existing.rowStart + 1
    let runningSum = existing.prefixSums[currentLength - 1] ?? 0
    let runningCount = existing.prefixCount[currentLength - 1] ?? 0
    let runningAverageCount = existing.prefixAverageCount[currentLength - 1] ?? 0
    let runningErrorCode = (existing.prefixErrorCodes[currentLength - 1] as ErrorCode | undefined) ?? ErrorCode.None
    let runningErrorCount = existing.prefixErrorCounts[currentLength - 1] ?? 0
    let runningMinimum = existing.prefixMinimums[currentLength - 1] ?? Number.POSITIVE_INFINITY
    let runningMaximum = existing.prefixMaximums[currentLength - 1] ?? Number.NEGATIVE_INFINITY
    for (let offset = 0; offset < deltaView.length; offset += 1) {
      const tag = decodeValueTag(rawTagAt(deltaView, offset))
      switch (tag) {
        case ValueTag.Number: {
          const numeric = numericAt(deltaView, offset)
          runningSum += numeric
          runningCount += 1
          runningAverageCount += 1
          runningMinimum = Math.min(runningMinimum, numeric)
          runningMaximum = Math.max(runningMaximum, numeric)
          break
        }
        case ValueTag.Boolean: {
          const booleanNumber = numericAt(deltaView, offset) !== 0 ? 1 : 0
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
          runningErrorCode ||= errorAt(deltaView, offset) ?? ErrorCode.None
          runningErrorCount += 1
          break
        case ValueTag.String:
        default:
          break
      }
      const targetOffset = currentLength + offset
      existing.prefixSums[targetOffset] = runningSum
      existing.prefixCount[targetOffset] = runningCount
      existing.prefixAverageCount[targetOffset] = runningAverageCount
      existing.prefixErrorCodes[targetOffset] = runningErrorCode
      existing.prefixErrorCounts[targetOffset] = runningErrorCount
      existing.prefixMinimums[targetOffset] = runningMinimum
      existing.prefixMaximums[targetOffset] = runningMaximum
    }
    existing.rowEnd = request.rowEnd
    existing.columnVersion = deltaView.columnVersion
    existing.structureVersion = deltaView.structureVersion
    return existing
  }

  return {
    getOrBuildPrefix(request) {
      const key = cacheKey(request.sheetName, request.col)
      const currentVersions = getCurrentVersions(request.sheetName, request.col)
      const recent = recentReusable.get(key)
      if (
        recent &&
        recent.columnVersion === currentVersions.columnVersion &&
        recent.structureVersion === currentVersions.structureVersion &&
        recent.rowStart <= request.rowStart
      ) {
        if (recent.rowEnd >= request.rowEnd) {
          return recent
        }
        const extended = extendPrefix(recent, {
          sheetName: request.sheetName,
          rowStart: recent.rowStart,
          rowEnd: request.rowEnd,
          col: request.col,
        })
        recentReusable.set(key, extended)
        return extended
      }
      const existingEntries = cache.get(key) ?? []
      const compatibleEntries: AggregatePrefixEntry[] = []
      let reusable: AggregatePrefixEntry | undefined
      for (const entry of existingEntries) {
        if (entry.columnVersion !== currentVersions.columnVersion || entry.structureVersion !== currentVersions.structureVersion) {
          continue
        }
        compatibleEntries.push(entry)
        if (entry.rowStart <= request.rowStart && (reusable === undefined || entry.rowStart < reusable.rowStart)) {
          reusable = entry
        }
      }
      if (compatibleEntries.length > 0) {
        cache.set(key, compatibleEntries)
      }
      if (reusable) {
        if (reusable.rowEnd >= request.rowEnd) {
          recentReusable.set(key, reusable)
          return reusable
        }
        const extended = extendPrefix(reusable, {
          sheetName: request.sheetName,
          rowStart: reusable.rowStart,
          rowEnd: request.rowEnd,
          col: request.col,
        })
        recentReusable.set(key, extended)
        return extended
      }
      const built = buildPrefix(request)
      compatibleEntries.push(built)
      cache.set(key, compatibleEntries)
      recentReusable.set(key, built)
      return built
    },
  }
}
