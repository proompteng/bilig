import { describe, expect, it, vi } from 'vitest'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { createAggregateStateStore } from '../deps/aggregate-state-store.js'
import type { SingleColumnRegionNode } from '../deps/region-node-store.js'
import type {
  EngineRuntimeColumnStoreService,
  RuntimeColumnSlice,
  RuntimeColumnView,
} from '../engine/services/runtime-column-store-service.js'
import { WorkbookStore } from '../workbook-store.js'

function createWorkbookVersions(columnVersion: number, structureVersion: number) {
  const workbook = new WorkbookStore('aggregate-state-store')
  workbook.createSheet('Sheet1')
  const sheet = workbook.getSheet('Sheet1')
  if (!sheet) {
    throw new Error('expected Sheet1 to exist')
  }
  sheet.columnVersions = Uint32Array.of(columnVersion)
  sheet.structureVersion = structureVersion
  return { workbook, sheet }
}

function region(rowEnd: number, rowStart = 0): SingleColumnRegionNode {
  return {
    id: 0,
    kind: 'single-column',
    sheetId: 1,
    sheetName: 'Sheet1',
    rowStart,
    rowEnd,
    col: 0,
    length: rowEnd - rowStart + 1,
  }
}

function makeRuntimeColumnView(args: {
  rawTags: readonly number[]
  numbers?: readonly number[]
  errors?: readonly number[]
  request: { sheetName: string; rowStart: number; rowEnd: number; col: number }
  columnVersion: number
  structureVersion: number
}): RuntimeColumnView {
  const length = args.request.rowEnd - args.request.rowStart + 1
  return {
    owner: {
      sheetName: args.request.sheetName,
      col: args.request.col,
      columnVersion: args.columnVersion,
      structureVersion: args.structureVersion,
      sheetColumnVersions: Uint32Array.of(args.columnVersion),
      pages: new Map(),
    },
    sheetName: args.request.sheetName,
    rowStart: args.request.rowStart,
    rowEnd: args.request.rowEnd,
    col: args.request.col,
    length,
    columnVersion: args.columnVersion,
    structureVersion: args.structureVersion,
    sheetColumnVersions: Uint32Array.of(args.columnVersion),
    readTagAt(offset) {
      return args.rawTags[args.request.rowStart + offset] ?? ValueTag.Empty
    },
    readNumberAt(offset) {
      return args.numbers?.[args.request.rowStart + offset] ?? 0
    },
    readStringIdAt() {
      return 0
    },
    readErrorAt(offset) {
      return args.errors?.[args.request.rowStart + offset] ?? ErrorCode.None
    },
    readCellValueAt() {
      return { tag: ValueTag.Empty }
    },
  }
}

function makeRuntimeColumnSlice(args: {
  values: readonly CellValue[]
  request: { sheetName: string; rowStart: number; rowEnd: number; col: number }
  columnVersion: number
  structureVersion: number
}): RuntimeColumnSlice {
  const length = args.request.rowEnd - args.request.rowStart + 1
  const tags = new Uint8Array(length)
  const numbers = new Float64Array(length)
  const stringIds = new Uint32Array(length)
  const errors = new Uint16Array(length)
  for (let offset = 0; offset < length; offset += 1) {
    const value = args.values[args.request.rowStart + offset] ?? { tag: ValueTag.Empty }
    tags[offset] = value.tag
    if (value.tag === ValueTag.Number) {
      numbers[offset] = value.value
    } else if (value.tag === ValueTag.Boolean) {
      numbers[offset] = value.value ? 1 : 0
    } else if (value.tag === ValueTag.String) {
      stringIds[offset] = value.stringId
    } else if (value.tag === ValueTag.Error) {
      errors[offset] = value.code
    }
  }
  return {
    sheetName: args.request.sheetName,
    rowStart: args.request.rowStart,
    rowEnd: args.request.rowEnd,
    col: args.request.col,
    length,
    columnVersion: args.columnVersion,
    structureVersion: args.structureVersion,
    sheetColumnVersions: Uint32Array.of(args.columnVersion),
    tags,
    numbers,
    stringIds,
    errors,
  }
}

function createStore(args: { workbook: WorkbookStore; runtimeColumnStore: EngineRuntimeColumnStoreService; rowEnd: number }) {
  return createAggregateStateStore({
    workbook: args.workbook,
    runtimeColumnStore: args.runtimeColumnStore,
    regionGraph: {
      getRegion: (regionId) => (regionId === 0 ? region(args.rowEnd) : undefined),
    },
  })
}

describe('AggregateStateStore', () => {
  it('builds direct column-view prefixes across numeric, error, string, empty, and unknown tags', () => {
    const { workbook } = createWorkbookVersions(4, 2)
    const rawTags = [ValueTag.Number, ValueTag.Boolean, ValueTag.Empty, ValueTag.Error, ValueTag.String, ValueTag.Empty, 99]
    const numbers = [4, 1, 0, 0, 0, 0, 0]
    const errors = [0, 0, 0, ErrorCode.NA, 0, 0, 0]
    const getColumnView = vi.fn((request: { sheetName: string; rowStart: number; rowEnd: number; col: number }) =>
      makeRuntimeColumnView({
        rawTags,
        numbers,
        errors,
        request,
        columnVersion: 4,
        structureVersion: 2,
      }),
    )
    const store = createStore({
      workbook,
      rowEnd: rawTags.length - 1,
      runtimeColumnStore: {
        getColumnOwner: () => {
          throw new Error('unexpected column owner request')
        },
        getColumnView,
        getColumnSlice: vi.fn(),
        readCellValue: () => ({ tag: ValueTag.Empty }),
        readRangeValues: () => [],
        normalizeStringId: () => '',
        normalizeLookupText: () => '',
      },
    })

    const entry = store.getOrBuildPrefixForRegion(0)

    expect(entry.prefixSums.subarray(0, 7)).toEqual(Float64Array.from([4, 5, 5, 5, 5, 5, 5]))
    expect(entry.prefixCount.subarray(0, 7)).toEqual(Uint32Array.from([1, 2, 2, 2, 2, 2, 2]))
    expect(entry.prefixAverageCount.subarray(0, 7)).toEqual(Uint32Array.from([1, 2, 3, 3, 3, 4, 5]))
    expect(entry.prefixErrorCodes.subarray(0, 7)).toEqual(
      Uint16Array.from([0, 0, 0, ErrorCode.NA, ErrorCode.NA, ErrorCode.NA, ErrorCode.NA]),
    )
    expect(entry.prefixErrorCounts.subarray(0, 7)).toEqual(Uint32Array.from([0, 0, 0, 1, 1, 1, 1]))
    expect(entry.prefixMinimums.subarray(0, 7)).toEqual(Float64Array.from([4, 1, 0, 0, 0, 0, 0]))
    expect(entry.prefixMaximums.subarray(0, 7)).toEqual(Float64Array.from([4, 4, 4, 4, 4, 4, 4]))
    expect(() => store.getOrBuildPrefixForRegion(1)).toThrow('Unknown region id: 1')
  })

  it('falls back to column slices when direct column views are unavailable', () => {
    const { workbook } = createWorkbookVersions(3, 1)
    const values: CellValue[] = [
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.String, value: 'ignored', stringId: 9 },
    ]
    const getColumnSlice = vi.fn((request: { sheetName: string; rowStart: number; rowEnd: number; col: number }) =>
      makeRuntimeColumnSlice({
        values,
        request,
        columnVersion: 3,
        structureVersion: 1,
      }),
    )
    const runtimeColumnStore: EngineRuntimeColumnStoreService = {
      getColumnOwner: () => {
        throw new Error('unexpected column owner request')
      },
      getColumnView: () => {
        throw new Error('expected aggregate state store to fall back to getColumnSlice')
      },
      getColumnSlice,
      readCellValue: () => ({ tag: ValueTag.String, value: 'fallback', stringId: 1 }),
      readRangeValues: () => [],
      normalizeStringId: () => '',
      normalizeLookupText: () => '',
    }
    Reflect.deleteProperty(runtimeColumnStore, 'getColumnView')
    const store = createStore({
      workbook,
      rowEnd: 1,
      runtimeColumnStore,
    })

    const entry = store.getOrBuildPrefixForRegion(0)

    expect(entry.prefixSums.subarray(0, 2)).toEqual(Float64Array.from([2, 2]))
    expect(entry.prefixCount.subarray(0, 2)).toEqual(Uint32Array.from([1, 1]))
    expect(getColumnSlice).toHaveBeenCalledOnce()
  })

  it('applies literal deltas in place and rebuilds when extrema are requested', () => {
    const { workbook, sheet } = createWorkbookVersions(1, 1)
    const rawTags = [ValueTag.Number, ValueTag.Number, ValueTag.Number]
    const numbers = [1, 2, 3]
    const getColumnView = vi.fn((request: { sheetName: string; rowStart: number; rowEnd: number; col: number }) =>
      makeRuntimeColumnView({
        rawTags,
        numbers,
        request,
        columnVersion: sheet.columnVersions[0] ?? 0,
        structureVersion: sheet.structureVersion,
      }),
    )
    const store = createStore({
      workbook,
      rowEnd: 2,
      runtimeColumnStore: {
        getColumnOwner: () => {
          throw new Error('unexpected column owner request')
        },
        getColumnView,
        getColumnSlice: vi.fn(),
        readCellValue: () => ({ tag: ValueTag.Empty }),
        readRangeValues: () => [],
        normalizeStringId: () => '',
        normalizeLookupText: () => '',
      },
    })

    const entry = store.getOrBuildPrefixForRegion(0)
    sheet.columnVersions[0] = 2
    numbers[1] = 5
    store.noteLiteralWrite({
      sheetName: 'Sheet1',
      row: 1,
      col: 0,
      oldValue: { tag: ValueTag.Number, value: 2 },
      newValue: { tag: ValueTag.Number, value: 5 },
    })

    expect(entry.prefixSums.subarray(0, 3)).toEqual(Float64Array.from([1, 6, 9]))
    expect(entry.prefixCount.subarray(0, 3)).toEqual(Uint32Array.from([1, 2, 3]))
    expect(entry.extremaValid).toBe(false)

    const rebuiltForExtrema = store.getOrBuildPrefixForRegion(0, 'max')
    expect(rebuiltForExtrema).not.toBe(entry)
    expect(rebuiltForExtrema.extremaValid).toBe(true)
    expect(rebuiltForExtrema.prefixMaximums.subarray(0, 3)).toEqual(Float64Array.from([1, 5, 5]))
    expect(getColumnView).toHaveBeenCalledTimes(2)
  })

  it('grows prefix buffers when extending past the initial capacity', () => {
    const { workbook, sheet } = createWorkbookVersions(1, 1)
    const rawTags = Array.from({ length: 33 }, () => ValueTag.Number)
    const numbers = Array.from({ length: 33 }, (_, index) => index + 1)
    let rowEnd = 15
    const getColumnView = vi.fn((request: { sheetName: string; rowStart: number; rowEnd: number; col: number }) =>
      makeRuntimeColumnView({
        rawTags,
        numbers,
        request,
        columnVersion: sheet.columnVersions[0] ?? 0,
        structureVersion: sheet.structureVersion,
      }),
    )
    const store = createAggregateStateStore({
      workbook,
      runtimeColumnStore: {
        getColumnOwner: () => {
          throw new Error('unexpected column owner request')
        },
        getColumnView,
        getColumnSlice: vi.fn(),
        readCellValue: () => ({ tag: ValueTag.Empty }),
        readRangeValues: () => [],
        normalizeStringId: () => '',
        normalizeLookupText: () => '',
      },
      regionGraph: {
        getRegion: () => region(rowEnd),
      },
    })

    const initial = store.getOrBuildPrefixForRegion(0)
    expect(initial.prefixSums.length).toBe(16)

    rowEnd = 32
    const extended = store.getOrBuildPrefixForRegion(0)

    expect(extended).toBe(initial)
    expect(extended.prefixSums.length).toBeGreaterThanOrEqual(33)
    expect(extended.prefixSums[32]).toBe(561)
    expect(getColumnView).toHaveBeenCalledTimes(2)
  })

  it('applies bounded prefix suffix deltas in place on early-row writes', () => {
    const { workbook, sheet } = createWorkbookVersions(1, 1)
    const rawTags = Array.from({ length: 256 }, () => ValueTag.Number)
    const numbers = Array.from({ length: 256 }, () => 1)
    const getColumnView = vi.fn((request: { sheetName: string; rowStart: number; rowEnd: number; col: number }) =>
      makeRuntimeColumnView({
        rawTags,
        numbers,
        request,
        columnVersion: sheet.columnVersions[0] ?? 0,
        structureVersion: sheet.structureVersion,
      }),
    )
    const store = createStore({
      workbook,
      rowEnd: 255,
      runtimeColumnStore: {
        getColumnOwner: () => {
          throw new Error('unexpected column owner request')
        },
        getColumnView,
        getColumnSlice: vi.fn(),
        readCellValue: () => ({ tag: ValueTag.Empty }),
        readRangeValues: () => [],
        normalizeStringId: () => '',
        normalizeLookupText: () => '',
      },
    })

    const entry = store.getOrBuildPrefixForRegion(0)
    sheet.columnVersions[0] = 2
    numbers[0] = 2
    store.noteLiteralWrite({
      sheetName: 'Sheet1',
      row: 0,
      col: 0,
      oldValue: { tag: ValueTag.Number, value: 1 },
      newValue: { tag: ValueTag.Number, value: 2 },
    })

    const rebuilt = store.getOrBuildPrefixForRegion(0)
    expect(rebuilt).toBe(entry)
    expect(rebuilt.prefixSums[0]).toBe(2)
    expect(rebuilt.prefixSums[255]).toBe(257)
    expect(getColumnView).toHaveBeenCalledTimes(1)
  })

  it('anchors prefixes at the requested row start to avoid cancellation from earlier rows', () => {
    const { workbook } = createWorkbookVersions(1, 1)
    const rawTags = Array.from({ length: 5 }, () => ValueTag.Number)
    const numbers = [663897216, 1327794496, 1327794496, 440759555901973500, 663897248]
    const getColumnView = vi.fn((request: { sheetName: string; rowStart: number; rowEnd: number; col: number }) =>
      makeRuntimeColumnView({
        rawTags,
        numbers,
        request,
        columnVersion: 1,
        structureVersion: 1,
      }),
    )
    const store = createAggregateStateStore({
      workbook,
      runtimeColumnStore: {
        getColumnOwner: () => {
          throw new Error('unexpected column owner request')
        },
        getColumnView,
        getColumnSlice: vi.fn(),
        readCellValue: () => ({ tag: ValueTag.Empty }),
        readRangeValues: () => [],
        normalizeStringId: () => '',
        normalizeLookupText: () => '',
      },
      regionGraph: {
        getRegion: () => region(4, 4),
      },
    })

    const entry = store.getOrBuildPrefixForRegion(0)

    expect(entry.rowStart).toBe(4)
    expect(entry.prefixSums[0]).toBe(663897248)
  })

  it('evicts stale or error-bearing prefixes and ignores neutral writes', () => {
    const { workbook, sheet } = createWorkbookVersions(1, 1)
    const rawTags = [ValueTag.Boolean, ValueTag.Boolean]
    const numbers = [1, 0]
    const getColumnView = vi.fn((request: { sheetName: string; rowStart: number; rowEnd: number; col: number }) =>
      makeRuntimeColumnView({
        rawTags,
        numbers,
        request,
        columnVersion: sheet.columnVersions[0] ?? 0,
        structureVersion: sheet.structureVersion,
      }),
    )
    const store = createStore({
      workbook,
      rowEnd: 1,
      runtimeColumnStore: {
        getColumnOwner: () => {
          throw new Error('unexpected column owner request')
        },
        getColumnView,
        getColumnSlice: vi.fn(),
        readCellValue: () => ({ tag: ValueTag.Empty }),
        readRangeValues: () => [],
        normalizeStringId: () => '',
        normalizeLookupText: () => '',
      },
    })

    const entry = store.getOrBuildPrefixForRegion(0)
    sheet.columnVersions[0] = 2
    numbers[0] = 0
    store.noteLiteralWrite({
      sheetName: 'Sheet1',
      row: 0,
      col: 0,
      oldValue: { tag: ValueTag.Boolean, value: true },
      newValue: { tag: ValueTag.Boolean, value: false },
    })
    expect(entry.prefixSums.subarray(0, 2)).toEqual(Float64Array.from([0, 0]))

    store.noteLiteralWrite({
      sheetName: 'Sheet1',
      row: 0,
      col: 0,
      oldValue: { tag: ValueTag.String, value: 'same', stringId: 1 },
      newValue: { tag: ValueTag.String, value: 'same', stringId: 1 },
    })
    expect(getColumnView).toHaveBeenCalledTimes(1)

    sheet.structureVersion = 9
    store.noteLiteralWrite({
      sheetName: 'Sheet1',
      row: 0,
      col: 0,
      oldValue: { tag: ValueTag.Boolean, value: false },
      newValue: { tag: ValueTag.Number, value: 4 },
    })
    store.getOrBuildPrefixForRegion(0)
    expect(getColumnView).toHaveBeenCalledTimes(2)

    store.noteLiteralWrite({
      sheetName: 'Sheet1',
      row: 0,
      col: 0,
      oldValue: { tag: ValueTag.Error, code: ErrorCode.Value },
      newValue: { tag: ValueTag.Number, value: 4 },
    })
    store.getOrBuildPrefixForRegion(0)
    expect(getColumnView).toHaveBeenCalledTimes(3)
  })
})
