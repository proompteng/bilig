import { describe, expect, it, vi } from 'vitest'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import { StringPool } from '../string-pool.js'
import { WorkbookStore } from '../workbook-store.js'
import { createColumnIndexStore } from '../indexes/column-index-store.js'
import type { EngineRuntimeState, PreparedApproximateVectorLookup } from '../engine/runtime-state.js'
import { createSortedColumnSearchService } from '../engine/services/sorted-column-search-service.js'
import { isLookupColumnOwner } from '../engine/services/lookup-column-owner.js'
import { createEngineRuntimeColumnStoreService } from '../engine/services/runtime-column-store-service.js'
import { createEngineCounters } from '../perf/engine-counters.js'
import type {
  EngineRuntimeColumnStoreService,
  RuntimeColumnOwner,
  RuntimeColumnView,
} from '../engine/services/runtime-column-store-service.js'

function setStoredNumber(workbook: WorkbookStore, strings: StringPool, address: string, value: number): void {
  const cellIndex = workbook.ensureCell('Sheet1', address)
  workbook.cellStore.setValue(cellIndex, { tag: ValueTag.Number, value }, 0)
  void strings
}

function setStoredString(workbook: WorkbookStore, strings: StringPool, address: string, value: string): void {
  const cellIndex = workbook.ensureCell('Sheet1', address)
  workbook.cellStore.setValue(cellIndex, { tag: ValueTag.String, value }, strings.intern(value))
}

function createSorted(workbook: WorkbookStore, strings: StringPool, counters = createEngineCounters()) {
  const runtimeColumnStore = createEngineRuntimeColumnStoreService({
    state: { workbook, strings, counters },
  })
  const columnIndexStore = createColumnIndexStore({
    state: { workbook, strings },
    runtimeColumnStore,
  })
  return createSortedColumnSearchService({
    state: { workbook, strings, counters },
    runtimeColumnStore,
    columnIndexStore,
  })
}

describe('createSortedColumnSearchService', () => {
  it('serves approximate matches from a primed sorted column and invalidates by column version', () => {
    const workbook = new WorkbookStore('sorted-index')
    const strings = new StringPool()
    workbook.createSheet('Sheet1')

    ;[1, 3, 5, 7].forEach((value, index) => {
      setStoredNumber(workbook, strings, `A${index + 1}`, value)
    })

    const sorted = createSorted(workbook, strings)

    sorted.primeColumnIndex({ sheetName: 'Sheet1', rowStart: 0, rowEnd: 3, col: 0 })

    expect(
      sorted.findVectorMatch({
        lookupValue: { tag: ValueTag.Number, value: 6 },
        sheetName: 'Sheet1',
        start: 'A1',
        end: 'A4',
        startRow: 0,
        endRow: 3,
        startCol: 0,
        endCol: 0,
        matchMode: 1,
      }),
    ).toEqual({ handled: true, position: 3 })

    setStoredNumber(workbook, strings, 'A3', 6)

    expect(
      sorted.findVectorMatch({
        lookupValue: { tag: ValueTag.Number, value: 6 },
        sheetName: 'Sheet1',
        start: 'A1',
        end: 'A4',
        startRow: 0,
        endRow: 3,
        startCol: 0,
        endCol: 0,
        matchMode: 1,
      }),
    ).toEqual({ handled: true, position: 3 })
  })

  it('updates cached sorted columns incrementally for monotonic literal writes without rematerializing the column', () => {
    const workbook = new WorkbookStore('sorted-index-incremental')
    const strings = new StringPool()
    workbook.createSheet('Sheet1')

    ;[1, 3, 5, 7].forEach((value, index) => {
      setStoredNumber(workbook, strings, `A${index + 1}`, value)
    })

    const runtimeColumnStore = createEngineRuntimeColumnStoreService({
      state: { workbook, strings },
    })
    const getColumnSliceSpy = vi.spyOn(runtimeColumnStore, 'getColumnSlice')
    const columnIndexStore = createColumnIndexStore({
      state: { workbook, strings },
      runtimeColumnStore,
    })
    const sorted = createSortedColumnSearchService({
      state: { workbook, strings },
      runtimeColumnStore,
      columnIndexStore,
    })

    const prepared = sorted.prepareVectorLookup({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 3,
      col: 0,
    })
    const sliceCallsAfterPrepare = getColumnSliceSpy.mock.calls.length

    setStoredNumber(workbook, strings, 'A4', 9)
    sorted.recordLiteralWrite({
      sheetName: 'Sheet1',
      row: 3,
      col: 0,
      oldValue: { tag: ValueTag.Number, value: 7 },
      newValue: { tag: ValueTag.Number, value: 9 },
    })

    expect(
      sorted.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.Number, value: 8 },
        prepared,
        matchMode: 1,
      }),
    ).toEqual({ handled: true, position: 3 })
    expect(getColumnSliceSpy.mock.calls.length).toBe(sliceCallsAfterPrepare)
  })

  it('prepares overlapping approximate lookups from shared column owners instead of copied slices', () => {
    const workbook = new WorkbookStore('sorted-index-column-owner')
    const strings = new StringPool()
    workbook.createSheet('Sheet1')

    ;[1, 3, 5, 7].forEach((value, index) => {
      setStoredNumber(workbook, strings, `A${index + 1}`, value)
    })

    const runtimeColumnStore = createEngineRuntimeColumnStoreService({
      state: { workbook, strings },
    })
    const getColumnOwnerSpy = vi.spyOn(runtimeColumnStore, 'getColumnOwner')
    const getColumnSliceSpy = vi.spyOn(runtimeColumnStore, 'getColumnSlice')
    const getColumnViewSpy = vi.spyOn(runtimeColumnStore, 'getColumnView')
    const columnIndexStore = createColumnIndexStore({
      state: { workbook, strings },
      runtimeColumnStore,
    })
    const sorted = createSortedColumnSearchService({
      state: { workbook, strings },
      runtimeColumnStore,
      columnIndexStore,
    })

    sorted.prepareVectorLookup({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 2,
      col: 0,
    })
    sorted.prepareVectorLookup({
      sheetName: 'Sheet1',
      rowStart: 1,
      rowEnd: 3,
      col: 0,
    })

    expect(getColumnOwnerSpy).toHaveBeenCalledTimes(1)
    expect(getColumnViewSpy).not.toHaveBeenCalled()
    expect(getColumnSliceSpy).not.toHaveBeenCalled()
  })

  it('primes owner-backed approximate ranges without building window indices', () => {
    const workbook = new WorkbookStore('sorted-index-owner-prime')
    const strings = new StringPool()
    const counters = createEngineCounters()
    workbook.createSheet('Sheet1')

    ;[1, 3, 5, 7].forEach((value, index) => {
      setStoredNumber(workbook, strings, `A${index + 1}`, value)
    })

    const runtimeColumnStore = createEngineRuntimeColumnStoreService({
      state: { workbook, strings, counters },
    })
    const getColumnViewSpy = vi.spyOn(runtimeColumnStore, 'getColumnView')
    const columnIndexStore = createColumnIndexStore({
      state: { workbook, strings },
      runtimeColumnStore,
    })
    const sorted = createSortedColumnSearchService({
      state: { workbook, strings, counters },
      runtimeColumnStore,
      columnIndexStore,
    })

    sorted.primeColumnIndex({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 3,
      col: 0,
    })

    expect(counters.approxIndexBuilds).toBe(0)
    expect(getColumnViewSpy).not.toHaveBeenCalled()

    const prepared = sorted.prepareVectorLookup({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 3,
      col: 0,
    })
    expect(prepared.internalOwner).toBeDefined()
  })

  it('uses owner-backed prepared approximate lookups for numeric and text subranges after writes', () => {
    const workbook = new WorkbookStore('sorted-index-owner-prepared')
    const strings = new StringPool()
    workbook.createSheet('Sheet1')

    ;[1, 3, 5, 7].forEach((value, index) => {
      setStoredNumber(workbook, strings, `A${index + 1}`, value)
    })
    ;['apple', 'banana', 'pear', 'plum'].forEach((value, index) => {
      setStoredString(workbook, strings, `B${index + 1}`, value)
    })

    const sorted = createSorted(workbook, strings)

    const numericPrepared = sorted.prepareVectorLookup({
      sheetName: 'Sheet1',
      rowStart: 1,
      rowEnd: 3,
      col: 0,
    })
    expect(numericPrepared.internalOwner).toBeDefined()
    expect(numericPrepared.comparableKind).toBe('numeric')
    expect(numericPrepared.sortedAscending).toBe(true)
    expect(
      sorted.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.Number, value: 6 },
        prepared: numericPrepared,
        matchMode: 1,
      }),
    ).toEqual({ handled: true, position: 2 })

    const textPrepared = sorted.prepareVectorLookup({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 2,
      col: 1,
    })
    expect(textPrepared.internalOwner).toBeDefined()
    expect(textPrepared.comparableKind).toBe('text')
    expect(
      sorted.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.String, value: 'PEACH' },
        prepared: textPrepared,
        matchMode: 1,
      }),
    ).toEqual({ handled: true, position: 2 })

    setStoredString(workbook, strings, 'B2', 'blueberry')
    sorted.recordLiteralWrite({
      sheetName: 'Sheet1',
      row: 1,
      col: 1,
      oldValue: { tag: ValueTag.String, value: 'banana' },
      newValue: { tag: ValueTag.String, value: 'blueberry' },
      oldStringId: strings.intern('banana'),
      newStringId: strings.intern('blueberry'),
    })

    expect(
      sorted.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.String, value: 'BLUEBERRY' },
        prepared: textPrepared,
        matchMode: 1,
      }),
    ).toEqual({ handled: true, position: 2 })
  })

  it('uses owner-backed uniform approximate summaries before probing owner numeric arrays', () => {
    const workbook = new WorkbookStore('sorted-index-owner-uniform-fast-path')
    const strings = new StringPool()
    workbook.createSheet('Sheet1')

    ;[10, 20, 30, 40].forEach((value, index) => {
      setStoredNumber(workbook, strings, `A${index + 1}`, value)
    })
    ;[40, 30, 20, 10].forEach((value, index) => {
      setStoredNumber(workbook, strings, `B${index + 1}`, value)
    })

    const sorted = createSorted(workbook, strings)

    const ascendingPrepared = sorted.prepareVectorLookup({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 3,
      col: 0,
    })
    expect(ascendingPrepared.internalOwner).toBeDefined()
    expect(ascendingPrepared.uniformStart).toBe(10)
    expect(ascendingPrepared.uniformStep).toBe(10)

    const ascendingOwner = ascendingPrepared.internalOwner
    if (!isLookupColumnOwner(ascendingOwner)) {
      throw new Error('expected ascending prepared lookup to use an owner')
    }
    Object.defineProperty(ascendingOwner, 'numericValues', {
      value: new Float64Array([Number.NaN, Number.NaN, Number.NaN, Number.NaN]),
    })

    expect(
      sorted.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.Number, value: 25 },
        prepared: ascendingPrepared,
        matchMode: 1,
      }),
    ).toEqual({ handled: true, position: 2 })
    expect(
      sorted.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.Empty, value: null },
        prepared: ascendingPrepared,
        matchMode: 1,
      }),
    ).toEqual({ handled: true, position: undefined })

    const descendingPrepared = sorted.prepareVectorLookup({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 3,
      col: 1,
    })
    expect(descendingPrepared.internalOwner).toBeDefined()
    expect(descendingPrepared.uniformStart).toBe(40)
    expect(descendingPrepared.uniformStep).toBe(-10)

    const descendingOwner = descendingPrepared.internalOwner
    if (!isLookupColumnOwner(descendingOwner)) {
      throw new Error('expected descending prepared lookup to use an owner')
    }
    Object.defineProperty(descendingOwner, 'numericValues', {
      value: new Float64Array([Number.NaN, Number.NaN, Number.NaN, Number.NaN]),
    })

    expect(
      sorted.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.Number, value: 25 },
        prepared: descendingPrepared,
        matchMode: -1,
      }),
    ).toEqual({ handled: true, position: 2 })
    expect(
      sorted.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.Boolean, value: true },
        prepared: descendingPrepared,
        matchMode: -1,
      }),
    ).toEqual({ handled: true, position: 4 })
  })

  it('covers fallback approximate indices when owner coverage is unavailable and invalidates on type-changing writes', () => {
    const strings = new StringPool()
    const columnVersions = new Uint32Array([0, 0])
    let structureVersion = 1
    let sheetAvailable = true
    const numericTags = [ValueTag.Number, ValueTag.Number, ValueTag.Number, ValueTag.Number]
    const numericNumbers = [1, 3, 5, 7]
    const textTags = [ValueTag.String, ValueTag.String, ValueTag.String, ValueTag.String]
    const textIds = ['apple', 'banana', 'pear', 'plum'].map((value) => strings.intern(value))

    const getColumnView = ({ col }: { col: number }): RuntimeColumnView =>
      ({
        owner: {
          sheetName: 'Sheet1',
          col,
          columnVersion: columnVersions[col] ?? 0,
          structureVersion: 1,
          sheetColumnVersions: columnVersions,
          pages: new Map(),
        },
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 3,
        col,
        length: 4,
        columnVersion: columnVersions[col] ?? 0,
        structureVersion: 1,
        sheetColumnVersions: columnVersions,
        readTagAt(offset) {
          return col === 0 ? numericTags[offset] : textTags[offset]
        },
        readNumberAt(offset) {
          return col === 0 ? numericNumbers[offset] : 0
        },
        readStringIdAt(offset) {
          return col === 1 ? textIds[offset] : 0
        },
        readErrorAt() {
          return 0
        },
        readCellValueAt(offset) {
          if (col === 0) {
            return { tag: ValueTag.Number, value: numericNumbers[offset] }
          }
          return { tag: ValueTag.String, value: strings.get(textIds[offset]), stringId: textIds[offset] }
        },
      }) satisfies RuntimeColumnView

    const runtimeColumnStore: EngineRuntimeColumnStoreService = {
      getColumnOwner({ col }): RuntimeColumnOwner {
        return {
          sheetName: 'Sheet1',
          col,
          columnVersion: columnVersions[col] ?? 0,
          structureVersion: 1,
          sheetColumnVersions: columnVersions,
          pages: new Map(),
        }
      },
      getColumnView(request) {
        return getColumnView(request)
      },
      getColumnSlice() {
        throw new Error('not needed')
      },
      readCellValue() {
        return { tag: ValueTag.Empty, value: null }
      },
      readRangeValues() {
        return []
      },
      readRangeValueMatrix() {
        return []
      },
      normalizeStringId(stringId) {
        return strings.get(stringId).toUpperCase()
      },
      normalizeLookupText(value) {
        return value.value.toUpperCase()
      },
    }

    const counters = createEngineCounters()
    const state: Pick<EngineRuntimeState, 'workbook' | 'strings'> & { counters: typeof counters } = {
      workbook: {
        getSheet() {
          return sheetAvailable ? { columnVersions, structureVersion } : undefined
        },
      },
      strings,
      counters,
    }
    const sorted = createSortedColumnSearchService({
      state,
      runtimeColumnStore,
      columnIndexStore: createColumnIndexStore({
        state,
        runtimeColumnStore,
      }),
    })

    const numericPrepared = sorted.prepareVectorLookup({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 3,
      col: 0,
    })
    expect(numericPrepared.internalOwner).toBeUndefined()
    expect(numericPrepared.comparableKind).toBe('numeric')
    expect(
      sorted.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.Number, value: 6 },
        prepared: numericPrepared,
        matchMode: 1,
      }),
    ).toEqual({ handled: true, position: 3 })
    expect(counters.approxIndexBuilds).toBe(1)

    const textPrepared = sorted.prepareVectorLookup({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 3,
      col: 1,
    })
    expect(textPrepared.internalOwner).toBeUndefined()
    expect(textPrepared.comparableKind).toBe('text')
    expect(
      sorted.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.String, value: 'PEACH' },
        prepared: textPrepared,
        matchMode: 1,
      }),
    ).toEqual({ handled: true, position: 2 })
    expect(counters.approxIndexBuilds).toBe(2)
    expect(
      sorted.findVectorMatch({
        lookupValue: { tag: ValueTag.Number, value: 6 },
        sheetName: 'Sheet1',
        start: 'A1',
        end: 'A4',
        startRow: 0,
        endRow: 3,
        startCol: 0,
        endCol: 0,
        matchMode: 1,
      }),
    ).toEqual({ handled: true, position: 3 })
    expect(
      sorted.findVectorMatch({
        lookupValue: { tag: ValueTag.Empty, value: null },
        sheetName: 'Sheet1',
        start: 'B1',
        end: 'B4',
        startRow: 0,
        endRow: 3,
        startCol: 1,
        endCol: 1,
        matchMode: 1,
      }),
    ).toEqual({ handled: true, position: undefined })
    expect(
      sorted.findVectorMatch({
        lookupValue: { tag: ValueTag.Error, code: ErrorCode.Div0 },
        sheetName: 'Sheet1',
        start: 'A1',
        end: 'A4',
        startRow: 0,
        endRow: 3,
        startCol: 0,
        endCol: 0,
        matchMode: 1,
      }),
    ).toEqual({ handled: false })
    expect(
      sorted.findVectorMatch({
        lookupValue: { tag: ValueTag.Number, value: 1 },
        sheetName: 'Sheet1',
        start: 'A1',
        end: 'B2',
        matchMode: 1,
      }),
    ).toEqual({ handled: false })

    columnVersions[1] += 1
    textTags[2] = ValueTag.Boolean
    sorted.recordLiteralWrite({
      sheetName: 'Sheet1',
      row: 2,
      col: 1,
      oldValue: { tag: ValueTag.String, value: 'pear' },
      newValue: { tag: ValueTag.Boolean, value: false },
      oldStringId: textIds[2],
    })

    expect(
      sorted.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.String, value: 'PEACH' },
        prepared: textPrepared,
        matchMode: 1,
      }),
    ).toEqual({ handled: false })
    structureVersion = 2
    sorted.recordLiteralWrite({
      sheetName: 'Sheet1',
      row: 1,
      col: 1,
      oldValue: { tag: ValueTag.String, value: 'blueberry' },
      newValue: { tag: ValueTag.String, value: 'blueberry' },
      oldStringId: strings.intern('blueberry'),
      newStringId: strings.intern('blueberry'),
    })
    sheetAvailable = false
    sorted.recordLiteralWrite({
      sheetName: 'Sheet1',
      row: 1,
      col: 1,
      oldValue: { tag: ValueTag.String, value: 'blueberry' },
      newValue: { tag: ValueTag.String, value: 'blueberry' },
      oldStringId: strings.intern('blueberry'),
      newStringId: strings.intern('blueberry'),
    })
    sheetAvailable = true
    sorted.invalidateColumn({ sheetName: 'Sheet1', col: 1 })
    sorted.invalidateColumn({ sheetName: 'Sheet1', col: 9 })
  })

  it('refreshes prepared lookups after structural remaps and rejects unsupported lookup shapes', () => {
    const workbook = new WorkbookStore('sorted-index-prepared')
    const strings = new StringPool()
    workbook.createSheet('Sheet1')

    ;[1, 3, 5].forEach((value, index) => {
      setStoredNumber(workbook, strings, `A${index + 1}`, value)
      setStoredNumber(workbook, strings, `B${index + 1}`, 7 - index * 2)
    })
    ;['apple', 'banana', 'pear'].forEach((value, index) => {
      setStoredString(workbook, strings, `C${index + 1}`, value)
    })
    setStoredNumber(workbook, strings, 'D1', 1)
    setStoredString(workbook, strings, 'D2', 'mixed')
    setStoredNumber(workbook, strings, 'D3', 3)

    const sorted = createSorted(workbook, strings)

    const ascendingPrepared = sorted.prepareVectorLookup({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 2,
      col: 0,
    })
    expect(
      sorted.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.Boolean, value: true },
        prepared: ascendingPrepared,
        matchMode: 1,
      }),
    ).toEqual({ handled: true, position: 1 })
    expect(
      sorted.findVectorMatch({
        lookupValue: { tag: ValueTag.Empty, value: null },
        sheetName: 'Sheet1',
        start: 'A1',
        end: 'A3',
        matchMode: 1,
      }),
    ).toEqual({ handled: true, position: undefined })
    expect(
      sorted.findVectorMatch({
        lookupValue: { tag: ValueTag.Boolean, value: true },
        sheetName: 'Sheet1',
        start: 'A1',
        end: 'A3',
        matchMode: 1,
      }),
    ).toEqual({ handled: true, position: 1 })
    expect(
      sorted.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.String, value: 'bad' },
        prepared: ascendingPrepared,
        matchMode: 1,
      }),
    ).toEqual({ handled: false })

    const descendingPrepared = sorted.prepareVectorLookup({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 2,
      col: 1,
    })
    const textPrepared = sorted.prepareVectorLookup({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 2,
      col: 2,
    })
    expect(
      sorted.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.String, value: 'peach' },
        prepared: textPrepared,
        matchMode: 1,
      }),
    ).toEqual({ handled: true, position: 2 })
    expect(
      sorted.findVectorMatch({
        lookupValue: { tag: ValueTag.Empty, value: null },
        sheetName: 'Sheet1',
        start: 'C1',
        end: 'C3',
        matchMode: 1,
      }),
    ).toEqual({ handled: true, position: undefined })
    expect(
      sorted.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.Error, code: ErrorCode.Div0 },
        prepared: textPrepared,
        matchMode: 1,
      }),
    ).toEqual({ handled: false })
    expect(
      sorted.findVectorMatch({
        lookupValue: { tag: ValueTag.Number, value: 6 },
        sheetName: 'Sheet1',
        start: 'B1',
        end: 'B3',
        matchMode: -1,
      }),
    ).toEqual({ handled: true, position: 1 })
    expect(
      sorted.findVectorMatch({
        lookupValue: { tag: ValueTag.Error, code: ErrorCode.Div0 },
        sheetName: 'Sheet1',
        start: 'C1',
        end: 'C3',
        matchMode: 1,
      }),
    ).toEqual({ handled: false })

    workbook.deleteRows('Sheet1', 0, 1)
    workbook.remapSheetCells('Sheet1', 'row', (index) => (index === 0 ? undefined : index - 1))

    expect(
      sorted.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.Empty, value: null },
        prepared: descendingPrepared,
        matchMode: -1,
      }),
    ).toEqual({ handled: true, position: 3 })
    expect(
      sorted.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.Empty, value: null },
        prepared: descendingPrepared,
        matchMode: -1,
      }),
    ).toEqual({ handled: true, position: 3 })
    expect(
      sorted.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.Number, value: 10 },
        prepared: descendingPrepared,
        matchMode: 1,
      }),
    ).toEqual({ handled: false })

    expect(
      sorted.findVectorMatch({
        lookupValue: { tag: ValueTag.Number, value: 2 },
        sheetName: 'Sheet1',
        start: 'D1',
        end: 'D3',
        matchMode: 1,
      }),
    ).toEqual({ handled: false })
    expect(
      sorted.findVectorMatch({
        lookupValue: { tag: ValueTag.String, value: 'pear' },
        sheetName: 'Sheet1',
        start: 'A1',
        end: 'B2',
        matchMode: 1,
      }),
    ).toEqual({ handled: false })
  })

  it('covers manual fallback prepared approximate descriptors for numeric and text branches', () => {
    const workbook = new WorkbookStore('sorted-index-manual-prepared')
    const strings = new StringPool()
    workbook.createSheet('Sheet1')
    const runtimeColumnStore = createEngineRuntimeColumnStoreService({
      state: { workbook, strings },
    })
    const columnIndexStore = createColumnIndexStore({
      state: { workbook, strings },
      runtimeColumnStore,
    })
    const sorted = createSortedColumnSearchService({
      state: { workbook, strings },
      runtimeColumnStore,
      columnIndexStore,
    })
    const sheetColumnVersions = workbook.getSheet('Sheet1')?.columnVersions
    expect(sheetColumnVersions).toBeDefined()

    const ascendingNumericPrepared: PreparedApproximateVectorLookup = {
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 2,
      col: 0,
      length: 3,
      columnVersion: 0,
      structureVersion: workbook.getSheetStructureVersion('Sheet1'),
      sheetColumnVersions: sheetColumnVersions!,
      comparableKind: 'numeric',
      uniformStart: 1,
      uniformStep: 2,
      sortedAscending: true,
      sortedDescending: false,
      numericValues: new Float64Array([1, 3, 5]),
      textValues: undefined,
      internalOwner: undefined,
    }
    expect(
      sorted.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.Boolean, value: true },
        prepared: ascendingNumericPrepared,
        matchMode: 1,
      }),
    ).toEqual({ handled: true, position: 1 })
    expect(
      sorted.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.Number, value: 0 },
        prepared: ascendingNumericPrepared,
        matchMode: 1,
      }),
    ).toEqual({ handled: true, position: undefined })
    expect(
      sorted.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.Number, value: 10 },
        prepared: ascendingNumericPrepared,
        matchMode: 1,
      }),
    ).toEqual({ handled: true, position: 3 })

    const ascendingBinaryPrepared: PreparedApproximateVectorLookup = {
      ...ascendingNumericPrepared,
      uniformStart: undefined,
      uniformStep: undefined,
    }
    expect(
      sorted.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.Number, value: 4 },
        prepared: ascendingBinaryPrepared,
        matchMode: 1,
      }),
    ).toEqual({ handled: true, position: 2 })

    const emptyUniformPrepared: PreparedApproximateVectorLookup = {
      ...ascendingNumericPrepared,
      length: 0,
      numericValues: new Float64Array(),
    }
    expect(
      sorted.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.Number, value: 4 },
        prepared: emptyUniformPrepared,
        matchMode: 1,
      }),
    ).toEqual({ handled: true, position: undefined })

    const descendingUniformPrepared: PreparedApproximateVectorLookup = {
      ...ascendingNumericPrepared,
      comparableKind: 'numeric',
      uniformStart: 5,
      uniformStep: -2,
      sortedAscending: false,
      sortedDescending: true,
      numericValues: new Float64Array([5, 3, 1]),
    }
    expect(
      sorted.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.Number, value: 10 },
        prepared: descendingUniformPrepared,
        matchMode: -1,
      }),
    ).toEqual({ handled: true, position: undefined })
    expect(
      sorted.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.Number, value: 0 },
        prepared: descendingUniformPrepared,
        matchMode: -1,
      }),
    ).toEqual({ handled: true, position: 3 })
    expect(
      sorted.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.Number, value: 4 },
        prepared: descendingUniformPrepared,
        matchMode: -1,
      }),
    ).toEqual({ handled: true, position: 1 })

    const descendingBinaryPrepared: PreparedApproximateVectorLookup = {
      ...ascendingNumericPrepared,
      uniformStart: undefined,
      uniformStep: undefined,
      sortedAscending: false,
      sortedDescending: true,
      numericValues: new Float64Array([7, 4, 1]),
    }
    expect(
      sorted.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.Number, value: 5 },
        prepared: descendingBinaryPrepared,
        matchMode: -1,
      }),
    ).toEqual({ handled: true, position: 1 })

    const mismatchedUniformDirectionPrepared: PreparedApproximateVectorLookup = {
      ...descendingBinaryPrepared,
      uniformStart: 1,
      uniformStep: 2,
    }
    expect(
      sorted.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.Number, value: 5 },
        prepared: mismatchedUniformDirectionPrepared,
        matchMode: -1,
      }),
    ).toEqual({ handled: true, position: 1 })

    const missingNumericValues: PreparedApproximateVectorLookup = {
      ...ascendingNumericPrepared,
      numericValues: undefined,
    }
    expect(
      sorted.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.Number, value: 5 },
        prepared: missingNumericValues,
        matchMode: 1,
      }),
    ).toEqual({ handled: false })

    const textPrepared: PreparedApproximateVectorLookup = {
      ...ascendingNumericPrepared,
      col: 1,
      comparableKind: 'text',
      sortedAscending: false,
      sortedDescending: true,
      numericValues: undefined,
      textValues: ['PLUM', 'PEAR', 'APPLE'],
      uniformStart: undefined,
      uniformStep: undefined,
    }
    expect(
      sorted.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.String, value: 'PEACH' },
        prepared: textPrepared,
        matchMode: -1,
      }),
    ).toEqual({ handled: true, position: 2 })

    const missingTextValues: PreparedApproximateVectorLookup = {
      ...textPrepared,
      textValues: undefined,
    }
    expect(
      sorted.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.Empty, value: null },
        prepared: missingTextValues,
        matchMode: -1,
      }),
    ).toEqual({ handled: false })
  })
})
