import { describe, expect, it, vi } from 'vitest'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import { StringPool } from '../string-pool.js'
import { WorkbookStore } from '../workbook-store.js'
import { createColumnIndexStore } from '../indexes/column-index-store.js'
import type { EngineRuntimeState, PreparedExactVectorLookup } from '../engine/runtime-state.js'
import { createExactColumnIndexService } from '../engine/services/exact-column-index-service.js'
import { createEngineRuntimeColumnStoreService } from '../engine/services/runtime-column-store-service.js'
import type {
  EngineRuntimeColumnStoreService,
  RuntimeColumnOwner,
  RuntimeColumnView,
} from '../engine/services/runtime-column-store-service.js'

function setStoredCellValue(
  workbook: WorkbookStore,
  strings: StringPool,
  sheetName: string,
  address: string,
  value: { tag: ValueTag.String; value: string } | { tag: ValueTag.Number; value: number } | { tag: ValueTag.Boolean; value: boolean },
): void {
  const cellIndex = workbook.ensureCell(sheetName, address)
  workbook.cellStore.setValue(cellIndex, value, value.tag === ValueTag.String ? strings.intern(value.value) : 0)
}

function createExact(workbook: WorkbookStore, strings: StringPool) {
  const runtimeColumnStore = createEngineRuntimeColumnStoreService({
    state: { workbook, strings },
  })
  const columnIndexStore = createColumnIndexStore({
    state: { workbook, strings },
    runtimeColumnStore,
  })
  return createExactColumnIndexService({
    state: { workbook, strings },
    runtimeColumnStore,
    columnIndexStore,
  })
}

describe('createExactColumnIndexService', () => {
  it('serves exact matches from a primed column index and invalidates by column version', () => {
    const workbook = new WorkbookStore('exact-index')
    const strings = new StringPool()
    workbook.createSheet('Sheet1')

    setStoredCellValue(workbook, strings, 'Sheet1', 'A1', { tag: ValueTag.String, value: 'pear' })
    setStoredCellValue(workbook, strings, 'Sheet1', 'A2', {
      tag: ValueTag.String,
      value: 'apple',
    })
    setStoredCellValue(workbook, strings, 'Sheet1', 'A3', { tag: ValueTag.String, value: 'pear' })

    const exact = createExact(workbook, strings)

    exact.primeColumnIndex({ sheetName: 'Sheet1', rowStart: 0, rowEnd: 2, col: 0 })

    expect(
      exact.findVectorMatch({
        lookupValue: { tag: ValueTag.String, value: 'PEAR' },
        sheetName: 'Sheet1',
        start: 'A1',
        end: 'A3',
        startRow: 0,
        endRow: 2,
        startCol: 0,
        endCol: 0,
        searchMode: 1,
      }),
    ).toEqual({ handled: true, position: 1 })

    setStoredCellValue(workbook, strings, 'Sheet1', 'A1', {
      tag: ValueTag.String,
      value: 'banana',
    })

    expect(
      exact.findVectorMatch({
        lookupValue: { tag: ValueTag.String, value: 'PEAR' },
        sheetName: 'Sheet1',
        start: 'A1',
        end: 'A3',
        startRow: 0,
        endRow: 2,
        startCol: 0,
        endCol: 0,
        searchMode: 1,
      }),
    ).toEqual({ handled: true, position: 3 })
  })

  it('refreshes prepared lookups after structural row remaps and handles text and mixed columns', () => {
    const workbook = new WorkbookStore('exact-index-prepared')
    const strings = new StringPool()
    workbook.createSheet('Sheet1')

    ;[1, 3, 5].forEach((value, index) => {
      setStoredCellValue(workbook, strings, 'Sheet1', `A${index + 1}`, {
        tag: ValueTag.Number,
        value,
      })
    })
    ;['apple', 'pear', 'plum'].forEach((value, index) => {
      setStoredCellValue(workbook, strings, 'Sheet1', `B${index + 1}`, {
        tag: ValueTag.String,
        value,
      })
    })
    workbook.ensureCell('Sheet1', 'C1')
    setStoredCellValue(workbook, strings, 'Sheet1', 'C2', {
      tag: ValueTag.String,
      value: 'pear',
    })
    setStoredCellValue(workbook, strings, 'Sheet1', 'C3', {
      tag: ValueTag.Boolean,
      value: false,
    })

    const exact = createExact(workbook, strings)

    const numericPrepared = exact.prepareVectorLookup({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 2,
      col: 0,
    })
    expect(
      exact.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.Number, value: 3 },
        prepared: numericPrepared,
        searchMode: 1,
      }),
    ).toEqual({ handled: true, position: 2 })
    expect(
      exact.findVectorMatch({
        lookupValue: { tag: ValueTag.Number, value: 3 },
        sheetName: 'Sheet1',
        start: 'A1',
        end: 'A3',
        searchMode: 1,
      }),
    ).toEqual({ handled: true, position: 2 })
    expect(
      exact.findVectorMatch({
        lookupValue: { tag: ValueTag.Boolean, value: false },
        sheetName: 'Sheet1',
        start: 'A1',
        end: 'A3',
        searchMode: 1,
      }),
    ).toEqual({ handled: true, position: undefined })
    expect(
      exact.findVectorMatch({
        lookupValue: { tag: ValueTag.Error, code: ErrorCode.Div0 },
        sheetName: 'Sheet1',
        start: 'A1',
        end: 'A3',
        searchMode: 1,
      }),
    ).toEqual({ handled: false })

    const pristineTextPrepared = exact.prepareVectorLookup({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 2,
      col: 1,
    })
    expect(
      exact.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.String, value: 'PEAR' },
        prepared: pristineTextPrepared,
        searchMode: 1,
      }),
    ).toEqual({ handled: true, position: 2 })
    expect(
      exact.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.Boolean, value: false },
        prepared: pristineTextPrepared,
        searchMode: 1,
      }),
    ).toEqual({ handled: true, position: undefined })
    expect(
      exact.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.Error, code: ErrorCode.Div0 },
        prepared: pristineTextPrepared,
        searchMode: 1,
      }),
    ).toEqual({ handled: false })
    expect(
      exact.findVectorMatch({
        lookupValue: { tag: ValueTag.String, value: 'PEAR' },
        sheetName: 'Sheet1',
        start: 'B1',
        end: 'B3',
        searchMode: 1,
      }),
    ).toEqual({ handled: true, position: 2 })
    expect(
      exact.findVectorMatch({
        lookupValue: { tag: ValueTag.Boolean, value: false },
        sheetName: 'Sheet1',
        start: 'B1',
        end: 'B3',
        searchMode: 1,
      }),
    ).toEqual({ handled: true, position: undefined })
    expect(
      exact.findVectorMatch({
        lookupValue: { tag: ValueTag.Error, code: ErrorCode.Div0 },
        sheetName: 'Sheet1',
        start: 'B1',
        end: 'B3',
        searchMode: 1,
      }),
    ).toEqual({ handled: false })

    workbook.deleteRows('Sheet1', 0, 1)
    workbook.remapSheetCells('Sheet1', 'row', (index) => (index === 0 ? undefined : index - 1))

    expect(
      exact.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.Number, value: 5 },
        prepared: numericPrepared,
        searchMode: 1,
      }),
    ).toEqual({ handled: true, position: 2 })

    const textPrepared = exact.prepareVectorLookup({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 2,
      col: 1,
    })
    expect(
      exact.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.String, value: 'PEAR' },
        prepared: textPrepared,
        searchMode: 1,
      }),
    ).toEqual({ handled: true, position: 1 })
    expect(
      exact.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.Boolean, value: false },
        prepared: textPrepared,
        searchMode: 1,
      }),
    ).toEqual({ handled: true, position: undefined })
    expect(
      exact.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.Error, code: ErrorCode.Div0 },
        prepared: textPrepared,
        searchMode: 1,
      }),
    ).toEqual({ handled: false })
    const mixedPrepared = exact.prepareVectorLookup({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 2,
      col: 2,
    })
    expect(
      exact.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.Boolean, value: false },
        prepared: mixedPrepared,
        searchMode: 1,
      }),
    ).toEqual({ handled: true, position: 2 })
    expect(
      exact.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.Error, code: ErrorCode.Div0 },
        prepared: mixedPrepared,
        searchMode: 1,
      }),
    ).toEqual({ handled: false })

    expect(
      exact.findVectorMatch({
        lookupValue: { tag: ValueTag.String, value: 'pear' },
        sheetName: 'Sheet1',
        start: 'A1',
        end: 'B2',
        searchMode: 1,
      }),
    ).toEqual({ handled: false })
    expect(
      exact.findVectorMatch({
        lookupValue: { tag: ValueTag.Error, code: ErrorCode.Div0 },
        sheetName: 'Sheet1',
        start: 'A1',
        end: 'A2',
        searchMode: 1,
      }),
    ).toEqual({ handled: false })
  })

  it('updates cached exact lookups incrementally for literal writes and explicit invalidation', () => {
    const workbook = new WorkbookStore('exact-index-incremental')
    const strings = new StringPool()
    workbook.createSheet('Sheet1')

    ;[1, 2, 3].forEach((value, index) => {
      setStoredCellValue(workbook, strings, 'Sheet1', `A${index + 1}`, {
        tag: ValueTag.Number,
        value,
      })
    })

    const exact = createExact(workbook, strings)

    const prepared = exact.prepareVectorLookup({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 2,
      col: 0,
    })

    setStoredCellValue(workbook, strings, 'Sheet1', 'A3', {
      tag: ValueTag.Number,
      value: 30,
    })
    exact.recordLiteralWrite({
      sheetName: 'Sheet1',
      row: 2,
      col: 0,
      oldValue: { tag: ValueTag.Number, value: 3 },
      newValue: { tag: ValueTag.Number, value: 30 },
    })

    expect(
      exact.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.Number, value: 30 },
        prepared,
        searchMode: 1,
      }),
    ).toEqual({ handled: true, position: 3 })

    exact.invalidateColumn({ sheetName: 'Sheet1', col: 0 })

    expect(
      exact.findVectorMatch({
        lookupValue: { tag: ValueTag.Number, value: 30 },
        sheetName: 'Sheet1',
        start: 'A1',
        end: 'A3',
        searchMode: 1,
      }),
    ).toEqual({ handled: true, position: 3 })
  })

  it('reuses prepared exact lookups after incremental writes without rematerializing the column', () => {
    const workbook = new WorkbookStore('exact-index-slice-reuse')
    const strings = new StringPool()
    workbook.createSheet('Sheet1')

    ;[1, 2, 3].forEach((value, index) => {
      setStoredCellValue(workbook, strings, 'Sheet1', `A${index + 1}`, {
        tag: ValueTag.Number,
        value,
      })
    })

    const runtimeColumnStore = createEngineRuntimeColumnStoreService({
      state: { workbook, strings },
    })
    const getColumnSliceSpy = vi.spyOn(runtimeColumnStore, 'getColumnSlice')
    const columnIndexStore = createColumnIndexStore({
      state: { workbook, strings },
      runtimeColumnStore,
    })
    const exact = createExactColumnIndexService({
      state: { workbook, strings },
      runtimeColumnStore,
      columnIndexStore,
    })

    const prepared = exact.prepareVectorLookup({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 2,
      col: 0,
    })
    const sliceCallsAfterPrepare = getColumnSliceSpy.mock.calls.length

    setStoredCellValue(workbook, strings, 'Sheet1', 'A3', {
      tag: ValueTag.Number,
      value: 30,
    })
    exact.recordLiteralWrite({
      sheetName: 'Sheet1',
      row: 2,
      col: 0,
      oldValue: { tag: ValueTag.Number, value: 3 },
      newValue: { tag: ValueTag.Number, value: 30 },
    })

    expect(
      exact.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.Number, value: 30 },
        prepared,
        searchMode: 1,
      }),
    ).toEqual({ handled: true, position: 3 })
    expect(getColumnSliceSpy.mock.calls.length).toBe(sliceCallsAfterPrepare)
  })

  it('prepares overlapping exact lookups from shared column owners instead of copied slices', () => {
    const workbook = new WorkbookStore('exact-index-column-owner')
    const strings = new StringPool()
    workbook.createSheet('Sheet1')

    ;[1, 2, 3, 4].forEach((value, index) => {
      setStoredCellValue(workbook, strings, 'Sheet1', `A${index + 1}`, {
        tag: ValueTag.Number,
        value,
      })
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
    const exact = createExactColumnIndexService({
      state: { workbook, strings },
      runtimeColumnStore,
      columnIndexStore,
    })

    exact.prepareVectorLookup({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 2,
      col: 0,
    })
    exact.prepareVectorLookup({
      sheetName: 'Sheet1',
      rowStart: 1,
      rowEnd: 3,
      col: 0,
    })

    expect(getColumnOwnerSpy).toHaveBeenCalledTimes(1)
    expect(getColumnViewSpy).not.toHaveBeenCalled()
    expect(getColumnSliceSpy).not.toHaveBeenCalled()
  })

  it('uses owner-backed prepared exact lookups for numeric and text subranges after writes', () => {
    const workbook = new WorkbookStore('exact-index-owner-prepared')
    const strings = new StringPool()
    workbook.createSheet('Sheet1')

    ;[1, 3, 5, 7].forEach((value, index) => {
      setStoredCellValue(workbook, strings, 'Sheet1', `A${index + 1}`, { tag: ValueTag.Number, value })
    })
    ;['apple', 'banana', 'pear', 'plum'].forEach((value, index) => {
      setStoredCellValue(workbook, strings, 'Sheet1', `B${index + 1}`, { tag: ValueTag.String, value })
    })

    const exact = createExact(workbook, strings)

    const numericPrepared = exact.prepareVectorLookup({
      sheetName: 'Sheet1',
      rowStart: 1,
      rowEnd: 3,
      col: 0,
    })
    expect(numericPrepared.internalOwner).toBeDefined()
    expect(numericPrepared.comparableKind).toBe('numeric')
    expect(
      exact.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.Number, value: 5 },
        prepared: numericPrepared,
        searchMode: 1,
      }),
    ).toEqual({ handled: true, position: 2 })

    const textPrepared = exact.prepareVectorLookup({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 2,
      col: 1,
    })
    expect(textPrepared.internalOwner).toBeDefined()
    expect(textPrepared.comparableKind).toBe('text')
    expect(
      exact.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.String, value: 'PEAR' },
        prepared: textPrepared,
        searchMode: 1,
      }),
    ).toEqual({ handled: true, position: 3 })

    setStoredCellValue(workbook, strings, 'Sheet1', 'B2', { tag: ValueTag.String, value: 'blueberry' })
    exact.recordLiteralWrite({
      sheetName: 'Sheet1',
      row: 1,
      col: 1,
      oldValue: { tag: ValueTag.String, value: 'banana' },
      newValue: { tag: ValueTag.String, value: 'blueberry' },
      oldStringId: strings.intern('banana'),
      newStringId: strings.intern('blueberry'),
    })

    expect(
      exact.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.String, value: 'BLUEBERRY' },
        prepared: textPrepared,
        searchMode: 1,
      }),
    ).toEqual({ handled: true, position: 2 })
  })

  it('covers fallback exact indices when owner coverage is unavailable and invalidates on type-changing writes', () => {
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
      normalizeStringId(stringId) {
        return strings.get(stringId).toUpperCase()
      },
      normalizeLookupText(value) {
        return value.value.toUpperCase()
      },
    }

    const state: Pick<EngineRuntimeState, 'workbook' | 'strings'> = {
      workbook: {
        getSheet() {
          return sheetAvailable ? { columnVersions, structureVersion } : undefined
        },
        getSheetStructureVersion() {
          return structureVersion
        },
      },
      strings,
    }
    const exact = createExactColumnIndexService({
      state,
      runtimeColumnStore,
      columnIndexStore: createColumnIndexStore({
        state,
        runtimeColumnStore,
      }),
    })

    const numericPrepared = exact.prepareVectorLookup({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 3,
      col: 0,
    })
    expect(numericPrepared.internalOwner).toBeUndefined()
    expect(numericPrepared.comparableKind).toBe('numeric')
    expect(
      exact.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.Number, value: 5 },
        prepared: numericPrepared,
        searchMode: 1,
      }),
    ).toEqual({ handled: true, position: 3 })

    const textPrepared = exact.prepareVectorLookup({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 3,
      col: 1,
    })
    expect(textPrepared.internalOwner).toBeUndefined()
    expect(textPrepared.comparableKind).toBe('text')
    expect(
      exact.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.String, value: 'PEAR' },
        prepared: textPrepared,
        searchMode: 1,
      }),
    ).toEqual({ handled: true, position: 3 })
    expect(
      exact.findVectorMatch({
        lookupValue: { tag: ValueTag.Number, value: 5 },
        sheetName: 'Sheet1',
        start: 'A1',
        end: 'A4',
        startRow: 0,
        endRow: 3,
        startCol: 0,
        endCol: 0,
        searchMode: 1,
      }),
    ).toEqual({ handled: true, position: 3 })
    expect(
      exact.findVectorMatch({
        lookupValue: { tag: ValueTag.String, value: 'PEAR' },
        sheetName: 'Sheet1',
        start: 'B1',
        end: 'B4',
        startRow: 0,
        endRow: 3,
        startCol: 1,
        endCol: 1,
        searchMode: -1,
      }),
    ).toEqual({ handled: true, position: 3 })
    expect(
      exact.findVectorMatch({
        lookupValue: { tag: ValueTag.Error, code: ErrorCode.Div0 },
        sheetName: 'Sheet1',
        start: 'A1',
        end: 'A4',
        startRow: 0,
        endRow: 3,
        startCol: 0,
        endCol: 0,
        searchMode: 1,
      }),
    ).toEqual({ handled: false })
    expect(
      exact.findVectorMatch({
        lookupValue: { tag: ValueTag.String, value: 'PEAR' },
        sheetName: 'Sheet1',
        start: 'A1',
        end: 'B2',
        searchMode: 1,
      }),
    ).toEqual({ handled: false })

    columnVersions[1] += 1
    textTags[2] = ValueTag.Boolean
    exact.recordLiteralWrite({
      sheetName: 'Sheet1',
      row: 2,
      col: 1,
      oldValue: { tag: ValueTag.String, value: 'pear' },
      newValue: { tag: ValueTag.Boolean, value: false },
      oldStringId: textIds[2],
    })

    expect(
      exact.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.String, value: 'PEAR' },
        prepared: textPrepared,
        searchMode: 1,
      }),
    ).toEqual({ handled: true, position: undefined })
    structureVersion = 2
    exact.recordLiteralWrite({
      sheetName: 'Sheet1',
      row: 1,
      col: 1,
      oldValue: { tag: ValueTag.String, value: 'blueberry' },
      newValue: { tag: ValueTag.String, value: 'blueberry' },
      oldStringId: strings.intern('blueberry'),
      newStringId: strings.intern('blueberry'),
    })
    exact.recordLiteralWrite({
      sheetName: 'Sheet1',
      row: 1,
      col: 1,
      oldValue: { tag: ValueTag.String, value: 'blueberry' },
      newValue: { tag: ValueTag.String, value: 'blueberry' },
      oldStringId: strings.intern('blueberry'),
      newStringId: strings.intern('blueberry'),
    })
    sheetAvailable = false
    exact.recordLiteralWrite({
      sheetName: 'Sheet1',
      row: 1,
      col: 1,
      oldValue: { tag: ValueTag.String, value: 'blueberry' },
      newValue: { tag: ValueTag.String, value: 'blueberry' },
      oldStringId: strings.intern('blueberry'),
      newStringId: strings.intern('blueberry'),
    })
    sheetAvailable = true
    exact.invalidateColumn({ sheetName: 'Sheet1', col: 1 })
    exact.invalidateColumn({ sheetName: 'Sheet1', col: 9 })
  })

  it('updates fallback cached exact indices in place for numeric and text writes', () => {
    const workbook = new WorkbookStore('exact-index-fallback-cache-updates')
    const strings = new StringPool()
    workbook.createSheet('Sheet1')

    ;[1, 3, 5].forEach((value, index) => {
      setStoredCellValue(workbook, strings, 'Sheet1', `A${index + 1}`, { tag: ValueTag.Number, value })
    })
    ;['apple', 'banana', 'pear'].forEach((value, index) => {
      setStoredCellValue(workbook, strings, 'Sheet1', `B${index + 1}`, { tag: ValueTag.String, value })
    })

    const baseRuntimeColumnStore = createEngineRuntimeColumnStoreService({
      state: { workbook, strings },
    })
    const runtimeColumnStore: EngineRuntimeColumnStoreService = {
      ...baseRuntimeColumnStore,
      getColumnOwner({ sheetName, col }): RuntimeColumnOwner {
        const baseOwner = baseRuntimeColumnStore.getColumnOwner({ sheetName, col })
        return {
          ...baseOwner,
          pages: new Map(),
        }
      },
    }
    const columnIndexStore = createColumnIndexStore({
      state: { workbook, strings },
      runtimeColumnStore,
    })
    const exact = createExactColumnIndexService({
      state: { workbook, strings },
      runtimeColumnStore,
      columnIndexStore,
    })

    const numericPrepared = exact.prepareVectorLookup({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 2,
      col: 0,
    })
    const textPrepared = exact.prepareVectorLookup({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 2,
      col: 1,
    })

    setStoredCellValue(workbook, strings, 'Sheet1', 'A2', { tag: ValueTag.Number, value: 1 })
    exact.recordLiteralWrite({
      sheetName: 'Sheet1',
      row: 1,
      col: 0,
      oldValue: { tag: ValueTag.Number, value: 3 },
      newValue: { tag: ValueTag.Number, value: 1 },
    })
    expect(
      exact.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.Number, value: 1 },
        prepared: numericPrepared,
        searchMode: -1,
      }),
    ).toEqual({ handled: true, position: 2 })

    setStoredCellValue(workbook, strings, 'Sheet1', 'B2', { tag: ValueTag.String, value: 'apple' })
    exact.recordLiteralWrite({
      sheetName: 'Sheet1',
      row: 1,
      col: 1,
      oldValue: { tag: ValueTag.String, value: 'banana' },
      newValue: { tag: ValueTag.String, value: 'apple' },
      oldStringId: strings.intern('banana'),
      newStringId: strings.intern('apple'),
    })
    expect(
      exact.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.String, value: 'APPLE' },
        prepared: textPrepared,
        searchMode: -1,
      }),
    ).toEqual({ handled: true, position: 2 })

    exact.recordLiteralWrite({
      sheetName: 'Sheet1',
      row: 0,
      col: 1,
      oldValue: { tag: ValueTag.String, value: 'apple' },
      newValue: { tag: ValueTag.String, value: 'apple' },
      oldStringId: strings.intern('apple'),
      newStringId: strings.intern('apple'),
    })
    exact.recordLiteralWrite({
      sheetName: 'Sheet1',
      row: 99,
      col: 0,
      oldValue: { tag: ValueTag.Number, value: 0 },
      newValue: { tag: ValueTag.Number, value: 0 },
    })
  })

  it('covers manual fallback prepared exact descriptors for numeric text and mixed branches', () => {
    const workbook = new WorkbookStore('exact-index-manual-prepared')
    const strings = new StringPool()
    workbook.createSheet('Sheet1')
    const exact = createExact(workbook, strings)
    const sheetColumnVersions = workbook.getSheet('Sheet1')?.columnVersions
    expect(sheetColumnVersions).toBeDefined()

    const numericPrepared: PreparedExactVectorLookup = {
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 2,
      col: 0,
      length: 3,
      columnVersion: 0,
      structureVersion: workbook.getSheetStructureVersion('Sheet1'),
      sheetColumnVersions: sheetColumnVersions!,
      comparableKind: 'numeric',
      uniformStart: undefined,
      uniformStep: undefined,
      firstPositions: new Map([['n:1', 0]]),
      lastPositions: new Map([['n:1', 2]]),
      firstNumericPositions: new Map([[1, 0]]),
      lastNumericPositions: new Map([[1, 2]]),
      firstTextPositions: undefined,
      lastTextPositions: undefined,
      internalOwner: undefined,
    }
    expect(
      exact.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.String, value: 'bad' },
        prepared: numericPrepared,
        searchMode: 1,
      }),
    ).toEqual({ handled: true, position: undefined })
    expect(
      exact.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.Number, value: 1 },
        prepared: numericPrepared,
        searchMode: -1,
      }),
    ).toEqual({ handled: true, position: 3 })

    const textPrepared: PreparedExactVectorLookup = {
      ...numericPrepared,
      col: 1,
      comparableKind: 'text',
      firstPositions: new Map([['s:PEAR', 1]]),
      lastPositions: new Map([['s:PEAR', 2]]),
      firstNumericPositions: undefined,
      lastNumericPositions: undefined,
      firstTextPositions: new Map([['PEAR', 1]]),
      lastTextPositions: new Map([['PEAR', 2]]),
    }
    expect(
      exact.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.Error, code: ErrorCode.Div0 },
        prepared: textPrepared,
        searchMode: 1,
      }),
    ).toEqual({ handled: false })
    expect(
      exact.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.Boolean, value: false },
        prepared: textPrepared,
        searchMode: 1,
      }),
    ).toEqual({ handled: true, position: undefined })

    const mixedPrepared: PreparedExactVectorLookup = {
      ...numericPrepared,
      col: 2,
      comparableKind: 'mixed',
      firstPositions: new Map([['b:1', 0]]),
      lastPositions: new Map([['b:1', 2]]),
      firstNumericPositions: undefined,
      lastNumericPositions: undefined,
      firstTextPositions: undefined,
      lastTextPositions: undefined,
    }
    expect(
      exact.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.Error, code: ErrorCode.Div0 },
        prepared: mixedPrepared,
        searchMode: 1,
      }),
    ).toEqual({ handled: false })
    expect(
      exact.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.Boolean, value: true },
        prepared: mixedPrepared,
        searchMode: -1,
      }),
    ).toEqual({ handled: true, position: 3 })
  })

  it('falls back to rebuilds when incremental literal writes invalidate comparable kinds or structure versions', () => {
    const workbook = new WorkbookStore('exact-index-rebuild-fallbacks')
    const strings = new StringPool()
    workbook.createSheet('Sheet1')

    ;[1, 2, 3].forEach((value, index) => {
      setStoredCellValue(workbook, strings, 'Sheet1', `A${index + 1}`, {
        tag: ValueTag.Number,
        value,
      })
    })

    const exact = createExact(workbook, strings)

    const prepared = exact.prepareVectorLookup({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 2,
      col: 0,
    })

    const pearId = strings.intern('pear')
    setStoredCellValue(workbook, strings, 'Sheet1', 'A2', {
      tag: ValueTag.String,
      value: 'pear',
    })
    exact.recordLiteralWrite({
      sheetName: 'Sheet1',
      row: 1,
      col: 0,
      oldValue: { tag: ValueTag.Number, value: 2 },
      newValue: { tag: ValueTag.String, value: 'pear' },
      newStringId: pearId,
    })

    expect(
      exact.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.String, value: 'PEAR' },
        prepared,
        searchMode: 1,
      }),
    ).toEqual({ handled: true, position: 2 })

    workbook.deleteRows('Sheet1', 0, 1)
    workbook.remapSheetCells('Sheet1', 'row', (index) => (index === 0 ? undefined : index - 1))
    setStoredCellValue(workbook, strings, 'Sheet1', 'A2', {
      tag: ValueTag.Number,
      value: 30,
    })

    exact.recordLiteralWrite({
      sheetName: 'Sheet1',
      row: 1,
      col: 0,
      oldValue: { tag: ValueTag.Number, value: 3 },
      newValue: { tag: ValueTag.Number, value: 30 },
    })

    expect(
      exact.findPreparedVectorMatch({
        lookupValue: { tag: ValueTag.Number, value: 30 },
        prepared,
        searchMode: 1,
      }),
    ).toEqual({ handled: true, position: 2 })
  })
})
