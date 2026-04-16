import { describe, expect, it, vi } from 'vitest'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import { StringPool } from '../string-pool.js'
import { WorkbookStore } from '../workbook-store.js'
import { createExactColumnIndexService } from '../engine/services/exact-column-index-service.js'
import { createEngineRuntimeColumnStoreService } from '../engine/services/runtime-column-store-service.js'

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

    const runtimeColumnStore = createEngineRuntimeColumnStoreService({
      state: { workbook, strings },
    })
    const exact = createExactColumnIndexService({
      state: { workbook, strings },
      runtimeColumnStore,
    })

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

    const runtimeColumnStore = createEngineRuntimeColumnStoreService({
      state: { workbook, strings },
    })
    const exact = createExactColumnIndexService({
      state: { workbook, strings },
      runtimeColumnStore,
    })

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

    const runtimeColumnStore = createEngineRuntimeColumnStoreService({
      state: { workbook, strings },
    })
    const exact = createExactColumnIndexService({
      state: { workbook, strings },
      runtimeColumnStore,
    })

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
    const exact = createExactColumnIndexService({
      state: { workbook, strings },
      runtimeColumnStore,
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

    const runtimeColumnStore = createEngineRuntimeColumnStoreService({
      state: { workbook, strings },
    })
    const exact = createExactColumnIndexService({
      state: { workbook, strings },
      runtimeColumnStore,
    })

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
