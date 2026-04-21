import { describe, expect, it } from 'vitest'
import { ValueTag } from '@bilig/protocol'
import { StringPool } from '../string-pool.js'
import { WorkbookStore } from '../workbook-store.js'
import { createEngineRuntimeColumnStoreService } from '../engine/services/runtime-column-store-service.js'
import { createEngineCounters } from '../perf/engine-counters.js'

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

describe('createEngineRuntimeColumnStoreService', () => {
  it('reuses typed column slices until the source column version changes', () => {
    const workbook = new WorkbookStore('runtime-column-store')
    const strings = new StringPool()
    const counters = createEngineCounters()
    workbook.createSheet('Sheet1')

    setStoredCellValue(workbook, strings, 'Sheet1', 'A1', { tag: ValueTag.Number, value: 1 })
    setStoredCellValue(workbook, strings, 'Sheet1', 'A2', { tag: ValueTag.Boolean, value: true })
    setStoredCellValue(workbook, strings, 'Sheet1', 'A3', { tag: ValueTag.String, value: 'pear' })

    const runtimeColumnStore = createEngineRuntimeColumnStoreService({
      state: { workbook, strings, counters },
    })

    const firstSlice = runtimeColumnStore.getColumnSlice({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 2,
      col: 0,
    })
    const reusedSlice = runtimeColumnStore.getColumnSlice({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 2,
      col: 0,
    })

    expect(reusedSlice).toBe(firstSlice)
    expect(counters.columnSliceBuilds).toBe(1)
    expect(Array.from(firstSlice.tags)).toEqual([ValueTag.Number, ValueTag.Boolean, ValueTag.String])
    expect(Array.from(firstSlice.numbers)).toEqual([1, 1, 0])
    expect(firstSlice.stringIds[2]).not.toBe(0)
    expect(runtimeColumnStore.normalizeStringId(firstSlice.stringIds[2])).toBe('PEAR')

    setStoredCellValue(workbook, strings, 'Sheet1', 'A2', { tag: ValueTag.Number, value: 5 })

    const refreshedSlice = runtimeColumnStore.getColumnSlice({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 2,
      col: 0,
    })

    expect(refreshedSlice).not.toBe(firstSlice)
    expect(Array.from(refreshedSlice.tags)).toEqual([ValueTag.Number, ValueTag.Number, ValueTag.String])
    expect(Array.from(refreshedSlice.numbers)).toEqual([1, 5, 0])
    expect(counters.columnSliceBuilds).toBe(2)
  })

  it('reuses a shared column owner across overlapping view requests', () => {
    const workbook = new WorkbookStore('runtime-column-store-owner-view')
    const strings = new StringPool()
    workbook.createSheet('Sheet1')

    setStoredCellValue(workbook, strings, 'Sheet1', 'A1', { tag: ValueTag.Number, value: 1 })
    setStoredCellValue(workbook, strings, 'Sheet1', 'A2', { tag: ValueTag.String, value: 'pear' })
    setStoredCellValue(workbook, strings, 'Sheet1', 'A4', { tag: ValueTag.Boolean, value: true })

    const runtimeColumnStore = createEngineRuntimeColumnStoreService({
      state: { workbook, strings },
    })

    const firstView = runtimeColumnStore.getColumnView({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 2,
      col: 0,
    })
    const overlappingView = runtimeColumnStore.getColumnView({
      sheetName: 'Sheet1',
      rowStart: 1,
      rowEnd: 3,
      col: 0,
    })

    expect(overlappingView.owner).toBe(firstView.owner)
    expect(firstView.readCellValueAt(0)).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(firstView.readCellValueAt(1)).toEqual({
      tag: ValueTag.String,
      value: 'pear',
      stringId: expect.any(Number),
    })
    expect(overlappingView.readCellValueAt(0)).toEqual({
      tag: ValueTag.String,
      value: 'pear',
      stringId: expect.any(Number),
    })
    expect(overlappingView.readCellValueAt(1)).toEqual({ tag: ValueTag.Empty })
    expect(overlappingView.readCellValueAt(2)).toEqual({ tag: ValueTag.Boolean, value: true })
  })

  it('materializes row-major cell values from cached column slices', () => {
    const workbook = new WorkbookStore('runtime-column-store-read')
    const strings = new StringPool()
    workbook.createSheet('Sheet1')

    setStoredCellValue(workbook, strings, 'Sheet1', 'A1', { tag: ValueTag.Number, value: 1 })
    setStoredCellValue(workbook, strings, 'Sheet1', 'B1', { tag: ValueTag.String, value: 'pear' })
    setStoredCellValue(workbook, strings, 'Sheet1', 'A2', { tag: ValueTag.Boolean, value: false })

    const runtimeColumnStore = createEngineRuntimeColumnStoreService({
      state: { workbook, strings },
    })

    expect(runtimeColumnStore.readCellValue('Sheet1', 0, 1)).toEqual({
      tag: ValueTag.String,
      value: 'pear',
      stringId: expect.any(Number),
    })
    expect(
      runtimeColumnStore.readRangeValues({
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 1,
        colStart: 0,
        colEnd: 1,
      }),
    ).toEqual([
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.String, value: 'pear', stringId: expect.any(Number) },
      { tag: ValueTag.Boolean, value: false },
      { tag: ValueTag.Empty },
    ])
    expect(
      runtimeColumnStore.readRangeValueMatrix({
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 1,
        colStart: 0,
        colEnd: 1,
      }),
    ).toEqual([
      [
        { tag: ValueTag.Number, value: 1 },
        { tag: ValueTag.String, value: 'pear', stringId: expect.any(Number) },
      ],
      [{ tag: ValueTag.Boolean, value: false }, { tag: ValueTag.Empty }],
    ])
  })

  it('invalidates cached column slices after structural row remaps', () => {
    const workbook = new WorkbookStore('runtime-column-store-structural')
    const strings = new StringPool()
    workbook.createSheet('Sheet1')

    setStoredCellValue(workbook, strings, 'Sheet1', 'C3', { tag: ValueTag.Boolean, value: false })
    setStoredCellValue(workbook, strings, 'Sheet1', 'C4', { tag: ValueTag.String, value: 'north' })

    const runtimeColumnStore = createEngineRuntimeColumnStoreService({
      state: { workbook, strings },
    })

    expect(runtimeColumnStore.readCellValue('Sheet1', 2, 2)).toEqual({
      tag: ValueTag.Boolean,
      value: false,
    })

    workbook.deleteRows('Sheet1', 0, 1)
    workbook.remapSheetCells('Sheet1', 'row', (index) => (index === 0 ? undefined : index - 1))

    expect(runtimeColumnStore.readCellValue('Sheet1', 1, 2)).toEqual({
      tag: ValueTag.Boolean,
      value: false,
    })
    expect(runtimeColumnStore.readCellValue('Sheet1', 2, 2)).toEqual({
      tag: ValueTag.String,
      value: 'north',
      stringId: expect.any(Number),
    })
  })

  it('treats missing sheets and unknown raw tags as empty cells', () => {
    const workbook = new WorkbookStore('runtime-column-store-empty')
    const strings = new StringPool()
    workbook.createSheet('Sheet1')

    const cellIndex = workbook.ensureCell('Sheet1', 'A1')
    workbook.cellStore.tags[cellIndex] = 99

    const runtimeColumnStore = createEngineRuntimeColumnStoreService({
      state: { workbook, strings },
    })

    expect(runtimeColumnStore.readCellValue('Missing', 0, 0)).toEqual({
      tag: ValueTag.Empty,
    })
    expect(runtimeColumnStore.readCellValue('Sheet1', 0, 0)).toEqual({
      tag: ValueTag.Empty,
    })
    expect(
      runtimeColumnStore.readRangeValueMatrix({
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 0,
        colStart: 0,
        colEnd: 0,
      }),
    ).toEqual([[{ tag: ValueTag.Empty }]])
  })

  it('materializes error cells and falls back to raw lookup text when no string id is present', () => {
    const workbook = new WorkbookStore('runtime-column-store-errors')
    const strings = new StringPool()
    workbook.createSheet('Sheet1')

    const cellIndex = workbook.ensureCell('Sheet1', 'A1')
    workbook.cellStore.tags[cellIndex] = ValueTag.Error
    workbook.cellStore.errors[cellIndex] = 42

    const runtimeColumnStore = createEngineRuntimeColumnStoreService({
      state: { workbook, strings },
    })

    expect(runtimeColumnStore.readCellValue('Sheet1', 0, 0)).toEqual({
      tag: ValueTag.Error,
      code: 42,
    })
    expect(
      runtimeColumnStore.readRangeValueMatrix({
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 0,
        colStart: 0,
        colEnd: 0,
      }),
    ).toEqual([[{ tag: ValueTag.Error, code: 42 }]])
    expect(
      runtimeColumnStore.normalizeLookupText({
        tag: ValueTag.String,
        value: 'pear',
        stringId: 0,
      }),
    ).toBe('PEAR')

    const slice = runtimeColumnStore.getColumnSlice({
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 0,
      col: 0,
    })
    slice.tags[0] = 99

    expect(runtimeColumnStore.readCellValue('Sheet1', 0, 0)).toEqual({
      tag: ValueTag.Error,
      code: 42,
    })
  })
})
