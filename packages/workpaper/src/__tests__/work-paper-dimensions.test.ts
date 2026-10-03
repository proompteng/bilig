import { describe, expect, it, vi } from 'vitest'
import { ValueTag } from '@bilig/protocol'

import { WorkPaper, type WorkPaperCellAddress } from '../index.js'

function cell(sheet: number, row: number, col: number): WorkPaperCellAddress {
  return { sheet, row, col }
}

interface TestSheetDimensionCache {
  scan(...args: unknown[]): unknown
  invalidate(sheetId: number): void
}

interface TestWorkPaperEngine {
  getCell(...args: unknown[]): unknown
  getCellByIndex(...args: unknown[]): unknown
}

function isTestWorkPaperEngine(value: unknown): value is TestWorkPaperEngine {
  return (
    typeof value === 'object' &&
    value !== null &&
    typeof Reflect.get(value, 'getCell') === 'function' &&
    typeof Reflect.get(value, 'getCellByIndex') === 'function'
  )
}

function hasDimensionScanner(value: unknown): value is TestSheetDimensionCache {
  return (
    typeof value === 'object' &&
    value !== null &&
    typeof Reflect.get(value, 'scan') === 'function' &&
    typeof Reflect.get(value, 'invalidate') === 'function'
  )
}

function getDimensionCache(workbook: WorkPaper): TestSheetDimensionCache {
  const cache: unknown = Reflect.get(workbook, 'sheetDimensionCache')
  if (!hasDimensionScanner(cache)) {
    throw new Error('Expected WorkPaper to expose a sheet dimension cache in tests')
  }
  return cache
}

function trackDimensionScans(workbook: WorkPaper): { readonly count: number; restore: () => void } {
  const cache = getDimensionCache(workbook)
  const spy = vi.spyOn(cache, 'scan')
  return {
    get count() {
      return spy.mock.calls.length
    },
    restore: () => {
      spy.mockRestore()
    },
  }
}

function getTestEngine(workbook: WorkPaper): TestWorkPaperEngine {
  const engine: unknown = Reflect.get(workbook, 'engine')
  if (!isTestWorkPaperEngine(engine)) {
    throw new Error('Expected WorkPaper to expose an engine in tests')
  }
  return engine
}

describe('WorkPaper sheet dimensions', () => {
  it('keeps scalar formula appends on the incremental dimension path', () => {
    const workbook = WorkPaper.buildFromArray([[1, 2, '=SUM(A1:B1)']])
    const sheetId = workbook.getSheetId('Sheet1')!
    expect(workbook.getSheetDimensions(sheetId)).toEqual({ width: 3, height: 1 })
    const scans = trackDimensionScans(workbook)

    try {
      workbook.batch(() => {
        workbook.addRows(sheetId, 1, 1)
        workbook.setCellContents(cell(sheetId, 1, 0), [[3, 4, '=SUM(A2:B2)']])
      })

      expect(workbook.getSheetDimensions(sheetId)).toEqual({ width: 3, height: 2 })
      expect(scans.count).toBe(0)
    } finally {
      scans.restore()
    }
  })

  it('keeps dynamic spill sheet dimensions fresh after dependency edits grow the spill', () => {
    const workbook = WorkPaper.buildFromSheets({
      Data: [[1], [0], [3]],
      Summary: [['=FILTER(Data!A1:A3,Data!A1:A3>1)']],
    })
    const dataSheet = workbook.getSheetId('Data')!
    const summarySheet = workbook.getSheetId('Summary')!

    expect(workbook.getSheetDimensions(summarySheet)).toEqual({ width: 1, height: 1 })
    expect(
      workbook.getRangeValues({
        start: cell(summarySheet, 0, 0),
        end: cell(summarySheet, 0, 0),
      }),
    ).toEqual([[{ tag: ValueTag.Number, value: 3 }]])

    workbook.setCellContents(cell(dataSheet, 1, 0), 2)

    expect(
      workbook.getRangeValues({
        start: cell(summarySheet, 0, 0),
        end: cell(summarySheet, 1, 0),
      }),
    ).toEqual([[{ tag: ValueTag.Number, value: 2 }], [{ tag: ValueTag.Number, value: 3 }]])
    expect(workbook.getSheetDimensions(summarySheet)).toEqual({ width: 1, height: 2 })
  })

  it('checks only formula cells for dynamic-resize risk while scanning dimensions', () => {
    const rows = Array.from({ length: 200 }, (_unused, row) => [row + 1, (row + 1) * 2])
    rows.push([401, '=SUM(A1:A200)'])
    const workbook = WorkPaper.buildFromArray(rows)
    const sheetId = workbook.getSheetId('Sheet1')!
    getDimensionCache(workbook).invalidate(sheetId)
    const engine = getTestEngine(workbook)
    const getCell = vi.spyOn(engine, 'getCell')
    const getCellByIndex = vi.spyOn(engine, 'getCellByIndex')

    try {
      expect(workbook.getSheetDimensions(sheetId)).toEqual({ width: 2, height: 201 })
      expect(getCell).not.toHaveBeenCalled()
      expect(getCellByIndex).toHaveBeenCalledTimes(1)
    } finally {
      getCell.mockRestore()
      getCellByIndex.mockRestore()
    }
  })
})
