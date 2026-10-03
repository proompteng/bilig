import { describe, expect, it, vi } from 'vitest'

import { WorkPaper, type WorkPaperCellAddress } from '../index.js'

function cell(sheet: number, row: number, col: number): WorkPaperCellAddress {
  return { sheet, row, col }
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

function getTestEngine(workbook: WorkPaper): TestWorkPaperEngine {
  const engine: unknown = Reflect.get(workbook, 'engine')
  if (!isTestWorkPaperEngine(engine)) {
    throw new Error('Expected WorkPaper to expose an engine in tests')
  }
  return engine
}

describe('WorkPaper formula range reads', () => {
  it('hydrates only formula cells while reading sparse formula ranges', () => {
    const rows = Array.from({ length: 200 }, (_unused, row) => [row + 1, (row + 1) * 2])
    rows.push([401, '=SUM(A1:A200)'])
    const workbook = WorkPaper.buildFromArray(rows)
    const sheetId = workbook.getSheetId('Sheet1')!
    const engine = getTestEngine(workbook)
    const getCell = vi.spyOn(engine, 'getCell')
    const getCellByIndex = vi.spyOn(engine, 'getCellByIndex')

    try {
      const formulas = workbook.getRangeFormulas({
        start: cell(sheetId, 0, 0),
        end: cell(sheetId, 200, 1),
      })

      expect(formulas[0]?.[0]).toBeUndefined()
      expect(formulas[200]?.[1]).toBe('=SUM(A1:A200)')
      expect(getCell).not.toHaveBeenCalled()
      expect(getCellByIndex).toHaveBeenCalledTimes(1)
    } finally {
      getCell.mockRestore()
      getCellByIndex.mockRestore()
    }
  })

  it('hydrates only resident cells while reading sparse serialized ranges', () => {
    const rows = Array.from({ length: 200 }, () => [null, null])
    rows.push([null, '=SUM(A1:A200)'])
    const workbook = WorkPaper.buildFromArray(rows)
    const sheetId = workbook.getSheetId('Sheet1')!
    const engine = getTestEngine(workbook)
    const getCell = vi.spyOn(engine, 'getCell')
    const getCellByIndex = vi.spyOn(engine, 'getCellByIndex')

    try {
      const serialized = workbook.getRangeSerialized({
        start: cell(sheetId, 0, 0),
        end: cell(sheetId, 200, 1),
      })

      expect(serialized[0]?.[0]).toBeNull()
      expect(serialized[200]?.[1]).toBe('=SUM(A1:A200)')
      expect(getCell).not.toHaveBeenCalled()
      expect(getCellByIndex).toHaveBeenCalledTimes(1)
    } finally {
      getCell.mockRestore()
      getCellByIndex.mockRestore()
    }
  })
})
