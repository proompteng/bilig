import { describe, expect, it } from 'vitest'
import { ValueTag, type CellValue } from '@bilig/protocol'

import { WorkPaper } from '../index.js'

type TestCell = string | number | null

function cellValue(workbook: WorkPaper, sheetName: string, row: number, col: number): CellValue {
  return workbook.getCellValue({ sheet: workbook.getSheetId(sheetName), row, col })
}

function expectNumber(value: CellValue, expected: number): void {
  expect(value).toEqual({ tag: ValueTag.Number, value: expected })
}

describe('Initial formula topology regressions', () => {
  it('orders formulas that depend on aggregate formulas over formula ranges', () => {
    const rows = Array.from({ length: 334 }, () => Array.from<TestCell>({ length: 16 }).fill(null))
    for (let row = 3; row <= 332; row += 1) {
      rows[row][4] = row === 8 ? 12 : 0
      rows[row][7] = 0
      rows[row][10] = `=E${row + 1}+H${row + 1}`
      rows[row][13] = `=IF(K334>0,ROUND((K${row + 1}/K334)*100,4),"")`
    }
    rows[333][10] = '=SUM(K2:K333)'
    rows[333][13] = '=ROUND(SUM(N2:N333),2)'

    const workbook = WorkPaper.buildFromSheets({ Sheet1: rows }, { maxRows: 400, maxColumns: 20, useColumnIndex: true })

    try {
      workbook.rebuildAndRecalculate()

      expectNumber(cellValue(workbook, 'Sheet1', 333, 10), 12)
      expectNumber(cellValue(workbook, 'Sheet1', 8, 13), 100)
      expectNumber(cellValue(workbook, 'Sheet1', 333, 13), 100)
    } finally {
      workbook.dispose()
    }
  })
})
