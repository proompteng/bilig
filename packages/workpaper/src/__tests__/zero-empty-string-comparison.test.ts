import { describe, expect, it } from 'vitest'
import { ValueTag, type CellValue } from '@bilig/protocol'

import { WorkPaper } from '../index.js'

type TestCell = string | number | null

function cellValue(workbook: WorkPaper, sheetName: string, row: number, col: number): CellValue {
  return workbook.getCellValue({ sheet: workbook.getSheetId(sheetName), row, col })
}

function expectBoolean(value: CellValue, expected: boolean): void {
  expect(value).toEqual({ tag: ValueTag.Boolean, value: expected })
}

function expectNumber(value: CellValue, expected: number): void {
  expect(value).toEqual({ tag: ValueTag.Number, value: expected })
}

function expectString(value: CellValue, expected: string): void {
  expect(value).toMatchObject({ tag: ValueTag.String, value: expected })
}

describe('zero versus empty-string comparisons', () => {
  it.each([false, true])('does not treat numeric zero as blank with useColumnIndex=%s', (useColumnIndex) => {
    const rows: TestCell[][] = [[0, '=A1=""', '=A1<>""', '=IF(A1="","blank","not blank")', '=IF(A1="","",A1)']]
    const workbook = WorkPaper.buildFromSheets({ Sheet1: rows }, { maxRows: 8, maxColumns: 8, useColumnIndex })

    expectNumber(cellValue(workbook, 'Sheet1', 0, 0), 0)
    expectBoolean(cellValue(workbook, 'Sheet1', 0, 1), false)
    expectBoolean(cellValue(workbook, 'Sheet1', 0, 2), true)
    expectString(cellValue(workbook, 'Sheet1', 0, 3), 'not blank')
    expectNumber(cellValue(workbook, 'Sheet1', 0, 4), 0)

    workbook.dispose()
  })
})
