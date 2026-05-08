import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'

import { WorkPaper } from '../index.js'

function cellValue(workbook: WorkPaper, sheetName: string, row: number, col: number): CellValue {
  return workbook.getCellValue({ sheet: workbook.getSheetId(sheetName), row, col })
}

function expectNumber(value: CellValue, expected: number): void {
  expect(value).toEqual({ tag: ValueTag.Number, value: expected })
}

function expectError(value: CellValue, code: ErrorCode): void {
  expect(value).toEqual({ tag: ValueTag.Error, code })
}

describe('Excel 3D sheet references', () => {
  it('evaluates aggregate functions across a contiguous sheet span', () => {
    const workbook = WorkPaper.buildFromSheets(
      {
        Jan: [
          [null, null, null],
          [null, 100, 400],
        ],
        Feb: [
          [null, null, null],
          [null, 200, 500],
        ],
        Mar: [
          [null, null, null],
          [null, 300, 600],
        ],
        Summary: [['=SUM(Jan:Mar!B2)', '=SUM(Jan:Mar!B2:C2)', '=AVERAGE(Jan:Mar!B2)', '=COUNT(Jan:Mar!B2)']],
      },
      { maxRows: 20, maxColumns: 8, useColumnIndex: true },
    )

    expectNumber(cellValue(workbook, 'Summary', 0, 0), 600)
    expectNumber(cellValue(workbook, 'Summary', 0, 1), 2100)
    expectNumber(cellValue(workbook, 'Summary', 0, 2), 200)
    expectNumber(cellValue(workbook, 'Summary', 0, 3), 3)
  })

  it('supports quoted sheet names and recalculates when an interior sheet changes', () => {
    const workbook = WorkPaper.buildFromSheets(
      {
        'Jan 2026': [[10]],
        'Feb 2026': [[20]],
        'Mar 2026': [[30]],
        Summary: [["=SUM('Jan 2026':'Mar 2026'!A1)"]],
      },
      { maxRows: 20, maxColumns: 8, useColumnIndex: true },
    )

    expectNumber(cellValue(workbook, 'Summary', 0, 0), 60)

    const feb = workbook.getSheetId('Feb 2026')
    if (feb === undefined) {
      throw new Error('Feb 2026 sheet was not created')
    }
    workbook.setCellContents({ sheet: feb, row: 0, col: 0 }, 200)

    expectNumber(cellValue(workbook, 'Summary', 0, 0), 240)
  })

  it('returns #REF! when a sheet range endpoint is missing or out of order', () => {
    const workbook = WorkPaper.buildFromSheets(
      {
        Jan: [[100]],
        Feb: [[200]],
        Mar: [[300]],
        Summary: [['=SUM(Jan:Missing!A1)', '=SUM(Mar:Jan!A1)']],
      },
      { maxRows: 20, maxColumns: 8, useColumnIndex: true },
    )

    expectError(cellValue(workbook, 'Summary', 0, 0), ErrorCode.Ref)
    expectError(cellValue(workbook, 'Summary', 0, 1), ErrorCode.Ref)
  })
})
