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

function buildOmittedIfWorkbook(useColumnIndex: boolean): WorkPaper {
  return WorkPaper.buildFromSheets(
    {
      Model: [
        [null, null, null],
        [null, 0, '=IF(B2>0,B2/$B$8,)'],
        [null, 0, '=IF(B3>-1,,B3-1)'],
        [null, null, null],
        [null, -1, '=IF((0-B5)<$B$9,IF(B5>-1,,B5-1),)'],
        [null, -10, '=IF((0-B6)<$B$9,IF(B6>-1,,B6-1),)'],
        [null, null, '=C2*100'],
        [null, 10, null],
        [null, 5, null],
      ] satisfies TestCell[][],
    },
    { maxRows: 20, maxColumns: 8, useColumnIndex },
  )
}

describe('GitHub issue #52 omitted IF arguments', () => {
  it.each([false, true])('evaluates omitted IF branches as zero-valued blanks with useColumnIndex=%s', (useColumnIndex) => {
    const workbook = buildOmittedIfWorkbook(useColumnIndex)

    expectNumber(cellValue(workbook, 'Model', 1, 2), 0)
    expectNumber(cellValue(workbook, 'Model', 2, 2), 0)
    expectNumber(cellValue(workbook, 'Model', 4, 2), -2)
    expectNumber(cellValue(workbook, 'Model', 5, 2), 0)
    expectNumber(cellValue(workbook, 'Model', 6, 2), 0)
  })
})
