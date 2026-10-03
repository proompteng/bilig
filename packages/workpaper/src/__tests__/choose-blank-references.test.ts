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

function buildIssue62Workbook(useColumnIndex: boolean): WorkPaper {
  const dcf = Array.from({ length: 5 }, () => Array.from<TestCell>({ length: 5 }).fill(null))
  dcf[0][1] = 2
  dcf[1][2] = 10
  dcf[1][4] = 30
  dcf[2][2] = '=CHOOSE($B$1,C2,D2,E2)'
  dcf[3][2] = '=100+CHOOSE($B$1,C2,D2,E2)'
  dcf[4][2] = '=CHOOSE($B$1,C2,D2,E2)+CHOOSE(1,1,2,3)'
  dcf[4][3] = '=IFERROR(D2,1)'

  return WorkPaper.buildFromSheets(
    {
      DCF: dcf,
    },
    { maxRows: 20, maxColumns: 10, useColumnIndex },
  )
}

describe('CHOOSE blank references', () => {
  it.each([false, true])('returns numeric zero for selected blank cell references with useColumnIndex=%s', (useColumnIndex) => {
    const workbook = buildIssue62Workbook(useColumnIndex)

    expectNumber(cellValue(workbook, 'DCF', 2, 2), 0)
    expectNumber(cellValue(workbook, 'DCF', 3, 2), 100)
    expectNumber(cellValue(workbook, 'DCF', 4, 2), 1)
    expectNumber(cellValue(workbook, 'DCF', 4, 3), 0)
  })
})
