import { describe, expect, it } from 'vitest'
import { ValueTag, type CellValue } from '@bilig/protocol'

import { WorkPaper } from '../index.js'

type TestCell = string | number | null

function cellValue(workbook: WorkPaper, sheetName: string, row: number, col: number): CellValue {
  return workbook.getCellValue({ sheet: workbook.getSheetId(sheetName), row, col })
}

function expectNumber(value: CellValue, expected: number): void {
  expect(value).toMatchObject({ tag: ValueTag.Number, value: expected })
}

function buildIssue123Workbook(useColumnIndex: boolean): WorkPaper {
  const summary = Array.from({ length: 8 }, () => Array.from<TestCell>({ length: 4 }).fill(null))
  summary[3][2] = '=SUM(C5:C7)'
  summary[4][2] = '=SUM(Deposits!$C:$C)'
  summary[5][2] = "=SUMIFS('Exchanges In'!$C:$C,'Exchanges In'!$B:$B,\"<>*Transferred*\")"
  summary[6][2] = "=SUM('Securities Transferred In'!$G:$G)"

  const deposits = Array.from({ length: 10 }, () => Array.from<TestCell>({ length: 3 }).fill(null))
  deposits[1][2] = 100

  const exchangesIn = Array.from({ length: 10 }, () => Array.from<TestCell>({ length: 3 }).fill(null))
  exchangesIn[1][1] = 'Regular deposit'
  exchangesIn[1][2] = 200

  const securities = Array.from({ length: 10 }, () => Array.from<TestCell>({ length: 7 }).fill(null))
  securities[1][6] = 300

  return WorkPaper.buildFromSheets(
    {
      Summary: summary,
      Deposits: deposits,
      'Exchanges In': exchangesIn,
      'Securities Transferred In': securities,
    },
    { maxRows: 1000, maxColumns: 20, useColumnIndex },
  )
}

describe('SUM over formula members', () => {
  it.each([false, true])('keeps initial direct aggregate dependencies on formula cells with useColumnIndex=%s', (useColumnIndex) => {
    const workbook = buildIssue123Workbook(useColumnIndex)

    expectNumber(cellValue(workbook, 'Summary', 3, 2), 600)
    expectNumber(cellValue(workbook, 'Summary', 4, 2), 100)
    expectNumber(cellValue(workbook, 'Summary', 5, 2), 200)
    expectNumber(cellValue(workbook, 'Summary', 6, 2), 300)

    workbook.setCellContents({ sheet: workbook.getSheetId('Deposits'), row: 1, col: 2 }, 150)

    expectNumber(cellValue(workbook, 'Summary', 3, 2), 650)
    expectNumber(cellValue(workbook, 'Summary', 4, 2), 150)
  })
})
