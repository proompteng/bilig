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

function buildIssue124Workbook(useColumnIndex: boolean): WorkPaper {
  const summary = [
    [
      '=SUMIFS(Data!$C$1:$C$10,Data!$B$1:$B$10,"*Transferred*")',
      '=SUMIFS(Data!$C$1:$C$10,Data!$B$1:$B$10,"*Transferred*")+1',
      '=SUMIFS(Data!$C$1:$C$10,Data!$B$1:$B$10,"*Transferred*")+SUMIFS(Data!$C$1:$C$10,Data!$B$1:$B$10,"*Transferred*")',
    ],
  ]

  const data = Array.from({ length: 10 }, () => Array.from<TestCell>({ length: 3 }).fill(null))
  data[1][1] = 'Transferred from broker'
  data[1][2] = -50
  data[2][1] = 'Not transferred'
  data[2][2] = 1000

  return WorkPaper.buildFromSheets(
    {
      Summary: summary,
      Data: data,
    },
    { maxRows: 1000, maxColumns: 20, useColumnIndex },
  )
}

describe('SUMIFS wildcard arithmetic', () => {
  it.each([false, true])('keeps wildcard SUMIFS numeric when used in arithmetic with useColumnIndex=%s', (useColumnIndex) => {
    const workbook = buildIssue124Workbook(useColumnIndex)

    expectNumber(cellValue(workbook, 'Summary', 0, 0), 950)
    expectNumber(cellValue(workbook, 'Summary', 0, 1), 951)
    expectNumber(cellValue(workbook, 'Summary', 0, 2), 1900)
  })

  it.each([false, true])('does not match blank or numeric cells for wildcard text criteria with useColumnIndex=%s', (useColumnIndex) => {
    const workbook = WorkPaper.buildFromSheets(
      {
        Data: [
          ['Value', 'Bucket'],
          [10, null],
          [20, 'Agency A'],
          [30, 42],
          [40, 'Agency C'],
          [50, 'Agency B'],
        ],
        Summary: [
          ['', '=SUMIF(Data!B2:B6,"*"&A1&"*",Data!A2:A6)'],
          ['', '=COUNTIF(Data!B2:B6,"*"&A1&"*")'],
        ],
      },
      { maxRows: 16, maxColumns: 8, useColumnIndex },
    )

    expectNumber(cellValue(workbook, 'Summary', 0, 1), 110)
    expectNumber(cellValue(workbook, 'Summary', 1, 1), 3)
  })
})
