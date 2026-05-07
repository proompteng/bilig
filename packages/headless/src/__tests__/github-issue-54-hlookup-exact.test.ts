import { describe, expect, it } from 'vitest'
import { ValueTag, type CellValue } from '@bilig/protocol'

import { WorkPaper } from '../index.js'

type TestCell = string | number | null

function cellValue(workbook: WorkPaper, sheetName: string, row: number, col: number): CellValue {
  return workbook.getCellValue({ sheet: workbook.getSheetId(sheetName), row, col })
}

function expectNumberClose(value: CellValue, expected: number): void {
  expect(value.tag).toBe(ValueTag.Number)
  if (value.tag !== ValueTag.Number) {
    throw new Error(`Expected number ${String(expected)}, received ${JSON.stringify(value)}`)
  }
  expect(value.value).toBeCloseTo(expected, 12)
}

function buildIssue54Workbook(useColumnIndex: boolean): WorkPaper {
  const debtGrid = Array.from({ length: 5 }, () => Array.from<TestCell>({ length: 11 }).fill(null))
  for (let col = 0; col < debtGrid[0].length; col += 1) {
    debtGrid[0][col] = col / 10
  }
  debtGrid[2][2] = 45616.01199262912
  debtGrid[3][2] = 1.0233773685605514
  debtGrid[4][2] = 0.10664524406185244

  return WorkPaper.buildFromSheets(
    {
      DebtGrid: debtGrid,
      Summary: [['=HLOOKUP(0.2,DebtGrid!A1:K5,3,FALSE)', '=HLOOKUP(0.2,DebtGrid!A1:K5,4,FALSE)', '=HLOOKUP(0.2,DebtGrid!A1:K5,5,FALSE)']],
    },
    { maxRows: 100, maxColumns: 20, useColumnIndex },
  )
}

describe('GitHub issue #54 HLOOKUP exact horizontal matches', () => {
  it.each([false, true])('returns the requested row from exact numeric horizontal matches with useColumnIndex=%s', (useColumnIndex) => {
    const workbook = buildIssue54Workbook(useColumnIndex)

    expectNumberClose(cellValue(workbook, 'Summary', 0, 0), 45616.01199262912)
    expectNumberClose(cellValue(workbook, 'Summary', 0, 1), 1.0233773685605514)
    expectNumberClose(cellValue(workbook, 'Summary', 0, 2), 0.10664524406185244)
  })
})
