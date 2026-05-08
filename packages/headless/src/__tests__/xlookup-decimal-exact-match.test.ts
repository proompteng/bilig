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

function expectString(value: CellValue, expected: string): void {
  expect(value).toMatchObject({ tag: ValueTag.String, value: expected })
}

function buildIssue125Workbook(useColumnIndex: boolean): WorkPaper {
  const comparison = Array.from({ length: 18 }, () => Array.from<TestCell>({ length: 7 }).fill(null))
  comparison[17][0] = '2026-04-06'
  comparison[17][1] =
    '=SUMIFS(Braintree!$D$2:$D$331,Braintree!$A$2:$A$331,$A18,Braintree!$B$2:$B$331,"American Express")+SUMIFS(Braintree!$D$2:$D$331,Braintree!$A$2:$A$331,$A18,Braintree!$B$2:$B$331,"Apple Pay - American Express")'
  comparison[17][2] = '=IFERROR(XLOOKUP(B18,\'Amex Settlement\'!$C$2:$C$31,\'Amex Settlement\'!$A$2:$A$31,"",0),"")'
  comparison[17][4] = 2374.28
  comparison[17][5] = '=B18-E18'
  comparison[17][6] = '=IF(ABS(F18)<0.01,"Match","Review")'

  const braintree = Array.from({ length: 331 }, () => Array.from<TestCell>({ length: 4 }).fill(null))
  braintree[1][0] = '2026-04-06'
  braintree[1][1] = 'American Express'
  braintree[1][3] = 1000
  braintree[2][0] = '2026-04-06'
  braintree[2][1] = 'Apple Pay - American Express'
  braintree[2][3] = 1374.28

  const settlement = Array.from({ length: 31 }, () => Array.from<TestCell>({ length: 3 }).fill(null))
  settlement[6][0] = '2026-04-06'
  settlement[6][2] = 2374.28

  return WorkPaper.buildFromSheets(
    {
      'BT & Amex Comparison': comparison,
      Braintree: braintree,
      'Amex Settlement': settlement,
    },
    { maxRows: 1000, maxColumns: 20, useColumnIndex },
  )
}

describe('XLOOKUP decimal exact match', () => {
  it.each([false, true])('matches calculated decimal lookup values with useColumnIndex=%s', (useColumnIndex) => {
    const workbook = buildIssue125Workbook(useColumnIndex)

    expectNumberClose(cellValue(workbook, 'BT & Amex Comparison', 17, 1), 2374.28)
    expectString(cellValue(workbook, 'BT & Amex Comparison', 17, 2), '2026-04-06')
    expectString(cellValue(workbook, 'BT & Amex Comparison', 17, 6), 'Match')
  })
})
