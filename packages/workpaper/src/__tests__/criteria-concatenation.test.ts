import { ValueTag, type CellValue } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'

import { WorkPaper } from '../index.js'

type TestCell = string | number | null

function cellValue(workbook: WorkPaper, ref: string): CellValue {
  const address = workbook.simpleCellAddressFromString(ref)
  if (!address) {
    throw new Error(`Expected ${ref} to resolve`)
  }
  return workbook.getCellValue(address)
}

function buildLedgerWorkbook(useColumnIndex: boolean): WorkPaper {
  return WorkPaper.buildFromSheets(
    {
      Ledger: [
        ['Date', 'Account', 'Department', 'Status', 'Amount', null, 'Start', 'End', 'Date range sum', 'Date range count'],
        ['=DATE(2026,1,1)', '4100 Sales', 'Ops', 'Open', 100, null, '=DATE(2026,1,10)', '=DATE(2026,2,15)', null, null],
        ['=DATE(2026,1,15)', '4100 Sales', 'Ops', 'Open', 900, null, null, null, null, null],
        ['=DATE(2026,2,1)', '4200 Services', 'Ops', 'Closed', 400, null, null, null, null, null],
        ['=DATE(2026,2,15)', '4300 Support', 'Finance', '', 800, null, null, null, null, null],
        ['=DATE(2026,3,1)', '5100 Expense', 'Ops', 'Open', 1000, null, null, null, null, null],
        [
          null,
          null,
          null,
          null,
          null,
          null,
          null,
          null,
          '=SUMIFS(E2:E6,A2:A6,">="&G2,A2:A6,"<="&H2)',
          '=COUNTIFS(A2:A6,">="&G2,A2:A6,"<="&H2)',
        ],
      ] satisfies TestCell[][],
    },
    { maxRows: 12, maxColumns: 12, useColumnIndex },
  )
}

describe('criteria concatenation', () => {
  it.each([false, true])('evaluates SUMIFS and COUNTIFS criteria built from date cells with useColumnIndex=%s', (useColumnIndex) => {
    const workbook = buildLedgerWorkbook(useColumnIndex)

    expect(cellValue(workbook, 'Ledger!I7')).toEqual({
      tag: ValueTag.Number,
      value: 2100,
    })
    expect(cellValue(workbook, 'Ledger!J7')).toEqual({
      tag: ValueTag.Number,
      value: 3,
    })

    workbook.dispose()
  })
})
