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

function expectNumber(value: CellValue, expected: number): void {
  expect(value).toEqual({ tag: ValueTag.Number, value: expected })
}

describe('direct criteria cycle regressions', () => {
  it('does not create a cycle through unmatched SUMIF aggregate rows', () => {
    const workbook = WorkPaper.buildFromSheets(
      {
        'P&L assumptions': [
          [null, 'Ignored', "='P&L'!C3"],
          [null, 'Revenue', 10],
        ] satisfies TestCell[][],
        'P&L': [
          [null, 'Revenue', "=SUMIF('P&L assumptions'!$B:$B,'P&L'!$B1,'P&L assumptions'!C:C)"],
          [null, null, null],
          [null, null, '=C1/2'],
        ] satisfies TestCell[][],
      },
      { maxRows: 8, maxColumns: 4, useColumnIndex: true },
    )

    expectNumber(cellValue(workbook, 'P&L!C1'), 10)
    expectNumber(cellValue(workbook, 'P&L!C3'), 5)
    expectNumber(cellValue(workbook, 'P&L assumptions!C1'), 5)

    workbook.dispose()
  })
})
