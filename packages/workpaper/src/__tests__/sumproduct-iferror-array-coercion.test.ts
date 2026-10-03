import { ValueTag, type CellValue } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'

import { WorkPaper } from '../index.js'

function cellValue(workbook: WorkPaper, ref: string): CellValue {
  const address = workbook.simpleCellAddressFromString(ref)
  if (!address) {
    throw new Error(`Expected ${ref} to resolve`)
  }
  return workbook.getCellValue(address)
}

describe('SUMPRODUCT over IFERROR array coercion', () => {
  it.each([false, true])('applies IFERROR per array cell before SUMPRODUCT with useColumnIndex=%s', (useColumnIndex) => {
    const workbook = WorkPaper.buildFromSheets(
      {
        Sheet1: [
          ['x', 2],
          ['3', 'bad'],
          ['=SUMPRODUCT(IFERROR(1*A1:B2,0))', '=SUMPRODUCT(IFERROR(1*A1,0))', '=SUMPRODUCT(IFERROR(1*B1,0))'],
        ],
      },
      { maxRows: 8, maxColumns: 8, useColumnIndex },
    )

    expect(cellValue(workbook, 'Sheet1!A3')).toEqual({
      tag: ValueTag.Number,
      value: 5,
    })
    expect(cellValue(workbook, 'Sheet1!B3')).toEqual({
      tag: ValueTag.Number,
      value: 0,
    })
    expect(cellValue(workbook, 'Sheet1!C3')).toEqual({
      tag: ValueTag.Number,
      value: 2,
    })
    expect(workbook.getCellDisplayValue(workbook.simpleCellAddressFromString('Sheet1!A3')!)).toBe('5')

    workbook.dispose()
  })
})
