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

function buildWorkbook(): WorkPaper {
  return WorkPaper.buildFromSheets(
    {
      Model: [
        ['Key', 'Value', 'Formula', 'Result'],
        [1, 10, '_xlfn.XLOOKUP', '=_xlfn.XLOOKUP(2,A2:A4,B2:B4)'],
        [2, 20, '_xlfn.LET', '=_xlfn.LET(x,2,x+3)'],
        [3, 30, '_xlfn._xlws.FILTER', '=SUM(_xlfn._xlws.FILTER(B2:B4,A2:A4>1))'],
      ] satisfies TestCell[][],
    },
    { maxRows: 8, maxColumns: 6 },
  )
}

describe('GitHub issue #30 Excel function compatibility prefixes', () => {
  it('evaluates supported functions with _xlfn and _xlws prefixes', () => {
    const workbook = buildWorkbook()

    expectNumber(cellValue(workbook, 'Model!D2'), 20)
    expectNumber(cellValue(workbook, 'Model!D3'), 5)
    expectNumber(cellValue(workbook, 'Model!D4'), 50)

    workbook.dispose()
  })
})
