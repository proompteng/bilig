import { describe, expect, it } from 'vitest'
import { ValueTag, type CellValue } from '@bilig/protocol'

import { WorkPaper } from '../index.js'

function expectNumberClose(value: CellValue, expected: number): void {
  expect(value.tag).toBe(ValueTag.Number)
  if (value.tag !== ValueTag.Number) {
    throw new Error(`Expected number ${String(expected)}, received ${JSON.stringify(value)}`)
  }
  expect(value.value).toBeCloseTo(expected, 12)
}

describe('INDIRECT range references', () => {
  it('evaluates direct and dynamically constructed range references without diagnostics', () => {
    const workbook = WorkPaper.buildFromSheets(
      {
        Data: [
          ['Name', 'Value', 'Cash Flow'],
          ['Year 1', 100, 210],
          ['Year 2', 150, 240],
          ['Year 3', 160, 250],
        ],
        Summary: [
          ['INDIRECT range', '=SUM(INDIRECT("Data!B2:B4"))'],
          ['OFFSET range', '=SUM(OFFSET(Data!B2,0,0,3,1))'],
          ['INDIRECT cell', '=INDIRECT("Data!C3")'],
          ['ADDRESS cell', '=INDIRECT(ADDRESS(3,3,1,TRUE,"Data"))'],
        ],
      },
      { maxRows: 12, maxColumns: 8 },
    )
    const summary = workbook.getSheetId('Summary')
    const expected = [410, 410, 240, 240]

    expected.forEach((value, row) => {
      const address = { sheet: summary, row, col: 1 }
      expectNumberClose(workbook.getCellValue(address), value)
      expect(workbook.getCellFormulaDiagnostics(address)).toEqual([])
    })
  })
})
