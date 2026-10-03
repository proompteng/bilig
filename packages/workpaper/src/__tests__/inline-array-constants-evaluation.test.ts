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

describe('inline array constants', () => {
  it('evaluates inline array constants for aggregate, lookup, and finance formulas', () => {
    const workbook = WorkPaper.buildFromSheets(
      {
        Summary: [
          ['SUM array', '=SUM({-1000,300,400,500})'],
          ['INDEX array', '=INDEX({-1000,300,400,500},2)'],
          ['MIRR array', '=ROUND(MIRR({-1000,300,400,500},0.1,0.12),4)'],
          ['XNPV array', '=ROUND(XNPV(0.1,{-1000,300,400,500},{46023,46115,46207,46388}),2)'],
          ['XIRR array', '=ROUND(XIRR({-1000,300,400,500},{46023,46115,46207,46388}),4)'],
        ],
      },
      { maxRows: 20, maxColumns: 6 },
    )
    const summary = workbook.getSheetId('Summary')
    const expected = [200, 300, 0.0982, 128.66, 0.3334]

    expected.forEach((value, row) => {
      expectNumberClose(workbook.getCellValue({ sheet: summary, row, col: 1 }), value)
      expect(workbook.getCellFormulaDiagnostics({ sheet: summary, row, col: 1 })).toEqual([])
    })
  })
})
