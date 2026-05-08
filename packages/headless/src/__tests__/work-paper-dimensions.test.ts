import { describe, expect, it } from 'vitest'
import { ValueTag } from '@bilig/protocol'

import { WorkPaper, type WorkPaperCellAddress } from '../index.js'

function cell(sheet: number, row: number, col: number): WorkPaperCellAddress {
  return { sheet, row, col }
}

describe('WorkPaper sheet dimensions', () => {
  it('keeps dynamic spill sheet dimensions fresh after dependency edits grow the spill', () => {
    const workbook = WorkPaper.buildFromSheets({
      Data: [[1], [0], [3]],
      Summary: [['=FILTER(Data!A1:A3,Data!A1:A3>1)']],
    })
    const dataSheet = workbook.getSheetId('Data')!
    const summarySheet = workbook.getSheetId('Summary')!

    expect(workbook.getSheetDimensions(summarySheet)).toEqual({ width: 1, height: 1 })
    expect(
      workbook.getRangeValues({
        start: cell(summarySheet, 0, 0),
        end: cell(summarySheet, 0, 0),
      }),
    ).toEqual([[{ tag: ValueTag.Number, value: 3 }]])

    workbook.setCellContents(cell(dataSheet, 1, 0), 2)

    expect(
      workbook.getRangeValues({
        start: cell(summarySheet, 0, 0),
        end: cell(summarySheet, 1, 0),
      }),
    ).toEqual([[{ tag: ValueTag.Number, value: 2 }], [{ tag: ValueTag.Number, value: 3 }]])
    expect(workbook.getSheetDimensions(summarySheet)).toEqual({ width: 1, height: 2 })
  })
})
