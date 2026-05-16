import { ValueTag } from '@bilig/protocol'
import { describe, expect, it, vi } from 'vitest'

import { WorkPaper, type WorkPaperCellAddress } from '../index.js'

function cell(sheet: number, row: number, col: number): WorkPaperCellAddress {
  return { sheet, row, col }
}

function hasCaptureVisibilitySnapshot(value: unknown): value is WorkPaper & { captureVisibilitySnapshot: () => unknown } {
  return typeof Reflect.get(value, 'captureVisibilitySnapshot') === 'function'
}

describe('work paper batched structural fast path', () => {
  it('keeps appended formula rows on the tracked batch path', () => {
    const workbook = WorkPaper.buildFromSheets({
      Data: [
        [1, 2, '=SUM(A1:B1)'],
        [3, 4, '=SUM(A2:B2)'],
      ],
    })
    const sheetId = workbook.getSheetId('Data')!
    expect(hasCaptureVisibilitySnapshot(workbook)).toBe(true)
    if (!hasCaptureVisibilitySnapshot(workbook)) {
      throw new Error('Expected WorkPaper to expose captureVisibilitySnapshot in tests')
    }
    const captureVisibilitySnapshot = vi.spyOn(workbook, 'captureVisibilitySnapshot').mockImplementation(() => {
      throw new Error('batched append formulas should not rebuild visibility snapshots')
    })

    const changes = workbook.batch(() => {
      expect(workbook.addRows(sheetId, 2, 2)).toEqual([])
      expect(
        workbook.setCellContents(cell(sheetId, 2, 0), [
          [5, 6, '=SUM(A3:B3)'],
          [7, 8, '=SUM(A4:B4)'],
        ]),
      ).toEqual([])
    })
    captureVisibilitySnapshot.mockRestore()

    expect(changes.length).toBeGreaterThan(0)
    expect(workbook.getCellValue(cell(sheetId, 2, 2))).toEqual({ tag: ValueTag.Number, value: 11 })
    expect(workbook.getCellValue(cell(sheetId, 3, 2))).toEqual({ tag: ValueTag.Number, value: 15 })

    workbook.setCellContents(cell(sheetId, 2, 0), 10)
    expect(workbook.getCellValue(cell(sheetId, 2, 2))).toEqual({ tag: ValueTag.Number, value: 16 })

    workbook.undo()
    expect(workbook.getCellValue(cell(sheetId, 2, 2))).toEqual({ tag: ValueTag.Number, value: 11 })
    workbook.undo()
    expect(workbook.getSheetDimensions(sheetId).height).toBe(2)
  })
})
