import { CellStore } from '@bilig/core'
import { ValueTag } from '@bilig/protocol'
import { describe, expect, it, vi } from 'vitest'
import { WorkPaper } from '../index.js'

function cell(sheet: number, row: number, col: number): { readonly sheet: number; readonly row: number; readonly col: number } {
  return { sheet, row, col }
}

describe('WorkPaper sheet initialization', () => {
  it('pre-reserves materialized cells across all initial sheets', () => {
    const ensureCapacity = vi.spyOn(CellStore.prototype, 'ensureCapacity')
    try {
      const workbook = WorkPaper.buildFromSheets({
        Sheet1: [
          [1, '=A1+1'],
          [2, '=A2+1'],
        ],
        Sheet2: [
          [3, '=A1+1'],
          [4, '=A2+1'],
        ],
      })
      const sheet1Id = workbook.getSheetId('Sheet1')!
      const sheet2Id = workbook.getSheetId('Sheet2')!

      expect(ensureCapacity.mock.calls[0]?.[0]).toBe(8)
      expect(workbook.getCellValue(cell(sheet1Id, 1, 1))).toEqual({ tag: ValueTag.Number, value: 3 })
      expect(workbook.getCellValue(cell(sheet2Id, 1, 1))).toEqual({ tag: ValueTag.Number, value: 5 })
      workbook.dispose()
    } finally {
      ensureCapacity.mockRestore()
    }
  })
})
