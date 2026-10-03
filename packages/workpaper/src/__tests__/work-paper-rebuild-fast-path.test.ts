import { describe, expect, it, vi } from 'vitest'
import { ValueTag } from '@bilig/protocol'

import type { WorkPaperCellAddress } from '../index.js'
import { WorkPaper } from '../index.js'

function cell(sheet: number, row: number, col: number): WorkPaperCellAddress {
  return { sheet, row, col }
}

function hasCaptureVisibilitySnapshot(value: unknown): value is WorkPaper & { captureVisibilitySnapshot: () => unknown } {
  return typeof Reflect.get(value, 'captureVisibilitySnapshot') === 'function'
}

function hasCaptureFormulaResultValueSnapshot(value: unknown): value is WorkPaper & { captureFormulaResultValueSnapshot: () => unknown } {
  return typeof Reflect.get(value, 'captureFormulaResultValueSnapshot') === 'function'
}

describe('WorkPaper rebuild fast path', () => {
  it('rebuilds stable formulas without full visibility or formula-value snapshots', () => {
    const workbook = WorkPaper.buildFromSheets({
      Bench: [
        [1, 2, '=A1+B1', '=C1*2'],
        [3, 4, '=A2+B2', '=C2*2'],
      ],
    })
    const sheetId = workbook.getSheetId('Bench')!
    if (!hasCaptureVisibilitySnapshot(workbook)) {
      throw new Error('Expected work paper runtime to expose captureVisibilitySnapshot in tests')
    }
    if (!hasCaptureFormulaResultValueSnapshot(workbook)) {
      throw new Error('Expected work paper runtime to expose captureFormulaResultValueSnapshot in tests')
    }
    const captureVisibilitySnapshot = vi.spyOn(workbook, 'captureVisibilitySnapshot').mockImplementation(() => {
      throw new Error('stable rebuilds should not require full visibility snapshots')
    })
    const captureFormulaResultValueSnapshot = vi.spyOn(workbook, 'captureFormulaResultValueSnapshot').mockImplementation(() => {
      throw new Error('stable rebuilds should not require formula-value snapshots')
    })

    try {
      expect(workbook.rebuildAndRecalculate()).toEqual([])
      expect(workbook.getCellValue(cell(sheetId, 1, 3))).toEqual({ tag: ValueTag.Number, value: 14 })
      expect(captureVisibilitySnapshot).not.toHaveBeenCalled()
      expect(captureFormulaResultValueSnapshot).not.toHaveBeenCalled()
    } finally {
      captureVisibilitySnapshot.mockRestore()
      captureFormulaResultValueSnapshot.mockRestore()
    }
  })

  it('returns volatile formula changes without full visibility snapshots', () => {
    const workbook = WorkPaper.buildFromSheets({
      Bench: [['=RAND()']],
    })
    const sheetId = workbook.getSheetId('Bench')!
    const before = workbook.getCellValue(cell(sheetId, 0, 0))
    if (!hasCaptureVisibilitySnapshot(workbook)) {
      throw new Error('Expected work paper runtime to expose captureVisibilitySnapshot in tests')
    }
    const captureVisibilitySnapshot = vi.spyOn(workbook, 'captureVisibilitySnapshot').mockImplementation(() => {
      throw new Error('volatile rebuild changes should not require full visibility snapshots')
    })

    try {
      const changes = workbook.rebuildAndRecalculate()
      const after = workbook.getCellValue(cell(sheetId, 0, 0))

      expect(after).not.toEqual(before)
      expect(changes).toEqual([
        {
          kind: 'cell',
          address: cell(sheetId, 0, 0),
          sheetName: 'Bench',
          a1: 'A1',
          newValue: after,
        },
      ])
      expect(captureVisibilitySnapshot).not.toHaveBeenCalled()
    } finally {
      captureVisibilitySnapshot.mockRestore()
    }
  })
})
