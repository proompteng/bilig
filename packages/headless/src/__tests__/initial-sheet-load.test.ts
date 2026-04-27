import { describe, expect, it, vi } from 'vitest'
import * as formula from '@bilig/formula'
import { SpreadsheetEngine } from '@bilig/core'
import { ValueTag } from '@bilig/protocol'
import { WorkPaper } from '../index.js'

describe('initial mixed sheet load', () => {
  it('builds mixed sheets without routing formulas through restore cell mutations', () => {
    const restoreMutationSpy = vi.spyOn(SpreadsheetEngine.prototype, 'applyCellMutationsAtWithOptions')
    try {
      const workbook = WorkPaper.buildFromSheets({
        Bench: [
          [1, '=A1*2'],
          [2, '=A2*3'],
        ],
      })
      const sheetId = workbook.getSheetId('Bench')!

      expect(workbook.getCellValue({ sheet: sheetId, row: 0, col: 1 })).toEqual({
        tag: ValueTag.Number,
        value: 2,
      })
      expect(workbook.getCellValue({ sheet: sheetId, row: 1, col: 1 })).toEqual({
        tag: ValueTag.Number,
        value: 6,
      })
      expect(restoreMutationSpy).not.toHaveBeenCalled()
    } finally {
      restoreMutationSpy.mockRestore()
    }
  })

  it('normalizes repeated row-template formulas during mixed-sheet initialization', () => {
    const compileSpy = vi.spyOn(formula, 'compileFormulaAst')
    const parseSpy = vi.spyOn(formula, 'parseFormula')
    try {
      const workbook = WorkPaper.buildFromSheets({
        Bench: [
          [1, 2, '=A1+B1', '=C1*2'],
          [2, 4, '=A2+B2', '=C2*2'],
          [3, 6, '=A3+B3', '=C3*2'],
        ],
      })
      const sheetId = workbook.getSheetId('Bench')!

      expect(workbook.getCellValue({ sheet: sheetId, row: 0, col: 2 })).toEqual({
        tag: ValueTag.Number,
        value: 3,
      })
      expect(workbook.getCellValue({ sheet: sheetId, row: 2, col: 3 })).toEqual({
        tag: ValueTag.Number,
        value: 18,
      })
      expect(workbook.getPerformanceCounters().formulasParsed).toBe(2)
      expect(compileSpy).not.toHaveBeenCalled()
      expect(parseSpy).not.toHaveBeenCalled()
    } finally {
      compileSpy.mockRestore()
      parseSpy.mockRestore()
    }
  })

  it('rebuilds from serialized sheets through the runtime-image fast path when available', () => {
    const source = WorkPaper.buildFromSheets({
      Bench: [
        [1, 2, '=A1+B1', '=C1*2'],
        [2, 4, '=A2+B2', '=C2*2'],
      ],
    })
    const serialized = source.getAllSheetsSerialized()
    source.dispose()

    const rebuilt = WorkPaper.buildFromSheets(serialized)
    const sheetId = rebuilt.getSheetId('Bench')!

    expect(rebuilt.getCellValue({ sheet: sheetId, row: 0, col: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 3,
    })
    expect(rebuilt.getCellValue({ sheet: sheetId, row: 1, col: 3 })).toEqual({
      tag: ValueTag.Number,
      value: 12,
    })
    expect(rebuilt.getPerformanceCounters().snapshotOpsReplayed).toBe(0)
    expect(rebuilt.getPerformanceCounters().topoRebuilds).toBe(0)
    expect(rebuilt.getPerformanceCounters().wasmFullUploads).toBe(0)

    rebuilt.setCellContents({ sheet: sheetId, row: 0, col: 0 }, 3)

    expect(rebuilt.getCellValue({ sheet: sheetId, row: 0, col: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 5,
    })
    expect(rebuilt.getCellValue({ sheet: sheetId, row: 0, col: 3 })).toEqual({
      tag: ValueTag.Number,
      value: 10,
    })
  })
})
