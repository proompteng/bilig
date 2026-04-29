import { describe, expect, it, vi } from 'vitest'
import * as formula from '@bilig/formula'
import { readRuntimeImage, readRuntimeSnapshot, SpreadsheetEngine, WorkbookStore } from '@bilig/core'
import { ValueTag } from '@bilig/protocol'
import { WorkPaper } from '../index.js'
import { WorkPaperSheetSizeLimitExceededError } from '../work-paper-errors.js'

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

  it('reserves mixed-sheet formula refs and attaches fresh cells without public per-cell attach calls', () => {
    const attachSpy = vi.spyOn(WorkbookStore.prototype, 'attachAllocatedCellWithLogicalAxisIds')
    const initSpy = vi.spyOn(SpreadsheetEngine.prototype, 'initializeFormulaSourcesAtNow')
    try {
      const workbook = WorkPaper.buildFromSheets({
        Bench: [
          [1, 10, '=A1+B1', '=C1*2'],
          [2, 20, '=A2+B2', '=C2*2'],
        ],
      })
      const sheetId = workbook.getSheetId('Bench')!
      const refs = initSpy.mock.calls[0]?.[0] ?? []

      expect(refs).toHaveLength(4)
      expect(refs.every((ref) => typeof ref.cellIndex === 'number')).toBe(true)
      expect(refs.map((ref) => ref.source)).toEqual(['A1+B1', 'C1*2', 'A2+B2', 'C2*2'])
      expect(attachSpy).not.toHaveBeenCalled()
      expect(workbook.getCellValue({ sheet: sheetId, row: 1, col: 3 })).toEqual({
        tag: ValueTag.Number,
        value: 44,
      })
    } finally {
      attachSpy.mockRestore()
      initSpy.mockRestore()
    }
  })

  it('recognizes padded formulas without treating ordinary strings as formulas during mixed-sheet initialization', () => {
    const workbook = WorkPaper.buildFromSheets({
      Bench: [[2, '  =A1*2  ', ' label ']],
    })
    const sheetId = workbook.getSheetId('Bench')!

    expect(workbook.getCellFormula({ sheet: sheetId, row: 0, col: 1 })).toBe('=A1*2')
    expect(workbook.getCellValue({ sheet: sheetId, row: 0, col: 1 })).toEqual({
      tag: ValueTag.Number,
      value: 4,
    })
    expect(workbook.getCellValue({ sheet: sheetId, row: 0, col: 2 })).toMatchObject({
      tag: ValueTag.String,
      value: ' label ',
    })
  })

  it('rebuilds from serialized sheets through the runtime-image fast path when available', () => {
    const source = WorkPaper.buildFromSheets({
      Bench: [
        [1, 2, '=A1+B1', '=C1*2'],
        [2, 4, '=A2+B2', '=C2*2'],
      ],
    })
    const serialized = source.getAllSheetsSerialized()
    const runtimeImage = readRuntimeImage(readRuntimeSnapshot(serialized))
    expect(runtimeImage?.sheetCells?.[0]?.dimensions).toEqual({ width: 4, height: 2 })
    expect(runtimeImage?.sheetCells?.[0]?.cellCount).toBe(8)
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

  it('imports compatible runtime snapshots without reading serialized sheet matrix entries', () => {
    const source = WorkPaper.buildFromSheets({
      Bench: [
        [1, 2, '=A1+B1'],
        [2, 4, '=A2+B2'],
      ],
    })
    const serialized = source.getAllSheetsSerialized()
    source.dispose()
    serialized.Bench = serialized.Bench.map(
      (row) =>
        new Proxy(row, {
          get(target, property, receiver) {
            if (typeof property === 'string' && /^\d+$/.test(property)) {
              throw new Error('snapshot fast path should not read serialized cell values')
            }
            return Reflect.get(target, property, receiver)
          },
        }),
    )

    const rebuilt = WorkPaper.buildFromSheets(serialized)
    const sheetId = rebuilt.getSheetId('Bench')!

    expect(rebuilt.getSheetDimensions(sheetId)).toEqual({ width: 3, height: 2 })
    expect(rebuilt.getCellValue({ sheet: sheetId, row: 1, col: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 6,
    })
  })

  it('imports compatible runtime snapshots without reading serialized sheet rows', () => {
    const source = WorkPaper.buildFromSheets({
      Bench: [
        [1, 2, '=A1+B1'],
        [2, 4, '=A2+B2'],
      ],
    })
    const serialized = source.getAllSheetsSerialized()
    source.dispose()
    const sheet = serialized.Bench
    expect(sheet).toBeDefined()
    serialized.Bench = new Proxy(sheet, {
      get(target, property, receiver) {
        if (property === 'length' || (typeof property === 'string' && /^\d+$/.test(property))) {
          throw new Error('snapshot fast path should not read serialized sheet rows')
        }
        return Reflect.get(target, property, receiver)
      },
    })

    const rebuilt = WorkPaper.buildFromSheets(serialized)
    const sheetId = rebuilt.getSheetId('Bench')!

    expect(rebuilt.getSheetDimensions(sheetId)).toEqual({ width: 3, height: 2 })
    expect(rebuilt.getCellValue({ sheet: sheetId, row: 1, col: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 6,
    })
  })

  it('rejects oversized sheets before importing compatible runtime snapshots', () => {
    const source = WorkPaper.buildFromSheets({
      Bench: [
        [1, 2],
        [3, 4],
      ],
    })
    const serialized = source.getAllSheetsSerialized()
    source.dispose()

    expect(() => WorkPaper.buildFromSheets(serialized, { maxRows: 1 })).toThrow(WorkPaperSheetSizeLimitExceededError)
    expect(() => WorkPaper.buildFromSheets(serialized, { maxColumns: 1 })).toThrow(WorkPaperSheetSizeLimitExceededError)
  })
})
