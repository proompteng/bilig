import { describe, expect, it } from 'vitest'
import { ValueTag } from '@bilig/protocol'

import {
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
  WorkPaper,
  type WorkPaperCellAddress,
  type WorkPaperCellRange,
  type WorkPaperConfig,
} from '../index.js'

function cell(sheet: number, row: number, col: number): WorkPaperCellAddress {
  return { sheet, row, col }
}

describe('WorkPaper', () => {
  it('matches the published README usage example', () => {
    const workbook = WorkPaper.buildFromSheets(
      {
        Sheet1: [
          [10, 20, '=A1+B1'],
          [7, '=A2*3', null],
        ],
      },
      {
        maxRows: 1_000,
        maxColumns: 100,
        useColumnIndex: true,
      },
    )

    const sheet = workbook.getSheetId('Sheet1')
    if (sheet === undefined) {
      throw new Error('Sheet1 was not created')
    }

    const at = (row: number, col: number): WorkPaperCellAddress => ({
      sheet,
      row,
      col,
    })

    expect(workbook.getCellValue(at(0, 2))).toEqual({
      tag: ValueTag.Number,
      value: 30,
    })

    workbook.setCellContents(at(1, 2), '=A2+B2')
    expect(workbook.getCellFormula(at(1, 2))).toBe('=A2+B2')
    expect(workbook.getCellSerialized(at(1, 2))).toBe('=A2+B2')
    expect(workbook.getCellValue(at(1, 2))).toEqual({
      tag: ValueTag.Number,
      value: 28,
    })

    const document = exportWorkPaperDocument(workbook)
    const json = serializeWorkPaperDocument(document)
    const restored = createWorkPaperFromDocument(parseWorkPaperDocument(json))
    const restoredSheet = restored.getSheetId('Sheet1')
    if (restoredSheet === undefined) {
      throw new Error('Sheet1 was not restored')
    }

    expect(restored.getCellValue({ sheet: restoredSheet, row: 1, col: 2 })).toEqual({
      tag: ValueTag.Number,
      value: 28,
    })
  })

  it('is the canonical top-level headless workbook runtime', () => {
    const config: WorkPaperConfig = {
      useArrayArithmetic: true,
      useColumnIndex: true,
    }

    const workbook = WorkPaper.buildFromSheets(
      {
        Data: [[2, '=FILTER(A2:A5,A2:A5>A1)'], [1], [2], [3], [4]],
      },
      config,
    )

    const sheetId = workbook.getSheetId('Data')!
    const spillRange: WorkPaperCellRange = {
      start: cell(sheetId, 0, 1),
      end: cell(sheetId, 1, 1),
    }

    expect(workbook.getRangeValues(spillRange)).toEqual([[{ tag: ValueTag.Number, value: 3 }], [{ tag: ValueTag.Number, value: 4 }]])
    expect(workbook.isCellPartOfArray(cell(sheetId, 0, 1))).toBe(true)
  })

  it('supports a production-style headless workflow through the WorkPaper entrypoint', () => {
    const workbook = WorkPaper.buildFromSheets({
      Revenue: [[125], [250], [375], ['=SUM(A1:A3)']],
    })
    const revenueId = workbook.getSheetId('Revenue')!

    const copied = workbook.copy({
      start: cell(revenueId, 0, 0),
      end: cell(revenueId, 2, 0),
    })

    expect(copied).toEqual([
      [{ tag: ValueTag.Number, value: 125 }],
      [{ tag: ValueTag.Number, value: 250 }],
      [{ tag: ValueTag.Number, value: 375 }],
    ])

    const changes = workbook.batch(() => {
      workbook.paste(cell(revenueId, 0, 1))
      workbook.setCellContents(cell(revenueId, 3, 1), '=SUM(B1:B3)')
    })

    expect(changes.length).toBeGreaterThan(0)
    expect(workbook.getCellValue(cell(revenueId, 3, 1))).toEqual({
      tag: ValueTag.Number,
      value: 750,
    })

    workbook.undo()
    expect(workbook.getCellSerialized(cell(revenueId, 3, 1))).toBeNull()
    workbook.redo()
    expect(workbook.getCellFormula(cell(revenueId, 3, 1))).toBe('=SUM(B1:B3)')
  })

  it('keeps compatibility adapters frozen and returns detached adapter results', () => {
    const workbook = WorkPaper.buildFromSheets({
      Data: [[1, '=A1*2']],
    })
    const dataId = workbook.getSheetId('Data')!

    expect(Object.isFrozen(workbook.internals)).toBe(true)
    expect(Object.isFrozen(workbook.graph)).toBe(true)
    expect(Object.isFrozen(workbook.rangeMapping)).toBe(true)
    expect(Object.isFrozen(workbook.arrayMapping)).toBe(true)
    expect(Object.isFrozen(workbook.sheetMapping)).toBe(true)
    expect(Object.isFrozen(workbook.addressMapping)).toBe(true)
    expect(Object.isFrozen(workbook.dependencyGraph)).toBe(true)
    expect(Object.isFrozen(workbook.evaluator)).toBe(true)
    expect(Object.isFrozen(workbook.columnSearch)).toBe(true)
    expect(Object.isFrozen(workbook.lazilyTransformingAstService)).toBe(true)

    const serialized = workbook.rangeMapping.getSerialized({
      start: cell(dataId, 0, 0),
      end: cell(dataId, 0, 1),
    })
    serialized[0][0] = 999
    serialized[0][1] = '=A1*999'

    expect(workbook.getCellSerialized(cell(dataId, 0, 0))).toBe(1)
    expect(workbook.getCellSerialized(cell(dataId, 0, 1))).toBe('=A1*2')
  })
})
