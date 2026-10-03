import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import type { SpreadsheetEngine } from '@bilig/core'

import { WorkPaper } from '../index.js'
import { readWorkPaperRangeValueBlock } from '../work-paper-cell-read.js'

function isSpreadsheetEngine(value: unknown): value is SpreadsheetEngine {
  return typeof value === 'object' && value !== null && typeof Reflect.get(value, 'getRangeValues') === 'function'
}

function getTestEngine(workbook: WorkPaper): SpreadsheetEngine {
  const engine: unknown = Reflect.get(workbook, 'engine')
  if (!isSpreadsheetEngine(engine)) {
    throw new Error('Expected WorkPaper to expose an engine in tests')
  }
  return engine
}

const failLargeValueBlockFallback: SpreadsheetEngine['getRangeValues'] = () => {
  throw new Error('CellValue matrix fallback should not run for large typed value blocks')
}

describe('WorkPaper range value blocks', () => {
  it('reads a typed row-major value block without materializing CellValue objects', () => {
    const workbook = WorkPaper.buildFromSheets({
      Sheet1: [
        [1, true, 'alpha', null, '=1/0'],
        [2, '=A1+A2', 'beta', false, '=A2*3'],
      ],
    })
    const sheetId = workbook.getSheetId('Sheet1')!

    const block = workbook.getRangeValueBlock({
      start: { sheet: sheetId, row: 0, col: 0 },
      end: { sheet: sheetId, row: 1, col: 4 },
    })

    const offset = (row: number, col: number): number => row * block.colCount + col
    expect(block.rowCount).toBe(2)
    expect(block.colCount).toBe(5)
    expect(block.tags[offset(0, 0)]).toBe(ValueTag.Number)
    expect(block.numbers[offset(0, 0)]).toBe(1)
    expect(block.tags[offset(0, 1)]).toBe(ValueTag.Boolean)
    expect(block.numbers[offset(0, 1)]).toBe(1)
    expect(block.tags[offset(0, 2)]).toBe(ValueTag.String)
    expect(block.strings?.get(block.stringIds[offset(0, 2)])).toBe('alpha')
    expect(block.tags[offset(0, 3)]).toBe(ValueTag.Empty)
    expect(block.tags[offset(0, 4)]).toBe(ValueTag.Error)
    expect(block.errors[offset(0, 4)]).toBe(ErrorCode.Div0)
    expect(block.tags[offset(1, 1)]).toBe(ValueTag.Number)
    expect(block.numbers[offset(1, 1)]).toBe(3)
    expect(block.tags[offset(1, 4)]).toBe(ValueTag.Number)
    expect(block.numbers[offset(1, 4)]).toBe(6)
  })

  it('keeps large typed value-block reads off the CellValue matrix fallback', () => {
    const rowCount = 513
    const colCount = 512
    const workbook = WorkPaper.buildFromSheets({
      Sheet1: Array.from({ length: rowCount }, (_row, row) => Array.from({ length: colCount }, (_col, col) => row * colCount + col)),
    })
    const sheetId = workbook.getSheetId('Sheet1')!
    const engine = getTestEngine(workbook)
    const originalGetRangeValues = engine.getRangeValues.bind(engine)
    try {
      engine.getRangeValues = failLargeValueBlockFallback
      const block = readWorkPaperRangeValueBlock({
        engine,
        range: {
          start: { sheet: sheetId, row: 0, col: 0 },
          end: { sheet: sheetId, row: rowCount - 1, col: colCount - 1 },
        },
        rangeRef: () => ({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' }),
      })
      const terminalOffset = (rowCount - 1) * colCount + colCount - 1
      expect(block.rowCount).toBe(rowCount)
      expect(block.colCount).toBe(colCount)
      expect(block.tags[terminalOffset]).toBe(ValueTag.Number)
      expect(block.numbers[terminalOffset]).toBe(rowCount * colCount - 1)
    } finally {
      engine.getRangeValues = originalGetRangeValues
    }
  })
})
