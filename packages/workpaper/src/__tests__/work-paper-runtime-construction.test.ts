import { describe, expect, it } from 'vitest'
import { SpreadsheetEngine } from '@bilig/core'
import { WorkPaper } from '../index.js'

const LARGE_ENGINE_CELL_COUNT = 16_385

function largeFormulaSheetRows(): (number | string)[][] {
  return Array.from({ length: LARGE_ENGINE_CELL_COUNT }, (_value, row) => (row === 0 ? [1, '=A1+1'] : [row + 1]))
}

function workbookEngine(workbook: WorkPaper): SpreadsheetEngine {
  const engine: unknown = Reflect.get(workbook, 'engine')
  if (!(engine instanceof SpreadsheetEngine)) {
    throw new Error('Expected workbook engine')
  }
  return engine
}

function engineFormulaCount(engine: SpreadsheetEngine): number {
  const formulas: unknown = Reflect.get(engine, 'formulas')
  const size = typeof formulas === 'object' && formulas !== null ? Reflect.get(formulas, 'size') : undefined
  if (typeof size !== 'number') {
    throw new Error('Expected formula table size')
  }
  return size
}

describe('work paper runtime construction', () => {
  it('fully resets oversized engines on dispose instead of pooling retained formula state', () => {
    const workbook = WorkPaper.buildFromArray(largeFormulaSheetRows())
    const engine = workbookEngine(workbook)

    expect(engine.workbook.cellStore.capacity).toBeGreaterThan(LARGE_ENGINE_CELL_COUNT)
    expect(engineFormulaCount(engine)).toBe(1)

    workbook.dispose()

    expect(engine.workbook.cellStore.capacity).toBe(0)
    expect(engineFormulaCount(engine)).toBe(0)
    expect(engine.workbook.getSheet('Sheet1')).toBeUndefined()
  })
})
