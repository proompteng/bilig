import { describe, expect, it, vi } from 'vitest'
import { SpreadsheetEngine } from '../engine.js'
import { ValueTag } from '@bilig/protocol'

const wholeSheet = { sheetName: 'Sheet1', rowStart: 0, rowEnd: 1048575, colStart: 0, colEnd: 16383 }

function guardBlankCoordinateScan(engine: SpreadsheetEngine) {
  const get = engine.workbook.cellKeyToIndex.get.bind(engine.workbook.cellKeyToIndex)
  let lookups = 0
  return vi.spyOn(engine.workbook.cellKeyToIndex, 'get').mockImplementation((key) => {
    if (++lookups > 32) throw new Error('Dirty recalculation scanned blank coordinates')
    return get(key)
  })
}

describe('sparse dirty-region recalculation', () => {
  it('recalculates a whole sheet from stored cells and preserves other sheets', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'sparse-dirty' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.createSheet('Other')
    engine.setCellValue('Sheet1', 'A1', 10)
    engine.setCellFormula('Sheet1', 'B1', 'A1*2')
    engine.setCellValue('Other', 'A1', 7)
    engine.setCellFormula('Other', 'B1', 'A1*2')
    engine.workbook.cellStore.setValue(engine.workbook.ensureCell('Sheet1', 'A1'), { tag: ValueTag.Number, value: 30 })
    engine.workbook.cellStore.setValue(engine.workbook.ensureCell('Other', 'A1'), { tag: ValueTag.Number, value: 50 })
    const guard = guardBlankCoordinateScan(engine)
    try {
      const changed = engine.recalculateDirty([wholeSheet])
      expect(changed).toContain(engine.workbook.getCellIndex('Sheet1', 'B1'))
      expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 60 })
      expect(engine.getCellValue('Other', 'B1')).toEqual({ tag: ValueTag.Number, value: 14 })
    } finally {
      guard.mockRestore()
    }
  })

  it('finishes whole-sheet recalculation for an empty sheet without creating cells', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'empty-dirty' })
    await engine.ready()
    engine.createSheet('Sheet1')
    const guard = guardBlankCoordinateScan(engine)
    try {
      expect(engine.recalculateDirty([wholeSheet])).toEqual([])
      expect(engine.workbook.cellStore.size).toBe(0)
    } finally {
      guard.mockRestore()
    }
  })

  it('uses logical coordinates after row and column inserts for a whole-column dirty region', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'shifted-dirty' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 10)
    engine.setCellFormula('Sheet1', 'B1', 'A1*2')
    engine.insertRows('Sheet1', 0, 1)
    engine.insertColumns('Sheet1', 0, 1)
    engine.workbook.cellStore.setValue(engine.workbook.ensureCell('Sheet1', 'B2'), { tag: ValueTag.Number, value: 50 })
    const guard = guardBlankCoordinateScan(engine)
    try {
      engine.recalculateDirty([{ ...wholeSheet, colStart: 1, colEnd: 1 }])
      expect(engine.getCellValue('Sheet1', 'C2')).toEqual({ tag: ValueTag.Number, value: 100 })
    } finally {
      guard.mockRestore()
    }
  })

  it('keeps small-region recalculation bounded to coordinates rather than scanning every stored cell', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'small-dirty' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 10)
    engine.setCellFormula('Sheet1', 'B1', 'A1*2')
    engine.workbook.cellStore.setValue(engine.workbook.ensureCell('Sheet1', 'A1'), { tag: ValueTag.Number, value: 15 })
    const sheet = engine.workbook.getSheet('Sheet1')!
    const scan = vi.spyOn(sheet.grid, 'forEachCellEntry').mockImplementation(() => {
      throw new Error('Small dirty region scanned the whole sheet')
    })
    try {
      engine.recalculateDirty([{ ...wholeSheet, rowEnd: 0, colEnd: 0 }])
      expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 30 })
    } finally {
      scan.mockRestore()
    }
  })
})
