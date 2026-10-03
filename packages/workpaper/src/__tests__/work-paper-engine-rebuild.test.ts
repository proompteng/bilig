import { describe, expect, it } from 'vitest'

import { WorkPaper, type WorkPaperCellAddress } from '../index.js'

function cell(sheet: number, row: number, col: number): WorkPaperCellAddress {
  return { sheet, row, col }
}

function runtimeEngine(workbook: WorkPaper): object {
  const engine = Reflect.get(workbook, 'engine')
  if (typeof engine !== 'object' || engine === null) {
    throw new Error('Expected WorkPaper to expose a runtime engine in tests')
  }
  return engine
}

function dimensionCacheEngine(workbook: WorkPaper): object {
  const cache = Reflect.get(workbook, 'sheetDimensionCache')
  if (typeof cache !== 'object' || cache === null) {
    throw new Error('Expected WorkPaper to expose a sheet dimension cache in tests')
  }
  const engine = Reflect.get(cache, 'engine')
  if (typeof engine !== 'object' || engine === null) {
    throw new Error('Expected sheet dimension cache to expose its engine in tests')
  }
  return engine
}

describe('WorkPaper engine rebuilds', () => {
  it('keeps the active workbook usable when a config rebuild fails during reinitialization', () => {
    const workbook = WorkPaper.buildFromArray([[1, '=A1+1']], { language: 'enGB' })
    const originalEngine = runtimeEngine(workbook)
    const sheetId = workbook.getSheetId('Sheet1')!

    expect(() => {
      workbook.updateConfig({ language: 'timeout-rebuild-language', evaluationTimeoutMs: 0 })
    }).toThrow('Workbook evaluation timed out after 0ms')

    expect(workbook.getConfig().language).toBe('enGB')
    expect(workbook.getConfig().evaluationTimeoutMs).toBeUndefined()
    const restoredEngine = runtimeEngine(workbook)
    expect(restoredEngine).not.toBe(originalEngine)
    expect(dimensionCacheEngine(workbook)).toBe(restoredEngine)
    expect(workbook.getSheetNames()).toEqual(['Sheet1'])
    expect(workbook.getRangeSerialized({ start: cell(sheetId, 0, 0), end: cell(sheetId, 0, 1) })).toEqual([[1, '=A1+1']])
    expect(workbook.getSheetDimensions(sheetId)).toEqual({ width: 2, height: 1 })
  })

  it('rebinds the sheet dimension cache after config rebuilds replace the engine', () => {
    const workbook = WorkPaper.buildFromArray([[1]], { language: 'enGB' })
    const originalEngine = runtimeEngine(workbook)

    workbook.updateConfig({ language: 'rebuilt-dimension-cache-language' })

    const rebuiltEngine = runtimeEngine(workbook)
    expect(rebuiltEngine).not.toBe(originalEngine)
    expect(dimensionCacheEngine(workbook)).toBe(rebuiltEngine)

    const sheetId = workbook.getSheetId('Sheet1')!
    expect(workbook.getSheetDimensions(sheetId)).toEqual({ width: 1, height: 1 })
    workbook.setCellContents(cell(sheetId, 2, 2), 5)
    expect(workbook.getSheetDimensions(sheetId)).toEqual({ width: 3, height: 3 })
  })

  it('keeps the sheet dimension cache bound after transaction rollback rebuilds the engine', () => {
    const workbook = WorkPaper.buildFromArray([[1]], { language: 'enGB' })
    const originalEngine = runtimeEngine(workbook)

    expect(() => {
      workbook.transaction(() => {
        workbook.updateConfig({ language: 'transaction-dimension-cache-language' })
        const transactionEngine = runtimeEngine(workbook)
        expect(transactionEngine).not.toBe(originalEngine)
        expect(dimensionCacheEngine(workbook)).toBe(transactionEngine)
        throw new Error('rollback rebuilt engine')
      })
    }).toThrow('rollback rebuilt engine')

    const restoredEngine = runtimeEngine(workbook)
    expect(restoredEngine).not.toBe(originalEngine)
    expect(dimensionCacheEngine(workbook)).toBe(restoredEngine)
    expect(workbook.getConfig().language).toBe('enGB')

    const sheetId = workbook.getSheetId('Sheet1')!
    expect(workbook.getSheetDimensions(sheetId)).toEqual({ width: 1, height: 1 })
    workbook.setCellContents(cell(sheetId, 3, 1), 9)
    expect(workbook.getSheetDimensions(sheetId)).toEqual({ width: 2, height: 4 })
  })
})
