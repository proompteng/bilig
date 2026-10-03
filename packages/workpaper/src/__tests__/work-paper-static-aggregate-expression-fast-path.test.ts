import { describe, expect, it } from 'vitest'
import { ValueTag, type CellValue } from '@bilig/protocol'
import { WorkPaper, type WorkPaperCellAddress } from '../index.js'

function cell(sheet: number, row: number, col: number): WorkPaperCellAddress {
  return { sheet, row, col }
}

function expectNumberClose(value: CellValue, expected: number): void {
  expect(value.tag).toBe(ValueTag.Number)
  if (value.tag !== ValueTag.Number) {
    throw new Error(`Expected number ${String(expected)}, received ${JSON.stringify(value)}`)
  }
  expect(value.value).toBeCloseTo(expected, 12)
}

function buildRangeStatsWorkbook(rowCount: number): WorkPaper {
  const rows = Array.from({ length: rowCount }, (_unused, rowIndex) => {
    const value = rowIndex + 1
    return [value, value, value]
  })
  rows[0][3] = `=AVERAGE(A1:A${String(rowCount)})+MAX(B1:B${String(rowCount)})-MIN(C1:C${String(rowCount)})`
  return WorkPaper.buildFromSheets({ Bench: rows })
}

describe('static aggregate expression leaf fast path', () => {
  it('updates composite AVERAGE MAX MIN formulas without JS or WASM recalc', () => {
    const rowCount = 2_000
    const workbook = buildRangeStatsWorkbook(rowCount)
    const sheetId = workbook.getSheetId('Bench')!

    expectNumberClose(workbook.getCellValue(cell(sheetId, 0, 3)), 2_999.5)

    workbook.resetPerformanceCounters()
    const changes = workbook.setCellContents(cell(sheetId, 999, 1), 9_999)

    expect(new Set(changes.map((change) => (change.kind === 'cell' ? change.a1 : change.kind)))).toEqual(new Set(['B1000', 'D1']))
    expectNumberClose(workbook.getCellValue(cell(sheetId, 0, 3)), 10_998.5)
    expect(workbook.getStats().lastMetrics).toMatchObject({
      dirtyFormulaCount: 0,
      jsFormulaCount: 0,
      wasmFormulaCount: 0,
    })
    expect(workbook.getPerformanceCounters()).toMatchObject({
      directFormulaKernelSyncOnlyRecalcSkips: 1,
      formulasBound: 0,
      topoRepairs: 0,
    })
  })
})
