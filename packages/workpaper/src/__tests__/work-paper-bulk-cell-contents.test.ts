import { describe, expect, it, vi } from 'vitest'
import type { SheetRecord, SpreadsheetEngine } from '@bilig/core'
import { ValueTag } from '@bilig/protocol'
import {
  WorkPaper,
  type WorkPaperCellAddress,
  type WorkPaperCellValueUpdate,
  type WorkPaperSheetCellValueUpdate,
  type WorkPaperSheetRangeValues,
} from '../index.js'
import { hasDeferredTrackedIndexChanges } from '../tracked-cell-index-changes.js'

function cell(sheet: number, row: number, col: number): WorkPaperCellAddress {
  return { sheet, row, col }
}

function buildTwoInputFormulaRows(rowCount: number): (number | string)[][] {
  return Array.from({ length: rowCount }, (_, row) => {
    const rowNumber = row + 1
    return [rowNumber, rowNumber * 2, `=A${rowNumber}+B${rowNumber}`, `=A${rowNumber}*B${rowNumber}`]
  })
}

function buildTwoColumnUpdates(sheetId: number, rowCount: number): WorkPaperCellValueUpdate[] {
  return Array.from({ length: rowCount * 2 }, (_value, index) => {
    const row = Math.floor(index / 2)
    const col = index % 2
    return {
      address: cell(sheetId, row, col),
      value: col === 0 ? row * 3 : row * 5,
    }
  })
}

function buildTwoColumnSheetUpdates(rowCount: number): WorkPaperSheetCellValueUpdate[] {
  return Array.from({ length: rowCount * 2 }, (_value, index) => {
    const row = Math.floor(index / 2)
    const col = index % 2
    return {
      row,
      col,
      value: col === 0 ? row * 3 : row * 5,
    }
  })
}

function buildTwoColumnRangeValues(rowCount: number): WorkPaperSheetRangeValues {
  return Array.from({ length: rowCount }, (_value, row) => [row * 3, row * 5])
}

function buildDenseNumericRange(rowCount: number, colCount: number, offset = 0): WorkPaperSheetRangeValues {
  return Array.from({ length: rowCount }, (_rowValue, row) =>
    Array.from({ length: colCount }, (_colValue, col) => (row + 1) * (col + 2) + offset),
  )
}

function hasWorkbookStore(value: unknown): value is Pick<SpreadsheetEngine, 'workbook'> {
  return typeof value === 'object' && value !== null && 'workbook' in value
}

interface NumericBatchOperations {
  applyExistingNumericCellMutationsAtNow: (...args: unknown[]) => boolean
}

interface ExistingNumericBatchEngine {
  tryApplyExistingNumericCellMutationsAt: (...args: unknown[]) => boolean
}

function hasNumericBatchOperations(value: unknown): value is NumericBatchOperations {
  return typeof value === 'object' && value !== null && typeof Reflect.get(value, 'applyExistingNumericCellMutationsAtNow') === 'function'
}

function hasExistingNumericBatchEngine(value: unknown): value is ExistingNumericBatchEngine {
  return typeof value === 'object' && value !== null && typeof Reflect.get(value, 'tryApplyExistingNumericCellMutationsAt') === 'function'
}

function numericBatchOperations(workbook: WorkPaper): NumericBatchOperations {
  const engine = Reflect.get(workbook, 'engine')
  const operations = Reflect.get(Reflect.get(engine, 'runtime'), 'operations')
  if (!hasNumericBatchOperations(operations)) {
    throw new Error('Expected existing numeric batch operation')
  }
  return operations
}

function existingNumericBatchEngine(workbook: WorkPaper): ExistingNumericBatchEngine {
  const engine = Reflect.get(workbook, 'engine')
  if (!hasExistingNumericBatchEngine(engine)) {
    throw new Error('Expected existing numeric batch engine')
  }
  return engine
}

function getWorkbookSheetRecord(workbook: WorkPaper, sheetId: number): SheetRecord {
  const engine: unknown = Reflect.get(workbook, 'engine')
  if (!hasWorkbookStore(engine)) {
    throw new Error('Missing workbook engine')
  }
  const sheet = engine.workbook.getSheetById(sheetId)
  if (sheet === undefined) {
    throw new Error(`Missing sheet ${sheetId}`)
  }
  return sheet
}

describe('bulk cell values', () => {
  it('applies sparse updates through one tracked engine mutation', () => {
    const rowCount = 160
    const workbook = WorkPaper.buildFromSheets({ Bench: buildTwoInputFormulaRows(rowCount) })
    const sheetId = workbook.getSheetId('Bench')!
    const captureVisibilitySnapshot = vi.spyOn(workbook, 'captureVisibilitySnapshot').mockImplementation(() => {
      throw new Error('bulk sparse updates should not rebuild visibility snapshots')
    })
    const typedNumericBatch = vi.spyOn(numericBatchOperations(workbook), 'applyExistingNumericCellMutationsAtNow')

    try {
      const changes = workbook.setCellValues(buildTwoColumnUpdates(sheetId, rowCount))

      expect(changes).toHaveLength(rowCount * 4)
      expect(hasDeferredTrackedIndexChanges(changes)).toBe(true)
      expect(typedNumericBatch).toHaveReturnedWith(true)
      expect(workbook.getCellValue(cell(sheetId, rowCount - 1, 2))).toEqual({
        tag: ValueTag.Number,
        value: (rowCount - 1) * 8,
      })
      expect(workbook.getCellValue(cell(sheetId, rowCount - 1, 3))).toEqual({
        tag: ValueTag.Number,
        value: (rowCount - 1) * 3 * ((rowCount - 1) * 5),
      })
      expect(workbook.getPerformanceCounters()).toMatchObject({
        directScalarDeltaApplications: rowCount * 2,
        directScalarDeltaOnlyRecalcSkips: 1,
      })
    } finally {
      captureVisibilitySnapshot.mockRestore()
      typedNumericBatch.mockRestore()
    }
  })

  it('applies same-sheet sparse updates without per-update sheet addresses', () => {
    const rowCount = 160
    const workbook = WorkPaper.buildFromSheets({ Bench: buildTwoInputFormulaRows(rowCount) })
    const sheetId = workbook.getSheetId('Bench')!
    const changes = workbook.setSheetCellValues(sheetId, buildTwoColumnSheetUpdates(rowCount))

    expect(changes).toHaveLength(rowCount * 4)
    expect(hasDeferredTrackedIndexChanges(changes)).toBe(true)
    expect(workbook.getCellValue(cell(sheetId, rowCount - 1, 2))).toEqual({
      tag: ValueTag.Number,
      value: (rowCount - 1) * 8,
    })
    expect(workbook.getPerformanceCounters()).toMatchObject({
      directScalarDeltaApplications: rowCount * 2,
      directScalarDeltaOnlyRecalcSkips: 1,
    })
  })

  it('keeps public literal batches on the same scalar-delta bulk queue', () => {
    const rowCount = 160
    const workbook = WorkPaper.buildFromSheets({
      Bench: Array.from({ length: rowCount }, (_, row) => [row + 1, `=A${row + 1}*2`]),
    })
    const sheetId = workbook.getSheetId('Bench')!
    const captureVisibilitySnapshot = vi.spyOn(workbook, 'captureVisibilitySnapshot').mockImplementation(() => {
      throw new Error('public literal batches should not rebuild visibility snapshots')
    })
    const typedNumericBatch = vi.spyOn(existingNumericBatchEngine(workbook), 'tryApplyExistingNumericCellMutationsAt')

    try {
      workbook.resetPerformanceCounters()
      const changes = workbook.batch(() => {
        for (let row = 0; row < rowCount; row += 1) {
          workbook.setCellContents(cell(sheetId, row, 0), row * 3)
        }
      })

      expect(changes).toHaveLength(rowCount * 2)
      expect(hasDeferredTrackedIndexChanges(changes)).toBe(true)
      expect(typedNumericBatch).toHaveReturnedWith(true)
      expect(workbook.getCellValue(cell(sheetId, rowCount - 1, 1))).toEqual({
        tag: ValueTag.Number,
        value: (rowCount - 1) * 6,
      })
      expect(workbook.getPerformanceCounters()).toMatchObject({
        changedCellPayloadsBuilt: 0,
        directScalarDeltaApplications: rowCount,
        directScalarDeltaOnlyRecalcSkips: 1,
      })
    } finally {
      captureVisibilitySnapshot.mockRestore()
      typedNumericBatch.mockRestore()
    }
  })

  it('applies dense same-sheet range values without caller-allocated update objects', () => {
    const rowCount = 160
    const workbook = WorkPaper.buildFromSheets({ Bench: buildTwoInputFormulaRows(rowCount) })
    const sheetId = workbook.getSheetId('Bench')!
    const changes = workbook.setSheetRangeValues(sheetId, 0, 0, buildTwoColumnRangeValues(rowCount))

    expect(changes).toHaveLength(rowCount * 4)
    expect(hasDeferredTrackedIndexChanges(changes)).toBe(true)
    expect(workbook.getCellValue(cell(sheetId, rowCount - 1, 2))).toEqual({
      tag: ValueTag.Number,
      value: (rowCount - 1) * 8,
    })
    expect(workbook.getCellValue(cell(sheetId, rowCount - 1, 3))).toEqual({
      tag: ValueTag.Number,
      value: (rowCount - 1) * 3 * ((rowCount - 1) * 5),
    })
    expect(workbook.getPerformanceCounters()).toMatchObject({
      directScalarDeltaApplications: rowCount * 2,
      directScalarDeltaOnlyRecalcSkips: 1,
    })
  })

  it('applies fresh dense numeric range values through one rectangular kernel-sync path', () => {
    const rowCount = 80
    const colCount = 8
    const values: WorkPaperSheetRangeValues = Array.from({ length: rowCount }, (_rowValue, row) =>
      Array.from({ length: colCount }, (_colValue, col) => (row + 1) * (col + 2)),
    )
    const workbook = WorkPaper.buildFromSheets({ Bench: [] })
    const sheetId = workbook.getSheetId('Bench')!
    const sheet = getWorkbookSheetRecord(workbook, sheetId)
    const captureVisibilitySnapshot = vi.spyOn(workbook, 'captureVisibilitySnapshot').mockImplementation(() => {
      throw new Error('fresh dense range values should not rebuild visibility snapshots')
    })
    const getPhysical = vi.spyOn(sheet.grid, 'getPhysical')
    const forEachPhysicalRangeEntry = vi.spyOn(sheet.grid, 'forEachPhysicalRangeEntry')

    try {
      workbook.resetPerformanceCounters()
      const changes = workbook.setSheetRangeValues(sheetId, 0, 0, values)
      const physicalLookupsDuringWrite = getPhysical.mock.calls.length

      expect(changes).toHaveLength(rowCount * colCount)
      expect(hasDeferredTrackedIndexChanges(changes)).toBe(true)
      expect(forEachPhysicalRangeEntry).not.toHaveBeenCalled()
      expect(physicalLookupsDuringWrite).toBeLessThanOrEqual(rowCount * colCount * 2 + 1)
      expect(workbook.getCellValue(cell(sheetId, rowCount - 1, colCount - 1))).toEqual({
        tag: ValueTag.Number,
        value: rowCount * (colCount + 1),
      })
      expect(workbook.getSheetDimensions(sheetId)).toEqual({ height: rowCount, width: colCount })
      expect(workbook.getPerformanceCounters()).toMatchObject({
        changedCellPayloadsBuilt: 0,
        kernelSyncOnlyRecalcSkips: 1,
        topoRebuilds: 0,
      })
    } finally {
      forEachPhysicalRangeEntry.mockRestore()
      getPhysical.mockRestore()
      captureVisibilitySnapshot.mockRestore()
    }
  })

  it('overwrites existing dense numeric range values through one rectangular kernel-sync path', () => {
    const rowCount = 80
    const colCount = 8
    const workbook = WorkPaper.buildFromSheets({ Bench: buildDenseNumericRange(rowCount, colCount) })
    const sheetId = workbook.getSheetId('Bench')!
    const sheet = getWorkbookSheetRecord(workbook, sheetId)
    const captureVisibilitySnapshot = vi.spyOn(workbook, 'captureVisibilitySnapshot').mockImplementation(() => {
      throw new Error('existing dense range values should not rebuild visibility snapshots')
    })
    const getPhysical = vi.spyOn(sheet.grid, 'getPhysical')
    const forEachPhysicalRangeEntry = vi.spyOn(sheet.grid, 'forEachPhysicalRangeEntry')

    try {
      workbook.resetPerformanceCounters()
      const changes = workbook.setSheetRangeValues(sheetId, 0, 0, buildDenseNumericRange(rowCount, colCount, 10_000))
      const physicalLookupsDuringWrite = getPhysical.mock.calls.length

      expect(changes).toHaveLength(rowCount * colCount)
      expect(hasDeferredTrackedIndexChanges(changes)).toBe(true)
      expect(forEachPhysicalRangeEntry).toHaveBeenCalledTimes(1)
      expect(physicalLookupsDuringWrite).toBe(0)
      expect(workbook.getCellValue(cell(sheetId, rowCount - 1, colCount - 1))).toEqual({
        tag: ValueTag.Number,
        value: rowCount * (colCount + 1) + 10_000,
      })
      expect(workbook.getSheetDimensions(sheetId)).toEqual({ height: rowCount, width: colCount })
      expect(workbook.getPerformanceCounters()).toMatchObject({
        changedCellPayloadsBuilt: 0,
        kernelSyncOnlyRecalcSkips: 1,
        topoRebuilds: 0,
      })

      getPhysical.mockClear()
      forEachPhysicalRangeEntry.mockClear()
      const undoChanges = workbook.undo()
      expect(undoChanges).toHaveLength(rowCount * colCount)
      expect(workbook.getCellValue(cell(sheetId, rowCount - 1, colCount - 1))).toEqual({
        tag: ValueTag.Number,
        value: rowCount * (colCount + 1),
      })
    } finally {
      forEachPhysicalRangeEntry.mockRestore()
      getPhysical.mockRestore()
      captureVisibilitySnapshot.mockRestore()
    }
  })

  it('pre-resolves physical range indexes across grid block boundaries', () => {
    const workbook = WorkPaper.buildFromSheets({ Bench: buildDenseNumericRange(132, 40) })
    const sheetId = workbook.getSheetId('Bench')!
    const sheet = getWorkbookSheetRecord(workbook, sheetId)
    const getPhysical = vi.spyOn(sheet.grid, 'getPhysical')
    const forEachPhysicalRangeEntry = vi.spyOn(sheet.grid, 'forEachPhysicalRangeEntry')

    try {
      const changes = workbook.setSheetRangeValues(sheetId, 126, 30, buildDenseNumericRange(4, 8, 20_000))

      expect(changes).toHaveLength(32)
      expect(forEachPhysicalRangeEntry).toHaveBeenCalledTimes(1)
      expect(getPhysical.mock.calls.length).toBe(0)
      expect(workbook.getCellValue(cell(sheetId, 126, 30))).toEqual({ tag: ValueTag.Number, value: 20_002 })
      expect(workbook.getCellValue(cell(sheetId, 129, 37))).toEqual({ tag: ValueTag.Number, value: 20_036 })
    } finally {
      forEachPhysicalRangeEntry.mockRestore()
      getPhysical.mockRestore()
    }
  })

  it('applies null clears and ragged range value rows in one public call', () => {
    const workbook = WorkPaper.buildFromSheets({
      Bench: [
        [1, 2, '=A1+B1'],
        [3, 4, '=A2+B2'],
      ],
    })
    const sheetId = workbook.getSheetId('Bench')!

    workbook.setSheetRangeValues(sheetId, 0, 0, [[null, 5], [7]])

    expect(workbook.getCellValue(cell(sheetId, 0, 0))).toEqual({ tag: ValueTag.Empty })
    expect(workbook.getCellValue(cell(sheetId, 0, 2))).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(workbook.getCellValue(cell(sheetId, 1, 1))).toEqual({ tag: ValueTag.Number, value: 4 })
    expect(workbook.getCellValue(cell(sheetId, 1, 2))).toEqual({ tag: ValueTag.Number, value: 11 })
  })

  it('defers sparse updates while evaluation is suspended and flushes once on resume', () => {
    const rowCount = 96
    const workbook = WorkPaper.buildFromSheets({ Bench: buildTwoInputFormulaRows(rowCount) })
    const sheetId = workbook.getSheetId('Bench')!

    workbook.suspendEvaluation()
    expect(workbook.setCellValues(buildTwoColumnUpdates(sheetId, rowCount))).toEqual([])
    const changes = workbook.resumeEvaluation()

    expect(changes).toHaveLength(rowCount * 4)
    expect(hasDeferredTrackedIndexChanges(changes)).toBe(true)
    expect(workbook.getCellValue(cell(sheetId, rowCount - 1, 2))).toEqual({
      tag: ValueTag.Number,
      value: (rowCount - 1) * 8,
    })
    expect(workbook.getPerformanceCounters()).toMatchObject({
      directScalarDeltaApplications: rowCount * 2,
      directScalarDeltaOnlyRecalcSkips: 1,
    })
  })

  it('keeps last-write-wins sparse value patches in one public call', () => {
    const workbook = WorkPaper.buildFromSheets({ Bench: [[1, 2, '=A1+B1']] })
    const sheetId = workbook.getSheetId('Bench')!

    const changes = workbook.setCellValues([
      { address: cell(sheetId, 0, 0), value: 4 },
      { address: cell(sheetId, 0, 0), value: 5 },
      { address: cell(sheetId, 0, 1), value: 7 },
    ])

    expect(changes.map((change) => (change.kind === 'cell' ? change.a1 : change.name))).toEqual(['A1', 'B1', 'C1'])
    expect(workbook.getCellValue(cell(sheetId, 0, 2))).toEqual({ tag: ValueTag.Number, value: 12 })
  })

  it('applies addressed sparse value patches across sheets in one public call', () => {
    const workbook = WorkPaper.buildFromSheets({
      First: [[1, '=A1*2']],
      Second: [[2, '=A1*3']],
    })
    const firstSheet = workbook.getSheetId('First')!
    const secondSheet = workbook.getSheetId('Second')!

    const changes = workbook.setCellValues([
      { address: cell(firstSheet, 0, 0), value: 5 },
      { address: cell(secondSheet, 0, 0), value: 7 },
      { address: cell(firstSheet, 0, 0), value: 11 },
    ])

    expect(changes.map((change) => (change.kind === 'cell' ? `${change.sheetName}!${change.a1}` : change.name))).toEqual([
      'First!A1',
      'First!B1',
      'Second!A1',
      'Second!B1',
    ])
    expect(workbook.getCellValue(cell(firstSheet, 0, 1))).toEqual({ tag: ValueTag.Number, value: 22 })
    expect(workbook.getCellValue(cell(secondSheet, 0, 1))).toEqual({ tag: ValueTag.Number, value: 21 })
  })

  it('rejects formula strings so formula rewrites keep using the semantic single-cell path', () => {
    const workbook = WorkPaper.buildFromSheets({ Bench: [[1, 2, '=A1+B1']] })
    const sheetId = workbook.getSheetId('Bench')!

    expect(() => {
      workbook.setCellValues([{ address: cell(sheetId, 0, 2), value: '=A1*B1' }])
    }).toThrow('Bulk cell value updates require literal values')
    expect(() => {
      workbook.setSheetCellValues(sheetId, [{ row: 0, col: 2, value: '=A1*B1' }])
    }).toThrow('Bulk cell value updates require literal values')
    expect(() => {
      workbook.setSheetRangeValues(sheetId, 0, 0, [[1, '=A1*B1']])
    }).toThrow('Bulk cell value updates require literal values')
  })
})
