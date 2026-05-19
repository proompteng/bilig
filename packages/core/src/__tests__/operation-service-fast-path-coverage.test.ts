import { describe, expect, it, vi } from 'vitest'
import { ValueTag } from '@bilig/protocol'
import type {
  EngineCellMutationRef,
  EngineExistingNumericCellMutationRef,
  EngineExistingNumericCellMutationResult,
} from '../cell-mutations-at.js'
import { SpreadsheetEngine } from '../engine.js'
import { DirectFormulaIndexCollection } from '../engine/services/direct-formula-index-collection.js'

function existingNumericMutationFastPath(
  engine: SpreadsheetEngine,
): (request: EngineExistingNumericCellMutationRef) => EngineExistingNumericCellMutationResult | null {
  const runtime = Reflect.get(engine, 'runtime')
  if (typeof runtime !== 'object' || runtime === null) {
    throw new TypeError('Expected engine runtime')
  }
  const operations = Reflect.get(runtime, 'operations')
  if (typeof operations !== 'object' || operations === null) {
    throw new TypeError('Expected operation service')
  }
  const fastPath = Reflect.get(operations, 'applyExistingNumericCellMutationAtNow')
  if (typeof fastPath !== 'function') {
    throw new TypeError('Expected existing numeric mutation fast path')
  }
  return (request: EngineExistingNumericCellMutationRef): EngineExistingNumericCellMutationResult | null =>
    Reflect.apply(fastPath, operations, [request])
}

function operationHook<Args extends readonly unknown[], Result>(engine: SpreadsheetEngine, name: string): (...args: Args) => Result {
  const runtime = Reflect.get(engine, 'runtime')
  if (typeof runtime !== 'object' || runtime === null) {
    throw new TypeError('Expected engine runtime')
  }
  const operations = Reflect.get(runtime, 'operations')
  if (typeof operations !== 'object' || operations === null) {
    throw new TypeError('Expected operation service')
  }
  const hooks = Reflect.get(operations, '__testHooks')
  if (typeof hooks !== 'object' || hooks === null) {
    throw new TypeError('Expected operation hooks')
  }
  const hook = Reflect.get(hooks, name)
  if (typeof hook !== 'function') {
    throw new TypeError(`Expected operation hook ${name}`)
  }
  return (...args: Args): Result => Reflect.apply(hook, hooks, args)
}

function runtimeFormula(engine: SpreadsheetEngine, cellIndex: number): unknown {
  const formulas = Reflect.get(engine, 'formulas')
  if (typeof formulas !== 'object' || formulas === null) {
    throw new TypeError('Expected formulas store')
  }
  const get = Reflect.get(formulas, 'get')
  if (typeof get !== 'function') {
    throw new TypeError('Expected formulas get method')
  }
  return get.call(formulas, cellIndex)
}

type FormulaScanTable = {
  readonly forEach: (callback: unknown) => void
}

function isFormulaScanTable(value: unknown): value is FormulaScanTable {
  if (typeof value !== 'object' || value === null) {
    return false
  }
  return typeof Reflect.get(value, 'forEach') === 'function'
}

function formulaScanTable(engine: SpreadsheetEngine): FormulaScanTable {
  const state = Reflect.get(engine, 'state')
  if (typeof state !== 'object' || state === null) {
    throw new TypeError('Expected engine state')
  }
  const formulas = Reflect.get(state, 'formulas')
  if (!isFormulaScanTable(formulas)) {
    throw new TypeError('Expected formula scan table')
  }
  return formulas
}

function columnLabel(col: number): string {
  let index = col
  let label = ''
  do {
    label = String.fromCharCode(65 + (index % 26)) + label
    index = Math.floor(index / 26) - 1
  } while (index >= 0)
  return label
}

describe('operation-service dense mutation fast paths', () => {
  it('applies dense row-pair direct scalar batches without general recalculation', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'dense-row-pair-fast-path' })
    await engine.ready()
    engine.createSheet('Sheet1')

    for (let row = 1; row <= 20; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
      engine.setCellValue('Sheet1', `B${row}`, row + 100)
      engine.setCellFormula('Sheet1', `C${row}`, `A${row}+B${row}`)
      engine.setCellFormula('Sheet1', `D${row}`, `A${row}*B${row}`)
    }
    const general = vi.fn()
    const tracked = vi.fn()
    const watched = vi.fn()
    const unsubscribeGeneral = engine.subscribe(general)
    const unsubscribeTracked = engine.events.subscribeTracked(tracked)
    const unsubscribeWatched = engine.subscribeCell('Sheet1', 'C20', watched)

    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const refs: EngineCellMutationRef[] = []
    for (let row = 0; row < 20; row += 1) {
      refs.push(
        {
          sheetId,
          cellIndex: engine.workbook.getCellIndex('Sheet1', `A${row + 1}`),
          mutation: { kind: 'setCellValue', row, col: 0, value: row + 10 },
        },
        {
          sheetId,
          cellIndex: engine.workbook.getCellIndex('Sheet1', `B${row + 1}`),
          mutation: { kind: 'setCellValue', row, col: 1, value: row + 200 },
        },
      )
    }

    engine.resetPerformanceCounters()
    const applied = operationHook<[readonly EngineCellMutationRef[], null, number], boolean>(
      engine,
      'tryApplyDenseRowPairDirectScalarLiteralBatch',
    )(refs, null, 0)

    expect(applied).toBe(true)
    expect(engine.getCellValue('Sheet1', 'C20')).toEqual({ tag: ValueTag.Number, value: 19 + 10 + 19 + 200 })
    expect(engine.getCellValue('Sheet1', 'D20')).toEqual({ tag: ValueTag.Number, value: (19 + 10) * (19 + 200) })
    expect(engine.getPerformanceCounters().directScalarDeltaOnlyRecalcSkips).toBeGreaterThan(0)
    expect(general).toHaveBeenCalledWith(expect.objectContaining({ kind: 'batch', explicitChangedCount: 40 }))
    expect(tracked).toHaveBeenCalledWith(expect.objectContaining({ kind: 'batch', explicitChangedCount: 40 }))
    expect(watched).toHaveBeenCalled()
    unsubscribeWatched()
    unsubscribeTracked()
    unsubscribeGeneral()
  })

  it('applies dense single-input direct scalar batches without general recalculation', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'dense-single-input-fast-path' })
    await engine.ready()
    engine.createSheet('Sheet1')

    for (let row = 1; row <= 40; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
      engine.setCellFormula('Sheet1', `B${row}`, `A${row}*2`)
      engine.setCellFormula('Sheet1', `C${row}`, `A${row}+5`)
    }
    const general = vi.fn()
    const tracked = vi.fn()
    const watched = vi.fn()
    const unsubscribeGeneral = engine.subscribe(general)
    const unsubscribeTracked = engine.events.subscribeTracked(tracked)
    const unsubscribeWatched = engine.subscribeCell('Sheet1', 'B40', watched)

    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const refs: EngineCellMutationRef[] = []
    for (let row = 0; row < 40; row += 1) {
      refs.push({
        sheetId,
        cellIndex: engine.workbook.getCellIndex('Sheet1', `A${row + 1}`),
        mutation: { kind: 'setCellValue', row, col: 0, value: row + 50 },
      })
    }

    engine.resetPerformanceCounters()
    const undoOps = engine.applyCellMutationsAt(refs, 0)

    expect(undoOps).not.toBeNull()
    expect(engine.getCellValue('Sheet1', 'B40')).toEqual({ tag: ValueTag.Number, value: (39 + 50) * 2 })
    expect(engine.getCellValue('Sheet1', 'C40')).toEqual({ tag: ValueTag.Number, value: 39 + 50 + 5 })
    expect(engine.getPerformanceCounters().directScalarDeltaOnlyRecalcSkips).toBeGreaterThan(0)
    expect(general).toHaveBeenCalledWith(expect.objectContaining({ kind: 'batch', explicitChangedCount: 40 }))
    expect(tracked).toHaveBeenCalledWith(expect.objectContaining({ kind: 'batch', explicitChangedCount: 40 }))
    expect(watched).toHaveBeenCalled()
    unsubscribeWatched()
    unsubscribeTracked()
    unsubscribeGeneral()
  })

  it('keeps descending single-dependent direct scalar column batches ordered and delta-only', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'dense-single-dependent-descending-fast-path' })
    await engine.ready()
    engine.createSheet('Sheet1')

    const rowCount = 40
    for (let row = 1; row <= rowCount; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
      engine.setCellFormula('Sheet1', `B${row}`, `A${row}*3+7`)
    }

    const general = vi.fn()
    const tracked = vi.fn()
    const watched = vi.fn()
    const unsubscribeGeneral = engine.subscribe(general)
    const unsubscribeTracked = engine.events.subscribeTracked(tracked)
    const unsubscribeWatched = engine.subscribeCell('Sheet1', 'B1', watched)

    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const refs: EngineCellMutationRef[] = []
    for (let row = rowCount - 1; row >= 0; row -= 1) {
      refs.push({
        sheetId,
        cellIndex: engine.workbook.getCellIndex('Sheet1', `A${row + 1}`),
        mutation: { kind: 'setCellValue', row, col: 0, value: row + 100 },
      })
    }

    engine.resetPerformanceCounters()
    const undoOps = engine.applyCellMutationsAt(refs, 0)

    expect(undoOps).not.toBeNull()
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 100 * 3 + 7 })
    expect(engine.getCellValue('Sheet1', 'B40')).toEqual({ tag: ValueTag.Number, value: 139 * 3 + 7 })
    expect(engine.getPerformanceCounters().directScalarDeltaApplications).toBe(rowCount)
    expect(engine.getPerformanceCounters().directScalarDeltaOnlyRecalcSkips).toBe(1)
    expect(engine.getLastMetrics()).toMatchObject({
      changedInputCount: rowCount,
      dirtyFormulaCount: 0,
      jsFormulaCount: 0,
      wasmFormulaCount: 0,
    })
    expect(general).toHaveBeenCalledWith(expect.objectContaining({ kind: 'batch', explicitChangedCount: rowCount }))
    expect(tracked).toHaveBeenCalledWith(expect.objectContaining({ kind: 'batch', explicitChangedCount: rowCount }))
    expect(watched).toHaveBeenCalled()
    unsubscribeWatched()
    unsubscribeTracked()
    unsubscribeGeneral()
  })

  it('handles dense aggregate-only numeric column coordinate batches', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'dense-aggregate-fast-path' })
    await engine.ready()
    engine.createSheet('Sheet1')

    for (let row = 1; row <= 40; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
    }
    engine.setCellFormula('Sheet1', 'C1', 'SUM(A1:A40)')
    engine.setCellFormula('Sheet1', 'C2', 'AVERAGE(A1:A40)')
    const general = vi.fn()
    const tracked = vi.fn()
    const watched = vi.fn()
    const unsubscribeGeneral = engine.subscribe(general)
    const unsubscribeTracked = engine.events.subscribeTracked(tracked)
    const unsubscribeWatched = engine.subscribeCell('Sheet1', 'C1', watched)

    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const refs: EngineCellMutationRef[] = []
    for (let row = 0; row < 40; row += 1) {
      refs.push({
        sheetId,
        cellIndex: engine.workbook.getCellIndex('Sheet1', `A${row + 1}`),
        mutation: { kind: 'setCellValue', row, col: 0, value: row + 100 },
      })
    }

    engine.resetPerformanceCounters()
    const undoOps = engine.applyCellMutationsAt(refs, 0)

    expect(undoOps).not.toBeNull()
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 4780 })
    expect(engine.getCellValue('Sheet1', 'C2')).toEqual({ tag: ValueTag.Number, value: 119.5 })
    expect(general).toHaveBeenCalledWith(expect.objectContaining({ kind: 'batch', explicitChangedCount: 40 }))
    expect(tracked).toHaveBeenCalledWith(expect.objectContaining({ kind: 'batch', explicitChangedCount: 40 }))
    expect(watched).toHaveBeenCalled()
    unsubscribeWatched()
    unsubscribeTracked()
    unsubscribeGeneral()
  })

  it('applies dense rectangular row-sum aggregate batches without region queries', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'dense-rectangular-row-sum-fast-path' })
    await engine.ready()
    engine.createSheet('Sheet1')

    const rowCount = 12
    const inputCols = 4
    for (let row = 1; row <= rowCount; row += 1) {
      for (let col = 0; col < inputCols; col += 1) {
        engine.setCellValue('Sheet1', `${String.fromCharCode(65 + col)}${row}`, row * (col + 1))
      }
      engine.setCellFormula('Sheet1', `E${row}`, `SUM(A${row}:D${row})`)
    }

    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const refs: EngineCellMutationRef[] = []
    for (let row = 0; row < rowCount; row += 1) {
      for (let col = 0; col < inputCols; col += 1) {
        refs.push({
          sheetId,
          cellIndex: engine.workbook.getCellIndex('Sheet1', `${String.fromCharCode(65 + col)}${row + 1}`),
          mutation: { kind: 'setCellValue', row, col, value: (row + 1) * (col + 2) },
        })
      }
    }

    engine.resetPerformanceCounters()
    const undoOps = engine.applyCellMutationsAt(refs, 0)

    expect(undoOps).not.toBeNull()
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 14 })
    expect(engine.getCellValue('Sheet1', 'E12')).toEqual({ tag: ValueTag.Number, value: 168 })
    expect(engine.getLastMetrics()).toMatchObject({
      changedInputCount: rowCount * inputCols,
      dirtyFormulaCount: 0,
      jsFormulaCount: 0,
      wasmFormulaCount: 0,
    })
    expect(engine.getPerformanceCounters().directAggregateDeltaApplications).toBe(rowCount)
    expect(engine.getPerformanceCounters().directAggregateDeltaOnlyRecalcSkips).toBe(1)
    expect(engine.getPerformanceCounters().regionQueryIndexBuilds).toBe(0)
    expect(engine.getPerformanceCounters().directAggregateScanEvaluations).toBe(0)
  })

  it('applies dense rectangular row-sum aggregate clears without region queries', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'dense-rectangular-row-sum-clear-fast-path' })
    await engine.ready()
    engine.createSheet('Sheet1')

    const rowCount = 12
    const inputCols = 4
    for (let row = 1; row <= rowCount; row += 1) {
      for (let col = 0; col < inputCols; col += 1) {
        engine.setCellValue('Sheet1', `${String.fromCharCode(65 + col)}${row}`, row * (col + 1))
      }
      engine.setCellFormula('Sheet1', `E${row}`, `SUM(A${row}:D${row})`)
    }

    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const refs: EngineCellMutationRef[] = []
    for (let row = 0; row < rowCount; row += 1) {
      for (let col = 0; col < inputCols; col += 1) {
        refs.push({
          sheetId,
          cellIndex: engine.workbook.getCellIndex('Sheet1', `${String.fromCharCode(65 + col)}${row + 1}`),
          mutation: { kind: 'clearCell', row, col },
        })
      }
    }

    engine.resetPerformanceCounters()
    const undoOps = engine.applyCellMutationsAt(refs, 0)

    expect(undoOps).not.toBeNull()
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Empty })
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(engine.getCellValue('Sheet1', 'E12')).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(engine.getLastMetrics()).toMatchObject({
      changedInputCount: rowCount * inputCols,
      dirtyFormulaCount: 0,
      jsFormulaCount: 0,
      wasmFormulaCount: 0,
    })
    expect(engine.getPerformanceCounters().directAggregateDeltaApplications).toBe(rowCount)
    expect(engine.getPerformanceCounters().directAggregateDeltaOnlyRecalcSkips).toBe(1)
    expect(engine.getPerformanceCounters().regionQueryIndexBuilds).toBe(0)
    expect(engine.getPerformanceCounters().directAggregateScanEvaluations).toBe(0)
  })

  it('applies fresh dense rectangular numeric batches below aggregate ranges without region queries', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'fresh-dense-rectangular-fast-path' })
    await engine.ready()
    engine.createSheet('Sheet1')

    const existingRows = 12
    const appendRows = 8
    const inputCols = 4
    for (let row = 1; row <= existingRows; row += 1) {
      for (let col = 0; col < inputCols; col += 1) {
        engine.setCellValue('Sheet1', `${String.fromCharCode(65 + col)}${row}`, row * (col + 1))
      }
      engine.setCellFormula('Sheet1', `E${row}`, `SUM(A${row}:D${row})`)
    }
    engine.insertRows('Sheet1', existingRows, appendRows)

    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const refs: EngineCellMutationRef[] = []
    for (let row = 0; row < appendRows; row += 1) {
      for (let col = 0; col < inputCols; col += 1) {
        refs.push({
          sheetId,
          mutation: { kind: 'setCellValue', row: existingRows + row, col, value: (row + 1) * (col + 2) },
        })
      }
    }

    const ensureCellAt = vi.spyOn(engine.workbook, 'ensureCellAt')
    const attachAllocatedCellWithLogicalAxisIds = vi.spyOn(engine.workbook, 'attachAllocatedCellWithLogicalAxisIds')
    const allocateReserved = vi.spyOn(engine.workbook.cellStore, 'allocateReserved')
    engine.resetPerformanceCounters()
    try {
      const undoOps = engine.applyCellMutationsAt(refs, refs.length)

      expect(undoOps).not.toBeNull()
      expect(ensureCellAt).not.toHaveBeenCalled()
      expect(attachAllocatedCellWithLogicalAxisIds).not.toHaveBeenCalled()
      expect(allocateReserved).not.toHaveBeenCalled()
      expect(engine.getCellValue('Sheet1', 'A13')).toEqual({ tag: ValueTag.Number, value: 2 })
      expect(engine.getCellValue('Sheet1', 'D20')).toEqual({ tag: ValueTag.Number, value: 40 })
      expect(engine.getCellValue('Sheet1', 'E12')).toEqual({ tag: ValueTag.Number, value: 120 })
      expect(engine.getLastMetrics()).toMatchObject({
        changedInputCount: appendRows * inputCols,
        dirtyFormulaCount: 0,
        jsFormulaCount: 0,
        wasmFormulaCount: 0,
      })
      expect(engine.getPerformanceCounters().kernelSyncOnlyRecalcSkips).toBe(1)
      expect(engine.getPerformanceCounters().regionQueryIndexBuilds).toBe(0)
      expect(engine.getPerformanceCounters().directAggregateScanEvaluations).toBe(0)
    } finally {
      allocateReserved.mockRestore()
      attachAllocatedCellWithLogicalAxisIds.mockRestore()
      ensureCellAt.mockRestore()
    }
  })

  it('validates fresh physical dense numeric rectangles without logical visibility lookups', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'fresh-physical-rectangular-fast-path' })
    await engine.ready()
    engine.createSheet('Sheet1')

    const rowCount = 8
    const inputCols = 4
    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const sheet = engine.workbook.getSheetById(sheetId)!
    const refs: EngineCellMutationRef[] = []
    for (let row = 0; row < rowCount; row += 1) {
      for (let col = 0; col < inputCols; col += 1) {
        refs.push({
          sheetId,
          mutation: { kind: 'setCellValue', row, col, value: (row + 1) * (col + 2) },
        })
      }
    }

    const getVisibleCell = vi.spyOn(sheet.logical, 'getVisibleCell').mockImplementation(() => {
      throw new Error('fresh physical dense rectangle should validate emptiness from the physical grid')
    })
    engine.resetPerformanceCounters()
    try {
      const undoOps = engine.applyCellMutationsAt(refs, refs.length)

      expect(undoOps).not.toBeNull()
      expect(getVisibleCell).not.toHaveBeenCalled()
      expect(engine.getLastMetrics()).toMatchObject({
        changedInputCount: rowCount * inputCols,
        dirtyFormulaCount: 0,
        jsFormulaCount: 0,
        wasmFormulaCount: 0,
      })
      expect(engine.getPerformanceCounters().kernelSyncOnlyRecalcSkips).toBe(1)
    } finally {
      getVisibleCell.mockRestore()
    }
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'D8')).toEqual({ tag: ValueTag.Number, value: 40 })
  })

  it('stores fresh row aggregate formula results without aggregate scans', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'fresh-row-aggregate-formula-results' })
    await engine.ready()
    engine.createSheet('Sheet1')

    const existingRows = 12
    const appendRows = 8
    const inputCols = 4
    for (let row = 1; row <= existingRows; row += 1) {
      for (let col = 0; col < inputCols; col += 1) {
        engine.setCellValue('Sheet1', `${String.fromCharCode(65 + col)}${row}`, row * (col + 1))
      }
      engine.setCellFormula('Sheet1', `E${row}`, `SUM(A${row}:D${row})`)
    }
    engine.insertRows('Sheet1', existingRows, appendRows)

    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const valueRefs: EngineCellMutationRef[] = []
    const formulaRefs: EngineCellMutationRef[] = []
    for (let row = 0; row < appendRows; row += 1) {
      const rowIndex = existingRows + row
      const rowNumber = rowIndex + 1
      for (let col = 0; col < inputCols; col += 1) {
        valueRefs.push({
          sheetId,
          mutation: { kind: 'setCellValue', row: rowIndex, col, value: rowNumber * (col + 1) },
        })
      }
      formulaRefs.push({
        sheetId,
        mutation: { kind: 'setCellFormula', row: rowIndex, col: inputCols, formula: `SUM(A${rowNumber}:D${rowNumber})` },
      })
    }

    engine.applyCellMutationsAt(valueRefs, valueRefs.length)
    engine.resetPerformanceCounters()
    engine.applyCellMutationsAt(formulaRefs, formulaRefs.length)

    expect(engine.getCellValue('Sheet1', 'E13')).toEqual({ tag: ValueTag.Number, value: 130 })
    expect(engine.getCellValue('Sheet1', 'E20')).toEqual({ tag: ValueTag.Number, value: 200 })
    expect(engine.getLastMetrics()).toMatchObject({
      changedInputCount: appendRows,
      dirtyFormulaCount: 0,
      jsFormulaCount: 0,
      wasmFormulaCount: 0,
    })
    expect(engine.getPerformanceCounters().formulasBound).toBe(0)
    expect(engine.getPerformanceCounters().directAggregateScanEvaluations).toBe(0)
    expect(engine.getPerformanceCounters().directAggregateScanCells).toBe(0)

    engine.setCellValue('Sheet1', 'A13', 99)
    expect(engine.getCellValue('Sheet1', 'E13')).toEqual({ tag: ValueTag.Number, value: 216 })
  })

  it('applies fresh row aggregate formula batches without coordinate re-ensure or dirty recalculation', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'fresh-row-aggregate-formula-batch-fast-path' })
    await engine.ready()
    engine.createSheet('Sheet1')

    const existingRows = 16
    const appendRows = 80
    const inputCols = 4
    for (let row = 1; row <= existingRows; row += 1) {
      for (let col = 0; col < inputCols; col += 1) {
        engine.setCellValue('Sheet1', `${String.fromCharCode(65 + col)}${row}`, row * (col + 1))
      }
      engine.setCellFormula('Sheet1', `E${row}`, `SUM(A${row}:D${row})`)
    }
    engine.insertRows('Sheet1', existingRows, appendRows)

    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const valueRefs: EngineCellMutationRef[] = []
    const formulaRefs: EngineCellMutationRef[] = []
    for (let row = 0; row < appendRows; row += 1) {
      const rowIndex = existingRows + row
      const rowNumber = rowIndex + 1
      for (let col = 0; col < inputCols; col += 1) {
        valueRefs.push({
          sheetId,
          mutation: { kind: 'setCellValue', row: rowIndex, col, value: rowNumber * (col + 1) },
        })
      }
      formulaRefs.push({
        sheetId,
        mutation: { kind: 'setCellFormula', row: rowIndex, col: inputCols, formula: `SUM(A${rowNumber}:D${rowNumber})` },
      })
    }

    engine.applyCellMutationsAt(valueRefs, valueRefs.length)
    const ensureCellAt = vi.spyOn(engine.workbook, 'ensureCellAt')
    const attachAllocatedCellWithLogicalAxisIds = vi.spyOn(engine.workbook, 'attachAllocatedCellWithLogicalAxisIds')
    const allocateReserved = vi.spyOn(engine.workbook.cellStore, 'allocateReserved')
    try {
      engine.resetPerformanceCounters()
      const undoOps = engine.applyCellMutationsAt(formulaRefs, formulaRefs.length)

      expect(undoOps).not.toBeNull()
      expect(ensureCellAt).not.toHaveBeenCalled()
      expect(attachAllocatedCellWithLogicalAxisIds).not.toHaveBeenCalled()
      expect(allocateReserved).not.toHaveBeenCalled()
      expect(engine.getCellValue('Sheet1', 'E17')).toEqual({ tag: ValueTag.Number, value: 170 })
      expect(engine.getCellValue('Sheet1', `E${existingRows + appendRows}`)).toEqual({
        tag: ValueTag.Number,
        value: (existingRows + appendRows) * 10,
      })
      expect(engine.getLastMetrics()).toMatchObject({
        changedInputCount: appendRows,
        dirtyFormulaCount: 0,
        jsFormulaCount: 0,
        wasmFormulaCount: 0,
      })
      expect(engine.getPerformanceCounters()).toMatchObject({
        calcChainFullScans: 0,
        directAggregateScanCells: 0,
        directAggregateScanEvaluations: 0,
        directFormulaKernelSyncOnlyRecalcSkips: 1,
        formulasBound: 0,
        topoRepairs: 0,
      })
    } finally {
      ensureCellAt.mockRestore()
      attachAllocatedCellWithLogicalAxisIds.mockRestore()
      allocateReserved.mockRestore()
    }

    engine.setCellValue('Sheet1', 'A17', 99)
    expect(engine.getCellValue('Sheet1', 'E17')).toEqual({ tag: ValueTag.Number, value: 252 })
    engine.undo()
    expect(engine.getCellValue('Sheet1', 'E17')).toEqual({ tag: ValueTag.Number, value: 170 })
    engine.undo()
    expect(engine.getCellValue('Sheet1', 'E17')).toEqual({ tag: ValueTag.Empty })
  })

  it('does not mark direct aggregate inputs covered for formula-only fresh aggregate batches', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'formula-only-fresh-row-aggregate-coverage' })
    await engine.ready()
    engine.createSheet('Sheet1')

    const existingRows = 12
    const appendRows = 8
    const inputCols = 4
    for (let row = 1; row <= existingRows; row += 1) {
      for (let col = 0; col < inputCols; col += 1) {
        engine.setCellValue('Sheet1', `${String.fromCharCode(65 + col)}${row}`, row * (col + 1))
      }
      engine.setCellFormula('Sheet1', `E${row}`, `SUM(A${row}:D${row})`)
    }
    engine.insertRows('Sheet1', existingRows, appendRows)

    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const valueRefs: EngineCellMutationRef[] = []
    const formulaRefs: EngineCellMutationRef[] = []
    for (let row = 0; row < appendRows; row += 1) {
      const rowIndex = existingRows + row
      const rowNumber = rowIndex + 1
      for (let col = 0; col < inputCols; col += 1) {
        valueRefs.push({
          sheetId,
          mutation: { kind: 'setCellValue', row: rowIndex, col, value: rowNumber * (col + 1) },
        })
      }
      formulaRefs.push({
        sheetId,
        mutation: { kind: 'setCellFormula', row: rowIndex, col: inputCols, formula: `SUM(A${rowNumber}:D${rowNumber})` },
      })
    }

    engine.applyCellMutationsAt(valueRefs, valueRefs.length)
    engine.resetPerformanceCounters()
    const markCovered = vi.spyOn(DirectFormulaIndexCollection.prototype, 'markDirectRangeInputCovered')
    try {
      engine.applyCellMutationsAt(formulaRefs, formulaRefs.length)
      expect(markCovered).not.toHaveBeenCalled()
    } finally {
      markCovered.mockRestore()
    }
    expect(engine.getCellValue('Sheet1', 'E13')).toEqual({ tag: ValueTag.Number, value: 130 })
    expect(engine.getCellValue('Sheet1', 'E20')).toEqual({ tag: ValueTag.Number, value: 200 })
    expect(engine.getLastMetrics()).toMatchObject({
      dirtyFormulaCount: 0,
      jsFormulaCount: 0,
      wasmFormulaCount: 0,
    })
    expect(engine.getPerformanceCounters().directAggregateScanEvaluations).toBe(0)
  })

  it('keeps combined appended row aggregates out of the dirty recalc path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'combined-fresh-row-aggregate-formula-results' })
    await engine.ready()
    engine.createSheet('Sheet1')

    const existingRows = 12
    const appendRows = 8
    const inputCols = 4
    for (let row = 1; row <= existingRows; row += 1) {
      for (let col = 0; col < inputCols; col += 1) {
        engine.setCellValue('Sheet1', `${String.fromCharCode(65 + col)}${row}`, row * (col + 1))
      }
      engine.setCellFormula('Sheet1', `E${row}`, `SUM(A${row}:D${row})`)
    }
    engine.insertRows('Sheet1', existingRows, appendRows)

    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const refs: EngineCellMutationRef[] = []
    for (let row = 0; row < appendRows; row += 1) {
      const rowIndex = existingRows + row
      const rowNumber = rowIndex + 1
      for (let col = 0; col < inputCols; col += 1) {
        refs.push({
          sheetId,
          mutation: { kind: 'setCellValue', row: rowIndex, col, value: rowNumber * (col + 1) },
        })
      }
      refs.push({
        sheetId,
        mutation: { kind: 'setCellFormula', row: rowIndex, col: inputCols, formula: `SUM(A${rowNumber}:D${rowNumber})` },
      })
    }

    engine.resetPerformanceCounters()
    engine.applyCellMutationsAt(refs, refs.length)

    expect(engine.getCellValue('Sheet1', 'E13')).toEqual({ tag: ValueTag.Number, value: 130 })
    expect(engine.getCellValue('Sheet1', 'E20')).toEqual({ tag: ValueTag.Number, value: 200 })
    expect(engine.getLastMetrics()).toMatchObject({
      changedInputCount: appendRows * (inputCols + 1),
      dirtyFormulaCount: 0,
      jsFormulaCount: 0,
      wasmFormulaCount: 0,
    })
    expect(engine.getPerformanceCounters()).toMatchObject({
      calcChainFullScans: 0,
      directAggregateScanCells: 0,
      directAggregateScanEvaluations: 0,
      directFormulaKernelSyncOnlyRecalcSkips: 1,
      formulasBound: 0,
      topoRepairs: 0,
    })

    engine.setCellValue('Sheet1', 'A13', 99)
    expect(engine.getCellValue('Sheet1', 'E13')).toEqual({ tag: ValueTag.Number, value: 216 })
  })

  it('applies phased fresh numeric and row aggregate formula matrices in one dense pass', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'phased-fresh-row-aggregate-matrix-fast-path' })
    await engine.ready()
    engine.createSheet('Sheet1')

    const existingRows = 16
    const appendRows = 80
    const inputCols = 4
    for (let row = 1; row <= existingRows; row += 1) {
      for (let col = 0; col < inputCols; col += 1) {
        engine.setCellValue('Sheet1', `${String.fromCharCode(65 + col)}${row}`, row * (col + 1))
      }
      engine.setCellFormula('Sheet1', `E${row}`, `SUM(A${row}:D${row})`)
    }
    engine.insertRows('Sheet1', existingRows, appendRows)

    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const valueRefs: EngineCellMutationRef[] = []
    const formulaRefs: EngineCellMutationRef[] = []
    for (let row = 0; row < appendRows; row += 1) {
      const rowIndex = existingRows + row
      const rowNumber = rowIndex + 1
      for (let col = 0; col < inputCols; col += 1) {
        valueRefs.push({
          sheetId,
          mutation: { kind: 'setCellValue', row: rowIndex, col, value: rowNumber * (col + 1) },
        })
      }
      formulaRefs.push({
        sheetId,
        mutation: { kind: 'setCellFormula', row: rowIndex, col: inputCols, formula: `SUM(A${rowNumber}:D${rowNumber})` },
      })
    }

    const refs = [...valueRefs, ...formulaRefs]
    const formulas = formulaScanTable(engine)
    const formulaScan = vi.spyOn(formulas, 'forEach')
    const ensureCellAt = vi.spyOn(engine.workbook, 'ensureCellAt')
    const attachAllocatedCellWithLogicalAxisIds = vi.spyOn(engine.workbook, 'attachAllocatedCellWithLogicalAxisIds')
    const allocateReserved = vi.spyOn(engine.workbook.cellStore, 'allocateReserved')
    try {
      engine.resetPerformanceCounters()
      const undoOps = engine.applyCellMutationsAt(refs, refs.length)

      expect(undoOps).not.toBeNull()
      expect(ensureCellAt).not.toHaveBeenCalled()
      expect(attachAllocatedCellWithLogicalAxisIds).not.toHaveBeenCalled()
      expect(allocateReserved).not.toHaveBeenCalled()
      expect(engine.getCellValue('Sheet1', 'E17')).toEqual({ tag: ValueTag.Number, value: 170 })
      expect(engine.getCellValue('Sheet1', `E${existingRows + appendRows}`)).toEqual({
        tag: ValueTag.Number,
        value: (existingRows + appendRows) * 10,
      })
      expect(engine.getLastMetrics()).toMatchObject({
        changedInputCount: appendRows * (inputCols + 1),
        dirtyFormulaCount: 0,
        jsFormulaCount: 0,
        wasmFormulaCount: 0,
      })
      expect(engine.getPerformanceCounters()).toMatchObject({
        calcChainFullScans: 0,
        directAggregateScanCells: 0,
        directAggregateScanEvaluations: 0,
        directFormulaKernelSyncOnlyRecalcSkips: 1,
        formulasBound: 0,
        kernelSyncOnlyRecalcSkips: 1,
        nativeDirectAggregatePrefixEvaluations: appendRows,
        topoRepairs: 0,
      })
      expect(formulaScan).not.toHaveBeenCalled()
    } finally {
      formulaScan.mockRestore()
      ensureCellAt.mockRestore()
      attachAllocatedCellWithLogicalAxisIds.mockRestore()
      allocateReserved.mockRestore()
    }

    engine.setCellValue('Sheet1', 'A17', 99)
    expect(engine.getCellValue('Sheet1', 'E17')).toEqual({ tag: ValueTag.Number, value: 252 })
    engine.undo()
    expect(engine.getCellValue('Sheet1', 'E17')).toEqual({ tag: ValueTag.Number, value: 170 })
    engine.undo()
    expect(engine.getCellValue('Sheet1', 'E17')).toEqual({ tag: ValueTag.Empty })
  })

  it('replaces same-dependency direct scalar formulas without graph rebuilds', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'same-dependency-formula-replacement-fast-path' })
    await engine.ready()
    engine.createSheet('Sheet1')

    const downstreamCount = 96
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'B1', 2)
    engine.setCellFormula('Sheet1', 'C1', 'A1+B1')
    for (let col = 3; col <= downstreamCount + 2; col += 1) {
      engine.setCellFormula('Sheet1', `${columnLabel(col)}1`, `${columnLabel(col - 1)}1+1`)
    }

    engine.resetPerformanceCounters()
    engine.setCellFormula('Sheet1', 'C1', 'A1*B1')

    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', `${columnLabel(downstreamCount + 2)}1`)).toEqual({
      tag: ValueTag.Number,
      value: downstreamCount + 2,
    })
    expect(engine.getLastMetrics()).toMatchObject({
      dirtyFormulaCount: 0,
      jsFormulaCount: 0,
      wasmFormulaCount: 0,
    })
    expect(engine.getPerformanceCounters().directScalarDeltaApplications).toBe(downstreamCount)
    expect(engine.getPerformanceCounters().directScalarDeltaOnlyRecalcSkips).toBe(1)
    expect(engine.getPerformanceCounters().topoRebuilds).toBe(0)
    expect(engine.getPerformanceCounters().topoRepairs).toBe(0)
    expect(engine.getPerformanceCounters().regionQueryIndexBuilds).toBe(0)
  })

  it('handles dense lookup-only numeric column coordinate batches', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'dense-lookup-fast-path', useColumnIndex: true })
    await engine.ready()
    engine.createSheet('Sheet1')

    for (let row = 1; row <= 40; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
    }
    engine.setCellFormula('Sheet1', 'C1', 'MATCH(120,A1:A40,0)')
    engine.setCellFormula('Sheet1', 'C2', 'MATCH(120.5,A1:A40,1)')
    const general = vi.fn()
    const tracked = vi.fn()
    const watched = vi.fn()
    const unsubscribeGeneral = engine.subscribe(general)
    const unsubscribeTracked = engine.events.subscribeTracked(tracked)
    const unsubscribeWatched = engine.subscribeCell('Sheet1', 'C1', watched)

    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const refs: EngineCellMutationRef[] = []
    for (let row = 0; row < 40; row += 1) {
      refs.push({
        sheetId,
        cellIndex: engine.workbook.getCellIndex('Sheet1', `A${row + 1}`),
        mutation: { kind: 'setCellValue', row, col: 0, value: row + 101 },
      })
    }

    engine.resetPerformanceCounters()
    const undoOps = engine.applyCellMutationsAt(refs, 0)

    expect(undoOps).not.toBeNull()
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 20 })
    expect(engine.getCellValue('Sheet1', 'C2')).toEqual({ tag: ValueTag.Number, value: 20 })
    expect(general).toHaveBeenCalledWith(expect.objectContaining({ kind: 'batch', explicitChangedCount: 40 }))
    expect(tracked).toHaveBeenCalledWith(expect.objectContaining({ kind: 'batch', explicitChangedCount: 40 }))
    expect(watched).toHaveBeenCalled()
    unsubscribeWatched()
    unsubscribeTracked()
    unsubscribeGeneral()
  })

  it('skips exact lookup-only numeric batches with one cached numeric dependent plan', async () => {
    const rowCount = 80
    const engine = new SpreadsheetEngine({ workbookName: 'dense-exact-lookup-batch-plan', useColumnIndex: true })
    await engine.ready()
    engine.createSheet('Sheet1')

    for (let row = 1; row <= rowCount; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
    }
    engine.setCellValue('Sheet1', 'D1', Math.floor(rowCount / 2))
    engine.setCellFormula('Sheet1', 'E1', `MATCH(D1,A1:A${rowCount},0)`)

    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const refs: EngineCellMutationRef[] = []
    for (let row = rowCount - 1; row >= rowCount - 40; row -= 1) {
      refs.push({
        sheetId,
        cellIndex: engine.workbook.getCellIndex('Sheet1', `A${row + 1}`),
        mutation: { kind: 'setCellValue', row, col: 0, value: row + 1_000 },
      })
    }

    const sheet = engine.workbook.getSheet('Sheet1')!
    const columnVersionBefore = sheet.columnVersions[0] ?? 0
    engine.resetPerformanceCounters()
    const applied = operationHook<[readonly EngineCellMutationRef[], null, number], boolean>(
      engine,
      'tryApplyLookupOnlyNumericColumnLiteralBatch',
    )(refs, null, 0)

    expect(applied).toBe(true)
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: Math.floor(rowCount / 2) })
    expect(engine.getLastMetrics()).toMatchObject({ dirtyFormulaCount: 0, jsFormulaCount: 0, wasmFormulaCount: 0 })
    expect(engine.getPerformanceCounters().kernelSyncOnlyRecalcSkips).toBe(1)
    expect(sheet.columnVersions[0]).toBe(columnVersionBefore + 1)
  })

  it('skips exact lookup-only numeric batches with multiple cached numeric dependent plans', async () => {
    const rowCount = 96
    const engine = new SpreadsheetEngine({ workbookName: 'dense-exact-lookup-batch-multi-plan', useColumnIndex: true })
    await engine.ready()
    engine.createSheet('Sheet1')

    for (let row = 1; row <= rowCount; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
    }
    engine.setCellValue('Sheet1', 'D1', Math.floor(rowCount / 4))
    engine.setCellValue('Sheet1', 'D2', Math.floor(rowCount / 3))
    engine.setCellFormula('Sheet1', 'E1', `MATCH(D1,A1:A${rowCount},0)`)
    engine.setCellFormula('Sheet1', 'E2', `MATCH(D2,A1:A${rowCount},0)`)

    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const refs: EngineCellMutationRef[] = []
    for (let row = rowCount - 1; row >= rowCount - 40; row -= 1) {
      refs.push({
        sheetId,
        cellIndex: engine.workbook.getCellIndex('Sheet1', `A${row + 1}`),
        mutation: { kind: 'setCellValue', row, col: 0, value: row + 1_000 },
      })
    }

    engine.resetPerformanceCounters()
    const applied = operationHook<[readonly EngineCellMutationRef[], null, number], boolean>(
      engine,
      'tryApplyLookupOnlyNumericColumnLiteralBatch',
    )(refs, null, 0)

    expect(applied).toBe(true)
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: Math.floor(rowCount / 4) })
    expect(engine.getCellValue('Sheet1', 'E2')).toEqual({ tag: ValueTag.Number, value: Math.floor(rowCount / 3) })
    expect(engine.getLastMetrics()).toMatchObject({ dirtyFormulaCount: 0, jsFormulaCount: 0, wasmFormulaCount: 0 })
    expect(engine.getPerformanceCounters().kernelSyncOnlyRecalcSkips).toBe(1)
  })

  it('rejects exact lookup-only numeric batches when a write can change the match', async () => {
    const rowCount = 80
    const engine = new SpreadsheetEngine({ workbookName: 'dense-exact-lookup-batch-plan-hit', useColumnIndex: true })
    await engine.ready()
    engine.createSheet('Sheet1')

    for (let row = 1; row <= rowCount; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
    }
    engine.setCellValue('Sheet1', 'D1', Math.floor(rowCount / 2))
    engine.setCellFormula('Sheet1', 'E1', `MATCH(D1,A1:A${rowCount},0)`)

    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const refs: EngineCellMutationRef[] = []
    for (let row = rowCount - 1; row >= rowCount - 41; row -= 1) {
      refs.push({
        sheetId,
        cellIndex: engine.workbook.getCellIndex('Sheet1', `A${row + 1}`),
        mutation: { kind: 'setCellValue', row, col: 0, value: row + 1_000 },
      })
    }

    const applied = operationHook<[readonly EngineCellMutationRef[], null, number], boolean>(
      engine,
      'tryApplyLookupOnlyNumericColumnLiteralBatch',
    )(refs, null, 0)

    expect(applied).toBe(false)
    expect(engine.getCellValue('Sheet1', `A${Math.floor(rowCount / 2)}`)).toEqual({ tag: ValueTag.Number, value: Math.floor(rowCount / 2) })
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: Math.floor(rowCount / 2) })
  })

  it('rejects exact lookup-only numeric batches when any cached dependent can change', async () => {
    const rowCount = 96
    const engine = new SpreadsheetEngine({ workbookName: 'dense-exact-lookup-batch-multi-plan-hit', useColumnIndex: true })
    await engine.ready()
    engine.createSheet('Sheet1')

    for (let row = 1; row <= rowCount; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
    }
    engine.setCellValue('Sheet1', 'D1', Math.floor(rowCount / 4))
    engine.setCellValue('Sheet1', 'D2', rowCount - 10)
    engine.setCellFormula('Sheet1', 'E1', `MATCH(D1,A1:A${rowCount},0)`)
    engine.setCellFormula('Sheet1', 'E2', `MATCH(D2,A1:A${rowCount},0)`)

    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const refs: EngineCellMutationRef[] = []
    for (let row = rowCount - 1; row >= rowCount - 40; row -= 1) {
      refs.push({
        sheetId,
        cellIndex: engine.workbook.getCellIndex('Sheet1', `A${row + 1}`),
        mutation: { kind: 'setCellValue', row, col: 0, value: row + 1_000 },
      })
    }

    const applied = operationHook<[readonly EngineCellMutationRef[], null, number], boolean>(
      engine,
      'tryApplyLookupOnlyNumericColumnLiteralBatch',
    )(refs, null, 0)

    expect(applied).toBe(false)
    expect(engine.getCellValue('Sheet1', `A${rowCount - 10}`)).toEqual({ tag: ValueTag.Number, value: rowCount - 10 })
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: Math.floor(rowCount / 4) })
    expect(engine.getCellValue('Sheet1', 'E2')).toEqual({ tag: ValueTag.Number, value: rowCount - 10 })
  })

  it('applies trusted existing numeric aggregate mutations through the narrow fast path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'trusted-existing-aggregate-fast-path' })
    await engine.ready()
    engine.createSheet('Sheet1')
    for (let row = 1; row <= 10; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
    }
    engine.setCellFormula('Sheet1', 'C1', 'SUM(A1:A10)')

    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const cellIndex = engine.workbook.getCellIndex('Sheet1', 'A1')!
    const result = existingNumericMutationFastPath(engine)({
      sheetId,
      row: 0,
      col: 0,
      cellIndex,
      value: 100,
      trustedExistingNumericLiteral: true,
      oldNumericValue: 1,
      emitTracked: false,
    })

    expect(result).not.toBeNull()
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 100 })
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 154 })
    expect(engine.getPerformanceCounters().directAggregateDeltaOnlyRecalcSkips).toBeGreaterThan(0)
  })

  it('applies trusted existing numeric scalar-closure mutations through the narrow fast path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'trusted-existing-scalar-fast-path' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellFormula('Sheet1', 'B1', 'A1*2')
    engine.setCellFormula('Sheet1', 'C1', 'B1+5')

    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const cellIndex = engine.workbook.getCellIndex('Sheet1', 'A1')!
    const result = existingNumericMutationFastPath(engine)({
      sheetId,
      row: 0,
      col: 0,
      cellIndex,
      value: 7,
      trustedExistingNumericLiteral: true,
      oldNumericValue: 2,
      emitTracked: false,
    })

    expect(result).not.toBeNull()
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 14 })
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 19 })
    expect(engine.getPerformanceCounters().directScalarDeltaOnlyRecalcSkips).toBeGreaterThan(0)
  })

  it('applies trusted existing numeric lookup operand mutations through the narrow fast path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'trusted-existing-lookup-fast-path', useColumnIndex: true })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 3)
    for (let row = 1; row <= 5; row += 1) {
      engine.setCellValue('Sheet1', `B${row}`, row)
    }
    engine.setCellFormula('Sheet1', 'D1', 'MATCH(A1,B1:B5,0)')

    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const cellIndex = engine.workbook.getCellIndex('Sheet1', 'A1')!
    const result = existingNumericMutationFastPath(engine)({
      sheetId,
      row: 0,
      col: 0,
      cellIndex,
      value: 5,
      trustedExistingNumericLiteral: true,
      oldNumericValue: 3,
      emitTracked: false,
    })

    expect(result).not.toBeNull()
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 5 })
  })

  it('applies non-uniform numeric exact lookup operand mutations without dirty traversal', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'trusted-existing-nonuniform-lookup-fast-path', useColumnIndex: true })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 3)
    ;[2, 3, 7, 11, 18].forEach((value, index) => {
      engine.setCellValue('Sheet1', `B${index + 1}`, value)
    })
    engine.setCellFormula('Sheet1', 'D1', 'MATCH(A1,B1:B5,0)')

    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const inputIndex = engine.workbook.getCellIndex('Sheet1', 'A1')!
    const formulaIndex = engine.workbook.getCellIndex('Sheet1', 'D1')!
    const formula = runtimeFormula(engine, formulaIndex)
    if (typeof formula !== 'object' || formula === null) {
      throw new TypeError('Expected runtime formula')
    }
    expect(Reflect.get(formula, 'directLookup')).toMatchObject({ kind: 'exact' })

    engine.resetPerformanceCounters()
    const result = existingNumericMutationFastPath(engine)({
      sheetId,
      row: 0,
      col: 0,
      cellIndex: inputIndex,
      value: 11,
      trustedExistingNumericLiteral: true,
      oldNumericValue: 3,
      emitTracked: false,
    })

    expect(result).toMatchObject({
      changedCellCount: 2,
      firstChangedCellIndex: inputIndex,
      secondChangedCellIndex: formulaIndex,
      secondChangedNumericValue: 4,
    })
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 11 })
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 4 })
    expect(engine.getLastMetrics()).toMatchObject({
      dirtyFormulaCount: 0,
      jsFormulaCount: 0,
      wasmFormulaCount: 0,
    })
    expect(engine.getPerformanceCounters()).toMatchObject({
      changedCellPayloadsBuilt: 0,
      directFormulaKernelSyncOnlyRecalcSkips: 1,
      formulasBound: 0,
      topoRepairs: 0,
    })
  })

  it('applies direct formula literal mutations without event plumbing', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'single-direct-formula-no-events' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellFormula('Sheet1', 'B1', 'A1*3')

    const inputIndex = engine.workbook.getCellIndex('Sheet1', 'A1')!
    const formulaIndex = engine.workbook.getCellIndex('Sheet1', 'B1')!
    const applied = operationHook<
      [
        {
          existingIndex: number
          formulaCellIndex: number
          value: number
          oldNumber: number
          newNumber: number
          exactLookupValue: number | undefined
          approximateLookupValue: number | undefined
        },
      ],
      boolean
    >(
      engine,
      'tryApplySingleDirectFormulaLiteralMutationWithoutEvents',
    )({
      existingIndex: inputIndex,
      formulaCellIndex: formulaIndex,
      value: 5,
      oldNumber: 2,
      newNumber: 5,
      exactLookupValue: undefined,
      approximateLookupValue: undefined,
    })

    expect(applied).toBe(true)
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 15 })
    expect(engine.getPerformanceCounters().directScalarDeltaApplications).toBe(1)
  })

  it('applies multi-dependent direct scalar literal mutations without events', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'single-direct-scalar-multiple-no-events' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellFormula('Sheet1', 'B1', 'A1+1')
    engine.setCellFormula('Sheet1', 'C1', 'A1+2')

    const inputIndex = engine.workbook.getCellIndex('Sheet1', 'A1')!
    const applied = operationHook<[{ existingIndex: number; value: number; oldNumber: number; newNumber: number }], boolean>(
      engine,
      'tryApplySingleDirectScalarLiteralMutationWithoutEvents',
    )({
      existingIndex: inputIndex,
      value: 5,
      oldNumber: 2,
      newNumber: 5,
    })

    expect(applied).toBe(true)
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 7 })
    expect(engine.getPerformanceCounters().directScalarDeltaApplications).toBe(2)
  })

  it('emits tracked events for kernel-sync-only literal mutation fast paths', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'kernel-sync-literal-hook-tracked' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    const tracked = vi.fn()
    const unsubscribe = engine.events.subscribeTracked(tracked)
    let afterWriteCalled = false

    const inputIndex = engine.workbook.getCellIndex('Sheet1', 'A1')!
    const applied = operationHook<[{ existingIndex: number; value: number; emitTracked: boolean; afterWrite: () => void }], boolean>(
      engine,
      'tryApplySingleKernelSyncOnlyLiteralMutationFastPath',
    )({
      existingIndex: inputIndex,
      value: 9,
      emitTracked: true,
      afterWrite: () => {
        afterWriteCalled = true
      },
    })

    expect(applied).toBe(true)
    expect(afterWriteCalled).toBe(true)
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 9 })
    expect(tracked).toHaveBeenCalledWith(
      expect.objectContaining({
        changedCellIndices: new Uint32Array([inputIndex]),
        explicitChangedCount: 1,
      }),
    )
    unsubscribe()
  })

  it('applies direct lookup operand mutation fast paths with tracked changes', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'direct-lookup-operand-hook-tracked', useColumnIndex: true })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 2)
    for (let row = 1; row <= 5; row += 1) {
      engine.setCellValue('Sheet1', `B${row}`, row)
    }
    engine.setCellFormula('Sheet1', 'D1', 'MATCH(A1,B1:B5,0)')
    const tracked = vi.fn()
    const unsubscribe = engine.events.subscribeTracked(tracked)

    const inputIndex = engine.workbook.getCellIndex('Sheet1', 'A1')!
    const formulaIndex = engine.workbook.getCellIndex('Sheet1', 'D1')!
    const result = operationHook<
      [
        {
          existingIndex: number
          formulaCellIndex: number
          value: number
          exactLookupValue: number | undefined
          approximateLookupValue: number | undefined
          emitTracked: boolean
          lookupSheetHint: unknown
        },
      ],
      unknown
    >(
      engine,
      'tryApplySingleDirectLookupOperandMutationFastPath',
    )({
      existingIndex: inputIndex,
      formulaCellIndex: formulaIndex,
      value: 5,
      exactLookupValue: 5,
      approximateLookupValue: undefined,
      emitTracked: true,
      lookupSheetHint: engine.workbook.getSheet('Sheet1'),
    })

    expect(result).not.toBeNull()
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(tracked).toHaveBeenCalledWith(
      expect.objectContaining({
        changedCellIndices: new Uint32Array([inputIndex, formulaIndex]),
        explicitChangedCount: 1,
      }),
    )
    unsubscribe()
  })

  it('applies direct aggregate literal mutation fast paths with tracked changes', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'direct-aggregate-hook-tracked' })
    await engine.ready()
    engine.createSheet('Sheet1')
    for (let row = 1; row <= 10; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
    }
    engine.setCellFormula('Sheet1', 'B1', 'SUM(A1:A10)')
    const tracked = vi.fn()
    const unsubscribe = engine.events.subscribeTracked(tracked)

    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const inputIndex = engine.workbook.getCellIndex('Sheet1', 'A1')!
    const formulaIndex = engine.workbook.getCellIndex('Sheet1', 'B1')!
    const result = operationHook<
      [
        {
          existingIndex: number
          sheetId: number
          sheetName: string
          row: number
          col: number
          value: number
          delta: number
          emitTracked: boolean
        },
      ],
      unknown
    >(
      engine,
      'tryApplySingleDirectAggregateLiteralMutationFastPath',
    )({
      existingIndex: inputIndex,
      sheetId,
      sheetName: 'Sheet1',
      row: 0,
      col: 0,
      value: 100,
      delta: 99,
      emitTracked: true,
    })

    expect(result).not.toBeNull()
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 100 })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 154 })
    expect(tracked).toHaveBeenCalledWith(
      expect.objectContaining({
        changedCellIndices: new Uint32Array([inputIndex, formulaIndex]),
        explicitChangedCount: 1,
      }),
    )
    unsubscribe()
  })

  it('applies single existing exact lookup column writes through the direct literal gate', async () => {
    const rowCount = 32
    const engine = new SpreadsheetEngine({ workbookName: 'single-existing-exact-lookup-gate', useColumnIndex: true })
    await engine.ready()
    engine.createSheet('Sheet1')
    for (let row = 1; row <= rowCount; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
    }
    engine.setCellValue('Sheet1', 'D1', Math.floor(rowCount / 2))
    engine.setCellFormula('Sheet1', 'E1', `MATCH(D1,A1:A${rowCount},0)`)

    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const inputIndex = engine.workbook.getCellIndex('Sheet1', `A${rowCount}`)!
    const applied = operationHook<[readonly EngineCellMutationRef[], null, 'local' | 'restore' | 'undo' | 'redo'], boolean>(
      engine,
      'tryApplySingleExistingDirectLiteralMutation',
    )(
      [
        {
          sheetId,
          cellIndex: inputIndex,
          mutation: { kind: 'setCellValue', row: rowCount - 1, col: 0, value: rowCount + 100 },
        },
      ],
      null,
      'local',
    )

    expect(applied).toBe(true)
    expect(engine.getCellValue('Sheet1', `A${rowCount}`)).toEqual({ tag: ValueTag.Number, value: rowCount + 100 })
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: Math.floor(rowCount / 2) })
    expect(engine.getPerformanceCounters().kernelSyncOnlyRecalcSkips).toBeGreaterThan(0)
  })

  it('applies single existing approximate lookup column writes through the direct literal gate', async () => {
    const rowCount = 32
    const engine = new SpreadsheetEngine({ workbookName: 'single-existing-approximate-lookup-gate' })
    await engine.ready()
    engine.createSheet('Sheet1')
    for (let row = 1; row <= rowCount; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
    }
    engine.setCellValue('Sheet1', 'D1', Math.floor(rowCount / 2) + 0.5)
    engine.setCellFormula('Sheet1', 'E1', `MATCH(D1,A1:A${rowCount},1)`)

    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const inputIndex = engine.workbook.getCellIndex('Sheet1', `A${rowCount}`)!
    const applied = operationHook<[readonly EngineCellMutationRef[], null, 'local' | 'restore' | 'undo' | 'redo'], boolean>(
      engine,
      'tryApplySingleExistingDirectLiteralMutation',
    )(
      [
        {
          sheetId,
          cellIndex: inputIndex,
          mutation: { kind: 'setCellValue', row: rowCount - 1, col: 0, value: rowCount + 1 },
        },
      ],
      null,
      'local',
    )

    expect(applied).toBe(true)
    expect(engine.getCellValue('Sheet1', `A${rowCount}`)).toEqual({ tag: ValueTag.Number, value: rowCount + 1 })
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: Math.floor(rowCount / 2) })
    expect(engine.getPerformanceCounters().kernelSyncOnlyRecalcSkips).toBeGreaterThan(0)
  })

  it('coalesces direct scalar and aggregate dependents through the single literal gate', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'single-existing-scalar-aggregate-gate' })
    await engine.ready()
    engine.createSheet('Sheet1')
    for (let row = 1; row <= 10; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
    }
    engine.setCellFormula('Sheet1', 'B1', 'A1+1')
    engine.setCellFormula('Sheet1', 'C1', 'SUM(A1:A10)')

    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const inputIndex = engine.workbook.getCellIndex('Sheet1', 'A1')!
    const applied = operationHook<[readonly EngineCellMutationRef[], null, 'local' | 'restore' | 'undo' | 'redo'], boolean>(
      engine,
      'tryApplySingleExistingDirectLiteralMutation',
    )(
      [
        {
          sheetId,
          cellIndex: inputIndex,
          mutation: { kind: 'setCellValue', row: 0, col: 0, value: 100 },
        },
      ],
      null,
      'local',
    )

    expect(applied).toBe(true)
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 101 })
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 154 })
    expect(engine.getPerformanceCounters().directScalarDeltaApplications).toBe(1)
    expect(engine.getPerformanceCounters().directAggregateDeltaApplications).toBe(1)
  })

  it('patches uniform lookup tail writes for single and multiple dependents', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'uniform-lookup-tail-patch-hook', useColumnIndex: true })
    await engine.ready()
    engine.createSheet('Sheet1')
    for (let row = 1; row <= 5; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
    }
    engine.setCellValue('Sheet1', 'D1', 9)
    engine.setCellValue('Sheet1', 'D2', 9)
    engine.setCellFormula('Sheet1', 'E1', 'MATCH(D1,A1:A5,0)')
    engine.setCellFormula('Sheet1', 'F1', 'XMATCH(D2,A1:A5,0)')

    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const patch = operationHook<
      [
        {
          sheetId: number
          col: number
          row: number
          oldNumeric: number
          newNumeric: number
          exact: boolean
          sorted: boolean
        },
      ],
      { exact: boolean; sorted: boolean }
    >(
      engine,
      'patchUniformLookupTailWrites',
    )({
      sheetId,
      col: 0,
      row: 4,
      oldNumeric: 5,
      newNumeric: 9,
      exact: true,
      sorted: false,
    })

    expect(patch).toEqual({ exact: true, sorted: false })
    for (const address of ['E1', 'F1']) {
      const formulaIndex = engine.workbook.getCellIndex('Sheet1', address)!
      const formula = runtimeFormula(engine, formulaIndex)
      const directLookup = typeof formula === 'object' && formula !== null ? Reflect.get(formula, 'directLookup') : undefined
      expect(directLookup).toEqual(
        expect.objectContaining({
          tailPatch: expect.objectContaining({ row: 4, oldNumeric: 5, newNumeric: 9 }),
        }),
      )
    }
  })

  it('applies direct lookup current-result branches when exact matches disappear', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'single-direct-lookup-current-result-no-events', useColumnIndex: true })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 3)
    for (let row = 1; row <= 5; row += 1) {
      engine.setCellValue('Sheet1', `B${row}`, row)
    }
    engine.setCellFormula('Sheet1', 'D1', 'MATCH(A1,B1:B5,0)')

    const inputIndex = engine.workbook.getCellIndex('Sheet1', 'A1')!
    const formulaIndex = engine.workbook.getCellIndex('Sheet1', 'D1')!
    const applied = operationHook<
      [
        {
          existingIndex: number
          formulaCellIndex: number
          value: number
          oldNumber: number
          newNumber: number
          exactLookupValue: number | undefined
          approximateLookupValue: number | undefined
        },
      ],
      boolean
    >(
      engine,
      'tryApplySingleDirectFormulaLiteralMutationWithoutEvents',
    )({
      existingIndex: inputIndex,
      formulaCellIndex: formulaIndex,
      value: 9,
      oldNumber: 3,
      newNumber: 9,
      exactLookupValue: 9,
      approximateLookupValue: undefined,
    })

    expect(applied).toBe(true)
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 9 })
    expect(engine.getCellValue('Sheet1', 'D1')).toMatchObject({ tag: ValueTag.Error })
    expect(engine.getPerformanceCounters().directFormulaKernelSyncOnlyRecalcSkips).toBe(1)
  })

  it('uses aggregate fast path for writes with no applicable range dependent', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'direct-aggregate-hook-no-dependent' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellFormula('Sheet1', 'B1', 'SUM(A1:A3)')

    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const inputIndex = engine.workbook.getCellIndex('Sheet1', 'A1')!
    const result = operationHook<
      [
        {
          existingIndex: number
          sheetId: number
          sheetName: string
          row: number
          col: number
          value: number
          delta: number
          emitTracked: boolean
        },
      ],
      unknown
    >(
      engine,
      'tryApplySingleDirectAggregateLiteralMutationFastPath',
    )({
      existingIndex: inputIndex,
      sheetId,
      sheetName: 'Sheet1',
      row: 20,
      col: 0,
      value: 5,
      delta: 4,
      emitTracked: false,
    })

    expect(result).not.toBeNull()
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getLastMetrics()).toMatchObject({
      changedInputCount: 1,
      dirtyFormulaCount: 0,
    })
  })

  it('applies aggregate delta fast path to a single dependent without tracked events', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'direct-aggregate-hook-single-dependent' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'A2', 2)
    engine.setCellValue('Sheet1', 'A3', 3)
    engine.setCellFormula('Sheet1', 'B1', 'SUM(A1:A3)')

    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const inputIndex = engine.workbook.getCellIndex('Sheet1', 'A2')!
    const result = operationHook<
      [
        {
          existingIndex: number
          sheetId: number
          sheetName: string
          row: number
          col: number
          value: number
          delta: number
          emitTracked: boolean
        },
      ],
      { changedCellIndices: Uint32Array } | null
    >(
      engine,
      'tryApplySingleDirectAggregateLiteralMutationFastPath',
    )({
      existingIndex: inputIndex,
      sheetId,
      sheetName: 'Sheet1',
      row: 1,
      col: 0,
      value: 20,
      delta: 18,
      emitTracked: false,
    })

    expect(result).not.toBeNull()
    expect(engine.getCellValue('Sheet1', 'A2')).toEqual({ tag: ValueTag.Number, value: 20 })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 24 })
    expect(engine.getPerformanceCounters().directAggregateDeltaOnlyRecalcSkips).toBe(1)
  })

  it('emits tracked aggregate delta changes for multiple direct dependents', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'direct-aggregate-hook-multiple-dependent' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'A2', 2)
    engine.setCellValue('Sheet1', 'A3', 3)
    engine.setCellFormula('Sheet1', 'B1', 'SUM(A1:A3)')
    engine.setCellFormula('Sheet1', 'C1', 'SUM(A1:A2)')

    const tracked = vi.fn()
    const unsubscribeTracked = engine.events.subscribeTracked(tracked)
    const inputIndex = engine.workbook.getCellIndex('Sheet1', 'A1')!
    const result = operationHook<
      [
        {
          existingIndex: number
          sheetName: string
          row: number
          col: number
          value: number
          delta: number
          emitTracked: boolean
        },
      ],
      { changedCellIndices: Uint32Array } | null
    >(
      engine,
      'tryApplySingleDirectAggregateLiteralMutationFastPath',
    )({
      existingIndex: inputIndex,
      sheetName: 'Sheet1',
      row: 0,
      col: 0,
      value: 10,
      delta: 9,
      emitTracked: true,
    })

    expect(result).not.toBeNull()
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 15 })
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 12 })
    expect(tracked).toHaveBeenCalledWith(
      expect.objectContaining({
        kind: 'batch',
        explicitChangedCount: 1,
        changedCellIndices: expect.any(Uint32Array),
      }),
    )
    unsubscribeTracked()
  })
})
