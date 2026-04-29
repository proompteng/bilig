import { describe, expect, it, vi } from 'vitest'
import { ValueTag } from '@bilig/protocol'
import { SpreadsheetEngine } from '../engine.js'
import type { EngineCellMutationRef } from '../cell-mutations-at.js'
import type { FormulaFamilyStore } from '../formula/formula-family-store.js'

function readRuntimeTemplateId(engine: SpreadsheetEngine, address: string): number | undefined {
  const cellIndex = engine.workbook.getCellIndex('Sheet1', address)
  if (cellIndex === undefined) {
    return undefined
  }
  const formulas = Reflect.get(engine, 'formulas')
  if (typeof formulas !== 'object' || formulas === null || typeof Reflect.get(formulas, 'get') !== 'function') {
    throw new TypeError('Expected internal formulas store')
  }
  const runtimeFormula = Reflect.get(formulas, 'get').call(formulas, cellIndex)
  const templateId = typeof runtimeFormula === 'object' && runtimeFormula !== null ? Reflect.get(runtimeFormula, 'templateId') : undefined
  return typeof templateId === 'number' ? templateId : undefined
}

function isFormulaFamilyStore(value: unknown): value is FormulaFamilyStore {
  return (
    typeof value === 'object' &&
    value !== null &&
    typeof Reflect.get(value, 'getStats') === 'function' &&
    typeof Reflect.get(value, 'listFamilies') === 'function'
  )
}

function getFormulaFamilyStore(engine: SpreadsheetEngine): FormulaFamilyStore {
  const runtime = Reflect.get(engine, 'runtime')
  if (typeof runtime !== 'object' || runtime === null) {
    throw new TypeError('Expected engine runtime')
  }
  const formulaFamilies = Reflect.get(runtime, 'formulaFamilies')
  if (!isFormulaFamilyStore(formulaFamilies)) {
    throw new TypeError('Expected formula family store')
  }
  return formulaFamilies
}

describe('SpreadsheetEngine formula initialization', () => {
  it('initializes formula refs without emitting watched events or batches', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'engine-formula-initialize' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 4)
    const sheetId = engine.workbook.getSheet('Sheet1')!.id

    const events: string[] = []
    const batches: unknown[] = []
    const unsubscribeEvents = engine.subscribe((event) => {
      events.push(event.kind)
    })
    const unsubscribeBatches = engine.subscribeBatches((batch) => {
      batches.push(batch)
    })
    events.length = 0
    batches.length = 0

    engine.initializeCellFormulasAt(
      [
        { sheetId, mutation: { kind: 'setCellFormula', row: 0, col: 1, formula: 'A1*2' } },
        { sheetId, mutation: { kind: 'setCellFormula', row: 0, col: 2, formula: 'B1+1' } },
      ],
      2,
    )

    unsubscribeEvents()
    unsubscribeBatches()

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 8 })
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 9 })
    expect(events).toEqual([])
    expect(batches).toEqual([])
    expect(readRuntimeTemplateId(engine, 'B1')).toBeDefined()
    expect(readRuntimeTemplateId(engine, 'C1')).toBeDefined()
  })

  it('initializes invalid formulas and propagates their errors through dependent formulas', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'engine-formula-initialize-errors' })
    await engine.ready()
    engine.createSheet('Sheet1')
    const sheetId = engine.workbook.getSheet('Sheet1')!.id

    engine.initializeCellFormulasAt(
      [
        { sheetId, mutation: { kind: 'setCellFormula', row: 0, col: 1, formula: 'SUM(' } },
        { sheetId, mutation: { kind: 'setCellFormula', row: 0, col: 2, formula: 'B1+1' } },
      ],
      2,
    )

    expect(engine.getCellValue('Sheet1', 'B1')).toMatchObject({
      tag: ValueTag.Error,
      code: expect.any(Number),
    })
    expect(engine.getCellValue('Sheet1', 'C1')).toMatchObject({
      tag: ValueTag.Error,
      code: expect.any(Number),
    })
  })

  it('shares template ownership across repeated initialized row families', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'engine-formula-initialize-families' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 4)
    engine.setCellValue('Sheet1', 'A2', 5)
    const sheetId = engine.workbook.getSheet('Sheet1')!.id

    engine.initializeCellFormulasAt(
      [
        { sheetId, mutation: { kind: 'setCellFormula', row: 0, col: 1, formula: 'A1*2' } },
        { sheetId, mutation: { kind: 'setCellFormula', row: 1, col: 1, formula: 'A2*2' } },
      ],
      2,
    )

    expect(readRuntimeTemplateId(engine, 'B1')).toBe(readRuntimeTemplateId(engine, 'B2'))
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 8 })
    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.Number, value: 10 })
  })

  it('uses reserved formula cell refs during initialization without re-ensuring target coordinates', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'engine-formula-initialize-reserved-refs' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'B1', 10)
    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const c1 = engine.workbook.ensureCellAt(sheetId, 0, 2).cellIndex
    const d1 = engine.workbook.ensureCellAt(sheetId, 0, 3).cellIndex
    const ensureCellAt = vi.spyOn(engine.workbook, 'ensureCellAt')
    try {
      engine.initializeCellFormulasAt(
        [
          { sheetId, cellIndex: c1, mutation: { kind: 'setCellFormula', row: 0, col: 2, formula: 'A1+B1' } },
          { sheetId, cellIndex: d1, mutation: { kind: 'setCellFormula', row: 0, col: 3, formula: 'C1*2' } },
        ],
        2,
      )

      expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 22 })
      expect(ensureCellAt.mock.calls).not.toContainEqual([sheetId, 0, 3])
    } finally {
      ensureCellAt.mockRestore()
    }
  })

  it('resolves initial direct-scalar dependencies from parsed coordinates', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'engine-formula-initialize-direct-scalar-coords' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'B1', 10)
    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const c1 = engine.workbook.ensureCellAt(sheetId, 0, 2).cellIndex
    const ensureCell = vi.spyOn(engine.workbook, 'ensureCell')
    try {
      engine.initializeCellFormulasAt(
        [{ sheetId, cellIndex: c1, mutation: { kind: 'setCellFormula', row: 0, col: 2, formula: 'A1+B1' } }],
        1,
      )

      expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 11 })
      expect(ensureCell).not.toHaveBeenCalled()
    } finally {
      ensureCell.mockRestore()
    }
  })

  it('skips the full topo rebuild for topologically ordered initial formula batches', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'engine-formula-initialize-topo-fast-path' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 4)
    engine.setCellValue('Sheet1', 'A2', 5)
    const sheetId = engine.workbook.getSheet('Sheet1')!.id

    engine.resetPerformanceCounters()
    engine.initializeCellFormulasAt(
      [
        { sheetId, mutation: { kind: 'setCellFormula', row: 0, col: 1, formula: 'A1*2' } },
        { sheetId, mutation: { kind: 'setCellFormula', row: 0, col: 2, formula: 'B1+1' } },
        { sheetId, mutation: { kind: 'setCellFormula', row: 1, col: 1, formula: 'A2*2' } },
        { sheetId, mutation: { kind: 'setCellFormula', row: 1, col: 2, formula: 'B2+1' } },
      ],
      4,
    )

    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 9 })
    expect(engine.getCellValue('Sheet1', 'C2')).toEqual({ tag: ValueTag.Number, value: 11 })
    expect(engine.getPerformanceCounters().topoRebuilds).toBe(0)
    expect(engine.getPerformanceCounters().cycleFormulaScans).toBe(0)
    expect(engine.getLastMetrics()).toMatchObject({
      dirtyFormulaCount: 0,
      wasmFormulaCount: 0,
      jsFormulaCount: 0,
      rangeNodeVisits: 0,
    })
    expect(engine.getPerformanceCounters().wasmFullUploads).toBe(0)
    expect(engine.getPerformanceCounters().directFormulaInitialEvaluations).toBe(0)
  })

  it('materializes anchored prefix aggregate families during initial direct evaluation', async () => {
    const rowCount = 128
    const engine = new SpreadsheetEngine({ workbookName: 'engine-formula-initialize-prefix-aggregate-family' })
    await engine.ready()
    engine.createSheet('Sheet1')
    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    for (let row = 1; row <= rowCount; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
    }

    engine.resetPerformanceCounters()
    engine.initializeCellFormulasAt(
      Array.from({ length: rowCount }, (_entry, row) => ({
        sheetId,
        mutation: {
          kind: 'setCellFormula' as const,
          row,
          col: 1,
          formula: `SUM(A1:A${row + 1})`,
        },
      })),
      rowCount,
    )

    expect(engine.getCellValue('Sheet1', `B${rowCount}`)).toEqual({
      tag: ValueTag.Number,
      value: (rowCount * (rowCount + 1)) / 2,
    })
    expect(engine.getLastMetrics()).toMatchObject({ dirtyFormulaCount: 0, wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(engine.getPerformanceCounters().directFormulaInitialEvaluations).toBe(rowCount)
    expect(engine.getPerformanceCounters().directAggregatePrefixEvaluations).toBe(0)
    expect(engine.getPerformanceCounters().directAggregateScanEvaluations).toBe(0)
    expect(engine.getPerformanceCounters().regionQueryIndexBuilds).toBe(0)
  })

  it('bulk-materializes parser-cache template sheets into column family runs', async () => {
    const rowCount = 12
    const engine = new SpreadsheetEngine({ workbookName: 'engine-formula-initialize-parser-template-families' })
    await engine.ready()
    engine.createSheet('Sheet1')
    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const refs: EngineCellMutationRef[] = []
    for (let row = 0; row < rowCount; row += 1) {
      const rowNumber = row + 1
      engine.setCellValue('Sheet1', `A${rowNumber}`, rowNumber)
      engine.setCellValue('Sheet1', `B${rowNumber}`, rowNumber * 2)
      refs.push({ sheetId, mutation: { kind: 'setCellFormula' as const, row, col: 2, formula: `A${rowNumber}+B${rowNumber}` } })
      refs.push({ sheetId, mutation: { kind: 'setCellFormula' as const, row, col: 3, formula: `C${rowNumber}*2` } })
      refs.push({ sheetId, mutation: { kind: 'setCellFormula' as const, row, col: 4, formula: `SUM(A1:A${rowNumber})` } })
      refs.push({ sheetId, mutation: { kind: 'setCellFormula' as const, row, col: 5, formula: `D${rowNumber}+E${rowNumber}` } })
    }

    engine.initializeCellFormulasAt(refs, refs.length)

    const store = getFormulaFamilyStore(engine)
    expect(store.getStats()).toEqual({ familyCount: 3, runCount: 4, memberCount: rowCount * 4 })
    expect(store.listFamilies().flatMap((family) => family.runs.map((run) => run.cellIndices.length))).toEqual([
      rowCount,
      rowCount,
      rowCount,
      rowCount,
    ])
  })

  it('bulk-notifies initialized direct formula value columns instead of per-cell value writes', async () => {
    const rowCount = 8
    const engine = new SpreadsheetEngine({ workbookName: 'engine-formula-initialize-bulk-value-columns' })
    await engine.ready()
    engine.createSheet('Sheet1')
    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const refs: EngineCellMutationRef[] = []
    for (let row = 0; row < rowCount; row += 1) {
      const rowNumber = row + 1
      engine.setCellValue('Sheet1', `A${rowNumber}`, rowNumber)
      engine.setCellValue('Sheet1', `B${rowNumber}`, rowNumber * 10)
      refs.push({ sheetId, mutation: { kind: 'setCellFormula' as const, row, col: 2, formula: `A${rowNumber}+B${rowNumber}` } })
      refs.push({ sheetId, mutation: { kind: 'setCellFormula' as const, row, col: 3, formula: `C${rowNumber}*2` } })
    }
    const notifyCellValueWritten = vi.spyOn(engine.workbook, 'notifyCellValueWritten')
    const notifyColumnsWritten = vi.spyOn(engine.workbook, 'notifyColumnsWritten')

    engine.initializeCellFormulasAt(refs, refs.length)

    expect(engine.getCellValue('Sheet1', `D${rowCount}`)).toEqual({
      tag: ValueTag.Number,
      value: rowCount * 11 * 2,
    })
    expect(notifyCellValueWritten).not.toHaveBeenCalled()
    expect(notifyColumnsWritten).toHaveBeenCalledWith(sheetId, Uint32Array.from([2, 3]))
    expect(engine.getPerformanceCounters().directFormulaInitialEvaluations).toBe(0)
    notifyCellValueWritten.mockRestore()
    notifyColumnsWritten.mockRestore()
  })

  it('bulk-materializes mixed parser-cache template sheets into strided family runs', async () => {
    const rowCount = 12
    const engine = new SpreadsheetEngine({ workbookName: 'engine-formula-initialize-mixed-template-families' })
    await engine.ready()
    engine.createSheet('Sheet1')
    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const refs: EngineCellMutationRef[] = []
    for (let row = 0; row < rowCount; row += 1) {
      const rowNumber = row + 1
      const valueMultiplier = rowNumber % 3 === 1 ? 2 : rowNumber % 3 === 2 ? 3 : 4
      engine.setCellValue('Sheet1', `A${rowNumber}`, rowNumber)
      engine.setCellValue('Sheet1', `B${rowNumber}`, rowNumber * valueMultiplier)
      if (rowNumber % 3 === 1) {
        refs.push({ sheetId, mutation: { kind: 'setCellFormula' as const, row, col: 2, formula: `A${rowNumber}+B${rowNumber}` } })
        refs.push({ sheetId, mutation: { kind: 'setCellFormula' as const, row, col: 3, formula: `C${rowNumber}*2` } })
        refs.push({ sheetId, mutation: { kind: 'setCellFormula' as const, row, col: 4, formula: `SUM(A1:A${rowNumber})` } })
        refs.push({ sheetId, mutation: { kind: 'setCellFormula' as const, row, col: 5, formula: `D${rowNumber}+E${rowNumber}` } })
      } else if (rowNumber % 3 === 2) {
        refs.push({ sheetId, mutation: { kind: 'setCellFormula' as const, row, col: 2, formula: `A${rowNumber}*B${rowNumber}` } })
        refs.push({ sheetId, mutation: { kind: 'setCellFormula' as const, row, col: 3, formula: `C${rowNumber}+A${rowNumber}` } })
        refs.push({ sheetId, mutation: { kind: 'setCellFormula' as const, row, col: 4, formula: `AVERAGE(A1:A${rowNumber})` } })
        refs.push({ sheetId, mutation: { kind: 'setCellFormula' as const, row, col: 5, formula: `D${rowNumber}-E${rowNumber}` } })
      } else {
        refs.push({ sheetId, mutation: { kind: 'setCellFormula' as const, row, col: 2, formula: `B${rowNumber}-A${rowNumber}` } })
        refs.push({ sheetId, mutation: { kind: 'setCellFormula' as const, row, col: 3, formula: `ABS(C${rowNumber})` } })
        refs.push({ sheetId, mutation: { kind: 'setCellFormula' as const, row, col: 4, formula: `MAX(A1:A${rowNumber})` } })
        refs.push({ sheetId, mutation: { kind: 'setCellFormula' as const, row, col: 5, formula: `D${rowNumber}+E${rowNumber}` } })
      }
    }

    engine.initializeCellFormulasAt(refs, refs.length)

    const store = getFormulaFamilyStore(engine)
    const stridedOutputRuns = store
      .listFamilies()
      .flatMap((family) => family.runs)
      .filter((run) => run.fixedIndex === 5 && run.step === 3 && run.cellIndices.length === 4)
      .map((run) => [run.start, run.end])
      .toSorted((left, right) => left[0]! - right[0]!)
    expect(store.getStats()).toEqual({ familyCount: 10, runCount: 12, memberCount: rowCount * 4 })
    expect(stridedOutputRuns).toEqual([
      [0, 9],
      [1, 10],
      [2, 11],
    ])
  })

  it('bulk-registers hydrated runtime snapshot formulas into strided family runs', async () => {
    const rowCount = 12
    const source = new SpreadsheetEngine({ workbookName: 'engine-formula-initialize-hydrated-families-source' })
    await source.ready()
    source.createSheet('Sheet1')
    const sheetId = source.workbook.getSheet('Sheet1')!.id
    const refs: EngineCellMutationRef[] = []
    for (let row = 0; row < rowCount; row += 1) {
      const rowNumber = row + 1
      const valueMultiplier = rowNumber % 3 === 1 ? 2 : rowNumber % 3 === 2 ? 3 : 4
      source.setCellValue('Sheet1', `A${rowNumber}`, rowNumber)
      source.setCellValue('Sheet1', `B${rowNumber}`, rowNumber * valueMultiplier)
      if (rowNumber % 3 === 1) {
        refs.push({ sheetId, mutation: { kind: 'setCellFormula' as const, row, col: 2, formula: `A${rowNumber}+B${rowNumber}` } })
        refs.push({ sheetId, mutation: { kind: 'setCellFormula' as const, row, col: 3, formula: `C${rowNumber}*2` } })
        refs.push({ sheetId, mutation: { kind: 'setCellFormula' as const, row, col: 4, formula: `SUM(A1:A${rowNumber})` } })
        refs.push({ sheetId, mutation: { kind: 'setCellFormula' as const, row, col: 5, formula: `D${rowNumber}+E${rowNumber}` } })
      } else if (rowNumber % 3 === 2) {
        refs.push({ sheetId, mutation: { kind: 'setCellFormula' as const, row, col: 2, formula: `A${rowNumber}*B${rowNumber}` } })
        refs.push({ sheetId, mutation: { kind: 'setCellFormula' as const, row, col: 3, formula: `C${rowNumber}+A${rowNumber}` } })
        refs.push({ sheetId, mutation: { kind: 'setCellFormula' as const, row, col: 4, formula: `AVERAGE(A1:A${rowNumber})` } })
        refs.push({ sheetId, mutation: { kind: 'setCellFormula' as const, row, col: 5, formula: `D${rowNumber}-E${rowNumber}` } })
      } else {
        refs.push({ sheetId, mutation: { kind: 'setCellFormula' as const, row, col: 2, formula: `B${rowNumber}-A${rowNumber}` } })
        refs.push({ sheetId, mutation: { kind: 'setCellFormula' as const, row, col: 3, formula: `ABS(C${rowNumber})` } })
        refs.push({ sheetId, mutation: { kind: 'setCellFormula' as const, row, col: 4, formula: `MAX(A1:A${rowNumber})` } })
        refs.push({ sheetId, mutation: { kind: 'setCellFormula' as const, row, col: 5, formula: `D${rowNumber}+E${rowNumber}` } })
      }
    }
    source.initializeCellFormulasAt(refs, refs.length)
    const snapshot = source.exportSnapshot()

    const restored = new SpreadsheetEngine({ workbookName: 'engine-formula-initialize-hydrated-families-restored' })
    await restored.ready()
    const restoredStore = getFormulaFamilyStore(restored)
    const registerRunSpy = vi.spyOn(restoredStore, 'registerFormulaRun')
    try {
      restored.importSnapshot(snapshot)

      const restoredSheetId = restored.workbook.getSheet('Sheet1')!.id
      const stridedOutputRuns = restoredStore
        .listFamilies()
        .flatMap((family) => family.runs)
        .filter((run) => run.fixedIndex === 5 && run.step === 3 && run.cellIndices.length === 4)
        .map((run) => [run.start, run.end])
        .toSorted((left, right) => left[0]! - right[0]!)
      expect(restored.getCellValue('Sheet1', 'F12')).toEqual({ tag: ValueTag.Number, value: 48 })
      expect(restored.workbook.getSheet('Sheet1')!.id).toBe(restoredSheetId)
      expect(registerRunSpy.mock.calls.length).toBeGreaterThan(0)
      expect(registerRunSpy.mock.calls.length).toBeLessThan(refs.length)
      expect(restoredStore.getStats()).toEqual({ familyCount: 10, runCount: 12, memberCount: rowCount * 4 })
      expect(stridedOutputRuns).toEqual([
        [0, 9],
        [1, 10],
        [2, 11],
      ])
    } finally {
      registerRunSpy.mockRestore()
    }
  })
})
