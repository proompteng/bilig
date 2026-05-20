import { describe, expect, it, vi } from 'vitest'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import { SpreadsheetEngine } from '../engine.js'
import { EngineEvaluationTimeoutError, EngineMutationError } from '../engine/errors.js'
import type { EngineCellMutationRef } from '../cell-mutations-at.js'
import type { FormulaFamilyStore } from '../formula/formula-family-store.js'
import { INITIAL_DIRECT_FORMULA_EVALUATION_LIMIT } from '../engine/services/formula-initialization-direct-formulas.js'
import { findErrorByName, getFormulaBindingNowService } from './operation-service-test-helpers.js'

const OVER_DIRECT_LIMIT_TEST_TIMEOUT_MS = 10_000

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
  const binding = Reflect.get(runtime, 'binding')
  const getFormulaFamilyStatsNow =
    typeof binding === 'object' && binding !== null ? Reflect.get(binding, 'getFormulaFamilyStatsNow') : undefined
  if (typeof getFormulaFamilyStatsNow === 'function') {
    getFormulaFamilyStatsNow.call(binding)
  }
  return formulaFamilies
}

function readFormulaFamilyStatsWithoutHelperEnsure(engine: SpreadsheetEngine): ReturnType<FormulaFamilyStore['getStats']> {
  const binding = getFormulaBindingNowService(engine)
  const getFormulaFamilyStatsNow = Reflect.get(binding, 'getFormulaFamilyStatsNow')
  if (typeof getFormulaFamilyStatsNow !== 'function') {
    throw new TypeError('Expected formula family stats service')
  }
  return getFormulaFamilyStatsNow.call(binding)
}

interface RuntimeFormulaStoreForTest {
  readonly forEach: (fn: (value: unknown, key: number) => void) => void
}

function isRuntimeFormulaStoreForTest(value: unknown): value is RuntimeFormulaStoreForTest {
  return typeof value === 'object' && value !== null && typeof Reflect.get(value, 'forEach') === 'function'
}

function getRuntimeFormulaStore(engine: SpreadsheetEngine): RuntimeFormulaStoreForTest {
  const formulas = Reflect.get(engine, 'formulas')
  if (!isRuntimeFormulaStoreForTest(formulas)) {
    throw new TypeError('Expected internal formula store')
  }
  return formulas
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

  it('keeps aggregate dependencies on formula cells in existing-workbook initialization batches', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'engine-formula-initialize-existing-aggregate-deps' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'A1', '1+1')
    engine.setCellValue('Sheet1', 'A2', 100)
    engine.setCellValue('Sheet1', 'A3', 200)
    const sheetId = engine.workbook.getSheet('Sheet1')!.id

    engine.initializeCellFormulasAt(
      [
        { sheetId, mutation: { kind: 'setCellFormula', row: 3, col: 2, formula: 'SUM(C5:C7)' } },
        { sheetId, mutation: { kind: 'setCellFormula', row: 4, col: 2, formula: 'A2' } },
        { sheetId, mutation: { kind: 'setCellFormula', row: 5, col: 2, formula: 'A3' } },
        { sheetId, mutation: { kind: 'setCellFormula', row: 6, col: 2, formula: '300' } },
      ],
      4,
    )

    expect(engine.getCellValue('Sheet1', 'C4')).toEqual({ tag: ValueTag.Number, value: 600 })

    engine.setCellValue('Sheet1', 'A2', 150)

    expect(engine.getCellValue('Sheet1', 'C5')).toEqual({ tag: ValueTag.Number, value: 150 })
    expect(engine.getCellValue('Sheet1', 'C4')).toEqual({ tag: ValueTag.Number, value: 650 })
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

  it('surfaces formula binding timeouts instead of storing fake invalid formulas during initialization', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'engine-formula-initialize-timeout' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const timeout = new EngineEvaluationTimeoutError(50)
    const binding = getFormulaBindingNowService(engine)
    const bindInitialSpy = vi.spyOn(binding, 'bindInitialFormulaNow').mockImplementation(() => {
      throw timeout
    })
    const bindPreparedSpy = vi.spyOn(binding, 'bindPreparedFormulaNow').mockImplementation(() => {
      throw timeout
    })

    let thrown: unknown
    try {
      engine.initializeCellFormulasAt([{ sheetId, mutation: { kind: 'setCellFormula', row: 0, col: 1, formula: 'A1+1' } }], 1)
    } catch (error) {
      thrown = error
    } finally {
      bindInitialSpy.mockRestore()
      bindPreparedSpy.mockRestore()
    }

    expect(thrown).toBeInstanceOf(EngineMutationError)
    expect(findErrorByName(thrown, 'EngineEvaluationTimeoutError')).toBe(timeout)
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
    expect(engine.getPerformanceCounters().directFormulaInitialEvaluations).toBe(4)
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

  it('uses native anchored prefix aggregate initialization beyond the JS direct limit', async () => {
    const rowCount = 20_000
    const engine = new SpreadsheetEngine({ workbookName: 'engine-formula-initialize-native-prefix-aggregate-family' })
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
    expect(engine.getPerformanceCounters()).toMatchObject({
      directFormulaInitialEvaluations: rowCount,
      nativeDirectAggregatePrefixEvaluations: rowCount,
      directAggregatePrefixEvaluations: 0,
      directAggregateScanEvaluations: 0,
      regionQueryIndexBuilds: 0,
    })
  })

  it('uses native direct-scalar initialization beyond the JS direct limit', async () => {
    const rowCount = 20_000
    const engine = new SpreadsheetEngine({ workbookName: 'engine-formula-initialize-native-direct-scalar-family' })
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
          formula: `A${row + 1}*2+1`,
        },
      })),
      rowCount,
    )

    expect(engine.getCellValue('Sheet1', `B${rowCount}`)).toEqual({
      tag: ValueTag.Number,
      value: rowCount * 2 + 1,
    })
    expect(engine.getLastMetrics()).toMatchObject({ dirtyFormulaCount: 0, wasmFormulaCount: 0, jsFormulaCount: 0 })
    expect(engine.getPerformanceCounters()).toMatchObject({
      directFormulaInitialEvaluations: rowCount,
      nativeDirectScalarInitialEvaluations: rowCount,
      nativeDirectAggregatePrefixEvaluations: 0,
      directAggregatePrefixEvaluations: 0,
      directAggregateScanEvaluations: 0,
      regionQueryIndexBuilds: 0,
    })
  })

  it('materializes out-of-order anchored prefix aggregates during initial direct evaluation', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'engine-formula-initialize-out-of-order-prefix-aggregate' })
    await engine.ready()
    engine.createSheet('Sheet1')
    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    for (let row = 1; row <= 4; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
    }

    engine.resetPerformanceCounters()
    engine.initializeCellFormulasAt(
      [
        { sheetId, mutation: { kind: 'setCellFormula', row: 3, col: 1, formula: 'SUM(A1:A4)' } },
        { sheetId, mutation: { kind: 'setCellFormula', row: 0, col: 1, formula: 'SUM(A1:A1)' } },
        { sheetId, mutation: { kind: 'setCellFormula', row: 2, col: 1, formula: 'SUM(A1:A3)' } },
        { sheetId, mutation: { kind: 'setCellFormula', row: 1, col: 1, formula: 'SUM(A1:A2)' } },
      ],
      4,
    )

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'B2')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'B3')).toEqual({ tag: ValueTag.Number, value: 6 })
    expect(engine.getCellValue('Sheet1', 'B4')).toEqual({ tag: ValueTag.Number, value: 10 })
    expect(engine.getPerformanceCounters().directFormulaInitialEvaluations).toBe(4)
    expect(engine.getPerformanceCounters().regionQueryIndexBuilds).toBe(0)
  })

  it('materializes initial aggregate variants and scalar fallback values in one pass', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'engine-formula-initialize-direct-variants' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'A2', true)
    engine.setCellValue('Sheet1', 'A3', 3)
    engine.setCellValue('Sheet1', 'A4', 'text')
    const sheetId = engine.workbook.getSheet('Sheet1')!.id

    engine.initializeCellFormulasAt(
      [
        { sheetId, mutation: { kind: 'setCellFormula', row: 0, col: 1, formula: 'SUM(A1:A3)' } },
        { sheetId, mutation: { kind: 'setCellFormula', row: 0, col: 2, formula: 'COUNT(A1:A3)' } },
        { sheetId, mutation: { kind: 'setCellFormula', row: 0, col: 3, formula: 'AVERAGE(A1:A3)' } },
        { sheetId, mutation: { kind: 'setCellFormula', row: 0, col: 4, formula: 'MIN(A1:A3)' } },
        { sheetId, mutation: { kind: 'setCellFormula', row: 0, col: 5, formula: 'MAX(A1:A3)' } },
        { sheetId, mutation: { kind: 'setCellFormula', row: 0, col: 6, formula: 'A4+1' } },
        { sheetId, mutation: { kind: 'setCellFormula', row: 0, col: 7, formula: 'A1/0' } },
        { sheetId, mutation: { kind: 'setCellFormula', row: 0, col: 8, formula: 'ABS(A4)' } },
        { sheetId, mutation: { kind: 'setCellFormula', row: 0, col: 9, formula: 'A3-A1' } },
        { sheetId, mutation: { kind: 'setCellFormula', row: 0, col: 10, formula: 'A1*A3' } },
      ],
      10,
    )

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 5 })
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 5 / 3 })
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 1 })
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getCellValue('Sheet1', 'G1')).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(engine.getCellValue('Sheet1', 'H1')).toEqual({ tag: ValueTag.Error, code: ErrorCode.Div0 })
    expect(engine.getCellValue('Sheet1', 'I1')).toEqual({ tag: ValueTag.Error, code: ErrorCode.Value })
    expect(engine.getCellValue('Sheet1', 'J1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'K1')).toEqual({ tag: ValueTag.Number, value: 3 })
    expect(engine.getPerformanceCounters().directFormulaInitialEvaluations).toBe(10)
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

  it(
    'replays over-cap fresh formula family runs without a full family-index rebuild',
    async () => {
      const rowCount = 16_385
      const engine = new SpreadsheetEngine({ workbookName: 'engine-formula-initialize-over-cap-family-runs' })
      await engine.ready()
      engine.createSheet('Sheet1')
      const sheetId = engine.workbook.getSheet('Sheet1')!.id
      const store = getFormulaFamilyStore(engine)
      const upsertFormula = vi.spyOn(store, 'upsertFormula')
      const registerFreshUniformRun = vi.spyOn(store, 'registerFreshUniformRun')
      const refs: EngineCellMutationRef[] = []
      for (let row = 0; row < rowCount; row += 1) {
        const rowNumber = row + 1
        engine.setCellValue('Sheet1', `A${rowNumber}`, rowNumber)
        refs.push({ sheetId, mutation: { kind: 'setCellFormula' as const, row, col: 1, formula: `A${rowNumber}+1` } })
      }

      try {
        engine.initializeCellFormulasAt(refs, refs.length)

        expect(upsertFormula).not.toHaveBeenCalled()
        expect(registerFreshUniformRun).not.toHaveBeenCalled()
        expect(engine.getCellValue('Sheet1', `B${rowCount}`)).toEqual({
          tag: ValueTag.Number,
          value: rowCount + 1,
        })
        const formulaForEach = vi.spyOn(getRuntimeFormulaStore(engine), 'forEach')
        try {
          expect(readFormulaFamilyStatsWithoutHelperEnsure(engine)).toEqual({ familyCount: 1, runCount: 1, memberCount: rowCount })
          expect(formulaForEach).not.toHaveBeenCalled()
          expect(registerFreshUniformRun).toHaveBeenCalledTimes(1)
          expect(upsertFormula).not.toHaveBeenCalled()
        } finally {
          formulaForEach.mockRestore()
        }
      } finally {
        upsertFormula.mockRestore()
        registerFreshUniformRun.mockRestore()
      }
    },
    OVER_DIRECT_LIMIT_TEST_TIMEOUT_MS,
  )

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
      refs.push({ sheetId, mutation: { kind: 'setCellFormula' as const, row, col: 35, formula: `A${rowNumber}+B${rowNumber}` } })
    }
    const notifyCellValueWritten = vi.spyOn(engine.workbook, 'notifyCellValueWritten')
    const notifyColumnsWritten = vi.spyOn(engine.workbook, 'notifyColumnsWritten')

    engine.initializeCellFormulasAt(refs, refs.length)

    expect(engine.getCellValue('Sheet1', `D${rowCount}`)).toEqual({
      tag: ValueTag.Number,
      value: rowCount * 11 * 2,
    })
    expect(engine.getCellValue('Sheet1', `AJ${rowCount}`)).toEqual({
      tag: ValueTag.Number,
      value: rowCount * 11,
    })
    expect(notifyCellValueWritten).not.toHaveBeenCalled()
    expect(notifyColumnsWritten).toHaveBeenCalledWith(sheetId, Uint32Array.from([2, 3, 35]))
    expect(engine.getPerformanceCounters().directFormulaInitialEvaluations).toBe(refs.length)
    expect(engine.getPerformanceCounters()).toMatchObject({ nativeDirectScalarInitialEvaluations: 0 })
    notifyCellValueWritten.mockRestore()
    notifyColumnsWritten.mockRestore()
  })

  it('uses native direct scalar batches for large fresh formula initialization', async () => {
    const rowCount = 6000
    const engine = new SpreadsheetEngine({ workbookName: 'engine-formula-initialize-native-bulk-value-columns' })
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

    engine.initializeCellFormulasAt(refs, refs.length)

    expect(engine.getCellValue('Sheet1', `D${rowCount}`)).toEqual({
      tag: ValueTag.Number,
      value: rowCount * 11 * 2,
    })
    expect(engine.getPerformanceCounters()).toMatchObject({
      directFormulaInitialEvaluations: refs.length,
      nativeDirectScalarInitialEvaluations: refs.length,
    })
  })

  it('chunks native direct scalar runs inside large mixed fresh formula initialization', async () => {
    const rowCount = 5500
    const engine = new SpreadsheetEngine({ workbookName: 'engine-formula-initialize-native-mixed-runs' })
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
      refs.push({ sheetId, mutation: { kind: 'setCellFormula' as const, row, col: 4, formula: `SUM(A1:A${rowNumber})` } })
    }

    engine.initializeCellFormulasAt(refs, refs.length)

    expect(engine.getCellValue('Sheet1', `D${rowCount}`)).toEqual({
      tag: ValueTag.Number,
      value: rowCount * 11 * 2,
    })
    expect(engine.getCellValue('Sheet1', `E${rowCount}`)).toEqual({
      tag: ValueTag.Number,
      value: (rowCount * (rowCount + 1)) / 2,
    })
    expect(getFormulaFamilyStore(engine).getStats()).toEqual({
      familyCount: 3,
      runCount: 3,
      memberCount: refs.length,
    })
    expect(engine.getPerformanceCounters()).toMatchObject({
      directFormulaInitialEvaluations: refs.length,
      nativeDirectScalarInitialEvaluations: rowCount * 2,
      nativeDirectAggregatePrefixEvaluations: 0,
      regionQueryIndexBuilds: 0,
    })
  })

  it('uses native uniform lookup batches for large fresh formula initialization', async () => {
    const lookupRows = 2000
    const formulaRows = 150
    const engine = new SpreadsheetEngine({ workbookName: 'engine-formula-initialize-native-lookup-columns' })
    await engine.ready()
    engine.createSheet('Sheet1')
    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    for (let row = 0; row < lookupRows; row += 1) {
      engine.setCellValue('Sheet1', `A${row + 1}`, row + 1)
    }
    const refs: EngineCellMutationRef[] = []
    for (let row = 0; row < formulaRows; row += 1) {
      const rowNumber = row + 1
      engine.setCellValue('Sheet1', `D${rowNumber}`, rowNumber + 100)
      engine.setCellValue('Sheet1', `E${rowNumber}`, rowNumber + 0.5)
      refs.push({
        sheetId,
        mutation: { kind: 'setCellFormula' as const, row, col: 5, formula: `MATCH(D${rowNumber},A1:A${lookupRows},0)` },
      })
      refs.push({
        sheetId,
        mutation: { kind: 'setCellFormula' as const, row, col: 6, formula: `MATCH(E${rowNumber},A1:A${lookupRows},1)` },
      })
    }

    engine.initializeCellFormulasAt(refs, refs.length)

    expect(engine.getCellValue('Sheet1', `F${formulaRows}`)).toEqual({
      tag: ValueTag.Number,
      value: formulaRows + 100,
    })
    expect(engine.getCellValue('Sheet1', `G${formulaRows}`)).toEqual({
      tag: ValueTag.Number,
      value: formulaRows,
    })
    expect(engine.getPerformanceCounters()).toMatchObject({
      directFormulaInitialEvaluations: refs.length,
      nativeDirectLookupInitialEvaluations: refs.length,
    })
  })

  it(
    'uses native uniform lookup initialization beyond the JS direct limit',
    async () => {
      const lookupRows = 512
      const formulaRows = INITIAL_DIRECT_FORMULA_EVALUATION_LIMIT + 1
      const engine = new SpreadsheetEngine({ workbookName: 'engine-formula-initialize-native-lookup-over-limit' })
      await engine.ready()
      engine.createSheet('Sheet1')
      const sheetId = engine.workbook.getSheet('Sheet1')!.id
      for (let row = 1; row <= lookupRows; row += 1) {
        engine.setCellValue('Sheet1', `A${row}`, row)
      }
      for (let row = 1; row <= formulaRows; row += 1) {
        engine.setCellValue('Sheet1', `D${row}`, ((row - 1) % lookupRows) + 1)
      }

      engine.resetPerformanceCounters()
      engine.initializeCellFormulasAt(
        Array.from({ length: formulaRows }, (_entry, row) => ({
          sheetId,
          mutation: {
            kind: 'setCellFormula' as const,
            row,
            col: 4,
            formula: `MATCH(D${row + 1},A1:A${lookupRows},0)`,
          },
        })),
        formulaRows,
      )

      expect(engine.getCellValue('Sheet1', `E${formulaRows}`)).toEqual({
        tag: ValueTag.Number,
        value: ((formulaRows - 1) % lookupRows) + 1,
      })
      expect(engine.getLastMetrics()).toMatchObject({ dirtyFormulaCount: 0, wasmFormulaCount: 0, jsFormulaCount: 0 })
      expect(engine.getPerformanceCounters()).toMatchObject({
        directFormulaInitialEvaluations: formulaRows,
        nativeDirectLookupInitialEvaluations: formulaRows,
        directAggregatePrefixEvaluations: 0,
        directAggregateScanEvaluations: 0,
        regionQueryIndexBuilds: 0,
      })
    },
    OVER_DIRECT_LIMIT_TEST_TIMEOUT_MS,
  )

  it('keeps mid-sized direct scalar initialization on the JS path', async () => {
    const rowCount = 1500
    const engine = new SpreadsheetEngine({ workbookName: 'engine-formula-initialize-mid-size-js-value-columns' })
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

    engine.initializeCellFormulasAt(refs, refs.length)

    expect(engine.getCellValue('Sheet1', `D${rowCount}`)).toEqual({
      tag: ValueTag.Number,
      value: rowCount * 11 * 2,
    })
    expect(engine.getPerformanceCounters()).toMatchObject({
      directFormulaInitialEvaluations: refs.length,
      nativeDirectScalarInitialEvaluations: 0,
    })
  })

  it('uses native direct scalar initialization for 12000 formula mixed-content builds', async () => {
    const rowCount = 6000
    const engine = new SpreadsheetEngine({ workbookName: 'engine-formula-initialize-12000-native-value-columns' })
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

    engine.initializeCellFormulasAt(refs, refs.length)

    expect(engine.getCellValue('Sheet1', `D${rowCount}`)).toEqual({
      tag: ValueTag.Number,
      value: rowCount * 11 * 2,
    })
    expect(engine.getPerformanceCounters()).toMatchObject({
      directFormulaInitialEvaluations: refs.length,
      nativeDirectScalarInitialEvaluations: refs.length,
    })
  })

  it('uses native row-chain initialization for larger direct scalar formula chains', async () => {
    const rowCount = 8193
    const engine = new SpreadsheetEngine({ workbookName: 'engine-formula-initialize-native-row-chain' })
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

    engine.initializeCellFormulasAt(refs, refs.length)

    expect(engine.getCellValue('Sheet1', `D${rowCount}`)).toEqual({
      tag: ValueTag.Number,
      value: rowCount * 11 * 2,
    })
    expect(engine.getPerformanceCounters()).toMatchObject({
      directFormulaInitialEvaluations: refs.length,
      nativeDirectScalarInitialEvaluations: refs.length,
    })
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

  it('lazily materializes hydrated runtime snapshot formulas into strided family runs', async () => {
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
      expect(registerRunSpy).not.toHaveBeenCalled()
      getFormulaFamilyStore(restored)

      const restoredSheetId = restored.workbook.getSheet('Sheet1')!.id
      const stridedOutputRuns = restoredStore
        .listFamilies()
        .flatMap((family) => family.runs)
        .filter((run) => run.fixedIndex === 5 && run.step === 3 && run.cellIndices.length === 4)
        .map((run) => [run.start, run.end])
        .toSorted((left, right) => left[0]! - right[0]!)
      expect(restored.getCellValue('Sheet1', 'F12')).toEqual({ tag: ValueTag.Number, value: 48 })
      expect(restored.workbook.getSheet('Sheet1')!.id).toBe(restoredSheetId)
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
