import { Effect } from 'effect'
import { describe, expect, it, vi } from 'vitest'
import { ErrorCode, FormulaMode, ValueTag } from '@bilig/protocol'
import { utcDateToExcelSerial } from '@bilig/formula'
import { SpreadsheetEngine } from '../engine.js'
import type { EngineTrackedEvent } from '../events.js'
import type { EngineRecalcService } from '../engine/services/recalc-service.js'
import type { RuntimeFormula } from '../engine/runtime-state.js'

function isEngineRecalcService(value: unknown): value is EngineRecalcService {
  if (typeof value !== 'object' || value === null) {
    return false
  }
  return (
    typeof Reflect.get(value, 'recalculateNow') === 'function' &&
    typeof Reflect.get(value, 'recalculateDirty') === 'function' &&
    typeof Reflect.get(value, 'recalculateDifferential') === 'function'
  )
}

function getRecalcService(engine: SpreadsheetEngine): EngineRecalcService {
  const runtime = Reflect.get(engine, 'runtime')
  if (typeof runtime !== 'object' || runtime === null) {
    throw new TypeError('Expected engine runtime')
  }
  const recalc = Reflect.get(runtime, 'recalc')
  if (!isEngineRecalcService(recalc)) {
    throw new TypeError('Expected engine recalc service')
  }
  return recalc
}

function isFormulaTable(value: unknown): value is { get(cellIndex: number): RuntimeFormula | undefined } {
  return typeof value === 'object' && value !== null && typeof Reflect.get(value, 'get') === 'function'
}

function getFormulaTable(engine: SpreadsheetEngine): { get(cellIndex: number): RuntimeFormula | undefined } {
  const formulas = Reflect.get(engine, 'formulas')
  if (!isFormulaTable(formulas)) {
    throw new TypeError('Expected engine formula table')
  }
  return formulas
}

describe('EngineRecalcService', () => {
  it('performs dirty-region recalculation through the extracted service boundary', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'recalc-dirty' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setCellValue('Sheet1', 'A1', 10)
    engine.setCellFormula('Sheet1', 'B1', 'A1*2')
    engine.setCellValue('Sheet1', 'C1', 5)
    engine.setCellFormula('Sheet1', 'D1', 'C1+10')

    const a1Index = engine.workbook.ensureCell('Sheet1', 'A1')
    const c1Index = engine.workbook.ensureCell('Sheet1', 'C1')
    engine.workbook.cellStore.setValue(a1Index, { tag: ValueTag.Number, value: 50 })
    engine.workbook.cellStore.setValue(c1Index, { tag: ValueTag.Number, value: 100 })

    const changed = Effect.runSync(
      getRecalcService(engine).recalculateDirty([{ sheetName: 'Sheet1', rowStart: 0, rowEnd: 0, colStart: 0, colEnd: 0 }]),
    )

    const b1Index = engine.workbook.getCellIndex('Sheet1', 'B1')
    const d1Index = engine.workbook.getCellIndex('Sheet1', 'D1')
    expect(changed).toContain(b1Index)
    expect(changed).not.toContain(d1Index)
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 100 })
    expect(engine.getCellValue('Sheet1', 'D1')).toEqual({ tag: ValueTag.Number, value: 15 })
  })

  it('recalculates all formulas through the extracted service boundary', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'recalc-now' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 10)
    engine.setCellFormula('Sheet1', 'B1', 'A1*2')
    const b1Index = engine.workbook.getCellIndex('Sheet1', 'B1')
    const formula = b1Index === undefined ? undefined : getFormulaTable(engine).get(b1Index)
    if (!formula) {
      throw new TypeError('Expected B1 formula')
    }
    formula.compiled.mode = FormulaMode.JsOnly

    const a1Index = engine.workbook.ensureCell('Sheet1', 'A1')
    engine.workbook.cellStore.setValue(a1Index, { tag: ValueTag.Number, value: 25 })

    const changed = Effect.runSync(getRecalcService(engine).recalculateNow())

    expect(changed).toContain(b1Index)
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 50 })
  })

  it('batches large direct scalar full recalculation through the native scalar kernel', async () => {
    const rowCount = 160
    const engine = new SpreadsheetEngine({ workbookName: 'recalc-native-direct-scalar' })
    await engine.ready()
    engine.createSheet('Sheet1')

    for (let row = 1; row <= rowCount; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
      engine.setCellFormula('Sheet1', `B${row}`, `A${row}*2+1`)
    }

    for (let row = 1; row <= rowCount; row += 1) {
      const cellIndex = engine.workbook.ensureCell('Sheet1', `A${row}`)
      engine.workbook.cellStore.setValue(cellIndex, { tag: ValueTag.Number, value: row + 10 })
    }

    engine.resetPerformanceCounters()
    const changed = Effect.runSync(getRecalcService(engine).recalculateNow())
    const counters = engine.getPerformanceCounters() as Record<string, number | undefined> & {
      readonly nativeDirectScalarRecalcEvaluations?: number
    }

    expect(changed).toContain(engine.workbook.getCellIndex('Sheet1', `B${rowCount}`))
    expect(engine.getCellValue('Sheet1', `B${rowCount}`)).toEqual({ tag: ValueTag.Number, value: (rowCount + 10) * 2 + 1 })
    expect(engine.getLastMetrics().wasmFormulaCount).toBeGreaterThanOrEqual(rowCount)
    expect(counters.nativeDirectScalarRecalcEvaluations).toBeGreaterThanOrEqual(rowCount)
  })

  it('keeps numeric text direct scalar recalculation on the JS oracle path', async () => {
    const rowCount = 80
    const engine = new SpreadsheetEngine({ workbookName: 'recalc-native-direct-scalar-text' })
    await engine.ready()
    engine.createSheet('Sheet1')

    for (let row = 1; row <= rowCount; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, `${row}`)
      engine.setCellFormula('Sheet1', `B${row}`, `A${row}*2+1`)
    }

    for (let row = 1; row <= rowCount; row += 1) {
      const textValue = `${row + 10}`
      const stringId = engine.strings.intern(textValue)
      const cellIndex = engine.workbook.ensureCell('Sheet1', `A${row}`)
      engine.workbook.cellStore.setValue(cellIndex, { tag: ValueTag.String, value: textValue, stringId }, stringId)
    }

    engine.resetPerformanceCounters()
    const changed = Effect.runSync(getRecalcService(engine).recalculateNow())

    expect(changed).toContain(engine.workbook.getCellIndex('Sheet1', `B${rowCount}`))
    expect(engine.getCellValue('Sheet1', `B${rowCount}`)).toEqual({ tag: ValueTag.Number, value: (rowCount + 10) * 2 + 1 })
    expect(engine.getPerformanceCounters().nativeDirectScalarRecalcEvaluations).toBe(0)
  })

  it('batches large direct uniform lookup full recalculation through the native lookup kernel', async () => {
    const lookupRowCount = 256
    const formulaRowCount = 80
    const engine = new SpreadsheetEngine({ workbookName: 'recalc-native-direct-lookup' })
    await engine.ready()
    engine.createSheet('Sheet1')

    for (let row = 1; row <= lookupRowCount; row += 1) {
      engine.setCellValue('Sheet1', `A${row}`, row)
    }
    for (let row = 1; row <= formulaRowCount; row += 1) {
      engine.setCellValue('Sheet1', `D${row}`, row)
    }
    for (let row = 1; row <= formulaRowCount; row += 1) {
      engine.setCellFormula('Sheet1', `B${row}`, `MATCH(D${row},$A$1:$A$${lookupRowCount},0)`)
    }
    for (let row = 1; row <= formulaRowCount; row += 1) {
      engine.setCellFormula('Sheet1', `C${row}`, `B${row}+10`)
    }

    for (let row = 1; row <= formulaRowCount; row += 1) {
      const cellIndex = engine.workbook.ensureCell('Sheet1', `D${row}`)
      engine.workbook.cellStore.setValue(cellIndex, { tag: ValueTag.Number, value: row + 1 })
    }

    engine.resetPerformanceCounters()
    const changed = Effect.runSync(getRecalcService(engine).recalculateNow())
    const counters = engine.getPerformanceCounters() as Record<string, number | undefined> & {
      readonly nativeDirectLookupRecalcEvaluations?: number
    }

    expect(changed).toContain(engine.workbook.getCellIndex('Sheet1', `B${formulaRowCount}`))
    expect(changed).toContain(engine.workbook.getCellIndex('Sheet1', `C${formulaRowCount}`))
    expect(engine.getCellValue('Sheet1', `B${formulaRowCount}`)).toEqual({
      tag: ValueTag.Number,
      value: formulaRowCount + 1,
    })
    expect(engine.getCellValue('Sheet1', `C${formulaRowCount}`)).toEqual({
      tag: ValueTag.Number,
      value: formulaRowCount + 11,
    })
    expect(counters.nativeDirectLookupRecalcEvaluations).toBeGreaterThanOrEqual(formulaRowCount)
  })

  it('preserves imported cached unsupported formula values during full recalculation', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'recalc-imported-unsupported-cache' })
    await engine.ready()
    engine.importSnapshot({
      version: 1,
      workbook: { name: 'Imported cached external formulas' },
      sheets: [
        {
          id: 1,
          name: 'Sheet1',
          order: 0,
          cells: [
            {
              address: 'A1',
              formula: '_xldudf_WISEPRICE(B1,"Shares Outstanding")',
              value: 14935800000,
            },
            { address: 'B1', value: 'AAPL' },
            { address: 'C1', formula: 'A1/1000000' },
          ],
        },
      ],
    })

    const a1Index = engine.workbook.getCellIndex('Sheet1', 'A1')
    const changed = Effect.runSync(getRecalcService(engine).recalculateNow())

    expect(changed).not.toContain(a1Index)
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Number, value: 14935800000 })
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 14935.8 })

    engine.setCellValue('Sheet1', 'B1', 'MSFT')

    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Error, code: ErrorCode.Name })
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Error, code: ErrorCode.Name })
  })

  it('emits tracked recalculation patches through the service event boundary', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'recalc-tracked-events' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 10)
    engine.setCellFormula('Sheet1', 'B1', 'A1*2')

    const tracked: EngineTrackedEvent[] = []
    const unsubscribe = engine.events.subscribeTracked((event) => {
      tracked.push(event)
    })

    const a1Index = engine.workbook.ensureCell('Sheet1', 'A1')
    engine.workbook.cellStore.setValue(a1Index, { tag: ValueTag.Number, value: 30 })
    const changed = Effect.runSync(
      getRecalcService(engine).recalculateDirty([{ sheetName: 'Sheet1', rowStart: 0, rowEnd: 0, colStart: 0, colEnd: 0 }]),
    )
    unsubscribe()

    expect(changed).toContain(engine.workbook.getCellIndex('Sheet1', 'B1'))
    expect(tracked).toHaveLength(1)
    expect(tracked[0]).toMatchObject({
      kind: 'batch',
      invalidation: 'cells',
      explicitChangedCount: 1,
    })
    expect(tracked[0]?.patches?.length).toBeGreaterThan(0)
  })

  it('recalculates volatile formulas from the current clock and random inputs', async () => {
    vi.useFakeTimers()
    const randomSpy = vi.spyOn(Math, 'random')
    vi.setSystemTime(new Date('2026-02-03T00:00:00Z'))
    randomSpy.mockReturnValue(0.125)

    const engine = new SpreadsheetEngine({ workbookName: 'recalc-volatile' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'A1', 'TODAY()')
    engine.setCellFormula('Sheet1', 'B1', 'RAND()')

    vi.setSystemTime(new Date('2026-02-04T00:00:00Z'))
    randomSpy.mockReturnValue(0.75)

    const changed = Effect.runSync(getRecalcService(engine).recalculateNow())

    expect(changed).toContain(engine.workbook.getCellIndex('Sheet1', 'A1'))
    expect(changed).toContain(engine.workbook.getCellIndex('Sheet1', 'B1'))
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({
      tag: ValueTag.Number,
      value: utcDateToExcelSerial(new Date('2026-02-04T00:00:00Z')),
    })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({
      tag: ValueTag.Number,
      value: 0.75,
    })
  })

  it('recalculates SUBTOTAL when row visibility metadata changes', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'recalc-subtotal-row-visibility' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 1)
    engine.setCellValue('Sheet1', 'A2', 2)
    engine.setCellValue('Sheet1', 'A3', 3)
    engine.setCellFormula('Sheet1', 'B1', 'SUBTOTAL(109,A1:A3)')

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 6 })

    engine.updateRowMetadata('Sheet1', 1, 1, null, true)

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 4 })

    engine.updateRowMetadata('Sheet1', 1, 1, null, false)

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 6 })
  })

  it('converges iterative circular formulas through the recalc service boundary', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'recalc-iterative-cycle' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCalculationSettings({ iterate: true, iterateCount: 100, iterateDelta: '0.0000000001' })

    engine.setCellValue('Sheet1', 'B2', 100000)
    engine.setCellValue('Sheet1', 'B3', 0.1)
    engine.setCellValue('Sheet1', 'B4', 5000)
    engine.setCellFormula('Sheet1', 'B5', 'B6*B3')
    engine.setCellFormula('Sheet1', 'B6', 'B2+B5-B4')

    const changed = Effect.runSync(getRecalcService(engine).recalculateNow())
    const b5Index = engine.workbook.getCellIndex('Sheet1', 'B5')
    const b6Index = engine.workbook.getCellIndex('Sheet1', 'B6')

    expect(changed).toContain(b5Index)
    expect(changed).toContain(b6Index)
    expect(engine.getCellValue('Sheet1', 'B5').tag).toBe(ValueTag.Number)
    expect(engine.getCellValue('Sheet1', 'B5')).toMatchObject({ tag: ValueTag.Number })
    expect(engine.getCellValue('Sheet1', 'B5').value).toBeCloseTo(10555.555555555555, 10)
    expect(engine.getCellValue('Sheet1', 'B6').tag).toBe(ValueTag.Number)
    expect(engine.getCellValue('Sheet1', 'B6')).toMatchObject({ tag: ValueTag.Number })
    expect(engine.getCellValue('Sheet1', 'B6').value).toBeCloseTo(105555.55555555555, 10)
  })

  it('keeps non-iterative circular formulas as cycle errors through the recalc service boundary', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'recalc-non-iterative-cycle' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setCellFormula('Sheet1', 'A1', 'B1+1')
    engine.setCellFormula('Sheet1', 'B1', 'A1+1')

    const changed = Effect.runSync(getRecalcService(engine).recalculateNow())
    const a1Index = engine.workbook.getCellIndex('Sheet1', 'A1')
    const b1Index = engine.workbook.getCellIndex('Sheet1', 'B1')

    expect(changed).toContain(a1Index)
    expect(changed).toContain(b1Index)
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: ValueTag.Error, code: ErrorCode.Cycle })
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Error, code: ErrorCode.Cycle })
  })
})
