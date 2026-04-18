import { describe, expect, it } from 'vitest'
import { ValueTag } from '@bilig/protocol'
import { SpreadsheetEngine } from '../engine.js'

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
})
