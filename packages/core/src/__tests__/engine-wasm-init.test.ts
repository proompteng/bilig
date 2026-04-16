import { FormulaMode, ValueTag } from '@bilig/protocol'
import { describe, expect, it } from 'vitest'
import { SpreadsheetEngine } from '../index.js'

describe('SpreadsheetEngine wasm initialization', () => {
  it('initializes the wasm kernel synchronously for immediate mutations in Node runtimes', () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })

    expect(engine.wasm.ready).toBe(true)

    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 10)
    engine.setCellFormula('Sheet1', 'B1', 'A1+1')
    engine.setCellValue('Sheet1', 'A1', 12)

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 13 })
    expect(engine.getLastMetrics().wasmFormulaCount).toBeGreaterThan(0)
  })

  it('keeps exact vector MATCH bindings on the JS lookup path', () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })

    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 'Key')
    engine.setCellValue('Sheet1', 'A2', 'KEY-00001')
    engine.setCellValue('Sheet1', 'A3', 'KEY-00002')
    engine.setCellValue('Sheet1', 'D1', 'KEY-00002')
    engine.setCellFormula('Sheet1', 'E1', 'MATCH(D1,A2:A3,0)')

    expect(engine.explainCell('Sheet1', 'E1').mode).toBe(FormulaMode.JsOnly)
    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({ tag: ValueTag.Number, value: 2 })
  })

  it('flushes deferred JS-only edits before the first wasm formula evaluation', () => {
    const engine = new SpreadsheetEngine({ workbookName: 'spec' })

    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 10)
    engine.setCellValue('Sheet1', 'A1', 12)
    engine.setCellFormula('Sheet1', 'B1', 'A1+1')

    expect(engine.explainCell('Sheet1', 'B1').mode).toBe(FormulaMode.WasmFastPath)
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 13 })

    engine.setCellValue('Sheet1', 'A1', 20)

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 21 })
  })
})
