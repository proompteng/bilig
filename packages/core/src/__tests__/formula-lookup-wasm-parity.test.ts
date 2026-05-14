import { describe, expect, it } from 'vitest'
import { ValueTag } from '@bilig/protocol'
import { SpreadsheetEngine } from '../engine.js'

describe('lookup wasm formula parity', () => {
  it('coerces blank VLOOKUP and HLOOKUP return cells to zero', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'lookup-blank-return-wasm' })
    await engine.ready()
    engine.createSheet('Sheet1')

    engine.setCellValue('Sheet1', 'A1', 'apple')
    engine.setCellValue('Sheet1', 'B1', 10)
    engine.setCellValue('Sheet1', 'A2', 'pear')
    engine.setCellValue('Sheet1', 'D1', 'Q1')
    engine.setCellValue('Sheet1', 'E1', 'Q2')
    engine.setCellValue('Sheet1', 'D2', 20)

    engine.setCellFormula('Sheet1', 'H1', 'VLOOKUP("pear",A1:B2,2,FALSE)')
    engine.setCellFormula('Sheet1', 'H2', 'HLOOKUP("Q2",D1:E2,2,FALSE)')

    expect(engine.getCellValue('Sheet1', 'H1')).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(engine.getCellValue('Sheet1', 'H2')).toEqual({ tag: ValueTag.Number, value: 0 })
    expect(engine.getLastMetrics()).toMatchObject({ jsFormulaCount: 0 })
  })
})
