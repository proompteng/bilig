import { describe, expect, it } from 'vitest'
import { ValueTag } from '@bilig/protocol'

import { SpreadsheetEngine, type EngineCellMutationRef } from '../engine.js'

describe('structural tail append direct aggregates', () => {
  it('does not defer unchanged existing row aggregate formulas after tail row inserts', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'tail-append-direct-aggregates' })
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

    engine.resetPerformanceCounters()
    engine.insertRows('Sheet1', existingRows, appendRows)

    expect([...engine.state.formulas.values()].filter((formula) => formula.structuralSourceTransform !== undefined)).toHaveLength(0)
    expect(engine.getPerformanceCounters()).toMatchObject({
      structuralFormulaImpactCandidates: 0,
      structuralFormulaRebindInputs: 0,
      structuralTransactions: 1,
    })

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
    expect(engine.getPerformanceCounters()).toMatchObject({
      directAggregateScanCells: 0,
      directAggregateScanEvaluations: 0,
      formulasBound: 0,
      structuralFormulaRebindInputs: 0,
    })
  })
})
