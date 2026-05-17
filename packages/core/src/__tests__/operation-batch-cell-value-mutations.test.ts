import { Effect } from 'effect'
import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import { createBatch } from '../replica-state.js'
import { SpreadsheetEngine } from '../engine.js'
import { applyBatchClearCellOp, applyBatchSetCellValueOp } from '../engine/services/operation-batch-cell-value-mutations.js'
import { getOperationService, getReplicaState } from './operation-service-test-helpers.js'

describe('operation batch cell value mutations', () => {
  it('keeps batch set-value and clear-cell application in a dedicated module', () => {
    expect(applyBatchSetCellValueOp).toBeTypeOf('function')
    expect(applyBatchClearCellOp).toBeTypeOf('function')
  })

  it('replaces existing formulas with generic batch literal writes and updates dependents', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'operation-batch-literal-over-formula' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 2)
    engine.setCellFormula('Sheet1', 'B1', 'A1*3')
    engine.setCellFormula('Sheet1', 'C1', 'B1+1')

    const batch = createBatch(getReplicaState(engine), [{ kind: 'setCellValue', sheetName: 'Sheet1', address: 'B1', value: 9 }])

    Effect.runSync(getOperationService(engine).applyBatch(batch, 'local'))

    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Number, value: 9 })
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 10 })
    expect(engine.getCell('Sheet1', 'B1').formula).toBeUndefined()
  })

  it('keeps lookup formulas current through generic batch literal writes', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'operation-batch-lookup-write',
      useColumnIndex: true,
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 10)
    engine.setCellValue('Sheet1', 'A2', 20)
    engine.setCellValue('Sheet1', 'A3', 30)
    engine.setCellValue('Sheet1', 'D1', 20)
    engine.setCellValue('Sheet1', 'D2', 25)
    engine.setCellFormula('Sheet1', 'E1', 'XMATCH(D1,A1:A3,0)')
    engine.setCellFormula('Sheet1', 'F1', 'MATCH(D2,A1:A3,1)')

    const batch = createBatch(getReplicaState(engine), [{ kind: 'setCellValue', sheetName: 'Sheet1', address: 'A2', value: 25 }])

    Effect.runSync(getOperationService(engine).applyBatch(batch, 'local'))

    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 2 })
  })

  it('keeps lookup and aggregate dependents current through generic batch clears', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'operation-batch-clear-lookup-and-aggregate',
      useColumnIndex: true,
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 10)
    engine.setCellValue('Sheet1', 'A2', 20)
    engine.setCellValue('Sheet1', 'A3', 30)
    engine.setCellValue('Sheet1', 'D1', 20)
    engine.setCellValue('Sheet1', 'D2', 25)
    engine.setCellFormula('Sheet1', 'E1', 'XMATCH(D1,A1:A3,0)')
    engine.setCellFormula('Sheet1', 'F1', 'MATCH(D2,A1:A3,1)')
    engine.setCellFormula('Sheet1', 'G1', 'SUM(A1:A3)')

    const batch = createBatch(getReplicaState(engine), [{ kind: 'clearCell', sheetName: 'Sheet1', address: 'A2' }])

    Effect.runSync(getOperationService(engine).applyBatch(batch, 'local'))

    expect(engine.getCellValue('Sheet1', 'E1')).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    })
    expect(engine.getCellValue('Sheet1', 'F1')).toEqual({ tag: ValueTag.Number, value: 2 })
    expect(engine.getCellValue('Sheet1', 'G1')).toEqual({ tag: ValueTag.Number, value: 40 })
  })
})
