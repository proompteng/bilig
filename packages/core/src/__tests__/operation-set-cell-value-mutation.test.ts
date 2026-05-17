import { Effect } from 'effect'
import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import { createBatch } from '../replica-state.js'
import { SpreadsheetEngine } from '../engine.js'
import { cellMutationRefToEngineOp, type EngineCellMutationRef } from '../cell-mutations-at.js'
import { applySetCellValueMutation } from '../engine/services/operation-set-cell-value-mutation.js'
import { getOperationService, getReplicaState } from './operation-service-test-helpers.js'

describe('operation set-cell-value mutations', () => {
  it('keeps set-cell-value mutation application in a dedicated module', () => {
    expect(applySetCellValueMutation).toBeTypeOf('function')
  })

  it('treats batched null writes to already-empty tracked dependency cells as no-ops', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'operation-batch-null-empty-noop' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'A1', 'A1+D4')

    const before = engine.exportSnapshot()
    const batch = createBatch(getReplicaState(engine), [{ kind: 'setCellValue', sheetName: 'Sheet1', address: 'D4', value: null }])

    Effect.runSync(getOperationService(engine).applyBatch(batch, 'local'))

    expect(engine.exportSnapshot()).toEqual(before)
  })

  it('treats batched null writes to missing cells as no-ops', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'operation-batch-null-missing-noop' })
    await engine.ready()
    engine.createSheet('Sheet1')

    const before = engine.exportSnapshot()
    const batch = createBatch(getReplicaState(engine), [{ kind: 'setCellValue', sheetName: 'Sheet1', address: 'D4', value: null }])

    Effect.runSync(getOperationService(engine).applyBatch(batch, 'local'))

    expect(engine.exportSnapshot()).toEqual(before)
  })

  it('prunes explicit null literal writes back out of tracked empty dependency cells', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'operation-null-literal-empty-prune',
      replicaId: 'a',
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'A1', 'A1+D4')
    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const refs: EngineCellMutationRef[] = [
      {
        sheetId,
        mutation: { kind: 'setCellValue', row: 3, col: 3, value: null },
      },
    ]
    const batch = createBatch(
      getReplicaState(engine),
      refs.map((ref) => cellMutationRefToEngineOp(engine.workbook, ref)),
    )

    Effect.runSync(getOperationService(engine).applyCellMutationsAt(refs, batch, 'local', 1))

    expect(engine.exportSnapshot().sheets[0].cells.some((cell) => cell.address === 'D4')).toBe(false)
  })

  it('treats mutation-ref null writes to missing cells as no-ops', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'operation-mutation-null-missing-noop',
      replicaId: 'a',
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const refs: EngineCellMutationRef[] = [
      {
        sheetId,
        mutation: { kind: 'setCellValue', row: 3, col: 3, value: null },
      },
    ]
    const forwardOps = refs.map((ref) => cellMutationRefToEngineOp(engine.workbook, ref))
    const batch = createBatch(getReplicaState(engine), forwardOps)
    const before = engine.exportSnapshot()

    Effect.runSync(getOperationService(engine).applyCellMutationsAt(refs, batch, 'local', 1))

    expect(engine.exportSnapshot()).toEqual(before)
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

  it('replaces existing formulas with generic batch literal writes', async () => {
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
})
