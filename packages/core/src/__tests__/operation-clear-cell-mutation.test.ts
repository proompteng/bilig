import { Effect } from 'effect'
import { describe, expect, it } from 'vitest'
import { ErrorCode, ValueTag } from '@bilig/protocol'
import { createBatch } from '../replica-state.js'
import { SpreadsheetEngine } from '../engine.js'
import { cellMutationRefToEngineOp, type EngineCellMutationRef } from '../cell-mutations-at.js'
import { getOperationService, getReplicaState } from './operation-service-test-helpers.js'

describe('operation clear-cell mutations', () => {
  it('treats batched clears of already-empty tracked dependency cells as no-ops', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'operation-batch-clear-empty-noop' })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'A1', 'A1+D4')

    const before = engine.exportSnapshot()
    const batch = createBatch(getReplicaState(engine), [{ kind: 'clearCell', sheetName: 'Sheet1', address: 'D4' }])

    Effect.runSync(getOperationService(engine).applyBatch(batch, 'local'))

    expect(engine.exportSnapshot()).toEqual(before)
  })

  it('treats mutation-ref clears of already-empty tracked dependency cells as no-ops', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'operation-mutation-clear-empty-noop',
      replicaId: 'a',
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellFormula('Sheet1', 'A1', 'A1+D4')
    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const refs: EngineCellMutationRef[] = [
      {
        sheetId,
        mutation: { kind: 'clearCell', row: 3, col: 3 },
      },
    ]
    const forwardOps = refs.map((ref) => cellMutationRefToEngineOp(engine.workbook, ref))
    const batch = createBatch(getReplicaState(engine), forwardOps)
    const before = engine.exportSnapshot()

    Effect.runSync(getOperationService(engine).applyCellMutationsAt(refs, batch, 'local', 1))

    expect(engine.exportSnapshot()).toEqual(before)
  })

  it('treats generic batch clears of missing cells as no-ops', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'operation-batch-clear-missing-noop' })
    await engine.ready()
    engine.createSheet('Sheet1')

    const before = engine.exportSnapshot()
    const batch = createBatch(getReplicaState(engine), [{ kind: 'clearCell', sheetName: 'Sheet1', address: 'G7' }])

    Effect.runSync(getOperationService(engine).applyBatch(batch, 'local'))

    expect(engine.exportSnapshot()).toEqual(before)
  })

  it('clears formula cells and recalculates dependent formulas through mutation refs', async () => {
    const engine = new SpreadsheetEngine({
      workbookName: 'operation-mutation-clear-formula-cell',
      replicaId: 'a',
    })
    await engine.ready()
    engine.createSheet('Sheet1')
    engine.setCellValue('Sheet1', 'A1', 7)
    engine.setCellFormula('Sheet1', 'B1', 'A1*3')
    engine.setCellFormula('Sheet1', 'C1', 'B1+2')
    const sheetId = engine.workbook.getSheet('Sheet1')!.id
    const refs: EngineCellMutationRef[] = [
      {
        sheetId,
        mutation: { kind: 'clearCell', row: 0, col: 1 },
      },
    ]
    const batch = createBatch(
      getReplicaState(engine),
      refs.map((ref) => cellMutationRefToEngineOp(engine.workbook, ref)),
    )

    Effect.runSync(getOperationService(engine).applyCellMutationsAt(refs, batch, 'local', 1))

    expect(engine.getCell('Sheet1', 'B1').formula).toBeUndefined()
    expect(engine.getCellValue('Sheet1', 'B1')).toEqual({ tag: ValueTag.Empty })
    expect(engine.getCellValue('Sheet1', 'C1')).toEqual({ tag: ValueTag.Number, value: 2 })
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
